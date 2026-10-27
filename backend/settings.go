package main

import (
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"
)

// settingsSchemaVersion exists for the same reason as the notes one: the file
// can evolve without guessing whether an old file is still readable.
const settingsSchemaVersion = 1

// settingsFile is the on-disk shape. The payload stays a free-form object so
// the UI can add preference keys without a protocol change; the sidecar never
// interprets them.
type settingsFile struct {
	Version  int            `json:"version"`
	Settings map[string]any `json:"settings"`
}

// settingsPath resolves <dataDir>/settings.json. The workbench namespace is the
// only legal key here: preferences belong to the plugin, not to a connection,
// so a per-connection file would fork the settings every time the user opens a
// different database.
func settingsPath(dataDir string) (string, error) {
	if dataDir == "" {
		return "", ErrNoDataDir
	}
	return filepath.Join(dataDir, "settings.json"), nil
}

// settingsStore owns the plugin preference file for one data directory. The
// access pattern is read-on-start, write-on-change, so a single file with one
// mutex is enough; per-connection handles like the notes store has would only
// multiply the readers of one file.
type settingsStore struct {
	dataDir string
	mutex   sync.Mutex
	cache   map[string]any
	loaded  bool
}

func newSettingsStore(dataDir string) *settingsStore {
	return &settingsStore{dataDir: dataDir}
}

// Persistence reports whether settings can survive a process restart.
func (store *settingsStore) Persistence() bool {
	return store.dataDir != ""
}

// Get returns the stored settings, reading the file on first use. A corrupt
// file is quarantined like a corrupt notes file: preferences are cheap to
// re-enter, but silently resetting them would look like the plugin forgot on
// purpose, so the old file is kept for inspection.
func (store *settingsStore) Get() (map[string]any, error) {
	store.mutex.Lock()
	defer store.mutex.Unlock()
	if store.loaded {
		return store.snapshotLocked(), nil
	}

	path, err := settingsPath(store.dataDir)
	if err != nil {
		return nil, err
	}
	data, err := os.ReadFile(path)
	switch {
	case errors.Is(err, os.ErrNotExist):
		store.cache, store.loaded = map[string]any{}, true
		return store.snapshotLocked(), nil
	case err != nil:
		return nil, err
	}

	var parsed settingsFile
	if err := json.Unmarshal(data, &parsed); err != nil || parsed.Settings == nil {
		quarantine := path + ".corrupt-" + time.Now().UTC().Format("20060102T150405")
		_ = os.Rename(path, quarantine)
		store.cache, store.loaded = map[string]any{}, true
		return store.snapshotLocked(), nil
	}
	store.cache, store.loaded = parsed.Settings, true
	return store.snapshotLocked(), nil
}

// Set merges the incoming keys over the stored ones. Partial writes are the
// point: the UI sends one changed key, not the whole panel, so an old caller
// that only knows about style cannot wipe a newer key another release added.
func (store *settingsStore) Set(incoming map[string]any) (map[string]any, error) {
	if _, err := store.Get(); err != nil {
		return nil, err
	}
	store.mutex.Lock()
	for key, value := range incoming {
		if value == nil {
			delete(store.cache, key)
			continue
		}
		store.cache[key] = value
	}
	snapshot := store.snapshotLocked()
	store.mutex.Unlock()

	return snapshot, store.flush(snapshot)
}

func (store *settingsStore) snapshotLocked() map[string]any {
	copied := make(map[string]any, len(store.cache))
	for key, value := range store.cache {
		copied[key] = value
	}
	return copied
}

// flush writes atomically, exactly like the notes store: temp file in the same
// directory, then a rename that retries while Windows holds the target open.
func (store *settingsStore) flush(payload settingsFile) error {
	data, err := json.MarshalIndent(payload, "", "  ")
	if err != nil {
		return err
	}
	data = append(data, '\n')

	path, err := settingsPath(store.dataDir)
	if err != nil {
		return err
	}
	if err := os.MkdirAll(filepath.Dir(path), 0o700); err != nil {
		return err
	}
	tmp := path + ".tmp"
	if err := os.WriteFile(tmp, data, 0o600); err != nil {
		return err
	}
	var lastErr error
	for attempt := 0; attempt < 5; attempt++ {
		if lastErr = os.Rename(tmp, path); lastErr == nil {
			return nil
		}
		time.Sleep(time.Duration(20*(attempt+1)) * time.Millisecond)
	}
	_ = os.Remove(tmp)
	return lastErr
}

// decodeSettings reads the settings object an update sends. Values stay
// untyped on purpose: the sidecar stores what the UI sent.
func decodeSettings(value any) (map[string]any, bool) {
	raw, ok := value.(map[string]any)
	if !ok {
		return nil, false
	}
	settings := make(map[string]any, len(raw))
	for key, entry := range raw {
		if key == "" || strings.ContainsAny(key, "\x00") {
			return nil, false
		}
		settings[key] = entry
	}
	return settings, true
}

// validSettingsKey is a light guard: keys are UI identifiers, so anything
// beyond a short identifier-shaped string is a caller bug worth surfacing.
func validSettingsKey(key string) error {
	if key == "" || len(key) > 64 {
		return fmt.Errorf("invalid settings key: %q", key)
	}
	return nil
}
