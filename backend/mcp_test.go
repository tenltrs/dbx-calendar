package main

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

// newToolPlugin builds a plugin whose notes live in a temp data directory.
func newToolPlugin(t *testing.T) *plugin {
	t.Helper()
	dataDir := t.TempDir()
	return &plugin{
		connections: map[string]*connNotes{},
		notes:       newNotesStore(dataDir),
		exports:     newExportStore(dataDir),
	}
}

// callTool drives one tool the way the host does and returns the decoded
// CallToolResult.
func callTool(t *testing.T, instance *plugin, tool string, arguments map[string]any) map[string]any {
	t.Helper()
	result, pluginError := instance.mcpCall(map[string]any{"tool": tool, "arguments": arguments})
	if pluginError != nil {
		t.Fatalf("tool %s failed at the protocol level: %v", tool, pluginError.Message)
	}
	decoded, ok := result.(map[string]any)
	if !ok {
		t.Fatalf("tool %s returned %T", tool, result)
	}
	return decoded
}

// toolText returns the text of the first content block.
func toolText(t *testing.T, result map[string]any) string {
	t.Helper()
	content, ok := result["content"].([]map[string]any)
	if !ok || len(content) == 0 {
		t.Fatalf("no content in %#v", result)
	}
	text, _ := content[0]["text"].(string)
	return text
}

// toolPayload decodes the JSON a successful tool returns.
func toolPayload(t *testing.T, result map[string]any) map[string]any {
	t.Helper()
	if result["isError"] == true {
		t.Fatalf("expected success, got error: %s", toolText(t, result))
	}
	var payload map[string]any
	if err := json.Unmarshal([]byte(toolText(t, result)), &payload); err != nil {
		t.Fatalf("payload is not JSON: %v (%s)", err, toolText(t, result))
	}
	return payload
}

// mustSet writes a note through the tool and fails the test if it is refused.
func mustSet(t *testing.T, instance *plugin, date, text string) {
	t.Helper()
	result := callTool(t, instance, toolSetNote, map[string]any{"date": date, "text": text})
	if result["isError"] == true {
		t.Fatalf("set_note(%s) failed: %s", date, toolText(t, result))
	}
}

func TestToolsDeclareReadOnlyOnlyWhereNothingIsWritten(t *testing.T) {
	byName := map[string]mcpTool{}
	for _, tool := range mcpTools() {
		byName[tool.Name] = tool
	}

	// These two only read, so the host runs them without asking.
	for _, name := range []string{toolListNotes, toolSearchNotes} {
		tool, ok := byName[name]
		if !ok {
			t.Fatalf("%s is missing", name)
		}
		if tool.Annotations["readOnlyHint"] != true {
			t.Fatalf("%s must be marked read-only", name)
		}
	}
	// These three change stored notes, so every call has to be confirmed.
	for _, name := range []string{toolAppendNote, toolSetNote, toolDeleteNote} {
		tool, ok := byName[name]
		if !ok {
			t.Fatalf("%s is missing", name)
		}
		if _, marked := tool.Annotations["readOnlyHint"]; marked {
			t.Fatalf("%s writes, so it must not be marked read-only", name)
		}
	}
}

// The host narrows the schema to a subset every provider accepts, so anything
// outside it is dropped — a required argument declared with it would quietly
// stop being required.
func TestToolSchemasStayInsideTheSupportedSubset(t *testing.T) {
	allowed := map[string]bool{
		"type": true, "description": true, "properties": true, "required": true,
		"items": true, "enum": true, "minLength": true, "maxLength": true,
		"minimum": true, "maximum": true, "minItems": true, "maxItems": true,
	}

	// walk checks a schema's keywords, and treats the names under `properties`
	// as names rather than as keywords: they are the model's argument names,
	// and the host keeps them.
	var walk func(path string, schema map[string]any)
	walk = func(path string, schema map[string]any) {
		for key, value := range schema {
			if !allowed[key] {
				t.Errorf("%s uses %q, which the host drops", path, key)
			}
			if key == "properties" {
				names, ok := value.(map[string]any)
				if !ok {
					t.Errorf("%s.properties is not an object", path)
					continue
				}
				for name, sub := range names {
					if !isPlainIdentifier(name, false) {
						t.Errorf("%s property %q must be letters, digits and underscores", path, name)
					}
					if nested, ok := sub.(map[string]any); ok {
						walk(path+"."+name, nested)
					}
				}
				continue
			}
			switch typed := value.(type) {
			case map[string]any:
				walk(path+"."+key, typed)
			case []any:
				for _, item := range typed {
					if nested, ok := item.(map[string]any); ok {
						walk(path+"."+key, nested)
					}
				}
			}
		}
	}

	for _, tool := range mcpTools() {
		if tool.Name == "" || tool.Description == "" {
			t.Errorf("a tool needs a name and a description: %#v", tool)
		}
		// Tools are exposed as <prefix>__<name>, so the name itself has to stay
		// plain.
		if !isPlainIdentifier(tool.Name, true) {
			t.Errorf("tool name %q must be letters, digits and underscores", tool.Name)
		}
		if _, ok := tool.InputSchema["properties"].(map[string]any); !ok {
			t.Errorf("%s has no properties", tool.Name)
			continue
		}
		walk(tool.Name, tool.InputSchema)
	}
}

// isPlainIdentifier reports whether a name is letters, digits, underscores and
// (for tool names) lowercase only.
func isPlainIdentifier(value string, lowerOnly bool) bool {
	return !strings.ContainsFunc(value, func(r rune) bool {
		switch {
		case r >= 'a' && r <= 'z', r >= '0' && r <= '9', r == '_':
			return false
		case r >= 'A' && r <= 'Z':
			return lowerOnly
		default:
			return true
		}
	})
}

func TestSetAndListRoundTrip(t *testing.T) {
	instance := newToolPlugin(t)

	saved := toolPayload(t, callTool(t, instance, toolSetNote, map[string]any{
		"date": "2026-09-22", "text": "renew the domain",
	}))
	if saved["saved"] != true || saved["date"] != "2026-09-22" {
		t.Fatalf("unexpected save result: %#v", saved)
	}
	mustSet(t, instance, "2026-09-05", "invoice 42")

	listed := toolPayload(t, callTool(t, instance, toolListNotes, map[string]any{}))
	notes, ok := listed["notes"].([]any)
	if !ok || len(notes) != 2 {
		t.Fatalf("expected two notes, got %v", listed["count"])
	}
	if notes[0].(map[string]any)["date"] != "2026-09-05" {
		t.Fatalf("expected the oldest note first, got %v", notes[0].(map[string]any)["date"])
	}
}

// A range query is the common case, and the bounds are inclusive.
func TestListNotesFiltersByRange(t *testing.T) {
	instance := newToolPlugin(t)
	for _, date := range []string{"2026-01-01", "2026-06-15", "2026-12-31"} {
		mustSet(t, instance, date, "note "+date)
	}

	listed := toolPayload(t, callTool(t, instance, toolListNotes, map[string]any{
		"from": "2026-06-01", "to": "2026-06-30",
	}))
	notes := listed["notes"].([]any)
	if len(notes) != 1 || notes[0].(map[string]any)["date"] != "2026-06-15" {
		t.Fatalf("unexpected range result, got %v notes", listed["count"])
	}
}

func TestSearchNotesIgnoresCaseAndFindsSubstrings(t *testing.T) {
	instance := newToolPlugin(t)
	mustSet(t, instance, "2026-02-14", "Dinner with Ana")
	mustSet(t, instance, "2026-03-01", "pay the rent")

	found := toolPayload(t, callTool(t, instance, toolSearchNotes, map[string]any{"query": "ana"}))
	notes := found["notes"].([]any)
	if len(notes) != 1 || notes[0].(map[string]any)["date"] != "2026-02-14" {
		t.Fatalf("unexpected search result, got %v notes", found["count"])
	}

	missing := toolPayload(t, callTool(t, instance, toolSearchNotes, map[string]any{"query": "nothing"}))
	if len(missing["notes"].([]any)) != 0 {
		t.Fatalf("expected no matches, got %v", missing["count"])
	}
}

// append_note is the quick-capture path, so what matters is that a second note
// on the same day does not wipe the first.
func TestAppendNoteKeepsWhatWasThere(t *testing.T) {
	instance := newToolPlugin(t)

	first := toolPayload(t, callTool(t, instance, toolAppendNote, map[string]any{
		"date": "2026-09-22", "text": "morning: the meeting moved",
	}))
	if first["appended"] != true {
		t.Fatalf("unexpected append result: %#v", first)
	}
	second := toolPayload(t, callTool(t, instance, toolAppendNote, map[string]any{
		"date": "2026-09-22", "text": "afternoon: rewrote the plan",
	}))
	if note, _ := second["note"].(string); !strings.Contains(note, "morning:") || !strings.Contains(note, "afternoon:") {
		t.Fatalf("the first line was lost: %q", note)
	}

	listed := toolPayload(t, callTool(t, instance, toolListNotes, map[string]any{}))
	notes := listed["notes"].([]any)
	if len(notes) != 1 {
		t.Fatalf("expected one day, got %d", len(notes))
	}
	text := notes[0].(map[string]any)["text"].(string)
	if !strings.Contains(text, "morning:") || !strings.Contains(text, "afternoon:") {
		t.Fatalf("both lines should survive: %q", text)
	}
}

// "Make a note of this" should not need the model to work out today's date.
func TestAppendNoteDefaultsToToday(t *testing.T) {
	instance := newToolPlugin(t)

	result := toolPayload(t, callTool(t, instance, toolAppendNote, map[string]any{"text": "quick one"}))
	today := time.Now().Format("2006-01-02")
	if result["date"] != today {
		t.Fatalf("expected today (%s), got %v", today, result["date"])
	}

	listed := toolPayload(t, callTool(t, instance, toolListNotes, map[string]any{}))
	if count := listed["count"].(float64); count != 1 {
		t.Fatalf("expected the note to land on today, got %v", count)
	}
}

// Replacing is the destructive one, so it must not inherit the convenience.
func TestSetNoteStillRequiresADay(t *testing.T) {
	instance := newToolPlugin(t)
	result := callTool(t, instance, toolSetNote, map[string]any{"text": "no day given"})
	if result["isError"] != true {
		t.Fatalf("set_note should refuse to guess the day: %#v", result)
	}
}

func TestAppendNoteRefusesEmptyText(t *testing.T) {
	instance := newToolPlugin(t)
	result := callTool(t, instance, toolAppendNote, map[string]any{"text": "   "})
	if result["isError"] != true {
		t.Fatalf("expected empty text to be refused: %#v", result)
	}
}

func TestDeleteNoteClearsADay(t *testing.T) {
	instance := newToolPlugin(t)
	mustSet(t, instance, "2026-09-22", "x")

	deleted := toolPayload(t, callTool(t, instance, toolDeleteNote, map[string]any{"date": "2026-09-22"}))
	if deleted["cleared"] != true {
		t.Fatalf("unexpected delete result: %#v", deleted)
	}
	listed := toolPayload(t, callTool(t, instance, toolListNotes, map[string]any{}))
	if count := listed["count"].(float64); count != 0 {
		t.Fatalf("expected the day to be gone, got %v notes", count)
	}
}

// The store owns the date format, so a bad one comes back as something the
// model can act on rather than as an internal error.
func TestToolsRejectBadDates(t *testing.T) {
	instance := newToolPlugin(t)
	for _, tool := range []string{toolSetNote, toolDeleteNote} {
		result := callTool(t, instance, tool, map[string]any{"date": "22/09/2026", "text": "x"})
		if result["isError"] != true {
			t.Fatalf("%s accepted a bad date: %#v", tool, result)
		}
	}
	result := callTool(t, instance, toolListNotes, map[string]any{"from": "not-a-date"})
	if result["isError"] != true {
		t.Fatalf("list_notes accepted a bad bound: %#v", result)
	}
}

func TestSetNoteRefusesEmptyText(t *testing.T) {
	instance := newToolPlugin(t)
	result := callTool(t, instance, toolSetNote, map[string]any{"date": "2026-09-22", "text": "   "})
	if result["isError"] != true {
		t.Fatalf("expected empty text to be refused: %#v", result)
	}
}

func TestSearchRequiresAQuery(t *testing.T) {
	instance := newToolPlugin(t)
	if result := callTool(t, instance, toolSearchNotes, map[string]any{}); result["isError"] != true {
		t.Fatalf("expected a missing query to be refused: %#v", result)
	}
}

func TestUnknownToolIsReportedNotPanicked(t *testing.T) {
	instance := newToolPlugin(t)
	result := callTool(t, instance, "no_such_tool", map[string]any{})
	if result["isError"] != true {
		t.Fatalf("expected an error for an unknown tool: %#v", result)
	}
}

// The host binds a connection and may report it in the arguments or in the
// lifecycle payload. Neither may partition storage: there is one notebook, so a
// note written while a connection is bound has to be the same note the UI shows
// without one. Following the connection here is what would make notes written
// before the connection existed look lost.
func TestBoundConnectionDoesNotPartitionStorage(t *testing.T) {
	instance := newToolPlugin(t)

	for _, call := range []map[string]any{
		{
			"tool":      toolSetNote,
			"arguments": map[string]any{"date": "2026-09-22", "text": "bound"},
			"lifecycle": map[string]any{"connection": map[string]any{"id": "conn-7"}},
		},
		{
			"tool":      toolAppendNote,
			"arguments": map[string]any{"connectionId": "conn-9", "date": "2026-09-23", "text": "also bound"},
		},
	} {
		if _, pluginError := instance.mcpCall(call); pluginError != nil {
			t.Fatal(pluginError.Message)
		}
	}

	// Read back with nothing bound at all — the UI's own path.
	listed := toolPayload(t, callTool(t, instance, toolListNotes, map[string]any{}))
	if count := listed["count"].(float64); count != 2 {
		t.Fatalf("expected both notes in the notebook, got %v", count)
	}

	// And the store on disk holds exactly one file, not one per connection.
	entries, err := os.ReadDir(filepath.Join(instance.notes.dataDir, "notes"))
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 1 || entries[0].Name() != workbenchNamespace+".json" {
		names := make([]string, 0, len(entries))
		for _, entry := range entries {
			names = append(names, entry.Name())
		}
		t.Fatalf("expected a single notebook file, got %v", names)
	}
}

// The host compresses results, but an unbounded answer is still unbounded work.
func TestListNotesIsCapped(t *testing.T) {
	dataDir := t.TempDir()
	handle, err := newNotesStore(dataDir).Open(workbenchNamespace)
	if err != nil {
		t.Fatal(err)
	}
	// Seeded directly: the cap is what is under test, not the write path.
	bulk := map[string]string{}
	for index := 0; index < maxToolNotes+5; index++ {
		key := fmt.Sprintf("2026-%02d-%02d", index/28+1, index%28+1)
		bulk[key] = "note"
	}
	if err := handle.Replace(bulk); err != nil {
		t.Fatal(err)
	}

	instance := &plugin{
		connections: map[string]*connNotes{},
		notes:       newNotesStore(dataDir),
		exports:     newExportStore(dataDir),
	}
	listed := toolPayload(t, callTool(t, instance, toolListNotes, map[string]any{}))
	if got := listed["count"].(float64); int(got) != maxToolNotes {
		t.Fatalf("expected the cap %d, got %v", maxToolNotes, got)
	}
	if listed["truncated"] != true {
		t.Fatalf("expected the answer to say it was cut short, got %v", listed["truncated"])
	}
}

// The RPC the host actually calls, driven over the real protocol.
func TestProtocolMcpTools(t *testing.T) {
	dataDir := t.TempDir()
	instance := &plugin{
		connections: map[string]*connNotes{},
		notes:       newNotesStore(dataDir),
		exports:     newExportStore(dataDir),
	}

	responses := serve(t, instance,
		requestLine(1, "plugin/initialize", map[string]any{
			"host": map[string]any{"protocolVersions": []int{1}},
		}),
		requestLine(2, "mcp/tools", map[string]any{"connectionId": workbenchNamespace}),
	)
	for _, id := range []int{1, 2} {
		if failure, failed := responses[id]["error"]; failed {
			t.Fatalf("request %d failed: %#v", id, failure)
		}
	}

	catalogue := responses[2]["result"].(map[string]any)
	tools, ok := catalogue["tools"].([]any)
	if !ok {
		t.Fatalf("mcp/tools returned no tool list: %#v", catalogue)
	}
	if len(tools) != 5 {
		t.Fatalf("expected five tools, got %d", len(tools))
	}
	if name := tools[0].(map[string]any)["name"]; name != toolListNotes {
		t.Fatalf("unexpected first tool: %v", name)
	}

	// The write is driven in its own round trip: the SDK dispatches each request
	// on its own goroutine, so a read in the same batch could race it.
	responses = serve(t, instance, requestLine(1, "mcp/call", map[string]any{
		"tool":      toolSetNote,
		"arguments": map[string]any{"date": "2026-09-22", "text": "from the assistant"},
	}))
	saved := responses[1]["result"].(map[string]any)
	if saved["isError"] != false {
		t.Fatalf("unexpected save result: %#v", saved)
	}

	responses = serve(t, instance, requestLine(1, "mcp/call", map[string]any{
		"tool":      toolListNotes,
		"arguments": map[string]any{},
	}))
	listed := responses[1]["result"].(map[string]any)
	content, ok := listed["content"].([]any)
	if !ok || len(content) == 0 {
		t.Fatalf("mcp/call returned no content: %#v", listed)
	}
	if text := content[0].(map[string]any)["text"].(string); !strings.Contains(text, "from the assistant") {
		t.Fatalf("the note did not come back: %s", text)
	}
}
