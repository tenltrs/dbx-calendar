package main

import (
	"encoding/json"
	"fmt"
	"sort"
	"strings"
	"time"

	dbxpluginsdk "github.com/t8y2/dbx/plugins/sdk/go/dbx-plugin-sdk"
)

// maxToolNotes bounds one tool answer. The host compresses results before they
// reach a model, but it cannot make a needlessly large one small: a range query
// over years of notes should say it was cut short rather than push everything
// across the bridge.
const maxToolNotes = 200

// mcpTool is one entry of the mcp/tools reply.
//
// The schema is deliberately plain: DBX narrows it to the subset every model
// provider accepts, so only type, description, properties, required, items,
// enum and numeric/length bounds survive. Parameter names stay within letters,
// digits and underscores for the same reason.
type mcpTool struct {
	Name        string         `json:"name"`
	Description string         `json:"description"`
	InputSchema map[string]any `json:"inputSchema"`
	Annotations map[string]any `json:"annotations,omitempty"`
}

// readOnlyHint is what lets a tool run without asking the user first. Every
// other call pauses for confirmation, so it is only set where nothing is
// written — the host's prompt guards against a model's mistake and is not a
// substitute for the checks in here.
var readOnlyHint = map[string]any{"readOnlyHint": true}

// connectionProperty is injected by the host when the schema declares it: the
// id of the plugin connection the assistant is bound to. It is stripped from
// the schema a model sees, so the model never sets it.
func connectionProperty() map[string]any {
	return map[string]any{
		"type":        "string",
		"description": "Filled in by DBX; do not set.",
	}
}

func dateProperty(description string) map[string]any {
	return map[string]any{
		"type":        "string",
		"description": description,
		"minLength":   10,
		"maxLength":   10,
	}
}

func textProperty(description string) map[string]any {
	return map[string]any{"type": "string", "description": description}
}

func objectSchema(properties map[string]any, required ...string) map[string]any {
	schema := map[string]any{"type": "object", "properties": properties}
	if len(required) > 0 {
		schema["required"] = required
	}
	return schema
}

// The tool names, declared once. The catalogue and the dispatch both read them
// from here, so a rename cannot leave the two halves disagreeing — which is how
// a tool ends up advertised under one name and answered under another.
const (
	toolListNotes   = "dbx_calendar_list_notes"
	toolSearchNotes = "dbx_calendar_search_notes"
	toolAppendNote  = "dbx_calendar_append_note"
	toolSetNote     = "dbx_calendar_set_note"
	toolDeleteNote  = "dbx_calendar_delete_note"
)

// mcpTools is the catalogue the host offers the assistant. Reads carry
// readOnlyHint so they run straight away; the writes do not, so the user sees
// the exact arguments and allows or refuses each one.
func mcpTools() []mcpTool {
	return []mcpTool{
		{
			Name: toolListNotes,
			Description: "List calendar notes in a date range, oldest first. " +
				"Use it for questions about what was written on or between particular days.",
			InputSchema: objectSchema(map[string]any{
				"from":         dateProperty("First day to include, as YYYY-MM-DD. Omit for no lower bound."),
				"to":           dateProperty("Last day to include, as YYYY-MM-DD. Omit for no upper bound."),
				"connectionId": connectionProperty(),
			}),
			Annotations: readOnlyHint,
		},
		{
			Name: toolSearchNotes,
			Description: "Find calendar notes whose text contains a phrase, ignoring case. " +
				"Use it for questions like \"what did I write about the invoice?\".",
			InputSchema: objectSchema(map[string]any{
				"query":        textProperty("The phrase to look for."),
				"connectionId": connectionProperty(),
			}, "query"),
			Annotations: readOnlyHint,
		},
		{
			// The quick-capture path: one line, on today unless told otherwise.
			Name: toolAppendNote,
			Description: "Add a line to a day's note, keeping whatever is already there. " +
				"This is the tool for \"make a note of this\". Omit 'date' to use today. " +
				"Prefer it over set_note, which replaces the day's whole note.",
			InputSchema: objectSchema(map[string]any{
				"date":         dateProperty("The day to add to, as YYYY-MM-DD. Defaults to today."),
				"text":         textProperty("The line to add."),
				"connectionId": connectionProperty(),
			}, "text"),
		},
		{
			// Replacing is destructive, so the day has to be named: quick
			// capture goes through append_note instead.
			Name: toolSetNote,
			Description: "Replace a day's whole note. The day must be named, because this " +
				"discards whatever was there. To add something, use append_note.",
			InputSchema: objectSchema(map[string]any{
				"date":         dateProperty("The day to write on, as YYYY-MM-DD."),
				"text":         textProperty("The note itself."),
				"connectionId": connectionProperty(),
			}, "date", "text"),
		},
		{
			Name:        toolDeleteNote,
			Description: "Clear the note on one day.",
			InputSchema: objectSchema(map[string]any{
				"date":         dateProperty("The day to clear, as YYYY-MM-DD."),
				"connectionId": connectionProperty(),
			}, "date"),
		},
	}
}

// toolResult and toolError build the MCP CallToolResult. A bad argument comes
// back as isError rather than as a protocol failure: the model wrote it, and it
// can correct itself if the reason is readable.
func toolResult(payload any) map[string]any {
	text, err := json.Marshal(payload)
	if err != nil {
		return toolError(fmt.Sprintf("could not encode the result: %v", err))
	}
	return map[string]any{
		"content": []map[string]any{{"type": "text", "text": string(text)}},
		"isError": false,
	}
}

func toolError(message string) map[string]any {
	return map[string]any{
		"content": []map[string]any{{"type": "text", "text": message}},
		"isError": true,
	}
}

// noteEntry is one note as the tools report it.
type noteEntry struct {
	Date string `json:"date"`
	Text string `json:"text"`
}

// mcpCall runs one tool.
//
// The arguments and the lifecycle payload are merged into the same shape the
// notes RPCs already take, so a tool resolves the same connection namespace
// they do: the host binds a connection and may report it in either place, and
// `optionalConnectionID` knows how to read both.
func (plugin *plugin) mcpCall(params map[string]any) (any, *dbxpluginsdk.PluginError) {
	name, _ := params["tool"].(string)

	values := map[string]any{}
	if arguments, ok := params["arguments"].(map[string]any); ok {
		for key, value := range arguments {
			values[key] = value
		}
	}
	// The lifecycle payload has the shape connection/connect uses, so anything
	// it carries is folded in without overwriting an injected argument.
	if lifecycle, ok := params["lifecycle"].(map[string]any); ok {
		for _, key := range []string{"connection", "connectionId"} {
			if _, taken := values[key]; taken {
				continue
			}
			if value, present := lifecycle[key]; present {
				values[key] = value
			}
		}
	}

	handle, pluginError := plugin.notesFor(values)
	if pluginError != nil {
		return toolError(pluginError.Message), nil
	}
	if !plugin.notes.Persistence() {
		return toolError("Notes cannot be read or written right now: the plugin data directory is unavailable."), nil
	}

	switch name {
	case toolListNotes:
		return listToolNotes(handle, values), nil
	case toolSearchNotes:
		query, ok := stringArgument(values, "query")
		if !ok {
			return toolError("search_notes needs a non-empty 'query'."), nil
		}
		return searchToolNotes(handle, query), nil
	case toolAppendNote:
		return appendToolNote(handle, values)
	case toolSetNote:
		return setToolNote(handle, values)
	case toolDeleteNote:
		dateKey, ok := stringArgument(values, "date")
		if !ok {
			return toolError("delete_note needs a 'date' of the form YYYY-MM-DD."), nil
		}
		if err := handle.Set(dateKey, ""); err != nil {
			return toolError(err.Error()), nil
		}
		return toolResult(map[string]any{"date": dateKey, "cleared": true}), nil
	default:
		return toolError(fmt.Sprintf("Unknown tool %q.", name)), nil
	}
}

func listToolNotes(handle *connNotes, values map[string]any) any {
	from, _ := stringArgument(values, "from")
	to, _ := stringArgument(values, "to")
	if from != "" && !isDateKey(from) {
		return toolError(fmt.Sprintf("'from' must be YYYY-MM-DD, got %q.", from))
	}
	if to != "" && !isDateKey(to) {
		return toolError(fmt.Sprintf("'to' must be YYYY-MM-DD, got %q.", to))
	}

	notes, err := handle.Notes()
	if err != nil {
		return toolError(err.Error())
	}
	// Date keys sort chronologically as plain strings, so no parsing is needed.
	dates := make([]string, 0, len(notes))
	for dateKey := range notes {
		if from != "" && dateKey < from {
			continue
		}
		if to != "" && dateKey > to {
			continue
		}
		dates = append(dates, dateKey)
	}
	sort.Strings(dates)

	entries, truncated := collectNotes(notes, dates)
	return toolResult(map[string]any{
		"notes":     entries,
		"count":     len(entries),
		"truncated": truncated,
	})
}

func searchToolNotes(handle *connNotes, query string) any {
	notes, err := handle.Notes()
	if err != nil {
		return toolError(err.Error())
	}
	needle := strings.ToLower(query)
	dates := make([]string, 0, len(notes))
	for dateKey, text := range notes {
		if strings.Contains(strings.ToLower(text), needle) {
			dates = append(dates, dateKey)
		}
	}
	sort.Strings(dates)

	entries, truncated := collectNotes(notes, dates)
	return toolResult(map[string]any{
		"notes":     entries,
		"count":     len(entries),
		"truncated": truncated,
	})
}

// collectNotes turns date keys into entries, stopping at the cap.
func collectNotes(notes map[string]string, dates []string) ([]noteEntry, bool) {
	truncated := len(dates) > maxToolNotes
	if truncated {
		dates = dates[:maxToolNotes]
	}
	entries := make([]noteEntry, 0, len(dates))
	for _, dateKey := range dates {
		entries = append(entries, noteEntry{Date: dateKey, Text: notes[dateKey]})
	}
	return entries, truncated
}

func setToolNote(handle *connNotes, values map[string]any) (any, *dbxpluginsdk.PluginError) {
	dateKey, ok := stringArgument(values, "date")
	if !ok {
		return toolError("set_note needs a 'date' of the form YYYY-MM-DD."), nil
	}
	text, _ := values["text"].(string)
	if strings.TrimSpace(text) == "" {
		return toolError("set_note needs non-empty 'text'. Use delete_note to clear a day."), nil
	}
	if err := handle.Set(dateKey, text); err != nil {
		return toolError(err.Error()), nil
	}
	return toolResult(map[string]any{"date": dateKey, "saved": true}), nil
}

// appendToolNote adds one line to a day, creating the day if it is empty.
//
// The day defaults to today, which is what makes "make a note of this" work
// without the model having to work out the date first.
func appendToolNote(handle *connNotes, values map[string]any) (any, *dbxpluginsdk.PluginError) {
	dateKey, dateErr := noteDate(values)
	if dateErr != nil {
		return toolError(dateErr.Error()), nil
	}
	text, _ := values["text"].(string)
	line := strings.TrimSpace(text)
	if line == "" {
		return toolError("append_note needs non-empty 'text'."), nil
	}

	existing, _, readErr := handle.Note(dateKey)
	if readErr != nil {
		return toolError(readErr.Error()), nil
	}
	combined := line
	if strings.TrimSpace(existing) != "" {
		combined = existing + "\n" + line
	}
	if err := handle.Set(dateKey, combined); err != nil {
		return toolError(err.Error()), nil
	}
	return toolResult(map[string]any{"date": dateKey, "appended": true, "note": combined}), nil
}

// noteDate reads the day a tool acts on, falling back to today.
func noteDate(values map[string]any) (string, error) {
	dateKey, ok := stringArgument(values, "date")
	if !ok {
		return time.Now().Format("2006-01-02"), nil
	}
	if err := validDateKey(dateKey); err != nil {
		return "", fmt.Errorf("'date' must be YYYY-MM-DD, got %q", dateKey)
	}
	return dateKey, nil
}

func stringArgument(values map[string]any, key string) (string, bool) {
	value, ok := values[key].(string)
	if !ok {
		return "", false
	}
	trimmed := strings.TrimSpace(value)
	return trimmed, trimmed != ""
}

// isDateKey reports whether a tool argument is a date the notes store accepts.
// The store validates again on write; this just lets a read fail with a
// sentence the model can act on.
func isDateKey(value string) bool {
	return validDateKey(value) == nil
}
