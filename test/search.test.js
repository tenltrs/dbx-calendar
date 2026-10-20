import assert from "node:assert/strict";
import test from "node:test";
import * as search from "../src/search.js";
const { highlightMatches, searchNotes } = search;

test("searchNotes matches note text case-insensitively and sorts dates newest first", () => {
  const notes = {
    "2026-02-01": "Planning meeting",
    "2026-03-09": "meeting follow-up",
    "2025-12-31": "Year end review",
  };

  assert.deepEqual(searchNotes(notes, "MEETING"), [
    { date: "2026-03-09", content: "meeting follow-up" },
    { date: "2026-02-01", content: "Planning meeting" },
  ]);
});

test("searchNotes returns no results for a blank query", () => {
  assert.deepEqual(searchNotes({ "2026-03-09": "meeting" }, "  "), []);
});

test("search keyboard navigation wraps through results and starts at the nearest end", () => {
  assert.deepEqual(search.navigateSearchResults(-1, 3, "ArrowDown"), { index: 0, activate: false });
  assert.deepEqual(search.navigateSearchResults(-1, 3, "ArrowUp"), { index: 2, activate: false });
  assert.deepEqual(search.navigateSearchResults(2, 3, "ArrowDown"), { index: 0, activate: false });
  assert.deepEqual(search.navigateSearchResults(0, 3, "ArrowUp"), { index: 2, activate: false });
});

test("search keyboard navigation activates only the selected result with Enter", () => {
  assert.deepEqual(search.navigateSearchResults(1, 3, "Enter"), { index: 1, activate: true });
  assert.deepEqual(search.navigateSearchResults(-1, 3, "Enter"), { index: -1, activate: false });
  assert.deepEqual(search.navigateSearchResults(-1, 0, "ArrowDown"), { index: -1, activate: false });
});

test("highlightMatches preserves the original content and marks every literal match", () => {
  assert.deepEqual(highlightMatches("Plan (Q1), then plan Q1", "(q1)"), [
    { text: "Plan ", highlighted: false },
    { text: "(Q1)", highlighted: true },
    { text: ", then plan Q1", highlighted: false },
  ]);
});

test("highlightMatches marks repeated matches without interpreting regex characters", () => {
  assert.deepEqual(highlightMatches("a+b A+B", "a+b"), [
    { text: "a+b", highlighted: true },
    { text: " ", highlighted: false },
    { text: "A+B", highlighted: true },
  ]);
});

test("highlightMatches keeps source boundaries when case folding expands a character", () => {
  assert.deepEqual(highlightMatches("İx", "I\u0307"), [
    { text: "İ", highlighted: true },
    { text: "x", highlighted: false },
  ]);
});
