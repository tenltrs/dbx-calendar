/** Returns matching note entries with the newest date first. */
export function searchNotes(notes, query) {
  const term = query.trim();
  if (!term) return [];

  const normalizedTerm = term.toLowerCase();
  return Object.entries(notes)
    .filter(([, content]) => String(content).toLowerCase().includes(normalizedTerm))
    .sort(([left], [right]) => right.localeCompare(left))
    .map(([date, content]) => ({ date, content: String(content) }));
}

/** Moves the active search result with the keyboard, wrapping at either end. */
export function navigateSearchResults(activeIndex, resultCount, key) {
  const count = Math.max(0, Math.trunc(Number(resultCount) || 0));
  if (!count) return { index: -1, activate: false };

  const currentIndex = Number.isInteger(activeIndex) && activeIndex >= 0 && activeIndex < count
    ? activeIndex
    : -1;

  if (key === "ArrowDown") {
    return { index: currentIndex < 0 ? 0 : (currentIndex + 1) % count, activate: false };
  }
  if (key === "ArrowUp") {
    return { index: currentIndex < 0 ? count - 1 : (currentIndex - 1 + count) % count, activate: false };
  }
  if (key === "Enter") return { index: currentIndex, activate: currentIndex >= 0 };
  return { index: currentIndex, activate: false };
}

/** Splits text into safe-to-render pieces, marking each literal match. */
export function highlightMatches(content, query) {
  const text = String(content);
  const term = query.trim();
  if (!term) return [{ text, highlighted: false }];

  let normalizedText = "";
  const sourceStarts = [];
  const sourceEnds = [];
  for (let index = 0; index < text.length; ) {
    const sourceCharacter = String.fromCodePoint(text.codePointAt(index));
    const normalizedCharacter = sourceCharacter.toLowerCase();
    normalizedText += normalizedCharacter;
    for (let offset = 0; offset < normalizedCharacter.length; offset += 1) {
      sourceStarts.push(index);
      sourceEnds.push(index + sourceCharacter.length);
    }
    index += sourceCharacter.length;
  }

  const normalizedTerm = term.toLowerCase();
  const segments = [];
  let cursor = 0;
  let normalizedCursor = 0;

  while (normalizedCursor < normalizedText.length) {
    const matchAt = normalizedText.indexOf(normalizedTerm, normalizedCursor);
    if (matchAt === -1) break;
    const sourceStart = sourceStarts[matchAt];
    const sourceEnd = sourceEnds[matchAt + normalizedTerm.length - 1];
    if (sourceStart > cursor) segments.push({ text: text.slice(cursor, sourceStart), highlighted: false });
    segments.push({ text: text.slice(sourceStart, sourceEnd), highlighted: true });
    cursor = sourceEnd;
    normalizedCursor = matchAt + normalizedTerm.length;
  }

  if (cursor < text.length) segments.push({ text: text.slice(cursor), highlighted: false });
  return segments.length ? segments : [{ text, highlighted: false }];
}
