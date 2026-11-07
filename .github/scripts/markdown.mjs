// Render values from the GitHub API as data inside Markdown job summaries.
//
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// Line breaks would end a table row; a pipe would split a table cell.
function flatten(value) {
  return String(value).replace(/\r\n|\r|\n/g, " ").replaceAll("|", "\\|");
}

// A code span whose fence is longer than any backtick run in the value, so a
// value with backticks can neither close the span early nor open another.
export function markdownCode(value) {
  const text = flatten(value);
  const longest = Math.max(0, ...(text.match(/`+/g) ?? []).map(run => run.length));
  const fence = "`".repeat(longest + 1);
  const pad = text === "" || text.startsWith("`") || text.endsWith("`") ? " " : "";
  return `${fence}${pad}${text}${pad}${fence}`;
}

// Plain text: backslashes are escaped first, so the value cannot escape the
// pipe added by flatten() or any character after it.
export function markdownText(value) {
  return flatten(String(value).replaceAll("\\", "\\\\"));
}
