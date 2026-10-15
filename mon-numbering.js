// Field monitoring work-scope numbering and strata synchronization helpers.
// Pure functions for parsing, formatting, keyboard handling, and non-destructive strata synchronization.

/**
 * Normalizes fullwidth characters (digits, parentheses, dots, colons, punctuation) to standard ASCII.
 */
export function normalizeFullwidth(str) {
  if (typeof str !== 'string') return '';
  return str
    .replace(/[\uFF10-\uFF19]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)) // ０-９ -> 0-9
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .replace(/【/g, '[')
    .replace(/】/g, ']')
    .replace(/［/g, '[')
    .replace(/］/g, ']')
    .replace(/．/g, '.')
    .replace(/：/g, ':');
}

/**
 * Normalizes a key representation, e.g. "1.(1)", "1.（1）", "(1)", "１．" into standard "1.(1)" or "1." format.
 */
export function normalizeKey(str) {
  if (!str) return '';
  const s = normalizeFullwidth(str.trim());
  // Compound or child: 1.(1), 1(1), (1)
  const comp = s.match(/^(\d+)?[\.]?\((\d+)\)$/);
  if (comp) {
    const main = comp[1] || '1';
    const child = comp[2];
    return `${main}.(${child})`;
  }
  // Main only: 1.
  const mainOnly = s.match(/^(\d+)\.?$/);
  if (mainOnly) {
    return `${mainOnly[1]}.`;
  }
  return s;
}

/**
 * Parses work-scope text into structured items, normalized keys, and detects duplicates/ambiguities.
 * Excludes ordinary dates (e.g. 2026.10.08), depth decimals (e.g. 1.5m), and GPS decimals.
 */
export function parseWorkScope(text) {
  if (typeof text !== 'string' || !text.trim()) {
    return {
      items: [],
      keys: [],
      duplicates: [],
      hasAmbiguous: false,
      warnings: []
    };
  }

  const rawLines = text.replace(/\r\n?/g, '\n').split('\n');
  const items = [];
  const seenKeys = new Map(); // key -> count
  const seenMainParents = new Map(); // mainNum -> count
  let currentMain = null;
  let currentMainHeading = '';

  for (let idx = 0; idx < rawLines.length; idx++) {
    const rawLine = rawLines[idx];
    const trimmedRaw = rawLine.trim();
    if (!trimmedRaw) continue;

    // 1. Check for compound line: e.g. 1.(1) 工項A or 1(1) 工項A or １．（１） 工項A
    const compoundMatch = rawLine.match(/^[ \t]*([0-9\uFF10-\uFF19]+)[ \t]*[\.．]?[ \t]*[（\(]([0-9\uFF10-\uFF19]+)[）\)][ \t]*[:：]?[ \t]*(.*)$/);
    if (compoundMatch) {
      const mainNum = parseInt(normalizeFullwidth(compoundMatch[1]), 10);
      const childNum = parseInt(normalizeFullwidth(compoundMatch[2]), 10);
      const key = `${mainNum}.(${childNum})`;
      const lineText = compoundMatch[3].trim();
      if (currentMain !== mainNum) {
        currentMain = mainNum;
        currentMainHeading = '';
      }

      items.push({
        key,
        mainNum,
        childNum,
        headingText: currentMainHeading,
        lineText,
        rawLine
      });
      seenKeys.set(key, (seenKeys.get(key) || 0) + 1);
      continue;
    }

    // 2. Check for child item: e.g. (1) 工項A or （１） 工項A
    const childMatch = rawLine.match(/^[ \t]*[（\(]([0-9\uFF10-\uFF19]+)[）\)][ \t]*[:：]?[ \t]*(.*)$/);
    if (childMatch) {
      const childNum = parseInt(normalizeFullwidth(childMatch[1]), 10);
      const mainNum = currentMain ?? 1;
      const key = `${mainNum}.(${childNum})`;
      const lineText = childMatch[2].trim();

      items.push({
        key,
        mainNum,
        childNum,
        headingText: currentMainHeading,
        lineText,
        rawLine
      });
      seenKeys.set(key, (seenKeys.get(key) || 0) + 1);
      continue;
    }

    // 3. Check for main heading line: e.g. 1. 公司： or 1. 滯洪池工程 or 1. standalone item
    // Dot MUST NOT be followed by a digit (excludes 1.5m, 2026.10.08, etc.)
    const mainMatch = rawLine.match(/^[ \t]*([0-9\uFF10-\uFF19]+)(?:[\.．](?![0-9\uFF10-\uFF19])|[:：])[ \t]*(.*)$/);
    if (mainMatch) {
      const mainNum = parseInt(normalizeFullwidth(mainMatch[1]), 10);
      const content = mainMatch[2].trim();
      currentMain = mainNum;
      currentMainHeading = content;
      seenMainParents.set(mainNum, (seenMainParents.get(mainNum) || 0) + 1);

      // Check if subsequent lines have children under this main item
      let hasChildren = false;
      for (let j = idx + 1; j < rawLines.length; j++) {
        const nextRaw = rawLines[j];
        if (!nextRaw.trim()) continue;

        // Next main heading encountered
        const nextMainMatch = nextRaw.match(/^[ \t]*([0-9\uFF10-\uFF19]+)(?:[\.．](?![0-9\uFF10-\uFF19])|[:：])/);
        if (nextMainMatch) {
          break;
        }

        // Next line is a compound item: e.g. 1.(1) or 2.(1)
        const nextCompMatch = nextRaw.match(/^[ \t]*([0-9\uFF10-\uFF19]+)[ \t]*[\.．]?[ \t]*[（\(]([0-9\uFF10-\uFF19]+)[）\)]/);
        if (nextCompMatch) {
          const compMain = parseInt(normalizeFullwidth(nextCompMatch[1]), 10);
          if (compMain === mainNum) {
            hasChildren = true;
          }
          break;
        }

        // Next line is a relative child: e.g. (1)
        const nextChildMatch = nextRaw.match(/^[ \t]*[（\(]([0-9\uFF10-\uFF19]+)[）\)]/);
        if (nextChildMatch) {
          hasChildren = true;
          break;
        }
      }

      // If it has NO children and has non-empty content (single-level standalone item), treat as item
      if (!hasChildren && content) {
        const key = `${mainNum}.`;
        items.push({
          key,
          mainNum,
          childNum: null,
          headingText: content,
          lineText: content,
          rawLine
        });
        seenKeys.set(key, (seenKeys.get(key) || 0) + 1);
      }
      continue;
    }
  }

  const duplicates = [];
  for (const [k, count] of seenKeys.entries()) {
    if (count > 1) duplicates.push(k);
  }
  const duplicateParents = [];
  for (const [m, count] of seenMainParents.entries()) {
    if (count > 1) {
      const k = `${m}.`;
      duplicateParents.push(k);
      if (!duplicates.includes(k)) {
        duplicates.push(k);
      }
    }
  }

  const warnings = [];
  if (duplicateParents.length > 0) {
    warnings.push(`主工項編號重複：${duplicateParents.join(', ')}，未進行推測指派；請修正編號。`);
  }
  const itemDuplicates = duplicates.filter(d => !duplicateParents.includes(d));
  if (itemDuplicates.length > 0) {
    warnings.push(`工項編號重複：${itemDuplicates.join(', ')}，未進行推測指派；請修正編號。`);
  }

  return {
    items,
    keys: items.map(it => it.key),
    duplicates,
    hasAmbiguous: duplicates.length > 0,
    warnings
  };
}

/**
 * Parses existing strata column text (strataSummary or strataDesc) into keyed entries,
 * unmatched entries, and any legacy unstructured lines.
 */
export function parseStrataEntries(text) {
  if (typeof text !== 'string' || !text.length) {
    return {
      hasKeyed: false,
      entries: [],
      unstructured: '',
      hasLeadingUnstructured: false,
      hasAmbiguous: false,
      duplicates: [],
      unmatched: []
    };
  }

  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  const entries = [];
  let current = null;
  let hasKeyed = false;
  const unstructuredLines = [];
  const seenKeys = new Map();

  for (const line of lines) {
    const unmatchedMatch = line.match(/^([ \t]*)\[未對應項目\][ \t]*(.*)$/);
    const isUnmatched = Boolean(unmatchedMatch);
    const lineToTest = unmatchedMatch ? unmatchedMatch[2] : line;
    const unmatchedIndent = unmatchedMatch ? unmatchedMatch[1] : '';

    // Check for compound key: e.g. 1.(1) or １．（１） or (1) or （１） or 1(1)
    const compMatch = lineToTest.match(/^([ \t]*)((?:[0-9\uFF10-\uFF19]+[ \t]*[\.．]?[ \t]*)?[（\(][0-9\uFF10-\uFF19]+[）\)])([ \t]*[:：]?[ \t]*)([\s\S]*)$/);

    // Check for main key: e.g. 1. or １． or 1: or 1：
    // Must NOT be followed by a digit (excludes 1.5m, 2026.10.08, coordinates)
    const mainMatch = lineToTest.match(/^([ \t]*)([0-9\uFF10-\uFF19]+(?:[\.．](?![0-9\uFF10-\uFF19])|[:：]))([ \t]*[:：]?[ \t]*)([\s\S]*)$/);

    const m = compMatch || mainMatch;
    if (m) {
      hasKeyed = true;
      const indent = unmatchedIndent + m[1];
      const rawKey = m[2];
      const normalizedKey = normalizeKey(rawKey);
      const sep = m[3];
      const firstLineContent = m[4];

      if (current) entries.push(current);
      current = {
        normalizedKey,
        rawKey,
        indent,
        sep,
        isUnmatched,
        lines: [firstLineContent]
      };
      seenKeys.set(normalizedKey, (seenKeys.get(normalizedKey) || 0) + 1);
    } else if (current) {
      // Continuation line for current entry - preserve multiline indentation and whitespace exactly
      current.lines.push(line);
    } else {
      unstructuredLines.push(line);
    }
  }
  if (current) entries.push(current);

  const duplicates = [];
  for (const [k, count] of seenKeys.entries()) {
    if (count > 1) duplicates.push(k);
  }

  return {
    hasKeyed,
    entries,
    unstructured: unstructuredLines.join('\n'),
    hasLeadingUnstructured: unstructuredLines.length > 0,
    hasAmbiguous: duplicates.length > 0,
    duplicates
  };
}

/**
 * Synchronizes a single strata column (strataSummary or strataDesc) nondestructively.
 * - Auto-adds EMPTY keyed stubs only.
 * - Preserves existing user soil/summary text and exact spacing.
 * - Preserves removed items visibly marked with [未對應項目].
 * - Reclaims unmatched items if key is re-added.
 * - Keeps legacy unstructured paragraphs intact without guessing.
 */
export function synchronizeStrataColumn(currentText, activeKeys) {
  const rawText = typeof currentText === 'string' ? currentText : '';
  // Parent headings organize work only; observations map to children, never to 1. / 2.
  const validActiveKeys = Array.isArray(activeKeys) ? activeKeys.filter(k => /^\d+\.\(\d+\)$/.test(k)) : [];

  // If strata field is completely empty:
  if (!rawText.trim()) {
    if (validActiveKeys.length > 0) {
      const text = validActiveKeys.map(k => `${k} `).join('\n');
      return {
        text,
        changed: true,
        warning: null,
        isUnstructured: false,
        unmatchedCount: 0,
        stubsAddedCount: validActiveKeys.length,
        hasAmbiguous: false,
        duplicates: []
      };
    }
    return {
      text: rawText,
      changed: false,
      warning: null,
      isUnstructured: false,
      unmatchedCount: 0,
      stubsAddedCount: 0,
      hasAmbiguous: false,
      duplicates: []
    };
  }

  const parsed = parseStrataEntries(rawText);
  // Remove only empty parent stubs generated by older versions. Never discard an observation.
  parsed.entries = parsed.entries.filter(entry => !(/^\d+\.$/.test(entry.normalizedKey) && entry.lines.every(line => !line.trim())));
  const counts = new Map();
  for (const entry of parsed.entries) counts.set(entry.normalizedKey, (counts.get(entry.normalizedKey) || 0) + 1);
  parsed.duplicates = [...counts].filter(([, count]) => count > 1).map(([key]) => key);
  parsed.hasAmbiguous = parsed.duplicates.length > 0;

  // If text has NO keyed entries at all, it's legacy unstructured paragraphs:
  // Must remain available and editable, with warning, not overwritten!
  if (!parsed.hasKeyed) {
    return {
      text: rawText,
      changed: false,
      warning: '地層欄位包含未編號之既有文字，保留原內容不自動覆寫；可手動加上編號進行同步。',
      isUnstructured: true,
      unmatchedCount: 0,
      stubsAddedCount: 0,
      hasAmbiguous: false,
      duplicates: []
    };
  }

  // Index existing entries by normalized key, preserving ALL entries as separate records
  const existingMap = new Map(); // key -> Array<Entry>
  for (const entry of parsed.entries) {
    if (!existingMap.has(entry.normalizedKey)) {
      existingMap.set(entry.normalizedKey, [entry]);
    } else {
      existingMap.get(entry.normalizedKey).push(entry);
    }
  }

  let stubsAddedCount = 0;
  let unmatchedCount = 0;
  let retainedParents = 0;
  const processedKeys = new Set();
  const outputBlocks = [];

  // If there was unstructured text / leading blank lines before the first keyed entry, preserve it
  if (parsed.hasLeadingUnstructured) {
    outputBlocks.push(parsed.unstructured);
  }

  // Helper to format an entry
  function formatEntry(entry, key, isOrphan = false) {
    const rawContent = entry.lines[0] ?? '';
    let sep = entry.sep || '';
    if (!sep && rawContent.length > 0 && !rawContent.startsWith(' ') && !rawContent.startsWith('\t')) {
      sep = ' ';
    }
    const indent = entry.indent || '';
    const prefix = isOrphan ? '[未對應項目] ' : '';
    const headerLine = `${prefix}${indent}${key}${sep}${rawContent}`;
    return [headerLine, ...entry.lines.slice(1)].join('\n');
  }

  // 1. Process active keys in order
  for (const key of validActiveKeys) {
    processedKeys.add(key);
    if (existingMap.has(key)) {
      const entries = existingMap.get(key);
      for (const entry of entries) {
        outputBlocks.push(formatEntry(entry, key, false));
      }
    } else {
      // Auto-add EMPTY keyed stub only! NEVER infer soil observations!
      outputBlocks.push(`${key} `);
      stubsAddedCount++;
    }
  }

  // 2. Process orphaned/removed entries that were in existing text
  for (const [key, entries] of existingMap.entries()) {
    if (!processedKeys.has(key)) {
      for (const entry of entries) {
        if (/^\d+\.$/.test(key)) {
          // Old observations are not empty stubs: retain without claiming a child assignment.
          retainedParents++;
          outputBlocks.push(formatEntry(entry, key, false));
        } else {
          unmatchedCount++;
          outputBlocks.push(formatEntry(entry, key, true));
        }
      }
    }
  }

  const updatedText = outputBlocks.join('\n');
  const changed = updatedText !== rawText;
  const warnings = [];
  if (retainedParents > 0) warnings.push('既有主項目下仍有觀察文字，已保留；請手動移至正確子項目，不會自動分配或刪除。');
  if (unmatchedCount > 0) {
    warnings.push(`偵測到 ${unmatchedCount} 項未對應工項的地層紀錄（已保留於欄位下方標示 [未對應項目]，未自動刪除）。`);
  }
  if (parsed.hasAmbiguous) {
    warnings.push(`地層欄位包含重複編號：${parsed.duplicates.join(', ')}，已保留個別紀錄項目；請修正編號避免歧義。`);
  }

  return {
    text: updatedText,
    changed,
    warning: warnings.length > 0 ? warnings.join('；') : null,
    isUnstructured: false,
    unmatchedCount,
    stubsAddedCount,
    hasAmbiguous: parsed.hasAmbiguous,
    duplicates: parsed.duplicates
  };
}

/**
 * Synchronizes BOTH strataSummary and strataDesc with work-scope keys.
 */
export function synchronizeBothStrata(workScopeText, strataSummaryText, strataDescText) {
  const parsedScope = parseWorkScope(workScopeText);

  // If duplicate keys exist, report duplicate labels and do not perform guessed assignment
  if (parsedScope.hasAmbiguous) {
    return {
      strataSummary: strataSummaryText,
      strataDesc: strataDescText,
      changed: false,
      warnings: parsedScope.warnings,
      duplicates: parsedScope.duplicates,
      hasAmbiguous: true,
      isUnstructured: false,
      stubsAdded: 0,
      unmatchedCount: 0
    };
  }

  const childKeys = parsedScope.items.filter(item => item.childNum !== null).map(item => item.key);
  const resSummary = synchronizeStrataColumn(strataSummaryText, childKeys);
  const resDesc = synchronizeStrataColumn(strataDescText, childKeys);

  const warnings = [];
  const hasUnstructuredScope = typeof workScopeText === 'string' && workScopeText.trim().length > 0 && parsedScope.items.length === 0;
  if (hasUnstructuredScope) {
    warnings.push('作業項目未包含編號結構，保留原內容不自動同步；可使用按鈕新增編號或手動維護。');
  }
  if (resSummary.warning) warnings.push(resSummary.warning);
  if (resDesc.warning && resDesc.warning !== resSummary.warning) warnings.push(resDesc.warning);

  const hasAmbiguous = resSummary.hasAmbiguous || resDesc.hasAmbiguous;
  const duplicates = [...new Set([...(resSummary.duplicates || []), ...(resDesc.duplicates || [])])];

  return {
    strataSummary: resSummary.text,
    strataDesc: resDesc.text,
    changed: resSummary.changed || resDesc.changed,
    warnings,
    duplicates,
    hasAmbiguous,
    isUnstructured: resSummary.isUnstructured || resDesc.isUnstructured || hasUnstructuredScope,
    stubsAdded: resSummary.stubsAddedCount + resDesc.stubsAddedCount,
    unmatchedCount: resSummary.unmatchedCount + resDesc.unmatchedCount
  };
}

/**
 * Finds the next main item number from work-scope text.
 */
export function getNextMainNumber(text) {
  if (typeof text !== 'string' || !text.trim()) return 1;
  const rawLines = text.replace(/\r\n?/g, '\n').split('\n');
  let maxMain = 0;
  for (const line of rawLines) {
    if (!line.trim()) continue;
    // Compound item: e.g. 2.(1)
    const compMatch = line.match(/^[ \t]*([0-9\uFF10-\uFF19]+)[ \t]*[\.．]?[ \t]*[（\(]([0-9\uFF10-\uFF19]+)[）\)]/);
    if (compMatch) {
      const num = parseInt(normalizeFullwidth(compMatch[1]), 10);
      if (num > maxMain) maxMain = num;
      continue;
    }
    // Main heading: e.g. 2. or 2:
    const mainMatch = line.match(/^[ \t]*([0-9\uFF10-\uFF19]+)(?:[\.．](?![0-9\uFF10-\uFF19])|[:：])/);
    if (mainMatch) {
      const num = parseInt(normalizeFullwidth(mainMatch[1]), 10);
      if (num > maxMain) maxMain = num;
      continue;
    }
  }
  return maxMain > 0 ? maxMain + 1 : 1;
}

/**
 * Finds the next child number under the current main item.
 * If no main item exists, returns 1.
 */
export function getNextChildNumber(text, caretPos = text?.length ?? 0) {
  if (typeof text !== 'string' || !text.trim()) return 1;
  const textBeforeCaret = text.slice(0, caretPos);
  if (!textBeforeCaret.trim()) return 1;

  const rawLines = textBeforeCaret.replace(/\r\n?/g, '\n').split('\n');
  let currentMain = null;
  let maxChildForCurrentMain = 0;

  for (const line of rawLines) {
    if (!line.trim()) continue;

    // Compound: e.g. 1.(1) or 2.(3)
    const compMatch = line.match(/^[ \t]*([0-9\uFF10-\uFF19]+)[ \t]*[\.．]?[ \t]*[（\(]([0-9\uFF10-\uFF19]+)[）\)]/);
    if (compMatch) {
      const mainNum = parseInt(normalizeFullwidth(compMatch[1]), 10);
      const childNum = parseInt(normalizeFullwidth(compMatch[2]), 10);
      if (mainNum !== currentMain) {
        currentMain = mainNum;
        maxChildForCurrentMain = childNum;
      } else {
        if (childNum > maxChildForCurrentMain) maxChildForCurrentMain = childNum;
      }
      continue;
    }

    // Main heading (including bare 2. or 2:): e.g. 2. or 2. 公司：
    const mainMatch = line.match(/^[ \t]*([0-9\uFF10-\uFF19]+)(?:[\.．](?![0-9\uFF10-\uFF19])|[:：])/);
    if (mainMatch) {
      const mainNum = parseInt(normalizeFullwidth(mainMatch[1]), 10);
      currentMain = mainNum;
      maxChildForCurrentMain = 0; // resets counter for new parent!
      continue;
    }

    // Relative child: e.g. (1)
    const childMatch = line.match(/^[ \t]*[（\(]([0-9\uFF10-\uFF19]+)[）\)]/);
    if (childMatch) {
      const childNum = parseInt(normalizeFullwidth(childMatch[1]), 10);
      if (currentMain === null) currentMain = 1;
      if (childNum > maxChildForCurrentMain) maxChildForCurrentMain = childNum;
      continue;
    }
  }

  return maxChildForCurrentMain + 1;
}

/**
 * Safely inserts text at the textarea's current selection/caret,
 * preserving the browser's undo/redo history where supported.
 */
export function insertTextAtCaret(textarea, textToInsert) {
  if (!textarea) return;
  textarea.focus();

  // Try document.execCommand first to preserve native undo stack
  let inserted = false;
  if (typeof document !== 'undefined' && typeof document.execCommand === 'function') {
    try {
      inserted = document.execCommand('insertText', false, textToInsert);
    } catch {
      inserted = false;
    }
  }

  // Fallback to setRangeText if execCommand is unavailable or failed
  if (!inserted) {
    const start = textarea.selectionStart ?? textarea.value.length;
    const end = textarea.selectionEnd ?? textarea.value.length;
    textarea.setRangeText(textToInsert, start, end, 'end');
    textarea.dispatchEvent(new Event('input', { bubbles: true }));
  }
}

/**
 * Handles keydown on the work-scope textarea:
 * - Enter after populated main heading (e.g. "1. 公司：") inserts "\n(1) "
 * - Enter after populated child(n) inserts "\n(n+1) "
 * - Shift+Enter creates a plain newline
 * - IME composition (isComposing, keyCode 229) is untouched
 * - Ordinary description lines are untouched (native Enter)
 */
export function handleWorkScopeKeyDown(e, textarea) {
  if (!textarea || e.key !== 'Enter') return;

  // Shift+Enter -> plain newline
  if (e.shiftKey) return;

  // IME composition in progress
  if (e.isComposing || e.keyCode === 229) return;

  const val = textarea.value;
  const selStart = textarea.selectionStart ?? val.length;
  const selEnd = textarea.selectionEnd ?? val.length;

  const beforeCaret = val.slice(0, selStart);
  const lastNewline = beforeCaret.lastIndexOf('\n');
  const lineStart = lastNewline === -1 ? 0 : lastNewline + 1;
  const nextNewline = val.indexOf('\n', selEnd);
  const lineEnd = nextNewline === -1 ? val.length : nextNewline;

  const fullLine = val.slice(lineStart, lineEnd);
  const normLine = normalizeFullwidth(fullLine).trim();

  // 1. Check for populated child item: e.g. "(1) 滯1區域開挖" or "1.(1) 開挖"
  const childMatch = normLine.match(/^(\d+)?\.?\((\d+)\)\s*[:：]?(.*)$/);
  if (childMatch) {
    const content = childMatch[3].trim();
    // Only intercept if the child line is POPULATED (has content typed by user)
    if (content.length > 0) {
      e.preventDefault();
      const curChild = parseInt(childMatch[2], 10);
      const nextChild = curChild + 1;
      const mainPrefix = childMatch[1] ? `${childMatch[1]}.` : '';
      const insertion = `\n${mainPrefix}(${nextChild}) `;
      insertTextAtCaret(textarea, insertion);
      return;
    }
    // If bare child without content, let native Enter happen
    return;
  }

  // 2. Check for populated main heading: e.g. "1. 公司："
  const mainMatch = normLine.match(/^(\d+)(?:\.(?!\d)|[:：])\s*(.*)$/);
  if (mainMatch) {
    const content = mainMatch[2].trim();
    // Only intercept if the main heading is POPULATED
    if (content.length > 0) {
      e.preventDefault();
      const insertion = `\n(1) `;
      insertTextAtCaret(textarea, insertion);
      return;
    }
    // If bare heading without content, let native Enter happen
    return;
  }

  // Ordinary line / description: do NOT hijack Enter!
}
