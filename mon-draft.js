export const MON_FIELDS = {
  projectName: 'monProject',
  date: 'monDate',
  constructionLocation: 'monLocation',
  siteName: 'monSite',
  watcherCount: 'monWatchers',
  workScope: 'monWorkScope',
  strataSummary: 'monStrataSummary',
  strataDesc: 'monStrata',
  culturalArtifacts: 'monArtifacts',
  culturalFeatures: 'monFeatures',
  otherEvents: 'monOtherEvents',
  dispositionAdvice: 'monAdvice'
};

export function monKey(day) {
  return `field-mon-form-${day}`;
}

export function formatMonDate(day) {
  if (typeof day !== 'string' || !day.includes('-')) return '';
  const [y, m, d] = day.split('-');
  return `${y}年${parseInt(m, 10)}月${parseInt(d, 10)}日`;
}

export function genericMonTemplate(day) {
  return {
    projectName: '',
    date: formatMonDate(day),
    constructionLocation: '',
    siteName: '',
    watcherCount: '1',
    workScope: '',
    strataSummary: '1.(1) 原堆積。',
    strataDesc: '',
    culturalArtifacts: '無。',
    culturalFeatures: '無。',
    otherEvents: '無。',
    dispositionAdvice: '無。'
  };
}

export function isMonFormDirty(currentInputs, baselineInputs) {
  if (!currentInputs) return false;
  if (!baselineInputs) {
    return Object.values(currentInputs).some(v => typeof v === 'string' && v.trim().length > 0);
  }
  for (const key of Object.keys(MON_FIELDS)) {
    const cur = currentInputs[key] ?? '';
    const base = baselineInputs[key] ?? '';
    if (cur !== base) return true;
  }
  return false;
}

export function safeGetLocalStorage() {
  try {
    if (typeof window !== 'undefined' && 'localStorage' in window) {
      return window.localStorage;
    }
    if (typeof globalThis !== 'undefined' && 'localStorage' in globalThis) {
      return globalThis.localStorage;
    }
    return null;
  } catch {
    return null;
  }
}

export class MonitoringDraftManager {
  constructor(storage = undefined) {
    if (storage !== undefined) {
      this.storage = storage;
    } else {
      this.storage = safeGetLocalStorage();
    }
    // Map<day, draftObject> (preserves unknown metadata and exact whitespace)
    this.inMemoryCopies = new Map();
    // Map<day, rawString> for corrupted / unparsable payloads
    this.corruptRawStorage = new Map();
    // Map<day, { status: string, isError: boolean, isUnsaved: boolean, isReadBlocked?: boolean, isCorrupt?: boolean, isConflict?: boolean }>
    this.dayStatus = new Map();
    // Set<day> for days where storage getItem failed/threw (read-blocked: unread, NOT empty)
    this.readBlockedDays = new Set();
    // Map<day, rawString | null> for base revision tracking (cross-tab conflict detection)
    this.baseRawStorage = new Map();
    // Map<day, { externalRaw: string, externalData: any, localData: any }>
    this.conflictStorage = new Map();
  }

  getDraft(day) {
    // Navigation must not resolve a blocked read or discard recovery buffers.
    if (this.inMemoryCopies.has(day)) return this.inMemoryCopies.get(day);

    if (!this.storage) {
      this.dayStatus.set(day, {
        status: '無可用本機儲存，草稿僅存於本次記憶體（重新載入將遺失）',
        isError: true,
        isUnsaved: false
      });
      return null;
    }

    let raw = null;
    try {
      raw = this.storage.getItem(monKey(day));
      this.readBlockedDays.delete(day);
      this.baseRawStorage.set(day, raw);
    } catch (e) {
      this.readBlockedDays.add(day);
      this.dayStatus.set(day, {
        status: `讀取草稿失敗（儲存區無法存取：${e.message || '存取受阻'}；已暫停自動覆寫現有儲存資料，現有變更保留於記憶體中）`,
        isError: true,
        isUnsaved: this.inMemoryCopies.has(day),
        isReadBlocked: true
      });
      return null;
    }

    if (raw === null || raw === undefined) {
      this.baseRawStorage.set(day, null);
      return null;
    }

    try {
      const parsed = JSON.parse(raw);
      if (parsed !== null && typeof parsed === 'object' && !Array.isArray(parsed)) {
        this.inMemoryCopies.set(day, parsed);
        return parsed;
      }
      // If it parsed to null, primitive, or array: treat as unexpected payload
      this.corruptRawStorage.set(day, raw);
      this.dayStatus.set(day, {
        status: '本機草稿資料格式異常（已保留原始資料，未自動覆寫；請檢查或匯出複製）',
        isError: true,
        isUnsaved: false,
        isCorrupt: true
      });
      return null;
    } catch {
      // JSON syntax error
      this.corruptRawStorage.set(day, raw);
      this.dayStatus.set(day, {
        status: '本機草稿損毀（已保留原始損毀字串，未自動覆寫；可匯出保留字串）',
        isError: true,
        isUnsaved: false,
        isCorrupt: true
      });
      return null;
    }
  }

  saveDraft(day, inputValues) {
    const existing = this.inMemoryCopies.get(day) || {};
    // Safe merge: preserve unknown metadata, overwrite with exact current inputs (exact whitespace preserved)
    const merged = { ...existing, ...inputValues };
    this.inMemoryCopies.set(day, merged);

    // 1. Guard against autosaving over unread / read-blocked data
    if (this.readBlockedDays.has(day)) {
      this.dayStatus.set(day, {
        status: '本機草稿讀取失敗／受阻，已保留記憶體中編輯，已暫停自動覆寫現有儲存資料；可重試讀取、匯出草稿或確認安全覆寫。',
        isError: true,
        isUnsaved: true,
        isReadBlocked: true
      });
      return { success: false, reason: 'read_blocked', data: merged };
    }

    // 2. Guard against autosaving over corrupt payload without confirmation
    if (this.corruptRawStorage.has(day)) {
      this.dayStatus.set(day, {
        status: '本機發現損毀／異常草稿，已暫停自動覆寫硬碟；請先點「匯出備份草稿」或「確認覆寫」以完成儲存。',
        isError: true,
        isUnsaved: true,
        isCorrupt: true
      });
      return { success: false, reason: 'corrupt_blocked', data: merged };
    }

    // 3. No storage available
    if (!this.storage) {
      this.dayStatus.set(day, {
        status: '無可用本機儲存，草稿僅存於本次記憶體（重新載入將遺失）',
        isError: true,
        isUnsaved: true
      });
      return { success: false, reason: 'no_storage', data: merged };
    }

    // 4. Cross-tab conflict detection using read base / raw revision
    // LocalStorage does not provide multi-process transactional locks; this optimistic check
    // prevents stale cached tabs from silently overwriting newer external modifications.
    let currentDiskRaw = null;
    try {
      currentDiskRaw = this.storage.getItem(monKey(day));
    } catch (e) {
      this.readBlockedDays.add(day);
      this.dayStatus.set(day, {
        status: `本機儲存區存取受阻（${e.message || '無法讀取'}）；已保留記憶體中編輯，暫停寫入。`,
        isError: true,
        isUnsaved: true,
        isReadBlocked: true
      });
      return { success: false, reason: 'read_blocked', data: merged };
    }

    const baseRaw = this.baseRawStorage.get(day);
    if ((baseRaw !== undefined && currentDiskRaw !== baseRaw) || this.hasConflict(day)) {
      return this.recordConflict(day, currentDiskRaw);
    }
    return this.persistDraft(day, merged, currentDiskRaw);
  }

  recordConflict(day, raw) {
    const previous = this.conflictStorage.get(day);
    const externalHistory = [...new Set([...(previous?.externalHistory || []), ...(previous ? [previous.externalRaw] : []), raw])];
    let externalData = null;
    try { externalData = JSON.parse(raw); } catch {}
    const data = this.inMemoryCopies.get(day) || {};
    this.conflictStorage.set(day, { externalRaw: raw, externalData, externalHistory, localData: data });
    this.dayStatus.set(day, {
      status: '偵測到其他分頁已更新或刪除草稿；已暫停覆寫。請先匯出備份，再另按「確認覆寫衝突草稿」決定是否取代外部資料。',
      isError: true, isUnsaved: true, isConflict: true
    });
    return { success: false, reason: 'conflict', externalRaw: raw, externalData, data };
  }

  // Compare again immediately before writing, then read back. This is optimistic
  // protection, not a cross-process atomic localStorage transaction.
  persistDraft(day, data, expectedRaw) {
    try {
      if (!this.storage) throw Error('無可用本機儲存');
      const raw = this.storage.getItem(monKey(day));
      if (raw !== expectedRaw) return this.recordConflict(day, raw);
      const serialized = JSON.stringify(data);
      this.storage.setItem(monKey(day), serialized);
      const persisted = this.storage.getItem(monKey(day));
      if (persisted !== serialized) return this.recordConflict(day, persisted);
      this.baseRawStorage.set(day, serialized);
      this.conflictStorage.delete(day);
      this.corruptRawStorage.delete(day);
      this.readBlockedDays.delete(day);
      this.dayStatus.set(day, { status: '草稿已自動儲存於本機。', isError: false, isUnsaved: false });
      return { success: true, data };
    } catch (e) {
      const quota = e?.name === 'QuotaExceededError' || /quota/i.test(e?.message || '');
      this.dayStatus.set(day, {
        status: `草稿儲存失敗（${quota ? '容量配額已滿' : e.message || '寫入失敗'}）；記憶體變更與異常／衝突複本仍保留，可重試或匯出；重新整理將遺失未存變更。`,
        isError: true, isUnsaved: true, isConflict: this.hasConflict(day), isCorrupt: this.hasCorrupt(day), isReadBlocked: this.isReadBlocked(day)
      });
      return { success: false, reason: quota ? 'quota_exceeded' : 'write_error', error: e, data };
    }
  }

  forceOverwriteCorrupt(day) {
    const inMem = this.inMemoryCopies.get(day) || {};

    if (!this.storage) {
      this.dayStatus.set(day, {
        status: '無可用本機儲存，無法覆寫硬碟。',
        isError: true,
        isUnsaved: true
      });
      return { success: false, reason: 'no_storage', data: inMem };
    }

    return this.persistDraft(day, inMem, this.corruptRawStorage.get(day));
  }

  forceOverwriteConflict(day, expectedRaw = this.conflictStorage.get(day)?.externalRaw) {
    if (!this.conflictStorage.has(day)) {
      return this.saveDraft(day, this.inMemoryCopies.get(day) || {});
    }
    return this.persistDraft(day, this.inMemoryCopies.get(day) || {}, expectedRaw);
  }

  resolveConflictKeepExternal(day) {
    this.conflictStorage.delete(day);
    this.inMemoryCopies.delete(day);
    this.baseRawStorage.delete(day);
    return this.getDraft(day);
  }

  resolveConflictMerge(day) {
    const conflict = this.conflictStorage.get(day);
    const ext = conflict?.externalData || {};
    const local = this.inMemoryCopies.get(day) || {};
    const merged = { ...ext, ...local };
    this.inMemoryCopies.set(day, merged);
    return this.persistDraft(day, merged, conflict ? conflict.externalRaw : this.baseRawStorage.get(day));
  }

  retryRead(day) {
    if (!this.storage) return { success: false, reason: 'no_storage' };
    try {
      const raw = this.storage.getItem(monKey(day));
      let parsed = null;
      if (raw !== null) {
        try {
          parsed = JSON.parse(raw);
          if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) throw Error('異常格式');
        } catch {
          this.corruptRawStorage.set(day, raw);
          this.dayStatus.set(day, { status: '讀取恢復，但草稿損毀／異常；原始字串與記憶體變更均保留，未覆寫。', isError: true, isUnsaved: this.inMemoryCopies.has(day), isCorrupt: true });
          return { success: false, reason: 'corrupt_blocked', data: this.inMemoryCopies.get(day) };
        }
      }
      this.readBlockedDays.delete(day);
      if (this.inMemoryCopies.has(day) && raw !== this.baseRawStorage.get(day) && (raw !== null || this.baseRawStorage.has(day))) return this.recordConflict(day, raw);
      if (!this.inMemoryCopies.has(day) && parsed) this.inMemoryCopies.set(day, parsed);
      this.baseRawStorage.set(day, raw);
      return this.saveDraft(day, this.inMemoryCopies.get(day) || {});
    } catch (e) {
      this.dayStatus.set(day, {
        status: `讀取草稿失敗（${e.message || '存取受阻'}）；已保留記憶體變更。`,
        isError: true,
        isUnsaved: true,
        isReadBlocked: true
      });
      return { success: false, reason: 'read_blocked', error: e };
    }
  }

  forceOverwriteReadBlocked(day) {
    this.readBlockedDays.delete(day);
    const inMem = this.inMemoryCopies.get(day) || {};
    return this.saveDraft(day, inMem);
  }

  isReadBlocked(day) {
    return this.readBlockedDays.has(day);
  }

  hasConflict(day) {
    return this.conflictStorage.has(day);
  }

  getConflict(day) {
    return this.conflictStorage.get(day) || null;
  }

  getRawCorrupt(day) {
    return this.corruptRawStorage.get(day) || null;
  }

  hasCorrupt(day) {
    return this.corruptRawStorage.has(day);
  }

  hasUnsaved(day) {
    const s = this.dayStatus.get(day);
    return Boolean(s && s.isUnsaved);
  }

  hasAnyUnsaved() {
    for (const [, s] of this.dayStatus.entries()) {
      if (s && s.isUnsaved) return true;
    }
    return false;
  }

  getUnsavedDays() {
    const list = [];
    for (const [day, s] of this.dayStatus.entries()) {
      if (s && s.isUnsaved) list.push(day);
    }
    return list;
  }

  getStatus(day) {
    return this.dayStatus.get(day) || null;
  }

  setStatus(day, statusObj) {
    this.dayStatus.set(day, statusObj);
  }

  clearCache(day) {
    if (this.hasUnsaved(day) || this.hasConflict(day) || this.hasCorrupt(day) || this.isReadBlocked(day)) return false;
    this.inMemoryCopies.delete(day);
    this.baseRawStorage.delete(day);
    this.corruptRawStorage.delete(day);
    this.readBlockedDays.delete(day);
    this.conflictStorage.delete(day);
    this.dayStatus.delete(day);
  }

  checkExternalChange(day) {
    if (!this.storage) return false;
    let diskRaw = null;
    try {
      diskRaw = this.storage.getItem(monKey(day));
    } catch {
      return false;
    }
    const baseRaw = this.baseRawStorage.get(day);
    if (baseRaw !== undefined && diskRaw !== baseRaw) {
      this.recordConflict(day, diskRaw);
      return true;
    }
    return false;
  }

  exportCopy(day) {
    const corrupt = this.corruptRawStorage.get(day);
    const mem = this.inMemoryCopies.get(day);
    const conflict = this.conflictStorage.get(day);

    if (corrupt !== undefined) {
      let combined = `=== 本機原始損毀草稿內容 ===\n${corrupt}\n\n=== 目前記憶體中未儲存編輯值與中繼資料 ===\n`;
      if (mem && Object.keys(mem).length > 0) {
        combined += JSON.stringify(mem, null, 2);
      } else {
        combined += '（目前無未儲存之記憶體編輯內容）';
      }
      return {
        type: 'corrupt_raw',
        raw: corrupt,
        rawCorrupt: corrupt,
        conflict: conflict || null,
        workingData: mem || null,
        workingJson: mem ? JSON.stringify(mem, null, 2) : null,
        combinedText: combined + (conflict ? '\n\n=== 外部衝突複本 ===\n' + JSON.stringify(conflict, null, 2) : '')
      };
    }

    if (conflict) {
      let combined = `=== 外部儲存區最新內容（null 代表已刪除） ===\n${conflict.externalRaw}\n\n=== 外部版本複本 ===\n${JSON.stringify(conflict.externalHistory)}\n\n=== 本地未儲存變更內容 ===\n`;
      if (mem) {
        combined += JSON.stringify(mem, null, 2);
      }
      return {
        type: 'conflict_backup',
        externalRaw: conflict.externalRaw,
        externalHistory: conflict.externalHistory,
        localData: mem,
        combinedText: combined
      };
    }

    if (mem) {
      return {
        type: 'json',
        data: mem,
        json: JSON.stringify(mem, null, 2),
        combinedText: JSON.stringify(mem, null, 2)
      };
    }

    return null;
  }

  applyTemplate(day, templateData, currentInputs = {}) {
    const existing = this.inMemoryCopies.get(day) || {};
    const filled = { ...existing };
    let hasOverwriteConflict = false;

    for (const [key, tVal] of Object.entries(templateData)) {
      const cur = currentInputs[key];
      // If user has already typed a non-whitespace value, do not silently overwrite
      if (typeof cur === 'string' && cur.trim().length > 0) {
        filled[key] = cur;
        if (cur !== tVal && typeof tVal === 'string' && tVal.trim().length > 0) {
          hasOverwriteConflict = true;
        }
      } else {
        filled[key] = tVal;
      }
    }

    this.inMemoryCopies.set(day, filled);
    const saveResult = this.saveDraft(day, filled);

    let statusMsg = '已帶入範本預設值並儲存草稿；請依當日現地觀察填寫作業與地層。';
    if (hasOverwriteConflict) {
      statusMsg = '已填欄位保留原內容不覆寫，其餘欄位已帶入範本預設值。';
    }
    if (!saveResult.success) {
      statusMsg += '（注意：本機儲存未成功，已暫存於記憶體）';
    }
    if (!saveResult.success) statusMsg += ' ' + (this.getStatus(day)?.status || '');
    this.dayStatus.set(day, {
      ...this.getStatus(day),
      status: statusMsg,
      isError: !saveResult.success,
      isUnsaved: !saveResult.success
    });

    return { filled, hasOverwriteConflict, saveResult, statusMsg };
  }
}
