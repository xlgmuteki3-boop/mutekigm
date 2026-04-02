// ==UserScript==
// @name         飞书多维表格-图片链接转附件
// @namespace    http://tampermonkey.net/
// @version      1.5.0
// @description  在飞书多维表格页面批量将「图片」列 URL 下载并上传为「附件」列图片（支持 UI 配置与本地保存）
// @author       you
// @match        *://*.feishu.cn/base/*
// @match        *://*.feishu.cn/wiki/*
// @grant        GM_xmlhttpRequest
// @connect      open.feishu.cn
// @connect      *
// @run-at       document-idle
// ==/UserScript==

/* eslint-disable no-unused-vars -- 油猴脚本顶层配置与工具函数 */

// =============================================================================
// 本地存储与状态文案（状态文案仍固定在脚本内，不入库）
// =============================================================================

/** 旧版单对象配置键（迁移后不再写入） */
const FEISHU_B2A_LEGACY_STORAGE_KEY = 'feishu_b2a_config_v1';

/** 多表配置档案存储键 */
const STORAGE_KEY = 'feishu_b2a_profiles_v1';

/** @returns {string} */
function generateProfileId() {
  return `p_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

/** 成功时写入状态列的文案 */
const STATUS_SUCCESS_TEXT = '已转为附件';

/** 部分成功时状态列文案 */
const STATUS_PARTIAL_TEXT = '部分成功（部分 URL 失败）';

/** 全部 URL 失败时状态列文案（仍会尝试写失败原因） */
const STATUS_FAIL_TEXT = '处理失败';

// =============================================================================
// 配置：默认值 / 读写 / 校验 / UI
// =============================================================================

/** @returns {Record<string, any>} 运行配置对象 */
function getDefaultConfig() {
  return {
    appId: '',
    appSecret: '',
    appToken: '',
    tableId: '',
    viewId: '',
    urlFieldName: '图片',
    attachmentFieldName: '附件',
    statusFieldName: '图片处理状态',
    errorFieldName: '失败原因',
    skipIfAttachmentExists: true,
    appendToExistingAttachments: false,
    recordConcurrency: 1,
    urlConcurrency: 1,
    processIntervalMs: 300,
    requestTimeoutMs: 120000,
    maxImageBytes: 20971520,
  };
}

/**
 * 将任意输入规整为配置对象（与默认值合并并做类型修正）
 * @param {any} raw
 */
function normalizeStoredConfig(raw) {
  const d = getDefaultConfig();
  if (!raw || typeof raw !== 'object') return { ...d };

  const num = (v, fallback) => {
    const n = Number(v);
    return Number.isFinite(n) ? n : fallback;
  };
  const bool = (v, fallback) => {
    if (v === true || v === false) return v;
    if (v === 'true' || v === '1') return true;
    if (v === 'false' || v === '0') return false;
    return fallback;
  };
  const str = (v, fallback) => (v == null ? fallback : String(v));

  return {
    ...d,
    appId: str(raw.appId, d.appId).trim(),
    appSecret: str(raw.appSecret, d.appSecret).trim(),
    appToken: str(raw.appToken, d.appToken).trim(),
    tableId: str(raw.tableId, d.tableId).trim(),
    viewId: str(raw.viewId, d.viewId).trim(),
    urlFieldName: str(raw.urlFieldName, d.urlFieldName).trim() || d.urlFieldName,
    attachmentFieldName: str(raw.attachmentFieldName, d.attachmentFieldName).trim() || d.attachmentFieldName,
    statusFieldName: str(raw.statusFieldName, d.statusFieldName),
    errorFieldName: str(raw.errorFieldName, d.errorFieldName),
    skipIfAttachmentExists: bool(raw.skipIfAttachmentExists, d.skipIfAttachmentExists),
    appendToExistingAttachments: bool(raw.appendToExistingAttachments, d.appendToExistingAttachments),
    recordConcurrency: Math.max(1, Math.floor(num(raw.recordConcurrency, d.recordConcurrency))),
    urlConcurrency: Math.max(1, Math.floor(num(raw.urlConcurrency, d.urlConcurrency))),
    processIntervalMs: Math.max(0, Math.floor(num(raw.processIntervalMs, d.processIntervalMs))),
    requestTimeoutMs: Math.max(1000, Math.floor(num(raw.requestTimeoutMs, d.requestTimeoutMs))),
    maxImageBytes: Math.max(1024, Math.floor(num(raw.maxImageBytes, d.maxImageBytes))),
  };
}

/**
 * @typedef {ReturnType<typeof getDefaultConfig> & {
 *   id: string;
 *   name: string;
 *   fieldMapJson: string;
 *   fieldEnabledJson: string;
 *   matchMode: 'table' | 'table_view' | 'manual';
 *   createdAt: number;
 *   updatedAt: number;
 *   tableLabel: string;
 *   tableDisplayName: string;
 *   viewDisplayName: string;
 *   bindSummary: string;
 *   lastMatchedAt: number;
 *   lastUsedAt: number;
 *   pinned: boolean;
 * }} FeishuB2AProfile
 */

/** @returns {FeishuB2AProfile} */
function getDefaultProfile() {
  const now = Date.now();
  return {
    id: generateProfileId(),
    name: '默认档案',
    ...getDefaultConfig(),
    fieldMapJson: '',
    fieldEnabledJson: '',
    matchMode: 'table',
    createdAt: now,
    updatedAt: now,
    tableLabel: '',
    tableDisplayName: '',
    viewDisplayName: '',
    bindSummary: '',
    lastMatchedAt: 0,
    lastUsedAt: 0,
    pinned: false,
  };
}

/**
 * @param {any} raw
 * @returns {FeishuB2AProfile}
 */
function normalizeProfile(raw) {
  const d = getDefaultProfile();
  if (!raw || typeof raw !== 'object') return { ...d };

  const base = normalizeStoredConfig(raw);
  const str = (v, fallback) => (v == null ? fallback : String(v));
  const mode = str(raw.matchMode, d.matchMode);
  const matchMode =
    mode === 'table_view' || mode === 'manual' || mode === 'table' ? mode : 'table';

  const numOpt = (v) => {
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  };

  return {
    ...d,
    ...base,
    id: str(raw.id, d.id).trim() || d.id,
    name: str(raw.name, d.name).trim() || d.name,
    fieldMapJson: str(raw.fieldMapJson, d.fieldMapJson),
    fieldEnabledJson: str(raw.fieldEnabledJson, d.fieldEnabledJson),
    matchMode,
    createdAt: typeof raw.createdAt === 'number' && Number.isFinite(raw.createdAt) ? raw.createdAt : d.createdAt,
    updatedAt: typeof raw.updatedAt === 'number' && Number.isFinite(raw.updatedAt) ? raw.updatedAt : d.updatedAt,
    tableLabel: str(raw.tableLabel, d.tableLabel),
    tableDisplayName: str(raw.tableDisplayName, d.tableDisplayName),
    viewDisplayName: str(raw.viewDisplayName, d.viewDisplayName),
    bindSummary: str(raw.bindSummary, d.bindSummary),
    lastMatchedAt: numOpt(raw.lastMatchedAt),
    lastUsedAt: numOpt(raw.lastUsedAt),
    pinned: raw.pinned === true,
  };
}

/**
 * @typedef {{ profiles: FeishuB2AProfile[]; lastProfileId: string }} FeishuB2AProfileStore
 */

/**
 * @param {any} raw
 * @returns {FeishuB2AProfileStore}
 */
function normalizeProfileStore(raw) {
  const def = getDefaultProfile();
  const empty = { profiles: /** @type {FeishuB2AProfile[]} */ ([]), lastProfileId: '' };
  if (!raw || typeof raw !== 'object') {
    empty.profiles = [{ ...def }];
    empty.lastProfileId = empty.profiles[0].id;
    return empty;
  }
  const arr = Array.isArray(raw.profiles) ? raw.profiles.map((p) => normalizeProfile(p)) : [];
  if (!arr.length) {
    empty.profiles = [{ ...def }];
    empty.lastProfileId = empty.profiles[0].id;
    return empty;
  }
  const last = typeof raw.lastProfileId === 'string' ? raw.lastProfileId.trim() : '';
  const lastOk = last && arr.some((p) => p.id === last);
  return {
    profiles: arr,
    lastProfileId: lastOk ? last : arr[0].id,
  };
}

/**
 * 从旧版单对象迁移为一条档案
 * @returns {FeishuB2AProfileStore | null}
 */
function migrateLegacyIfNeeded() {
  try {
    const s = localStorage.getItem(FEISHU_B2A_LEGACY_STORAGE_KEY);
    if (!s) return null;
    const j = JSON.parse(s);
    if (j && typeof j === 'object' && Array.isArray(j.profiles)) return null;
    const cfg = normalizeStoredConfig(j);
    const p = normalizeProfile({
      ...cfg,
      id: generateProfileId(),
      name: '旧配置迁移',
      fieldMapJson: '',
      fieldEnabledJson: '',
      matchMode: 'table',
      createdAt: Date.now(),
      updatedAt: Date.now(),
    });
    localStorage.removeItem(FEISHU_B2A_LEGACY_STORAGE_KEY);
    return { profiles: [p], lastProfileId: p.id };
  } catch {
    return null;
  }
}

/**
 * @returns {FeishuB2AProfileStore}
 */
function loadProfileStore() {
  try {
    const s = localStorage.getItem(STORAGE_KEY);
    if (s) {
      const j = JSON.parse(s);
      return normalizeProfileStore(j);
    }
  } catch {
    /* fallthrough */
  }
  const migrated = migrateLegacyIfNeeded();
  if (migrated) {
    saveProfileStore(migrated);
    return migrated;
  }
  const fresh = normalizeProfileStore(null);
  saveProfileStore(fresh);
  return fresh;
}

/**
 * @param {FeishuB2AProfileStore} store
 */
function saveProfileStore(store) {
  const normalized = normalizeProfileStore(store);
  localStorage.setItem(STORAGE_KEY, JSON.stringify(normalized));
}

/**
 * @param {Partial<FeishuB2AProfile>} partial
 * @returns {FeishuB2AProfile}
 */
function createProfile(partial) {
  const store = loadProfileStore();
  const base = getDefaultProfile();
  const p = normalizeProfile({
    ...base,
    ...partial,
    id: generateProfileId(),
    name: partial && partial.name != null ? String(partial.name).trim() || '新档案' : '新档案',
    createdAt: Date.now(),
    updatedAt: Date.now(),
  });
  store.profiles.push(p);
  store.lastProfileId = p.id;
  saveProfileStore(store);
  return p;
}

/**
 * @param {string} profileId
 * @param {Partial<FeishuB2AProfile>} patch
 * @returns {FeishuB2AProfile | null}
 */
function updateProfile(profileId, patch) {
  const store = loadProfileStore();
  const idx = store.profiles.findIndex((p) => p.id === profileId);
  if (idx < 0) return null;
  const cur = store.profiles[idx];
  const next = normalizeProfile({
    ...cur,
    ...patch,
    id: cur.id,
    updatedAt: Date.now(),
  });
  store.profiles[idx] = next;
  saveProfileStore(store);
  return next;
}

/**
 * @param {string} profileId
 * @returns {boolean}
 */
function deleteProfile(profileId) {
  const store = loadProfileStore();
  if (store.profiles.length <= 1) return false;
  const next = store.profiles.filter((p) => p.id !== profileId);
  if (next.length === store.profiles.length) return false;
  store.profiles = next;
  if (store.lastProfileId === profileId) {
    store.lastProfileId = next[0].id;
  }
  saveProfileStore(store);
  return true;
}

/**
 * @param {string} profileId
 * @returns {FeishuB2AProfile | null}
 */
function duplicateProfile(profileId) {
  const p = getProfileById(profileId);
  if (!p) return null;
  return createProfile({
    ...p,
    name: `${p.name}（副本）`,
    matchMode: p.matchMode,
    fieldMapJson: p.fieldMapJson,
    fieldEnabledJson: p.fieldEnabledJson,
  });
}

/**
 * @param {string} profileId
 * @returns {FeishuB2AProfile | null}
 */
function getProfileById(profileId) {
  if (!profileId) return null;
  const store = loadProfileStore();
  return store.profiles.find((p) => p.id === profileId) || null;
}

/**
 * 档案与当前 URL 推断值是否匹配（manual 不参与）
 * @param {FeishuB2AProfile} p
 * @param {{ app_token: string | null; table_id: string | null; view_id: string | null }} inferred
 */
function profileMatchesInferred(p, inferred) {
  const mode = p.matchMode || 'table';
  if (mode === 'manual') return false;
  const a = inferred.app_token ? String(inferred.app_token).trim() : '';
  const t = inferred.table_id ? String(inferred.table_id).trim() : '';
  const v = inferred.view_id ? String(inferred.view_id).trim() : '';
  const pa = String(p.appToken || '').trim();
  const pt = String(p.tableId || '').trim();
  const pv = String(p.viewId || '').trim();
  if (!pa || !pt) return false;
  if (pa !== a || pt !== t) return false;
  if (mode === 'table_view') {
    return pv === v || (!pv && !v);
  }
  return mode === 'table';
}

/**
 * 按优先级解析当前页面对应的档案
 * @param {FeishuB2AProfileStore} [store]
 * @returns {{ profile: FeishuB2AProfile; reason: string }}
 */
function resolveMatchedProfileByCurrentUrl(store) {
  const st = store || loadProfileStore();
  const inferred = inferAppAndTableFromUrl();

  const tv = st.profiles.find((p) => (p.matchMode || 'table') === 'table_view' && profileMatchesInferred(p, inferred));
  if (tv) {
    return { profile: normalizeProfile(tv), reason: 'appToken+tableId+viewId' };
  }
  const tb = st.profiles.find((p) => (p.matchMode || 'table') === 'table' && profileMatchesInferred(p, inferred));
  if (tb) {
    return { profile: normalizeProfile(tb), reason: 'appToken+tableId' };
  }

  const last = getProfileById(st.lastProfileId);
  if (last) {
    return { profile: normalizeProfile(last), reason: 'lastProfileId' };
  }
  if (st.profiles.length) {
    return { profile: normalizeProfile(st.profiles[0]), reason: 'first' };
  }
  const np = createProfile({ name: '默认档案' });
  return { profile: np, reason: 'created' };
}

/**
 * 供任务/API 使用的「当前表」配置（来自 URL 解析命中的档案）
 * @returns {ReturnType<typeof getDefaultConfig>}
 */
function loadConfig() {
  const { profile } = resolveMatchedProfileByCurrentUrl();
  return normalizeStoredConfig(profile);
}

/**
 * 将表单配置写回指定档案（仅更新该条）
 * @param {string} profileId
 * @param {FeishuB2AProfile} profile
 */
function saveProfileForm(profileId, profile) {
  const withSummary = normalizeProfile({
    ...profile,
    id: profileId,
    bindSummary: computeBindSummary(profile),
  });
  updateProfile(profileId, withSummary);
  const store = loadProfileStore();
  store.lastProfileId = profileId;
  saveProfileStore(store);
}

/**
 * @param {string} [tid]
 * @returns {string}
 */
function shortTableId(tid) {
  const s = String(tid || '').trim();
  if (!s) return '';
  return s.length > 18 ? `${s.slice(0, 10)}…${s.slice(-4)}` : s;
}

/**
 * @param {FeishuB2AProfile} prof
 * @returns {string}
 */
function computeBindSummary(prof) {
  const p = normalizeProfile(prof);
  const label = (p.tableLabel || '').trim() || (p.tableDisplayName || '').trim();
  const base = label || (p.tableId || '').trim() || '未绑定表';
  const mode = p.matchMode || 'table';
  if (mode === 'manual') return `${base}｜手动档案`;
  const modeText = mode === 'table_view' ? '按表+视图' : '按表匹配';
  const vid = (p.viewId || '').trim();
  if (mode === 'table_view' && vid) return `${base}｜${modeText}｜${vid}`;
  return `${base}｜${modeText}`;
}

/**
 * @param {FeishuB2AProfile} p
 * @returns {string}
 */
function formatProfileOptionLine(p) {
  const pr = normalizeProfile(p);
  const mode =
    pr.matchMode === 'table_view' ? '按表+视图' : pr.matchMode === 'manual' ? 'manual' : '按表匹配';
  const tid = shortTableId(pr.tableId) || '—';
  const show = (pr.tableLabel || '').trim() || pr.name;
  return `${show} ｜ ${tid} ｜ ${mode}`;
}

/**
 * 尝试从页面 DOM 识别当前表格名、视图名（飞书改版频繁，失败时返回空）
 * @returns {{ tableName: string; viewName: string }}
 */
function inferTableDisplayNamesFromDom() {
  let tableName = '';
  let viewName = '';
  try {
    const tryText = (el) => {
      const t = el && el.textContent ? String(el.textContent).trim() : '';
      return t.length > 80 ? '' : t;
    };
    const s1 = document.querySelector(
      '[class*="sheet-title"] span, [class*="base-title"] span, [class*="bitable"] [class*="title"] span:first-child, [class*="toolbar"] [class*="title"]',
    );
    tableName = tryText(s1);
    const s2 = document.querySelector(
      '[class*="view-tab"] [class*="active"], [class*="view-bar"] [class*="selected"], [class*="tabItem"][class*="active"]',
    );
    viewName = tryText(s2);
    if (!tableName && document.title) {
      const t = document.title
        .replace(/\s*[·•]\s*飞书多维表格.*$/i, '')
        .replace(/\s*-\s*飞书.*$/i, '')
        .replace(/\s*—\s*飞书.*$/i, '')
        .trim();
      if (t && t.length < 100) tableName = t;
    }
  } catch {
    /* ignore */
  }
  return { tableName: tableName || '', viewName: viewName || '' };
}

/**
 * @param {FeishuB2AProfile} edit
 * @param {{ app_token: string | null; table_id: string | null; view_id: string | null }} inferred
 * @returns {{ code: string; color: 'green' | 'yellow' | 'red' | 'gray'; message: string }}
 */
function getBindStatusForEdit(edit, inferred) {
  const mode = edit.matchMode || 'table';
  if (mode === 'manual') {
    return { code: 'manual', color: 'gray', message: 'ℹ️ 当前档案为 manual 模式（不按 URL 自动匹配）' };
  }

  const pa = String(edit.appToken || '').trim();
  const pt = String(edit.tableId || '').trim();
  const pv = String(edit.viewId || '').trim();
  const ia = inferred.app_token ? String(inferred.app_token).trim() : '';
  const it = inferred.table_id ? String(inferred.table_id).trim() : '';
  const iv = inferred.view_id ? String(inferred.view_id).trim() : '';

  if (!pt) {
    return { code: 'unbound', color: 'red', message: '❌ 当前档案未绑定 tableId（请填写或从当前页带入）' };
  }

  const appMismatch = ia && pa && pa !== ia;
  const tableMismatch = it && pt && pt !== it;
  const viewMismatch = mode === 'table_view' && iv && pv && pv !== iv;

  if (profileMatchesInferred(edit, inferred)) {
    return { code: 'ok', color: 'green', message: '✅ 当前档案已匹配本页 URL（appToken / tableId / viewId 一致）' };
  }

  if (appMismatch) {
    return { code: 'app', color: 'red', message: '⚠️ appToken 与当前页面不一致' };
  }
  if (tableMismatch) {
    return { code: 'table', color: 'yellow', message: '⚠️ tableId 与当前页面不一致' };
  }
  if (viewMismatch) {
    return { code: 'view', color: 'yellow', message: '⚠️ viewId 与当前页面不一致（按表+视图匹配时）' };
  }

  if (!ia || !it) {
    return { code: 'partial', color: 'yellow', message: '⚠️ 当前页面未能识别完整 URL 参数，请核对档案内绑定' };
  }

  return { code: 'mismatch', color: 'red', message: '❌ 当前档案与页面 URL 不一致' };
}

/**
 * @param {string} reason
 * @returns {boolean}
 */
function isUrlAutoMatchReason(reason) {
  return reason === 'appToken+tableId+viewId' || reason === 'appToken+tableId';
}

/**
 * @param {string} profileId
 * @param {string} reason
 */
function touchLastMatchedIfUrlMatch(profileId, reason) {
  if (!profileId || !isUrlAutoMatchReason(reason)) return;
  updateProfile(profileId, { lastMatchedAt: Date.now() });
}

/**
 * @param {string} profileId
 */
function touchLastUsed(profileId) {
  if (!profileId) return;
  updateProfile(profileId, { lastUsedAt: Date.now() });
}

/**
 * 执行任务：同时更新最近使用；若 URL 自动命中则更新最近匹配时间（单次写入）
 * @param {string} profileId
 * @param {string} reason
 */
function touchProfileRunUsage(profileId, reason) {
  if (!profileId) return;
  const patch = { lastUsedAt: Date.now() };
  if (isUrlAutoMatchReason(reason)) patch.lastMatchedAt = Date.now();
  updateProfile(profileId, patch);
}

/**
 * 保存前校验（必填 APP_ID / APP_SECRET）
 * @param {ReturnType<typeof getDefaultConfig>} config
 * @returns {{ valid: boolean; errors: string[] }}
 */
function validateConfig(config) {
  const errors = [];
  const c = normalizeStoredConfig(config);
  if (!String(c.appId || '').trim()) errors.push('APP_ID 必填');
  if (!String(c.appSecret || '').trim()) errors.push('APP_SECRET 必填');
  if (c.recordConcurrency < 1) errors.push('记录并发数须 ≥ 1');
  if (c.urlConcurrency < 1) errors.push('URL 并发数须 ≥ 1');
  if (c.requestTimeoutMs < 1000) errors.push('请求超时须 ≥ 1000 毫秒');
  if (c.maxImageBytes < 1024) errors.push('图片最大体积过小');
  return { valid: errors.length === 0, errors };
}

// =============================================================================
// 内部状态
// =============================================================================

let __stopRequested = false;
/** 是否有主流程正在执行（用于防重复开始、迷你条与根级停止按钮） */
let __taskRunning = false;
let __uiEnsured = false;

/** @type {HTMLElement | null} */
let __configOverlay = null;
/** @type {HTMLElement | null} */
let __configBanner = null;
/** @type {HTMLElement | null} */
let __configErrorBox = null;

/** 配置弹窗当前正在编辑的档案 id */
let __editingProfileId = '';

/** @type {HTMLSelectElement | null} */
let __profileSelectEl = null;

/** @type {HTMLElement | null} */
let __urlCtxEl = null;

/** @type {HTMLElement | null} */
let __bindCheckPanelEl = null;

/** @type {HTMLElement | null} */
let __bindStatusEl = null;

/** @type {HTMLInputElement | null} */
let __profileSearchInput = null;

/** @type {HTMLInputElement | null} */
let __filterMatchOnlyEl = null;

/** @type {HTMLInputElement | null} */
let __filterBoundOnlyEl = null;

/** @type {HTMLInputElement | null} */
let __filterManualOnlyEl = null;

/** @type {HTMLElement | null} */
let __profilePickerEl = null;

/** @type {HTMLButtonElement | null} */
let __btnProfilePickerToggle = null;

// =============================================================================
// 工具：日志 / UI
// =============================================================================

/** @type {HTMLTextAreaElement | null} */
let __logEl = null;
/** @type {HTMLElement | null} */
let __panelEl = null;
/** @type {HTMLButtonElement | null} */
let __btnStart = null;
/** @type {HTMLButtonElement | null} */
let __btnStop = null;
/** @type {HTMLButtonElement | null} */
let __btnRootStop = null;
/** @type {HTMLElement | null} */
let __runMiniEl = null;
/** @type {HTMLElement | null} */
let __statEl = null;

/**
 * 追加一行日志
 * @param {string} message
 * @param {'info'|'warn'|'err'} [level]
 */
function log(message, level = 'info') {
  const ts = new Date().toISOString().replace('T', ' ').replace(/\.\d+Z$/, '');
  const line = `[${ts}] ${level === 'err' ? '[错误] ' : level === 'warn' ? '[警告] ' : ''}${message}`;
  // eslint-disable-next-line no-console
  console.log(line);
  if (__logEl) {
    __logEl.value += `${line}\n`;
    __logEl.scrollTop = __logEl.scrollHeight;
  }
}

/**
 * 从配置弹窗表单收集对象
 * @returns {ReturnType<typeof getDefaultConfig>}
 */
function collectConfigFromForm() {
  if (!__configOverlay) return getDefaultConfig();
  const $ = (id) => /** @type {HTMLInputElement | null} */ (__configOverlay.querySelector(`#${id}`));
  return normalizeStoredConfig({
    appId: $('fs-b2a-f-app-id')?.value,
    appSecret: $('fs-b2a-f-app-secret')?.value,
    appToken: $('fs-b2a-f-app-token')?.value,
    tableId: $('fs-b2a-f-table-id')?.value,
    viewId: $('fs-b2a-f-view-id')?.value,
    urlFieldName: $('fs-b2a-f-url-field')?.value,
    attachmentFieldName: $('fs-b2a-f-attach-field')?.value,
    statusFieldName: $('fs-b2a-f-status-field')?.value,
    errorFieldName: $('fs-b2a-f-error-field')?.value,
    skipIfAttachmentExists: $('fs-b2a-f-skip-attach')?.checked,
    appendToExistingAttachments: $('fs-b2a-f-append')?.checked,
    recordConcurrency: $('fs-b2a-f-rec-conc')?.value,
    urlConcurrency: $('fs-b2a-f-url-conc')?.value,
    processIntervalMs: $('fs-b2a-f-interval')?.value,
    requestTimeoutMs: $('fs-b2a-f-timeout')?.value,
    maxImageBytes: $('fs-b2a-f-max-bytes')?.value,
  });
}

/**
 * 将配置写入表单
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 */
function fillConfigForm(cfg) {
  if (!__configOverlay) return;
  const c = normalizeStoredConfig(cfg);
  const set = (id, v) => {
    const el = /** @type {HTMLInputElement | null} */ (__configOverlay.querySelector(`#${id}`));
    if (el) el.value = v == null ? '' : String(v);
  };
  set('fs-b2a-f-app-id', c.appId);
  set('fs-b2a-f-app-secret', c.appSecret);
  set('fs-b2a-f-app-token', c.appToken);
  set('fs-b2a-f-table-id', c.tableId);
  set('fs-b2a-f-view-id', c.viewId);
  set('fs-b2a-f-url-field', c.urlFieldName);
  set('fs-b2a-f-attach-field', c.attachmentFieldName);
  set('fs-b2a-f-status-field', c.statusFieldName);
  set('fs-b2a-f-error-field', c.errorFieldName);
  const skip = /** @type {HTMLInputElement | null} */ (__configOverlay.querySelector('#fs-b2a-f-skip-attach'));
  if (skip) skip.checked = !!c.skipIfAttachmentExists;
  const app = /** @type {HTMLInputElement | null} */ (__configOverlay.querySelector('#fs-b2a-f-append'));
  if (app) app.checked = !!c.appendToExistingAttachments;
  set('fs-b2a-f-rec-conc', c.recordConcurrency);
  set('fs-b2a-f-url-conc', c.urlConcurrency);
  set('fs-b2a-f-interval', c.processIntervalMs);
  set('fs-b2a-f-timeout', c.requestTimeoutMs);
  set('fs-b2a-f-max-bytes', c.maxImageBytes);
}

/**
 * @param {string} s
 */
function escHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * @param {FeishuB2AProfile} profile
 */
function fillProfileForm(profile) {
  if (!__configOverlay) return;
  const p = normalizeProfile(profile);
  __editingProfileId = p.id;
  fillConfigForm(p);
  const mm = /** @type {HTMLSelectElement | null} */ (__configOverlay.querySelector('#fs-b2a-f-match-mode'));
  if (mm) mm.value = p.matchMode || 'table';
  const fm = /** @type {HTMLTextAreaElement | null} */ (__configOverlay.querySelector('#fs-b2a-f-field-map-json'));
  if (fm) fm.value = p.fieldMapJson || '';
  const fe = /** @type {HTMLTextAreaElement | null} */ (__configOverlay.querySelector('#fs-b2a-f-field-enabled-json'));
  if (fe) fe.value = p.fieldEnabledJson || '';
  const tl = /** @type {HTMLInputElement | null} */ (__configOverlay.querySelector('#fs-b2a-f-table-label'));
  if (tl) tl.value = p.tableLabel || '';
  const pin = /** @type {HTMLInputElement | null} */ (__configOverlay.querySelector('#fs-b2a-f-pinned'));
  if (pin) pin.checked = !!p.pinned;
  const dom = inferTableDisplayNamesFromDom();
  const tdn = /** @type {HTMLInputElement | null} */ (__configOverlay.querySelector('#fs-b2a-f-table-display-name'));
  const vdn = /** @type {HTMLInputElement | null} */ (__configOverlay.querySelector('#fs-b2a-f-view-display-name'));
  if (tdn) tdn.value = (p.tableDisplayName || '').trim() || dom.tableName || '';
  if (vdn) vdn.value = (p.viewDisplayName || '').trim() || dom.viewName || '';
  refreshProfileSelectOptions();
  if (__profileSelectEl) __profileSelectEl.value = p.id;
  updateBindCheckPanel();
  refreshProfilePickerList();
}

/**
 * @returns {FeishuB2AProfile}
 */
function collectProfileFromForm() {
  const base = collectConfigFromForm();
  const overlay = __configOverlay;
  if (!overlay) return normalizeProfile(getDefaultProfile());
  const mm = /** @type {HTMLSelectElement | null} */ (overlay.querySelector('#fs-b2a-f-match-mode'));
  const mv = (mm && mm.value) || 'table';
  const matchMode = mv === 'table_view' || mv === 'manual' || mv === 'table' ? mv : 'table';
  const fm = /** @type {HTMLTextAreaElement | null} */ (overlay.querySelector('#fs-b2a-f-field-map-json'));
  const fe = /** @type {HTMLTextAreaElement | null} */ (overlay.querySelector('#fs-b2a-f-field-enabled-json'));
  const tl = /** @type {HTMLInputElement | null} */ (overlay.querySelector('#fs-b2a-f-table-label'));
  const pin = /** @type {HTMLInputElement | null} */ (overlay.querySelector('#fs-b2a-f-pinned'));
  const tdn = /** @type {HTMLInputElement | null} */ (overlay.querySelector('#fs-b2a-f-table-display-name'));
  const vdn = /** @type {HTMLInputElement | null} */ (overlay.querySelector('#fs-b2a-f-view-display-name'));
  const existing = getProfileById(__editingProfileId);
  return normalizeProfile({
    ...(existing || getDefaultProfile()),
    ...base,
    id: __editingProfileId || (existing && existing.id) || generateProfileId(),
    name: existing ? existing.name : '默认档案',
    matchMode,
    fieldMapJson: fm ? fm.value : '',
    fieldEnabledJson: fe ? fe.value : '',
    tableLabel: tl ? tl.value : '',
    tableDisplayName: tdn ? tdn.value : '',
    viewDisplayName: vdn ? vdn.value : '',
    pinned: pin ? pin.checked : false,
  });
}

/**
 * @returns {FeishuB2AProfile[]}
 */
function getFilteredProfiles() {
  const store = loadProfileStore();
  const inferred = inferAppAndTableFromUrl();
  const q = (__profileSearchInput && __profileSearchInput.value.trim().toLowerCase()) || '';
  const matchOnly = __filterMatchOnlyEl && __filterMatchOnlyEl.checked;
  const boundOnly = __filterBoundOnlyEl && __filterBoundOnlyEl.checked;
  const manualOnly = __filterManualOnlyEl && __filterManualOnlyEl.checked;

  const searchHit = (p) => {
    if (!q) return true;
    const x = normalizeProfile(p);
    const hay = [x.name, x.tableLabel, x.tableId, x.viewId, x.bindSummary, x.tableDisplayName, x.viewDisplayName]
      .join('\n')
      .toLowerCase();
    return hay.includes(q);
  };

  let list = store.profiles.filter((p) => {
    if (!searchHit(p)) return false;
    const x = normalizeProfile(p);
    const mode = x.matchMode || 'table';
    if (matchOnly) {
      if (mode === 'manual') return false;
      if (!profileMatchesInferred(x, inferred)) return false;
    }
    if (boundOnly && !String(x.tableId || '').trim()) return false;
    if (manualOnly && mode !== 'manual') return false;
    return true;
  });

  list.sort((a, b) => {
    const pa = normalizeProfile(a);
    const pb = normalizeProfile(b);
    if (pa.pinned !== pb.pinned) return pa.pinned ? -1 : 1;
    return (pb.lastUsedAt || 0) - (pa.lastUsedAt || 0);
  });

  if (__editingProfileId && !list.some((x) => x.id === __editingProfileId)) {
    const cur = getProfileById(__editingProfileId);
    if (cur) list.unshift(normalizeProfile(cur));
  }
  return list;
}

function refreshProfileSelectOptions() {
  if (!__profileSelectEl) return;
  const list = getFilteredProfiles();
  const sel = __profileSelectEl;
  sel.innerHTML = '';
  if (!list.length) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = '（无匹配结果，请调整搜索或筛选）';
    opt.disabled = true;
    sel.appendChild(opt);
    return;
  }
  for (const p of list) {
    const opt = document.createElement('option');
    opt.value = p.id;
    opt.textContent = formatProfileOptionLine(p);
    sel.appendChild(opt);
  }
  if (__editingProfileId && list.some((x) => x.id === __editingProfileId)) {
    sel.value = __editingProfileId;
  } else if (list.length) {
    sel.value = list[0].id;
  }
}

function refreshProfilePickerList() {
  if (!__profilePickerEl) return;
  const list = getFilteredProfiles();
  __profilePickerEl.innerHTML = '';
  for (const p of list) {
    const pr = normalizeProfile(p);
    const row = document.createElement('button');
    row.type = 'button';
    row.className = 'fs-b2a-picker-row';
    const mode =
      pr.matchMode === 'table_view' ? '按表+视图' : pr.matchMode === 'manual' ? '手动' : '按表匹配';
    const st = getBindStatusForEdit(pr, inferAppAndTableFromUrl());
    const badgeColor =
      st.color === 'green' ? '#e8f5e9' : st.color === 'yellow' ? '#fff8e1' : st.color === 'red' ? '#ffebee' : '#f5f5f5';
    const badgeShort =
      {
        ok: '已匹配',
        manual: '手动',
        unbound: '未绑定',
        view: '视图',
        table: '表',
        app: 'APP',
        mismatch: '不符',
        partial: '部分',
        none: '—',
      }[st.code] || '—';
    const title = (pr.tableLabel || '').trim() || pr.name;
    const sub = (pr.bindSummary || '').trim() || formatProfileOptionLine(pr);
    const lu = pr.lastUsedAt ? new Date(pr.lastUsedAt).toLocaleString() : '—';
    row.innerHTML = `<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px;text-align:left;">
<div style="flex:1;min-width:0;">
<div style="font-weight:600;color:#1f2329;font-size:13px;">${escHtml(title)}</div>
<div style="font-size:11px;color:#646a73;margin-top:2px;">${escHtml(sub)}</div>
<div style="font-size:11px;color:#8f959e;margin-top:4px;">${escHtml(pr.tableId || '（无 tableId）')} · ${escHtml(mode)}</div>
</div>
<span style="flex-shrink:0;font-size:11px;padding:2px 6px;border-radius:4px;background:${badgeColor};color:#1f2329;">${escHtml(badgeShort)}</span>
</div><div style="font-size:10px;color:#a8abb2;margin-top:6px;">最近使用：${escHtml(lu)}</div>`;
    row.addEventListener('click', () => {
      if (__profileSelectEl) __profileSelectEl.value = pr.id;
      const picked = getProfileById(pr.id);
      if (picked) {
        __editingProfileId = pr.id;
        fillProfileForm(picked);
      }
      if (__profilePickerEl) __profilePickerEl.style.display = 'none';
      if (__btnProfilePickerToggle) __btnProfilePickerToggle.setAttribute('aria-expanded', 'false');
    });
    __profilePickerEl.appendChild(row);
  }
  if (!list.length) {
    const empty = document.createElement('div');
    empty.className = 'fs-b2a-picker-empty';
    empty.textContent = '无匹配档案，请调整搜索或筛选';
    __profilePickerEl.appendChild(empty);
  }
}

function updateBindCheckPanel() {
  if (!__bindCheckPanelEl || !__bindStatusEl) return;
  const inferred = inferAppAndTableFromUrl();
  const dom = inferTableDisplayNamesFromDom();
  const { profile: autoProf, reason } = resolveMatchedProfileByCurrentUrl();
  const edit = getProfileById(__editingProfileId);
  const editNorm = edit ? normalizeProfile(edit) : null;

  const pageTableName = (dom.tableName || '').trim() || '（未能从页面识别）';
  const pageViewName = (dom.viewName || '').trim() || '（未能从页面识别）';

  const left = `
<div class="fs-b2a-bind-col">
<div class="fs-b2a-bind-h">当前页面识别</div>
<div class="fs-b2a-bind-line"><span class="k">APP_TOKEN</span><code>${escHtml(inferred.app_token || '（无）')}</code></div>
<div class="fs-b2a-bind-line"><span class="k">TABLE_ID</span><code>${escHtml(inferred.table_id || '（无）')}</code></div>
<div class="fs-b2a-bind-line"><span class="k">VIEW_ID</span><code>${escHtml(inferred.view_id || '（无）')}</code></div>
<div class="fs-b2a-bind-line"><span class="k">表格名</span><span class="v">${escHtml(pageTableName)}</span></div>
<div class="fs-b2a-bind-line"><span class="k">视图名</span><span class="v">${escHtml(pageViewName)}</span></div>
</div>`;

  const reasonText =
    {
      'appToken+tableId+viewId': '按表+视图',
      'appToken+tableId': '按表',
      lastProfileId: '上次使用',
      first: '首条档案',
      created: '新建',
    }[reason] || reason;

  const right = editNorm
    ? `
<div class="fs-b2a-bind-col">
<div class="fs-b2a-bind-h">当前编辑档案</div>
<div class="fs-b2a-bind-line"><span class="k">档案名</span><span class="v">${escHtml(editNorm.name)}</span></div>
<div class="fs-b2a-bind-line"><span class="k">表格备注</span><span class="v">${escHtml((editNorm.tableLabel || '').trim() || '（空）')}</span></div>
<div class="fs-b2a-bind-line"><span class="k">APP_TOKEN</span><code>${escHtml(editNorm.appToken || '（空）')}</code></div>
<div class="fs-b2a-bind-line"><span class="k">TABLE_ID</span><code>${escHtml(editNorm.tableId || '（空）')}</code></div>
<div class="fs-b2a-bind-line"><span class="k">VIEW_ID</span><code>${escHtml(editNorm.viewId || '（空）')}</code></div>
<div class="fs-b2a-bind-line"><span class="k">匹配模式</span><span class="v">${escHtml(editNorm.matchMode || 'table')}</span></div>
<div class="fs-b2a-bind-line"><span class="k">归属摘要</span><span class="v sm">${escHtml((editNorm.bindSummary || computeBindSummary(editNorm)).slice(0, 120))}</span></div>
<div class="fs-b2a-bind-line"><span class="k">最近使用</span><span class="v">${escHtml(editNorm.lastUsedAt ? new Date(editNorm.lastUsedAt).toLocaleString() : '—')}</span></div>
<div class="fs-b2a-bind-line"><span class="k">最近匹配</span><span class="v">${escHtml(editNorm.lastMatchedAt ? new Date(editNorm.lastMatchedAt).toLocaleString() : '—')}</span></div>
</div>`
    : `<div class="fs-b2a-bind-col"><div class="fs-b2a-bind-h">当前编辑档案</div><div class="fs-b2a-bind-line">（无）</div></div>`;

  __bindCheckPanelEl.innerHTML = `<div class="fs-b2a-bind-grid">${left}${right}</div>`;

  const st = editNorm
    ? getBindStatusForEdit(editNorm, inferred)
    : { color: 'gray', message: '（无编辑档案）', code: 'none' };
  const bg =
    st.color === 'green' ? '#e8f5e9' : st.color === 'yellow' ? '#fff8e1' : st.color === 'red' ? '#ffebee' : '#f0f0f0';
  const border =
    st.color === 'green' ? '#a5d6a7' : st.color === 'yellow' ? '#ffe082' : st.color === 'red' ? '#ffcdd2' : '#e0e0e0';

  __bindStatusEl.style.background = bg;
  __bindStatusEl.style.borderColor = border;
  __bindStatusEl.innerHTML = `<div style="font-size:13px;font-weight:600;color:#1f2329;">${escHtml(st.message)}</div>
<div style="font-size:11px;color:#646a73;margin-top:6px;">执行任务默认档案：<strong>${escHtml(autoProf.name)}</strong>（${escHtml(reasonText)}） · bindSummary：<span style="word-break:break-all;">${escHtml((autoProf.bindSummary || computeBindSummary(autoProf)).slice(0, 160))}</span></div>`;

  if (__urlCtxEl) {
    __urlCtxEl.innerHTML = '';
    __urlCtxEl.style.display = 'none';
  }
}

function updateUrlContextHtml() {
  updateBindCheckPanel();
}

/**
 * 将 URL 识别到的 appToken / tableId / viewId 写入当前编辑档案表单（不自动保存）
 */
function applyInferToCurrentProfileForm() {
  if (!__configOverlay) return;
  const inferred = inferAppAndTableFromUrl();
  const dom = inferTableDisplayNamesFromDom();
  const set = (id, v) => {
    const el = /** @type {HTMLInputElement | null} */ (__configOverlay.querySelector(`#${id}`));
    if (el) el.value = v == null ? '' : String(v);
  };
  if (inferred.app_token) set('fs-b2a-f-app-token', inferred.app_token);
  if (inferred.table_id) set('fs-b2a-f-table-id', inferred.table_id);
  if (inferred.view_id) set('fs-b2a-f-view-id', inferred.view_id);
  if (dom.tableName) set('fs-b2a-f-table-display-name', dom.tableName);
  if (dom.viewName) set('fs-b2a-f-view-display-name', dom.viewName);
  updateBindCheckPanel();
  log('已从当前页面填入 URL 与识别到的表名/视图名（保存后生效）');
}

/**
 * 从当前页自动设置表格备注名（tableLabel）
 */
function autoNameTableLabelFromPage() {
  if (!__configOverlay) return;
  const dom = inferTableDisplayNamesFromDom();
  const inferred = inferAppAndTableFromUrl();
  let label = '';
  if (dom.tableName && dom.viewName) label = `${dom.tableName} - ${dom.viewName}`;
  else if (dom.tableName) label = dom.tableName;
  else if (inferred.table_id) label = inferred.table_id;
  const tl = /** @type {HTMLInputElement | null} */ (__configOverlay.querySelector('#fs-b2a-f-table-label'));
  if (tl) tl.value = label;
  updateBindCheckPanel();
  log(`表格备注名已填入：${label || '（空）'}`);
}

/**
 * 将当前页 appToken/tableId/viewId 绑定到当前档案并立即保存
 */
function bindCurrentPageToProfileAndSave() {
  if (!__configOverlay || !__editingProfileId) return;
  const draft = collectProfileFromForm();
  const inferred = inferAppAndTableFromUrl();
  const dom = inferTableDisplayNamesFromDom();
  const merged = normalizeProfile({
    ...draft,
    appToken: inferred.app_token ? String(inferred.app_token) : draft.appToken,
    tableId: inferred.table_id ? String(inferred.table_id) : draft.tableId,
    viewId: inferred.view_id ? String(inferred.view_id) : draft.viewId,
    tableDisplayName: (draft.tableDisplayName || '').trim() || dom.tableName || draft.tableDisplayName,
    viewDisplayName: (draft.viewDisplayName || '').trim() || dom.viewName || draft.viewDisplayName,
  });
  const v = validateConfig(merged);
  if (!v.valid) {
    if (__configErrorBox) {
      __configErrorBox.style.display = 'block';
      __configErrorBox.textContent = v.errors.join('；');
    }
    log(v.errors.join('；'), 'warn');
    return;
  }
  if (__configErrorBox) __configErrorBox.style.display = 'none';
  saveProfileForm(__editingProfileId, merged);
  fillProfileForm(getProfileById(__editingProfileId) || merged);
  log(`已将当前页绑定到档案「${merged.name}」并保存`);
}

function toggleProfilePicker() {
  if (!__profilePickerEl || !__btnProfilePickerToggle) return;
  const open = __profilePickerEl.style.display !== 'block';
  __profilePickerEl.style.display = open ? 'block' : 'none';
  __btnProfilePickerToggle.setAttribute('aria-expanded', open ? 'true' : 'false');
  if (open) {
    refreshProfilePickerList();
    const close = (e) => {
      const t = /** @type {Node | null} */ (e.target);
      if (!t) return;
      if (__profilePickerEl && __profilePickerEl.contains(t)) return;
      if (__btnProfilePickerToggle && __btnProfilePickerToggle.contains(t)) return;
      __profilePickerEl.style.display = 'none';
      __btnProfilePickerToggle.setAttribute('aria-expanded', 'false');
      document.removeEventListener('click', close, true);
    };
    setTimeout(() => document.addEventListener('click', close, true), 0);
  }
}

/**
 * 构建配置弹窗 DOM（仅一次）
 */
function renderConfigModal() {
  if (__configOverlay) return;

  const style = document.createElement('style');
  style.textContent += `
    #fs-b2a-config-overlay { position: fixed; right: 16px; bottom: 120px; z-index: 2147483647; width: min(520px, calc(100vw - 32px)); max-height: min(78vh, 720px); background: #fff; border: 1px solid #e5e6eb; border-radius: 10px; box-shadow: 0 8px 28px rgba(0,0,0,.14); display: none; flex-direction: column; overflow: hidden; font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif; }
    #fs-b2a-config-overlay header { flex-shrink: 0; display: flex; align-items: center; justify-content: space-between; padding: 10px 12px; border-bottom: 1px solid #e5e6eb; font-size: 14px; font-weight: 600; color: #1f2329; }
    #fs-b2a-config-overlay header button { border: none; background: transparent; cursor: pointer; font-size: 18px; line-height: 1; color: #646a73; padding: 4px 8px; }
    #fs-b2a-config-banner { display: none; flex-shrink: 0; padding: 8px 12px; font-size: 13px; color: #d93026; background: #fdecea; border-bottom: 1px solid #f5c6cb; }
    #fs-b2a-config-err { display: none; flex-shrink: 0; padding: 8px 12px; font-size: 12px; color: #d93026; background: #fff7f7; border-bottom: 1px solid #f0f0f0; }
    #fs-b2a-config-scroll { flex: 1; overflow-y: auto; padding: 10px 12px 12px; }
    .fs-b2a-cfg-section { margin-bottom: 14px; }
    .fs-b2a-cfg-section h3 { margin: 0 0 8px; font-size: 13px; font-weight: 600; color: #1f2329; }
    .fs-b2a-cfg-row { margin-bottom: 10px; }
    .fs-b2a-cfg-row label { display: block; font-size: 13px; color: #646a73; margin-bottom: 4px; }
    .fs-b2a-cfg-row input[type="text"],
    .fs-b2a-cfg-row input[type="password"],
    .fs-b2a-cfg-row input[type="number"] { width: 100%; box-sizing: border-box; padding: 8px 10px; font-size: 14px; border: 1px solid #d0d3d6; border-radius: 6px; color: #1f2329; }
    .fs-b2a-cfg-row input:focus { outline: none; border-color: #3366ff; }
    .fs-b2a-cfg-check { display: flex; align-items: center; gap: 8px; font-size: 14px; color: #1f2329; cursor: pointer; user-select: none; }
    .fs-b2a-cfg-check input { width: 16px; height: 16px; cursor: pointer; }
    #fs-b2a-config-footer { flex-shrink: 0; display: flex; flex-wrap: wrap; gap: 8px; padding: 10px 12px; border-top: 1px solid #e5e6eb; background: #fafbfc; }
    #fs-b2a-config-footer button { flex: 1; min-width: 100px; padding: 9px 10px; border-radius: 6px; border: 1px solid #d0d3d6; background: #f5f6f7; cursor: pointer; font-size: 13px; color: #1f2329; }
    #fs-b2a-config-footer button.fs-b2a-primary { background: #3366ff; color: #fff; border-color: #3366ff; }
    #fs-b2a-url-ctx { display: none; }
    .fs-b2a-profile-search-row { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin-bottom: 10px; }
    .fs-b2a-profile-search-row input[type="text"] { flex: 1; min-width: 140px; padding: 6px 8px; font-size: 12px; border: 1px solid #d0d3d6; border-radius: 6px; box-sizing: border-box; }
    .fs-b2a-filter-chk { display: flex; align-items: center; gap: 4px; font-size: 11px; color: #646a73; user-select: none; cursor: pointer; }
    .fs-b2a-filter-chk input { cursor: pointer; }
    .fs-b2a-select-row { display: flex; gap: 6px; align-items: stretch; }
    #fs-b2a-profile-select { flex: 1; min-width: 0; box-sizing: border-box; padding: 8px 10px; font-size: 12px; border: 1px solid #d0d3d6; border-radius: 6px; color: #1f2329; background: #fff; }
    .fs-b2a-profile-picker-wrap { position: relative; margin-bottom: 10px; }
    #fs-b2a-profile-picker { display: none; position: absolute; left: 0; right: 0; top: 100%; margin-top: 4px; max-height: 280px; overflow-y: auto; z-index: 20; background: #fff; border: 1px solid #e5e6eb; border-radius: 8px; box-shadow: 0 8px 24px rgba(0,0,0,.12); padding: 6px; }
    .fs-b2a-picker-row { width: 100%; border: 1px solid #f0f0f0; border-radius: 8px; padding: 8px; margin-bottom: 4px; background: #fafbfc; cursor: pointer; font: inherit; }
    .fs-b2a-picker-row:hover { border-color: #3366ff; background: #f5f8ff; }
    .fs-b2a-picker-empty { padding: 12px; text-align: center; font-size: 12px; color: #8f959e; }
    .fs-b2a-bind-wrap { margin-bottom: 12px; }
    .fs-b2a-bind-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
    @media (max-width: 520px) { .fs-b2a-bind-grid { grid-template-columns: 1fr; } }
    .fs-b2a-bind-col { border: 1px solid #e8eaed; border-radius: 8px; padding: 8px; background: #fafbfc; }
    .fs-b2a-bind-h { font-weight: 600; font-size: 12px; margin-bottom: 6px; color: #1f2329; }
    .fs-b2a-bind-line { font-size: 11px; margin-bottom: 4px; line-height: 1.45; word-break: break-word; }
    .fs-b2a-bind-line .k { color: #8f959e; font-weight: 500; margin-right: 6px; }
    .fs-b2a-bind-line code { font-size: 10px; background: #fff; padding: 1px 4px; border-radius: 4px; border: 1px solid #e8eaed; }
    .fs-b2a-bind-line .v { color: #1f2329; }
    .fs-b2a-bind-line .v.sm { font-size: 10px; color: #646a73; }
    #fs-b2a-bind-status { margin-top: 8px; padding: 10px 12px; border-radius: 8px; border: 1px solid #e0e0e0; }
    .fs-b2a-profile-btns { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 10px; }
    .fs-b2a-profile-btns button { padding: 6px 8px; font-size: 12px; border: 1px solid #d0d3d6; border-radius: 6px; background: #fff; cursor: pointer; color: #1f2329; }
    .fs-b2a-profile-btns button:hover { border-color: #3366ff; color: #3366ff; }
  `;
  document.documentElement.appendChild(style);

  const overlay = document.createElement('div');
  overlay.id = 'fs-b2a-config-overlay';

  const header = document.createElement('header');
  const title = document.createElement('span');
  title.textContent = '配置';
  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.textContent = '×';
  closeBtn.title = '关闭';
  closeBtn.addEventListener('click', () => closeConfigModal());
  header.appendChild(title);
  header.appendChild(closeBtn);

  const banner = document.createElement('div');
  banner.id = 'fs-b2a-config-banner';

  const errBox = document.createElement('div');
  errBox.id = 'fs-b2a-config-err';

  const scroll = document.createElement('div');
  scroll.id = 'fs-b2a-config-scroll';

  function addInput(section, titleText, id, inputType, placeholder) {
    const row = document.createElement('div');
    row.className = 'fs-b2a-cfg-row';
    const lab = document.createElement('label');
    lab.htmlFor = id;
    lab.textContent = titleText;
    const inp = document.createElement('input');
    inp.type = inputType;
    inp.id = id;
    if (placeholder) inp.placeholder = placeholder;
    row.appendChild(lab);
    row.appendChild(inp);
    section.appendChild(row);
  }

  function addTextarea(section, titleText, id, placeholder, rows) {
    const row = document.createElement('div');
    row.className = 'fs-b2a-cfg-row';
    const lab = document.createElement('label');
    lab.htmlFor = id;
    lab.textContent = titleText;
    const ta = document.createElement('textarea');
    ta.id = id;
    ta.rows = rows || 3;
    if (placeholder) ta.placeholder = placeholder;
    ta.style.cssText =
      'width:100%;box-sizing:border-box;padding:8px 10px;font-size:12px;border:1px solid #d0d3d6;border-radius:6px;font-family:ui-monospace,Consolas,monospace;resize:vertical;min-height:52px;';
    row.appendChild(lab);
    row.appendChild(ta);
    section.appendChild(row);
  }

  const secProfile = document.createElement('div');
  secProfile.className = 'fs-b2a-cfg-section';
  const hProfile = document.createElement('h3');
  hProfile.textContent = '档案与多表配置';
  secProfile.appendChild(hProfile);

  const rowSearch = document.createElement('div');
  rowSearch.className = 'fs-b2a-profile-search-row';
  __profileSearchInput = document.createElement('input');
  __profileSearchInput.type = 'text';
  __profileSearchInput.id = 'fs-b2a-profile-search';
  __profileSearchInput.placeholder = '搜索：档案名、备注、tableId、viewId、摘要…';
  __profileSearchInput.addEventListener('input', () => {
    refreshProfileSelectOptions();
    refreshProfilePickerList();
  });
  rowSearch.appendChild(__profileSearchInput);

  function addFilterChk(container, id, labelText) {
    const w = document.createElement('label');
    w.className = 'fs-b2a-filter-chk';
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.id = id;
    cb.addEventListener('change', () => {
      refreshProfileSelectOptions();
      refreshProfilePickerList();
      updateBindCheckPanel();
    });
    w.appendChild(cb);
    w.appendChild(document.createTextNode(labelText));
    container.appendChild(w);
    return cb;
  }
  __filterMatchOnlyEl = addFilterChk(rowSearch, 'fs-b2a-f-filter-match', '只看当前匹配');
  __filterBoundOnlyEl = addFilterChk(rowSearch, 'fs-b2a-f-filter-bound', '只看已绑定表');
  __filterManualOnlyEl = addFilterChk(rowSearch, 'fs-b2a-f-filter-manual', '只看 manual');
  secProfile.appendChild(rowSearch);

  const pickerWrap = document.createElement('div');
  pickerWrap.className = 'fs-b2a-profile-picker-wrap';
  const labP = document.createElement('label');
  labP.htmlFor = 'fs-b2a-profile-select';
  labP.textContent = '当前档案';
  labP.style.display = 'block';
  labP.style.fontSize = '13px';
  labP.style.color = '#646a73';
  labP.style.marginBottom = '4px';
  const selRow = document.createElement('div');
  selRow.className = 'fs-b2a-select-row';
  __profileSelectEl = document.createElement('select');
  __profileSelectEl.id = 'fs-b2a-profile-select';
  __btnProfilePickerToggle = document.createElement('button');
  __btnProfilePickerToggle.type = 'button';
  __btnProfilePickerToggle.textContent = '展开列表';
  __btnProfilePickerToggle.title = '可视化档案列表';
  __btnProfilePickerToggle.style.cssText =
    'flex-shrink:0;padding:8px 10px;font-size:12px;border:1px solid #d0d3d6;border-radius:6px;background:#fff;cursor:pointer;';
  __btnProfilePickerToggle.addEventListener('click', (e) => {
    e.stopPropagation();
    toggleProfilePicker();
  });
  __profilePickerEl = document.createElement('div');
  __profilePickerEl.id = 'fs-b2a-profile-picker';
  selRow.appendChild(__profileSelectEl);
  selRow.appendChild(__btnProfilePickerToggle);
  pickerWrap.appendChild(labP);
  pickerWrap.appendChild(selRow);
  pickerWrap.appendChild(__profilePickerEl);
  secProfile.appendChild(pickerWrap);

  const rowPb = document.createElement('div');
  rowPb.className = 'fs-b2a-profile-btns';
  const mkBtn = (text, title, fn) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.textContent = text;
    b.title = title;
    b.addEventListener('click', fn);
    return b;
  };
  rowPb.appendChild(
    mkBtn('新建', '新建空白档案', () => {
      const np = createProfile({ name: '新档案' });
      __editingProfileId = np.id;
      fillProfileForm(np);
      log(`已新建档案「${np.name}」，填写后请点击保存`);
    }),
  );
  rowPb.appendChild(
    mkBtn('复制', '复制当前档案为新档案', () => {
      if (!__editingProfileId) return;
      const np = duplicateProfile(__editingProfileId);
      if (np) {
        __editingProfileId = np.id;
        fillProfileForm(np);
        log(`已复制为「${np.name}」，请点击保存`);
      }
    }),
  );
  rowPb.appendChild(
    mkBtn('重命名', '', () => {
      const cur = getProfileById(__editingProfileId);
      if (!cur) return;
      const name = window.prompt('档案名称', cur.name);
      if (name == null) return;
      const trimmed = String(name).trim();
      if (!trimmed) {
        log('名称不能为空', 'warn');
        return;
      }
      updateProfile(__editingProfileId, { name: trimmed });
      refreshProfileSelectOptions();
      updateBindCheckPanel();
      log(`已重命名为「${trimmed}」`);
    }),
  );
  rowPb.appendChild(
    mkBtn('删除', '', () => {
      if (!window.confirm('确定删除当前档案？（至少保留一个档案）')) return;
      if (!deleteProfile(__editingProfileId)) {
        log('无法删除：至少保留一个档案', 'warn');
        return;
      }
      const store = loadProfileStore();
      const p = getProfileById(store.lastProfileId);
      if (p) {
        __editingProfileId = p.id;
        fillProfileForm(p);
      }
      log('已删除档案');
    }),
  );
  rowPb.appendChild(
    mkBtn('从当前页面带入', '将 URL 中的 appToken / tableId / viewId 写入下方可选项', () => {
      applyInferToCurrentProfileForm();
    }),
  );
  rowPb.appendChild(
    mkBtn('从当前表自动识别', '同上：识别 wiki / base 场景下的 table、view 等参数', () => {
      applyInferToCurrentProfileForm();
    }),
  );
  rowPb.appendChild(
    mkBtn('从当前表自动命名', '用页面识别到的表名/视图名或 tableId 填写「表格备注名」', () => {
      autoNameTableLabelFromPage();
    }),
  );
  rowPb.appendChild(
    mkBtn('当前表使用此档案', '绑定当前页 appToken/tableId/viewId 并保存', () => {
      bindCurrentPageToProfileAndSave();
    }),
  );
  secProfile.appendChild(rowPb);

  const rowMatch = document.createElement('div');
  rowMatch.className = 'fs-b2a-cfg-row';
  const labM = document.createElement('label');
  labM.htmlFor = 'fs-b2a-f-match-mode';
  labM.textContent = '自动匹配模式';
  const selM = document.createElement('select');
  selM.id = 'fs-b2a-f-match-mode';
  for (const [val, label] of [
    ['table', '按表匹配（appToken + tableId）'],
    ['table_view', '按表 + 视图匹配'],
    ['manual', '手动档案（不自动匹配）'],
  ]) {
    const o = document.createElement('option');
    o.value = val;
    o.textContent = label;
    selM.appendChild(o);
  }
  rowMatch.appendChild(labM);
  rowMatch.appendChild(selM);
  secProfile.appendChild(rowMatch);

  const bindWrap = document.createElement('div');
  bindWrap.className = 'fs-b2a-bind-wrap';
  const bindH = document.createElement('div');
  bindH.style.cssText = 'font-size:12px;font-weight:600;color:#1f2329;margin-bottom:6px;';
  bindH.textContent = '归属检查（当前页 vs 当前编辑档案）';
  __bindCheckPanelEl = document.createElement('div');
  __bindCheckPanelEl.id = 'fs-b2a-bind-check';
  __bindStatusEl = document.createElement('div');
  __bindStatusEl.id = 'fs-b2a-bind-status';
  bindWrap.appendChild(bindH);
  bindWrap.appendChild(__bindCheckPanelEl);
  bindWrap.appendChild(__bindStatusEl);
  secProfile.appendChild(bindWrap);

  __urlCtxEl = document.createElement('div');
  __urlCtxEl.id = 'fs-b2a-url-ctx';
  secProfile.appendChild(__urlCtxEl);

  addInput(
    secProfile,
    '表格备注名（tableLabel）',
    'fs-b2a-f-table-label',
    'text',
    '如：女装主表、鞋包商品库；列表中优先显示此名称',
  );

  const rowDisp = document.createElement('div');
  rowDisp.className = 'fs-b2a-cfg-row';
  const labDisp = document.createElement('label');
  labDisp.textContent = '页面识别 · 表格显示名（只读，可随 DOM 刷新）';
  labDisp.style.display = 'block';
  labDisp.style.fontSize = '13px';
  labDisp.style.color = '#646a73';
  labDisp.style.marginBottom = '4px';
  const inpDispT = document.createElement('input');
  inpDispT.type = 'text';
  inpDispT.id = 'fs-b2a-f-table-display-name';
  inpDispT.readOnly = true;
  inpDispT.style.cssText =
    'width:100%;box-sizing:border-box;padding:8px 10px;font-size:13px;border:1px solid #e8eaed;border-radius:6px;background:#f7f8fa;color:#646a73;';
  const labDisp2 = document.createElement('label');
  labDisp2.textContent = '页面识别 · 视图显示名（只读）';
  labDisp2.style.display = 'block';
  labDisp2.style.fontSize = '13px';
  labDisp2.style.color = '#646a73';
  labDisp2.style.marginTop = '8px';
  labDisp2.style.marginBottom = '4px';
  const inpDispV = document.createElement('input');
  inpDispV.type = 'text';
  inpDispV.id = 'fs-b2a-f-view-display-name';
  inpDispV.readOnly = true;
  inpDispV.style.cssText =
    'width:100%;box-sizing:border-box;padding:8px 10px;font-size:13px;border:1px solid #e8eaed;border-radius:6px;background:#f7f8fa;color:#646a73;';
  rowDisp.appendChild(labDisp);
  rowDisp.appendChild(inpDispT);
  rowDisp.appendChild(labDisp2);
  rowDisp.appendChild(inpDispV);
  secProfile.appendChild(rowDisp);

  const rowPin = document.createElement('label');
  rowPin.className = 'fs-b2a-cfg-check';
  const cbPin = document.createElement('input');
  cbPin.type = 'checkbox';
  cbPin.id = 'fs-b2a-f-pinned';
  rowPin.appendChild(cbPin);
  rowPin.appendChild(document.createTextNode('置顶该档案（列表排序优先）'));
  secProfile.appendChild(rowPin);

  __profileSelectEl.addEventListener('change', () => {
    if (!__profileSelectEl) return;
    const id = __profileSelectEl.value;
    if (id === __editingProfileId) return;
    const p = getProfileById(id);
    if (p) {
      __editingProfileId = id;
      fillProfileForm(p);
    }
  });

  selM.addEventListener('change', () => {
    updateBindCheckPanel();
  });

  const secReq = document.createElement('div');
  secReq.className = 'fs-b2a-cfg-section';
  const hReq = document.createElement('h3');
  hReq.textContent = '必填';
  secReq.appendChild(hReq);
  addInput(secReq, 'APP_ID', 'fs-b2a-f-app-id', 'text', '飞书开放平台应用 App ID');
  addInput(secReq, 'APP_SECRET', 'fs-b2a-f-app-secret', 'password', '飞书开放平台应用 App Secret');

  const secOpt = document.createElement('div');
  secOpt.className = 'fs-b2a-cfg-section';
  const hOpt = document.createElement('h3');
  hOpt.textContent = '可选项（优先自动识别 URL）';
  secOpt.appendChild(hOpt);
  addInput(
    secOpt,
    'APP_TOKEN',
    'fs-b2a-f-app-token',
    'text',
    '留空：读记录优先 URL；上传 parent 会做 get_node 规范化。填写：上传 parent 固定用此处（优先于 URL）',
  );
  addInput(secOpt, 'TABLE_ID', 'fs-b2a-f-table-id', 'text', '留空则尝试从 table= 或 tbl 识别');
  addInput(
    secOpt,
    'VIEW_ID（表格视图，可选）',
    'fs-b2a-f-view-id',
    'text',
    '留空则尝试从 URL 的 view= 识别；填写后仅拉取该视图下可见记录，减少与其它视图混淆',
  );

  const secFields = document.createElement('div');
  secFields.className = 'fs-b2a-cfg-section';
  const hFields = document.createElement('h3');
  hFields.textContent = '字段名';
  secFields.appendChild(hFields);
  addInput(secFields, '图片链接字段名', 'fs-b2a-f-url-field', 'text', '');
  addInput(secFields, '附件字段名', 'fs-b2a-f-attach-field', 'text', '');
  addInput(secFields, '状态字段名（留空则不写）', 'fs-b2a-f-status-field', 'text', '');
  addInput(secFields, '失败原因字段名（留空则不写）', 'fs-b2a-f-error-field', 'text', '');

  const secExt = document.createElement('div');
  secExt.className = 'fs-b2a-cfg-section';
  const hExt = document.createElement('h3');
  hExt.textContent = '扩展（百应写飞书 / 列映射，按档案保存）';
  secExt.appendChild(hExt);
  addTextarea(
    secExt,
    'fieldMapJson（列名映射 JSON，可选）',
    'fs-b2a-f-field-map-json',
    '与其它脚本共用列映射时可粘贴在此；本「图片转附件」主流程可不填',
    4,
  );
  addTextarea(
    secExt,
    'fieldEnabledJson（字段勾选 JSON，可选）',
    'fs-b2a-f-field-enabled-json',
    '与其它脚本共用字段启用状态',
    4,
  );

  const secSw = document.createElement('div');
  secSw.className = 'fs-b2a-cfg-section';
  const hSw = document.createElement('h3');
  hSw.textContent = '开关';
  secSw.appendChild(hSw);
  const rowSkip = document.createElement('label');
  rowSkip.className = 'fs-b2a-cfg-check';
  const cbSkip = document.createElement('input');
  cbSkip.type = 'checkbox';
  cbSkip.id = 'fs-b2a-f-skip-attach';
  rowSkip.appendChild(cbSkip);
  rowSkip.appendChild(document.createTextNode('已有附件是否跳过'));
  secSw.appendChild(rowSkip);
  const rowApp = document.createElement('label');
  rowApp.className = 'fs-b2a-cfg-check';
  const cbApp = document.createElement('input');
  cbApp.type = 'checkbox';
  cbApp.id = 'fs-b2a-f-append';
  rowApp.appendChild(cbApp);
  rowApp.appendChild(document.createTextNode('追加到已有附件后（关闭则覆盖）'));
  secSw.appendChild(rowApp);

  const secPerf = document.createElement('div');
  secPerf.className = 'fs-b2a-cfg-section';
  const hPerf = document.createElement('h3');
  hPerf.textContent = '性能与限制';
  secPerf.appendChild(hPerf);
  addInput(secPerf, '记录并发数', 'fs-b2a-f-rec-conc', 'number', '');
  addInput(secPerf, '单条记录内 URL 并发数', 'fs-b2a-f-url-conc', 'number', '');
  addInput(secPerf, '每条记录处理间隔（毫秒）', 'fs-b2a-f-interval', 'number', '');
  addInput(secPerf, '请求超时（毫秒）', 'fs-b2a-f-timeout', 'number', '');
  addInput(secPerf, '图片最大体积（字节）', 'fs-b2a-f-max-bytes', 'number', '默认 20971520（20MB）');

  scroll.appendChild(secProfile);
  scroll.appendChild(secReq);
  scroll.appendChild(secOpt);
  scroll.appendChild(secFields);
  scroll.appendChild(secExt);
  scroll.appendChild(secSw);
  scroll.appendChild(secPerf);

  scroll.addEventListener('input', (e) => {
    const t = /** @type {HTMLElement} */ (e.target);
    if (!t || !t.id) return;
    if (
      /^fs-b2a-f-(table-label|app-token|table-id|view-id|table-display-name|view-display-name)$/.test(t.id)
    ) {
      updateBindCheckPanel();
    }
  });

  const footer = document.createElement('div');
  footer.id = 'fs-b2a-config-footer';

  const btnSave = document.createElement('button');
  btnSave.type = 'button';
  btnSave.className = 'fs-b2a-primary';
  btnSave.textContent = '保存配置';
  btnSave.addEventListener('click', () => {
    const draft = collectProfileFromForm();
    const v = validateConfig(draft);
    if (!v.valid) {
      if (__configErrorBox) {
        __configErrorBox.style.display = 'block';
        __configErrorBox.textContent = v.errors.join('；');
      }
      return;
    }
    if (__configErrorBox) __configErrorBox.style.display = 'none';
    if (!__editingProfileId) {
      log('内部错误：未选择档案', 'err');
      return;
    }
    saveProfileForm(__editingProfileId, draft);
    log(`档案「${draft.name}」已保存`);
    closeConfigModal();
  });

  const btnCancel = document.createElement('button');
  btnCancel.type = 'button';
  btnCancel.textContent = '取消';
  btnCancel.addEventListener('click', () => closeConfigModal());

  const btnReset = document.createElement('button');
  btnReset.type = 'button';
  btnReset.textContent = '恢复默认值';
  btnReset.addEventListener('click', () => {
    const cur = getProfileById(__editingProfileId);
    const d = getDefaultConfig();
    const merged = normalizeProfile({
      ...(cur || getDefaultProfile()),
      ...d,
      id: cur ? cur.id : __editingProfileId || generateProfileId(),
      name: cur ? cur.name : '默认档案',
      fieldMapJson: '',
      fieldEnabledJson: '',
      matchMode: cur ? cur.matchMode : 'table',
      tableLabel: '',
      tableDisplayName: '',
      viewDisplayName: '',
      bindSummary: '',
      pinned: false,
    });
    fillProfileForm(merged);
    if (__configErrorBox) {
      __configErrorBox.style.display = 'none';
    }
  });

  footer.appendChild(btnSave);
  footer.appendChild(btnCancel);
  footer.appendChild(btnReset);

  overlay.appendChild(header);
  overlay.appendChild(banner);
  overlay.appendChild(errBox);
  overlay.appendChild(scroll);
  overlay.appendChild(footer);

  document.documentElement.appendChild(overlay);

  __configOverlay = overlay;
  __configBanner = banner;
  __configErrorBox = errBox;
}

/**
 * @param {{ runHint?: boolean }} [opts]
 */
function openConfigModal(opts) {
  renderConfigModal();
  if (!__configOverlay) return;
  const store = loadProfileStore();
  const { profile, reason } = resolveMatchedProfileByCurrentUrl(store);
  touchLastMatchedIfUrlMatch(profile.id, reason);
  const fresh = getProfileById(profile.id) || profile;
  __editingProfileId = fresh.id;
  fillProfileForm(fresh);
  if (__configErrorBox) __configErrorBox.style.display = 'none';
  if (__configBanner) {
    if (opts && opts.runHint) {
      __configBanner.style.display = 'block';
      __configBanner.textContent = '请先填写 APP_ID 和 APP_SECRET';
    } else {
      __configBanner.style.display = 'none';
      __configBanner.textContent = '';
    }
  }
  __configOverlay.style.display = 'flex';
}

function closeConfigModal() {
  if (__configOverlay) __configOverlay.style.display = 'none';
  if (__configBanner) {
    __configBanner.style.display = 'none';
    __configBanner.textContent = '';
  }
}

function syncRunMiniVisibility() {
  if (!__runMiniEl || !__panelEl) return;
  const panelOpen = __panelEl.style.display === 'flex';
  __runMiniEl.style.display = __taskRunning && !panelOpen ? 'flex' : 'none';
}

function updateTaskRunningUI() {
  if (__btnStart) {
    if (__taskRunning) {
      __btnStart.textContent = '展开执行日志';
      __btnStart.title = '任务正在运行，点此打开日志与停止按钮';
    } else {
      __btnStart.textContent = '图片链接转附件';
      __btnStart.title = '';
    }
  }
  if (__btnRootStop) {
    __btnRootStop.style.display = __taskRunning ? 'inline-block' : 'none';
  }
  syncRunMiniVisibility();
}

/**
 * 创建右下角按钮与日志面板
 */
function ensureUI() {
  if (__uiEnsured) return;
  __uiEnsured = true;

  const style = document.createElement('style');
  style.textContent = `
    #fs-b2a-root { position: fixed; right: 16px; bottom: 16px; z-index: 2147483646; display: flex; flex-direction: row; align-items: center; gap: 8px; font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif; }
    #fs-b2a-btn { padding: 10px 14px; border-radius: 8px; border: 1px solid #3366ff; background: #3366ff; color: #fff; cursor: pointer; font-size: 13px; box-shadow: 0 2px 8px rgba(0,0,0,.15); }
    #fs-b2a-btn:disabled { opacity: .6; cursor: not-allowed; }
    #fs-b2a-btn-cfg { padding: 10px 14px; border-radius: 8px; border: 1px solid #d0d3d6; background: #fff; color: #1f2329; cursor: pointer; font-size: 13px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
    #fs-b2a-btn-cfg:hover { border-color: #3366ff; color: #3366ff; }
    #fs-b2a-panel { position: fixed; right: 16px; bottom: 64px; width: min(520px, calc(100vw - 32px)); height: min(360px, 45vh); z-index: 2147483646; background: #fff; border: 1px solid #e5e6eb; border-radius: 8px; box-shadow: 0 8px 24px rgba(0,0,0,.12); display: none; flex-direction: column; overflow: hidden; }
    #fs-b2a-panel header { display: flex; align-items: center; justify-content: space-between; padding: 8px 10px; border-bottom: 1px solid #e5e6eb; font-size: 13px; font-weight: 600; color: #1f2329; }
    #fs-b2a-panel header button { border: none; background: transparent; cursor: pointer; font-size: 16px; line-height: 1; color: #646a73; padding: 4px 8px; }
    #fs-b2a-stats { padding: 6px 10px; font-size: 12px; color: #646a73; border-bottom: 1px solid #f0f0f0; }
    #fs-b2a-log { flex: 1; width: 100%; border: none; resize: none; padding: 8px 10px; font-size: 12px; line-height: 1.45; color: #1f2329; box-sizing: border-box; outline: none; }
    #fs-b2a-actions { display: flex; gap: 8px; padding: 8px 10px; border-top: 1px solid #e5e6eb; }
    #fs-b2a-actions button { flex: 1; padding: 8px 10px; border-radius: 6px; border: 1px solid #d0d3d6; background: #f5f6f7; cursor: pointer; font-size: 12px; }
    #fs-b2a-actions button.primary { background: #3366ff; color: #fff; border-color: #3366ff; }
    #fs-b2a-btn-stop-root { display: none; padding: 10px 14px; border-radius: 8px; border: 1px solid #d93026; background: #fff; color: #d93026; cursor: pointer; font-size: 13px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
    #fs-b2a-btn-stop-root:hover { background: #fff5f5; }
    #fs-b2a-run-mini { display: none; position: fixed; right: 16px; bottom: 72px; z-index: 2147483646; align-items: center; gap: 10px; padding: 8px 12px; background: #1f2329; color: #fff; border-radius: 8px; font-size: 12px; box-shadow: 0 6px 20px rgba(0,0,0,.2); font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif; max-width: min(420px, calc(100vw - 32px)); }
    #fs-b2a-run-mini .fs-b2a-mini-text { flex: 1; min-width: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    #fs-b2a-run-mini button { flex-shrink: 0; padding: 6px 10px; border-radius: 6px; border: none; cursor: pointer; font-size: 12px; }
    #fs-b2a-run-mini .fs-b2a-mini-expand { background: #3366ff; color: #fff; }
    #fs-b2a-run-mini .fs-b2a-mini-stop { background: transparent; color: #ffb4b4; border: 1px solid rgba(255,255,255,.35); }
  `;
  document.documentElement.appendChild(style);

  const root = document.createElement('div');
  root.id = 'fs-b2a-root';

  const btnCfg = document.createElement('button');
  btnCfg.id = 'fs-b2a-btn-cfg';
  btnCfg.type = 'button';
  btnCfg.textContent = '配置';
  btnCfg.addEventListener('click', () => openConfigModal());

  __btnRootStop = document.createElement('button');
  __btnRootStop.id = 'fs-b2a-btn-stop-root';
  __btnRootStop.type = 'button';
  __btnRootStop.textContent = '停止任务';
  __btnRootStop.title = '请求停止：当前记录内 URL 与当前条处理完后结束';
  __btnRootStop.addEventListener('click', () => {
    __stopRequested = true;
    log('已请求停止任务（当前这条处理完后结束）…', 'warn');
  });

  __btnStart = document.createElement('button');
  __btnStart.id = 'fs-b2a-btn';
  __btnStart.type = 'button';
  __btnStart.textContent = '图片链接转附件';
  __btnStart.addEventListener('click', () => {
    if (__taskRunning) {
      if (__panelEl) __panelEl.style.display = 'flex';
      syncRunMiniVisibility();
      return;
    }
    if (__panelEl) __panelEl.style.display = 'flex';
    void startProcess();
  });

  root.appendChild(btnCfg);
  root.appendChild(__btnRootStop);
  root.appendChild(__btnStart);

  __panelEl = document.createElement('div');
  __panelEl.id = 'fs-b2a-panel';

  const header = document.createElement('header');
  const title = document.createElement('span');
  title.textContent = '执行日志';
  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.textContent = '×';
  closeBtn.title = '收起面板（任务仍在后台运行；运行中会显示「运行中」提示条，可再展开）';
  closeBtn.addEventListener('click', () => {
    if (__panelEl) __panelEl.style.display = 'none';
    syncRunMiniVisibility();
  });
  header.appendChild(title);
  header.appendChild(closeBtn);

  __statEl = document.createElement('div');
  __statEl.id = 'fs-b2a-stats';
  __statEl.textContent = '成功: 0  跳过: 0  失败: 0';

  __logEl = document.createElement('textarea');
  __logEl.id = 'fs-b2a-log';
  __logEl.readOnly = true;
  __logEl.spellcheck = false;

  const actions = document.createElement('div');
  actions.id = 'fs-b2a-actions';
  __btnStop = document.createElement('button');
  __btnStop.type = 'button';
  __btnStop.textContent = '停止任务';
  __btnStop.addEventListener('click', () => {
    __stopRequested = true;
    log('已请求停止任务（当前这条处理完后结束）…', 'warn');
  });
  const clearBtn = document.createElement('button');
  clearBtn.type = 'button';
  clearBtn.textContent = '清空日志';
  clearBtn.addEventListener('click', () => {
    if (__logEl) __logEl.value = '';
  });
  actions.appendChild(__btnStop);
  actions.appendChild(clearBtn);

  __panelEl.appendChild(header);
  __panelEl.appendChild(__statEl);
  __panelEl.appendChild(__logEl);
  __panelEl.appendChild(actions);

  document.documentElement.appendChild(root);
  document.documentElement.appendChild(__panelEl);

  __runMiniEl = document.createElement('div');
  __runMiniEl.id = 'fs-b2a-run-mini';
  const miniText = document.createElement('span');
  miniText.className = 'fs-b2a-mini-text';
  miniText.textContent = '图片转附件 · 运行中';
  const miniExpand = document.createElement('button');
  miniExpand.type = 'button';
  miniExpand.className = 'fs-b2a-mini-expand';
  miniExpand.textContent = '展开日志';
  miniExpand.addEventListener('click', () => {
    if (__panelEl) __panelEl.style.display = 'flex';
    syncRunMiniVisibility();
  });
  const miniStop = document.createElement('button');
  miniStop.type = 'button';
  miniStop.className = 'fs-b2a-mini-stop';
  miniStop.textContent = '停止';
  miniStop.addEventListener('click', () => {
    __stopRequested = true;
    log('已请求停止任务（当前这条处理完后结束）…', 'warn');
  });
  __runMiniEl.appendChild(miniText);
  __runMiniEl.appendChild(miniExpand);
  __runMiniEl.appendChild(miniStop);
  document.documentElement.appendChild(__runMiniEl);

  renderConfigModal();
}

/**
 * @param {{ ok: number; skip: number; fail: number }} stats
 */
function updateStatsUI(stats) {
  if (__statEl) {
    __statEl.textContent = `成功: ${stats.ok}  跳过: ${stats.skip}  失败: ${stats.fail}`;
  }
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

/**
 * 限制并发执行异步任务
 * @template T
 * @param {T[]} items
 * @param {number} limit
 * @param {(item: T, index: number) => Promise<void>} worker
 */
async function runPool(items, limit, worker) {
  if (limit <= 1) {
    for (let i = 0; i < items.length; i += 1) {
      if (__stopRequested) break;
      await worker(items[i], i);
    }
    return;
  }
  let idx = 0;
  const runners = new Array(Math.min(limit, items.length)).fill(0).map(async () => {
    while (!__stopRequested) {
      const my = idx;
      idx += 1;
      if (my >= items.length) break;
      await worker(items[my], my);
    }
  });
  await Promise.all(runners);
}

// =============================================================================
// GM_xmlhttpRequest Promise 封装
// =============================================================================

/**
 * @param {Parameters<typeof GM_xmlhttpRequest>[0]} opts
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 * @returns {Promise<GMXMLHttpRequestResponse>}
 */
function gmXhr(opts, cfg) {
  const timeout = opts.timeout != null ? opts.timeout : cfg.requestTimeoutMs;
  return new Promise((resolve, reject) => {
    GM_xmlhttpRequest({
      method: opts.method || 'GET',
      url: opts.url,
      headers: opts.headers || {},
      data: opts.data,
      binary: opts.binary,
      responseType: opts.responseType || 'text',
      timeout,
      onload: (res) => resolve(res),
      onerror: (err) => reject(err),
      ontimeout: () => reject(new Error('网络超时')),
    });
  });
}

/**
 * @param {GMXMLHttpRequestResponse} res
 * @returns {any}
 */
function parseJsonResponse(res) {
  const text = res.responseText || '';
  try {
    return JSON.parse(text);
  } catch {
    throw new Error(`响应非 JSON（HTTP ${res.status}）：${text.slice(0, 200)}`);
  }
}

/**
 * @param {string} raw
 */
function feishuApiErrorMessage(raw) {
  try {
    const j = JSON.parse(raw);
    if (j && typeof j.code === 'number' && j.code !== 0) {
      return `code=${j.code} msg=${j.msg || ''}`;
    }
  } catch {
    /* ignore */
  }
  return raw.slice(0, 400);
}

/**
 * 兼容 Tampermonkey 下 responseHeaders 为字符串或对象
 * @param {GMXMLHttpRequestResponse} res
 * @param {string} name
 */
function getResponseHeader(res, name) {
  const h = res.responseHeaders;
  if (!h) return '';
  if (typeof h === 'string') {
    const re = new RegExp(`^${name}:\\s*(.+)$`, 'im');
    const m = h.match(re);
    return m ? m[1].trim() : '';
  }
  if (typeof h === 'object') {
    const key = Object.keys(h).find((k) => k.toLowerCase() === name.toLowerCase());
    return key ? String(h[key]) : '';
  }
  return '';
}

// =============================================================================
// URL 推断
// =============================================================================

/**
 * 从当前页面 URL / hash 尝试解析 app_token 与 table_id
 * @param {string} [href]
 * @returns {{ app_token: string | null; table_id: string | null; view_id: string | null }}
 */
function inferAppAndTableFromUrl(href) {
  const h = href || `${location.pathname}${location.search}${location.hash || ''}`;
  let app_token = null;
  let table_id = null;
  let view_id = null;

  const baseMatch = h.match(/\/base\/([A-Za-z0-9_-]+)/);
  if (baseMatch) app_token = baseMatch[1];

  const tableMatch = h.match(/[?&#]table=([A-Za-z0-9_-]+)/i) || h.match(/[?&#]tableId=([A-Za-z0-9_-]+)/i);
  if (tableMatch) table_id = tableMatch[1];

  if (!table_id) {
    const tbl = h.match(/\btbl[A-Za-z0-9_-]{4,}\b/);
    if (tbl) table_id = tbl[0];
  }

  const viewMatch = h.match(/[?&#]view=([A-Za-z0-9_-]+)/i) || h.match(/[?&#]viewId=([A-Za-z0-9_-]+)/i);
  if (viewMatch) view_id = viewMatch[1];

  return { app_token, table_id, view_id };
}

/**
 * 当前页面是否为知识库 wiki 路径（非独立 /base/ 打开）
 * @param {string} [href]
 */
function isWikiPage(href) {
  const h = href || `${location.pathname}${location.search}${location.hash || ''}`;
  return /\/wiki\//i.test(h);
}

/**
 * 从 URL 提取 wiki 节点 token（/wiki/ 后第一段）
 * @param {string} [href]
 * @returns {string | null}
 */
function inferWikiNodeTokenFromUrl(href) {
  const h = href || `${location.pathname}${location.search}${location.hash || ''}`;
  const m = h.match(/\/wiki\/([A-Za-z0-9_-]+)/i);
  return m ? m[1] : null;
}

/**
 * 调用知识库「获取节点信息」，用于 wiki 场景解析真实 bitable app_token
 * @param {string} tenant
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 * @param {string} token
 * @param {string} [objType] wiki | bitable | ...
 */
async function fetchWikiSpaceGetNode(tenant, cfg, token, objType) {
  const qs = new URLSearchParams();
  qs.set('token', token);
  if (objType) qs.set('obj_type', objType);
  const url = `https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?${qs.toString()}`;
  const res = await gmXhr(
    {
      method: 'GET',
      url,
      headers: { Authorization: `Bearer ${tenant}` },
      responseType: 'text',
    },
    cfg,
  );
  if (res.status < 200 || res.status >= 300) {
    throw new Error(`Wiki get_node HTTP ${res.status}: ${feishuApiErrorMessage(res.responseText || '')}`);
  }
  return parseJsonResponse(res);
}

/**
 * 尝试用 get_node(token+obj_type=bitable) 将「当前可用的 base token」规范为云文档侧 obj_token（改善 upload_all parent_node）
 * @param {string} tenant
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 * @param {string} token
 * @returns {Promise<string>}
 */
async function tryCanonicalBitableTokenViaGetNode(tenant, cfg, token) {
  if (!token) return token;
  try {
    const body = await fetchWikiSpaceGetNode(tenant, cfg, token, 'bitable');
    if (body.code !== 0) return token;
    const node = body.data && body.data.node;
    const ot = node && node.obj_token;
    if (typeof ot === 'string' && ot.trim()) return ot.trim();
  } catch {
    /* 忽略，回退原 token */
  }
  return token;
}

/**
 * 解析「读记录 / 更新记录」使用的多维表格 app_token
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 * @param {string} tenant
 * @param {{ app_token: string | null; table_id: string | null }} inferred
 * @param {boolean} wikiScene
 */
async function resolveRecordAppToken(cfg, tenant, inferred, wikiScene) {
  if (wikiScene) {
    const wikiTok = inferWikiNodeTokenFromUrl();
    if (!wikiTok) {
      throw new Error('当前为 Wiki 场景，但无法从 URL 解析 wiki 节点 token，请在地址栏打开对应知识库页面或于配置填写 APP_TOKEN');
    }
    const body = await fetchWikiSpaceGetNode(tenant, cfg, wikiTok, 'wiki');
    if (body.code !== 0) {
      throw new Error(`Wiki get_node 失败：code=${body.code} msg=${body.msg || ''}`);
    }
    const node = body.data && body.data.node;
    if (!node) throw new Error('Wiki get_node 返回无 node 数据');
    const objType = node.obj_type;
    const objToken = node.obj_token;
    if (objType === 'bitable' && objToken) {
      return String(objToken).trim();
    }
    throw new Error(
      `Wiki 节点文档类型为「${objType || '未知'}」，非多维表格(bitable)。请在飞书中点开目标多维表格使 URL 含 bitable，或在配置中填写正确的 APP_TOKEN`,
    );
  }

  const fromUrl = inferred.app_token && String(inferred.app_token).trim();
  const fromCfg = cfg.appToken && String(cfg.appToken).trim();
  const t = fromUrl || fromCfg;
  if (!t) {
    throw new Error('无法解析读记录用 APP_TOKEN：请打开 *.feishu.cn/base/xxx 页面，或在「配置」中填写 APP_TOKEN');
  }
  return t;
}

/**
 * 同步：配置中手填的 APP_TOKEN 优先作为上传父节点；否则使用读记录 token（后续再由 resolveActualAppTokenForUpload 规范化）
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 * @param {string} recordAppToken
 * @returns {{ manualUploadToken: string | null; fallbackRecordToken: string }}
 */
function resolveUploadParentNode(cfg, recordAppToken) {
  const manual = cfg.appToken && String(cfg.appToken).trim();
  return {
    manualUploadToken: manual || null,
    fallbackRecordToken: recordAppToken,
  };
}

/**
 * 最终用于 upload_all 的 parent_node：手填配置不被 URL 覆盖；未手填则对 record token 做 get_node(bitable) 规范化
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 * @param {string} tenant
 * @param {string} recordAppToken
 * @returns {Promise<{ uploadParentNode: string; uploadSource: string }>}
 */
async function resolveActualAppTokenForUpload(cfg, tenant, recordAppToken) {
  const { manualUploadToken, fallbackRecordToken } = resolveUploadParentNode(cfg, recordAppToken);
  if (manualUploadToken) {
    return { uploadParentNode: manualUploadToken, uploadSource: '配置 APP_TOKEN（手动，优先于 URL）' };
  }
  const canon = await tryCanonicalBitableTokenViaGetNode(tenant, cfg, fallbackRecordToken);
  if (canon !== fallbackRecordToken) {
    return { uploadParentNode: canon, uploadSource: 'get_node(obj_type=bitable) 规范化后的 obj_token' };
  }
  return { uploadParentNode: fallbackRecordToken, uploadSource: '与读记录 token 相同（未手填配置且 get_node 未返回新 token）' };
}

/**
 * 任务开始时打印 token 来源，便于排查 1061044 parent node not exist
 * @param {object} p
 */
function logResolvedTokens(p) {
  const {
    inferredApp,
    inferredTable,
    inferredView,
    configAppToken,
    configTableId,
    configViewId,
    finalViewId,
    recordAppToken,
    uploadParentNode,
    uploadSource,
    wikiScene,
  } = p;
  log(`页面 Wiki 场景：${wikiScene ? '是（/wiki/）' : '否'}`);
  log(`URL 自动识别 APP_TOKEN：${inferredApp || '（无）'}`);
  log(`URL 自动识别 TABLE_ID：${inferredTable || '（无）'}`);
  log(`URL 自动识别 VIEW_ID：${inferredView || '（无）'}`);
  log(`配置中 APP_TOKEN：${configAppToken ? configAppToken : '（空）'}`);
  log(`配置中 TABLE_ID：${configTableId ? configTableId : '（空）'}`);
  log(`配置中 VIEW_ID：${configViewId ? configViewId : '（空）'}`);
  log(
    `最终拉取记录使用的 VIEW_ID：${
      finalViewId ? finalViewId : '（未指定，拉取整张表全部记录）'
    }`,
  );
  log(`最终读记录 / 更新附件列 recordAppToken：${recordAppToken}`);
  log(`最终上传素材 parent_node（upload_all）：${uploadParentNode}（来源：${uploadSource}）`);
}

/**
 * 是否字段名不存在（1254045）
 * @param {string} msg
 */
function isFieldNameNotFoundError(msg) {
  return typeof msg === 'string' && (msg.includes('1254045') || msg.includes('FieldNameNotFound'));
}

/**
 * 是否上传父节点不存在（1061044）
 * @param {string} msg
 */
function isParentNodeNotExistError(msg) {
  return typeof msg === 'string' && (msg.includes('1061044') || /parent node not exist/i.test(msg));
}

/**
 * 状态 / 失败原因可选列：字段不存在时本轮任务仅提示一次并自动禁用
 * @param {'status'|'error'} kind
 * @param {string} fieldName
 * @param {Record<string, any>} fields
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 * @param {{ statusDisabled: boolean; errorDisabled: boolean; statusNotFoundLogged: boolean; errorNotFoundLogged: boolean }} writer
 */
async function safeOptionalFieldWriter(token, recordAppToken, tableId, recordId, kind, fieldName, fields, cfg, writer) {
  const name = String(fieldName || '').trim();
  if (!name) return;
  if (kind === 'status' && writer.statusDisabled) return;
  if (kind === 'error' && writer.errorDisabled) return;

  try {
    await updateRecord(token, recordAppToken, tableId, recordId, fields, cfg);
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    if (isFieldNameNotFoundError(msg)) {
      if (kind === 'status' && !writer.statusNotFoundLogged) {
        log('状态字段不存在，后续已自动跳过写入状态列', 'warn');
        writer.statusNotFoundLogged = true;
        writer.statusDisabled = true;
      }
      if (kind === 'error' && !writer.errorNotFoundLogged) {
        log('失败原因字段不存在，后续已自动跳过写入失败原因列', 'warn');
        writer.errorNotFoundLogged = true;
        writer.errorDisabled = true;
      }
      return;
    }
    log(`${kind === 'status' ? '状态' : '失败原因'} 写入失败：${msg}`, 'warn');
  }
}

// =============================================================================
// 飞书 API
// =============================================================================

/**
 * 获取 tenant_access_token
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 * @returns {Promise<string>}
 */
async function getTenantAccessToken(cfg) {
  if (!cfg.appId || !cfg.appSecret) {
    throw new Error('请先在「配置」中填写 APP_ID 与 APP_SECRET');
  }
  const res = await gmXhr(
    {
      method: 'POST',
      url: 'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal',
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      data: JSON.stringify({ app_id: cfg.appId, app_secret: cfg.appSecret }),
      responseType: 'text',
    },
    cfg,
  );
  if (res.status < 200 || res.status >= 300) {
    throw new Error(`获取 tenant_access_token HTTP ${res.status}: ${feishuApiErrorMessage(res.responseText || '')}`);
  }
  const body = parseJsonResponse(res);
  if (body.code !== 0) {
    throw new Error(`获取 tenant_access_token 失败：code=${body.code} msg=${body.msg || ''}`);
  }
  if (!body.tenant_access_token) {
    throw new Error('获取 tenant_access_token 失败：响应缺少 tenant_access_token');
  }
  return body.tenant_access_token;
}

/**
 * 分页拉取全部记录
 * @param {string} token
 * @param {string} appToken
 * @param {string} tableId
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 * @param {string} [viewId] 非空时传给接口 view_id，仅列举该表格视图下可见记录
 * @returns {Promise<{ items: any[]; total: number | null }>}
 */
async function listAllRecords(token, appToken, tableId, cfg, viewId) {
  const items = [];
  let page_token = undefined;
  let total = null;
  const v = viewId && String(viewId).trim();

  for (;;) {
    if (__stopRequested) break;
    const qs = new URLSearchParams();
    qs.set('page_size', '500');
    if (page_token) qs.set('page_token', page_token);
    if (v) qs.set('view_id', v);
    const url = `https://open.feishu.cn/open-apis/bitable/v1/apps/${encodeURIComponent(appToken)}/tables/${encodeURIComponent(tableId)}/records?${qs.toString()}`;
    const res = await gmXhr(
      {
        method: 'GET',
        url,
        headers: { Authorization: `Bearer ${token}` },
        responseType: 'text',
      },
      cfg,
    );
    if (res.status < 200 || res.status >= 300) {
      throw new Error(`读取记录失败 HTTP ${res.status}: ${feishuApiErrorMessage(res.responseText || '')}`);
    }
    const body = parseJsonResponse(res);
    if (body.code !== 0) {
      throw new Error(`读取记录失败：code=${body.code} msg=${body.msg || ''}`);
    }
    const data = body.data || {};
    if (typeof data.total === 'number' && total === null) total = data.total;
    const batch = Array.isArray(data.items) ? data.items : [];
    items.push(...batch);
    if (!data.has_more || !data.page_token) break;
    page_token = data.page_token;
  }
  if (__stopRequested && items.length && (total == null || items.length < total)) {
    log('读取记录过程中收到停止请求，已返回已拉取的部分记录', 'warn');
  }
  return { items, total };
}

/**
 * 从字段原始值中提取 URL 字符串列表
 * @param {any} raw
 * @returns {string[]}
 */
function extractUrls(raw) {
  if (raw == null) return [];
  const chunks = [];

  const pushText = (t) => {
    if (t == null) return;
    const s = String(t).trim();
    if (s) chunks.push(s);
  };

  if (typeof raw === 'string' || typeof raw === 'number') {
    pushText(raw);
  } else if (Array.isArray(raw)) {
    for (const x of raw) {
      if (typeof x === 'string' || typeof x === 'number') pushText(x);
      else if (x && typeof x === 'object') {
        if (typeof x.link === 'string') pushText(x.link);
        if (typeof x.url === 'string') pushText(x.url);
        if (typeof x.text === 'string') pushText(x.text);
      }
    }
  } else if (typeof raw === 'object') {
    if (typeof raw.link === 'string') pushText(raw.link);
    if (typeof raw.url === 'string') pushText(raw.url);
    if (typeof raw.text === 'string') pushText(raw.text);
  }

  const text = chunks.join('\n');
  const re = /https?:\/\/[^\s"'<>[\]{}]+/gi;
  const found = text.match(re) || [];
  const cleaned = [];
  const seen = new Set();
  for (let u of found) {
    u = u.replace(/[),.;]+$/g, '');
    if (!/^https?:\/\//i.test(u)) continue;
    if (seen.has(u)) continue;
    seen.add(u);
    cleaned.push(u);
  }
  return cleaned;
}

/**
 * 猜测文件名与扩展名
 * @param {string} url
 * @param {string} [contentType]
 */
function guessFileName(url, contentType) {
  let name = 'image';
  try {
    const u = new URL(url);
    const last = u.pathname.split('/').filter(Boolean).pop() || '';
    if (last && /\.[a-z0-9]{2,5}$/i.test(last)) {
      name = decodeURIComponent(last.split('?')[0]);
    }
  } catch {
    /* ignore */
  }
  if (!/\.[a-z0-9]{2,5}$/i.test(name) && contentType) {
    const ct = contentType.split(';')[0].trim().toLowerCase();
    const map = {
      'image/jpeg': '.jpg',
      'image/jpg': '.jpg',
      'image/png': '.png',
      'image/gif': '.gif',
      'image/webp': '.webp',
      'image/bmp': '.bmp',
      'image/svg+xml': '.svg',
    };
    if (map[ct]) name += map[ct];
    else if (ct.startsWith('image/')) name += `.${ct.replace('image/', '').replace('+xml', '')}`;
  }
  if (!/\.[a-z0-9]{2,5}$/i.test(name)) name += '.jpg';
  if (name.length > 200) name = name.slice(-200);
  return name;
}

/**
 * 使用 GM_xmlhttpRequest 下载图片为 Blob
 * @param {string} url
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 * @returns {Promise<{ blob: Blob; contentType: string }>}
 */
async function downloadImage(url, cfg) {
  if (!/^https?:\/\//i.test(url)) {
    throw new Error(`图片 URL 无效：${url}`);
  }
  const res = await gmXhr(
    {
      method: 'GET',
      url,
      responseType: 'arraybuffer',
    },
    cfg,
  );
  if (res.status < 200 || res.status >= 300) {
    throw new Error(`图片下载失败 HTTP ${res.status}：${url}`);
  }
  const buf = res.response;
  if (!buf || !(buf.byteLength >= 0)) {
    throw new Error(`图片下载失败：无响应体 ${url}`);
  }
  if (buf.byteLength > cfg.maxImageBytes) {
    throw new Error(`图片超过 ${cfg.maxImageBytes} 字节：${url}`);
  }
  const ct = getResponseHeader(res, 'content-type') || 'application/octet-stream';
  const blob = new Blob([buf], { type: ct.split(';')[0].trim() });
  return { blob, contentType: ct };
}

/**
 * 拼接 multipart/form-data（Uint8Array）
 * @param {Record<string, string>} fields
 * @param {string} fileFieldName
 * @param {string} fileName
 * @param {Blob} fileBlob
 */
async function buildMultipartBody(fields, fileFieldName, fileName, fileBlob) {
  const boundary = `----feishuB2A${Date.now()}${Math.random().toString(16).slice(2)}`;
  const enc = new TextEncoder();
  const chunks = [];

  const pushText = (s) => chunks.push(enc.encode(s));

  for (const [k, v] of Object.entries(fields)) {
    pushText(`--${boundary}\r\n`);
    pushText(`Content-Disposition: form-data; name="${k}"\r\n\r\n`);
    pushText(String(v));
    pushText('\r\n');
  }

  pushText(`--${boundary}\r\n`);
  pushText(
    `Content-Disposition: form-data; name="${fileFieldName}"; filename="${fileName.replace(/"/g, '')}"\r\n`,
  );
  pushText('Content-Type: application/octet-stream\r\n\r\n');

  const head = concatUint8Arrays(chunks);
  const fileBuf = new Uint8Array(await fileBlob.arrayBuffer());
  const tail = enc.encode(`\r\n--${boundary}--\r\n`);

  const out = new Uint8Array(head.length + fileBuf.length + tail.length);
  out.set(head, 0);
  out.set(fileBuf, head.length);
  out.set(tail, head.length + fileBuf.length);
  return { body: out, contentType: `multipart/form-data; boundary=${boundary}` };
}

/**
 * @param {Uint8Array[]} parts
 */
function concatUint8Arrays(parts) {
  let len = 0;
  for (const p of parts) len += p.length;
  const out = new Uint8Array(len);
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

/**
 * 单次上传（内部）
 * @param {string} token
 * @param {string} uploadParentNode
 * @param {Blob} blob
 * @param {string} fileName
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 */
async function uploadToFeishuOnce(token, uploadParentNode, blob, fileName, cfg) {
  const size = blob.size;
  if (size > cfg.maxImageBytes) {
    throw new Error(`图片超过 ${cfg.maxImageBytes} 字节，无法上传：${fileName}`);
  }
  const { body, contentType } = await buildMultipartBody(
    {
      file_name: fileName,
      parent_type: 'bitable_image',
      parent_node: uploadParentNode,
      size: String(size),
    },
    'file',
    fileName,
    blob,
  );

  const uploadBuf = body.buffer.slice(body.byteOffset, body.byteOffset + body.byteLength);
  const res = await gmXhr(
    {
      method: 'POST',
      url: 'https://open.feishu.cn/open-apis/drive/v1/medias/upload_all',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': contentType,
      },
      data: uploadBuf,
      binary: true,
      responseType: 'text',
    },
    cfg,
  );

  if (res.status < 200 || res.status >= 300) {
    const detail = feishuApiErrorMessage(res.responseText || '');
    throw new Error(
      `上传飞书失败 HTTP ${res.status}: ${detail} | parent_type=bitable_image parent_node=${uploadParentNode}`,
    );
  }
  const json = parseJsonResponse(res);
  if (json.code !== 0) {
    throw new Error(
      `上传飞书失败：code=${json.code} msg=${json.msg || ''} | parent_type=bitable_image parent_node=${uploadParentNode}`,
    );
  }
  const file_token = json.data && json.data.file_token;
  if (!file_token) {
    throw new Error(`上传飞书失败：响应缺少 file_token | parent_type=bitable_image parent_node=${uploadParentNode}`);
  }
  return file_token;
}

/**
 * 上传素材到飞书；1061044 时打印完整 parent 信息；若与读记录 token 不同则尝试用 recordAppToken 重试一次
 * @param {string} token
 * @param {string} uploadParentNode
 * @param {string} recordAppToken
 * @param {Blob} blob
 * @param {string} fileName
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 */
async function uploadToFeishu(token, uploadParentNode, recordAppToken, blob, fileName, cfg) {
  try {
    return await uploadToFeishuOnce(token, uploadParentNode, blob, fileName, cfg);
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    if (isParentNodeNotExistError(msg)) {
      log(
        `[上传失败] parent_type=bitable_image parent_node=${uploadParentNode} | 接口返回：${msg}`,
        'err',
      );
      if (recordAppToken && uploadParentNode !== recordAppToken) {
        log(
          `[上传] 尝试改用读记录 token 作为 parent_node 重试一次：parent_type=bitable_image parent_node=${recordAppToken}`,
          'warn',
        );
        return await uploadToFeishuOnce(token, recordAppToken, blob, fileName, cfg);
      }
    }
    throw e;
  }
}

/**
 * 归一化附件列为 {file_token} 数组
 * @param {any} val
 * @returns {{ file_token: string }[]}
 */
function normalizeAttachmentTokens(val) {
  if (!val) return [];
  if (!Array.isArray(val)) return [];
  const out = [];
  for (const x of val) {
    if (x && typeof x === 'object' && x.file_token) {
      out.push({ file_token: String(x.file_token) });
    }
  }
  return out;
}

/**
 * 判断附件列是否“已有内容”（用于跳过）
 * @param {any} val
 */
function attachmentLooksEmpty(val) {
  const arr = normalizeAttachmentTokens(val);
  return arr.length === 0;
}

/**
 * 更新记录（核心字段）；失败则抛出
 * @param {string} token
 * @param {string} appToken
 * @param {string} tableId
 * @param {string} recordId
 * @param {Record<string, any>} fields
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 */
async function updateRecord(token, appToken, tableId, recordId, fields, cfg) {
  const url = `https://open.feishu.cn/open-apis/bitable/v1/apps/${encodeURIComponent(appToken)}/tables/${encodeURIComponent(tableId)}/records/${encodeURIComponent(recordId)}`;
  const res = await gmXhr(
    {
      method: 'PUT',
      url,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json; charset=utf-8',
      },
      data: JSON.stringify({ fields }),
      responseType: 'text',
    },
    cfg,
  );
  if (res.status < 200 || res.status >= 300) {
    throw new Error(`更新记录失败 HTTP ${res.status}: ${feishuApiErrorMessage(res.responseText || '')}`);
  }
  const body = parseJsonResponse(res);
  if (body.code !== 0) {
    throw new Error(`更新记录失败：code=${body.code} msg=${body.msg || ''}`);
  }
}

// =============================================================================
// 主流程
// =============================================================================

/**
 * 拉取全表记录后预检：无需网络即可判定的跳过/失败一次性计入统计，避免逐条 processIntervalMs 与刷屏日志。
 * @param {ReturnType<typeof getDefaultConfig>} cfg
 * @param {any[]} items
 * @param {{ ok: number; skip: number; fail: number }} stats
 * @returns {number[]} 仍需执行 processOneRecord 的 items 下标（保持原顺序）
 */
function partitionRecordsForProcessing(cfg, items, stats) {
  const buckets = { emptyUrl: 0, noUrls: 0, hasAttachment: 0 };
  let noId = 0;
  const indices = [];
  for (let idx = 0; idx < items.length; idx += 1) {
    if (__stopRequested) break;
    const record = items[idx];
    const recordId = record.record_id || record.id || record.recordId;
    if (!recordId) {
      noId += 1;
      continue;
    }
    const fields = record.fields || {};
    const urlRaw = fields[cfg.urlFieldName];
    const attachRaw = fields[cfg.attachmentFieldName];
    if (urlRaw == null || urlRaw === '') {
      buckets.emptyUrl += 1;
      continue;
    }
    if (!extractUrls(urlRaw).length) {
      buckets.noUrls += 1;
      continue;
    }
    if (cfg.skipIfAttachmentExists && !attachmentLooksEmpty(attachRaw)) {
      buckets.hasAttachment += 1;
      continue;
    }
    indices.push(idx);
  }
  stats.skip += buckets.emptyUrl + buckets.noUrls + buckets.hasAttachment;
  stats.fail += noId;
  updateStatsUI(stats);
  const bits = [];
  if (buckets.hasAttachment) bits.push(`已有附件 ${buckets.hasAttachment} 条`);
  if (buckets.emptyUrl) bits.push(`「${cfg.urlFieldName}」为空 ${buckets.emptyUrl} 条`);
  if (buckets.noUrls) bits.push(`无有效 http(s) 链接 ${buckets.noUrls} 条`);
  if (bits.length) log(`预检快速跳过（不占逐条处理间隔）：${bits.join('；')}`);
  if (noId) log(`预检：${noId} 条缺少 record_id，已计入失败`, 'err');
  return indices;
}

/**
 * 处理单条记录
 * @param {object} args
 */
async function processOneRecord(args) {
  const { token, recordAppToken, uploadParentNode, tableId, record, index, total, stats, cfg, fieldWriter } = args;

  const recordId = record.record_id || record.id || record.recordId;
  if (!recordId) {
    stats.fail += 1;
    updateStatsUI(stats);
    log(`第 ${index + 1}/${total} 条：缺少 record_id，失败`, 'err');
    return;
  }

  const fields = record.fields || {};
  const urlRaw = fields[cfg.urlFieldName];
  const attachRaw = fields[cfg.attachmentFieldName];

  const prefix = `第 ${index + 1}/${total} 条（${recordId}）`;

  if (__stopRequested) {
    log(`${prefix}：已停止，跳过`, 'warn');
    return;
  }

  if (urlRaw == null || urlRaw === '') {
    stats.skip += 1;
    updateStatsUI(stats);
    log(`${prefix}：「${cfg.urlFieldName}」为空，跳过`);
    return;
  }

  const urls = extractUrls(urlRaw);
  if (!urls.length) {
    stats.skip += 1;
    updateStatsUI(stats);
    log(`${prefix}：未解析到有效 http(s) 链接，跳过`);
    return;
  }

  if (cfg.skipIfAttachmentExists && !attachmentLooksEmpty(attachRaw)) {
    stats.skip += 1;
    updateStatsUI(stats);
    log(`${prefix}：「${cfg.attachmentFieldName}」已有附件且开启跳过，跳过`);
    return;
  }

  const existing = cfg.appendToExistingAttachments ? normalizeAttachmentTokens(attachRaw) : [];
  const newTokens = [];
  const errors = [];

  await runPool(urls, Math.max(1, cfg.urlConcurrency), async (imageUrl) => {
    if (__stopRequested) return;
    try {
      log(`${prefix}：下载中… ${imageUrl}`);
      const { blob, contentType } = await downloadImage(imageUrl, cfg);
      const fileName = guessFileName(imageUrl, contentType);
      log(`${prefix}：上传中… ${fileName}（${blob.size} 字节）`);
      const file_token = await uploadToFeishu(token, uploadParentNode, recordAppToken, blob, fileName, cfg);
      newTokens.push({ file_token });
    } catch (e) {
      const msg = e && e.message ? e.message : String(e);
      errors.push(`${imageUrl} -> ${msg}`);
      log(`${prefix}：${msg}`, 'err');
    }
  });

  if (__stopRequested) {
    log(`${prefix}：任务已停止，本条未回写`, 'warn');
    return;
  }

  const merged = cfg.appendToExistingAttachments ? existing.concat(newTokens) : newTokens;

  const errField = String(cfg.errorFieldName || '').trim();
  const statusField = String(cfg.statusFieldName || '').trim();

  if (!merged.length) {
    stats.fail += 1;
    updateStatsUI(stats);
    const errText = errors.length ? errors.join(' | ') : '无可用附件';
    log(`${prefix}：失败（无任何成功上传）`, 'err');
    if (errField) {
      await safeOptionalFieldWriter(
        token,
        recordAppToken,
        tableId,
        recordId,
        'error',
        errField,
        { [errField]: errText.slice(0, 4900) },
        cfg,
        fieldWriter,
      );
    }
    if (statusField) {
      await safeOptionalFieldWriter(
        token,
        recordAppToken,
        tableId,
        recordId,
        'status',
        statusField,
        { [statusField]: STATUS_FAIL_TEXT },
        cfg,
        fieldWriter,
      );
    }
    return;
  }

  try {
    log(`${prefix}：更新中… 写入「${cfg.attachmentFieldName}」`);
    const payload = { [cfg.attachmentFieldName]: merged };
    await updateRecord(token, recordAppToken, tableId, recordId, payload, cfg);
  } catch (e) {
    stats.fail += 1;
    updateStatsUI(stats);
    const msg = e && e.message ? e.message : String(e);
    log(`${prefix}：更新附件列失败：${msg}`, 'err');
    if (errField) {
      await safeOptionalFieldWriter(
        token,
        recordAppToken,
        tableId,
        recordId,
        'error',
        errField,
        { [errField]: msg.slice(0, 4900) },
        cfg,
        fieldWriter,
      );
    }
    return;
  }

  const hasErr = errors.length > 0;
  if (errField) {
    await safeOptionalFieldWriter(
      token,
      recordAppToken,
      tableId,
      recordId,
      'error',
      errField,
      { [errField]: hasErr ? errors.join(' | ').slice(0, 4900) : '' },
      cfg,
      fieldWriter,
    );
  }
  if (statusField) {
    const st = hasErr ? STATUS_PARTIAL_TEXT : STATUS_SUCCESS_TEXT;
    await safeOptionalFieldWriter(
      token,
      recordAppToken,
      tableId,
      recordId,
      'status',
      statusField,
      { [statusField]: st },
      cfg,
      fieldWriter,
    );
  }

  stats.ok += 1;
  updateStatsUI(stats);
  log(`${prefix}：成功（附件 ${merged.length} 个）${hasErr ? '（部分 URL 失败）' : ''}`);
}

async function startProcess() {
  ensureUI();

  if (__taskRunning) {
    log('已有任务正在执行，请等待结束或点击「停止任务」。可通过「展开执行日志」查看进度。', 'warn');
    if (__panelEl) __panelEl.style.display = 'flex';
    syncRunMiniVisibility();
    return;
  }

  __stopRequested = false;

  const { profile: runProfile, reason: runReason } = resolveMatchedProfileByCurrentUrl();
  touchProfileRunUsage(runProfile.id, runReason);
  const cfg = normalizeStoredConfig(runProfile);

  if (!String(cfg.appId || '').trim() || !String(cfg.appSecret || '').trim()) {
    if (__panelEl && __panelEl.style.display !== 'flex') __panelEl.style.display = 'flex';
    log('未配置 APP_ID / APP_SECRET，已打开配置窗口', 'warn');
    openConfigModal({ runHint: true });
    return;
  }

  __taskRunning = true;
  updateTaskRunningUI();

  const stats = { ok: 0, skip: 0, fail: 0 };
  updateStatsUI(stats);

  if (__panelEl && __panelEl.style.display !== 'flex') __panelEl.style.display = 'flex';
  log(`开始执行：图片链接转附件（档案：${runProfile.name}，匹配：${runReason}）`);

  try {
    const inferred = inferAppAndTableFromUrl();
    const wikiScene = isWikiPage();
    const tableIdGuess = inferred.table_id || cfg.tableId || null;
    const finalViewId = String(cfg.viewId || '').trim() || (inferred.view_id ? String(inferred.view_id).trim() : '');

    const tenant = await getTenantAccessToken(cfg);
    log('已获取 tenant_access_token');

    const recordAppToken = await resolveRecordAppToken(cfg, tenant, inferred, wikiScene);
    const { uploadParentNode, uploadSource } = await resolveActualAppTokenForUpload(cfg, tenant, recordAppToken);

    logResolvedTokens({
      inferredApp: inferred.app_token,
      inferredTable: inferred.table_id,
      inferredView: inferred.view_id,
      configAppToken: cfg.appToken && String(cfg.appToken).trim() ? cfg.appToken : '',
      configTableId: cfg.tableId && String(cfg.tableId).trim() ? cfg.tableId : '',
      configViewId: cfg.viewId && String(cfg.viewId).trim() ? String(cfg.viewId).trim() : '',
      finalViewId,
      recordAppToken,
      uploadParentNode,
      uploadSource,
      wikiScene,
    });

    const finalTable = tableIdGuess;
    if (!finalTable) {
      throw new Error('TABLE_ID 无法从 URL 识别且配置中未填写，请在地址栏包含 table= 或于配置填写 TABLE_ID');
    }

    const { items, total } = await listAllRecords(tenant, recordAppToken, finalTable, cfg, finalViewId || undefined);
    const n = items.length;
    const totalHint = total != null ? String(total) : '未知';
    log(`记录拉取完成：本脚本合并 ${n} 条（接口 total=${totalHint}）`);

    if (!n) {
      log('没有可处理记录，结束');
      return;
    }

    const limit = Math.max(1, cfg.recordConcurrency);
    if (limit > 1) {
      log('记录并发数>1 可能导致飞书写冲突（1254291），已按你的配置执行', 'warn');
    }

    const fieldWriter = {
      statusDisabled: false,
      errorDisabled: false,
      statusNotFoundLogged: false,
      errorNotFoundLogged: false,
    };

    const workIndices = partitionRecordsForProcessing(cfg, items, stats);
    const nWork = workIndices.length;
    if (!nWork) {
      if (!__stopRequested) log('预检后无待处理记录，结束');
      log(
        `全部结束。统计 — 成功: ${stats.ok}，跳过: ${stats.skip}，失败: ${stats.fail}${
          __stopRequested ? '（用户停止）' : ''
        }`,
      );
      return;
    }
    log(`预检后待处理 ${nWork} 条（共拉取 ${n} 条，已跳过 ${n - nWork} 条无需下载的记录）`);

    let i = 0;
    const runners = new Array(Math.min(limit, nWork)).fill(0).map(async () => {
      while (!__stopRequested) {
        const my = i;
        i += 1;
        if (my >= nWork) break;
        const recIdx = workIndices[my];
        await processOneRecord({
          token: tenant,
          recordAppToken,
          uploadParentNode,
          tableId: finalTable,
          record: items[recIdx],
          index: recIdx,
          total: n,
          stats,
          cfg,
          fieldWriter,
        });
        if (__stopRequested) break;
        if (cfg.processIntervalMs > 0) await sleep(cfg.processIntervalMs);
      }
    });
    await Promise.all(runners);

    log(
      `全部结束。统计 — 成功: ${stats.ok}，跳过: ${stats.skip}，失败: ${stats.fail}${
        __stopRequested ? '（用户停止）' : ''
      }`,
    );
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    log(`任务异常终止：${msg}`, 'err');
  } finally {
    __taskRunning = false;
    updateTaskRunningUI();
  }
}

// =============================================================================
// 启动
// =============================================================================

function boot() {
  ensureUI();
  loadProfileStore();
  const { profile } = resolveMatchedProfileByCurrentUrl();
  const c = normalizeStoredConfig(profile);
  if (!String(c.appId || '').trim() || !String(c.appSecret || '').trim()) {
    log('首次使用：请点击「配置」填写 APP_ID 与 APP_SECRET（按「档案」分别保存于本机浏览器）', 'warn');
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', boot, { once: true });
} else {
  boot();
}
