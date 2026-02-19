// ============================================================
// Module: アロン・パナ分析
// ============================================================
// KaientaiM.registerModule() で自己登録するパターン。
// 新モジュール追加時はこのファイルをコピーしてカスタマイズ。
// ============================================================

(function () {
    'use strict';

    const MODULE_ID = 'aron-pana';
    const { util } = KaientaiM;
    const { $, $$, fmt, fmtYen, fmtPct, toNum, toStr, COL, parseExcel, exportCSV, destroyChart, createEl } = util;

    const DEFAULT_SETTINGS = Object.freeze({
        rebateAron: 0,
        rebatePana: 0,
        warehouseFee: 0,
        warehouseOutFee: 50,
        monthlyRebates: {},
        defaultShippingSmall: 100,
        keywordAron: ['アロン'],
        keywordPana: ['パナソニック', 'パナ', 'panasonic']
    });

    // ── Module-local state ──
    const state = {
        shippingData: [],
        salesData: [],
        productData: [],
        progressItems: [],
        progressSeq: 1,
        appliedSettings: cloneDefaultSettings(),
        results: null,
        charts: {},
        storeBaseCache: {},
        storeViewRuntime: null,
        storeDetailRuntime: null,
        storeCurrentPage: 1,
        storeCurrentPageTotal: 1
    };

    let logLines = [];
    let currentTab = 'overview';
    const AUTO_STATE_STORAGE_KEY = 'kaientai-aron-pana-autostate-v1';
    const TAB_UNLOCK_PASSWORD = 'ogura';
    let autoPersistTimer = null;
    let cloudPersistTimer = null;
    let cloudPersistInFlight = false;
    let cloudPersistPending = false;
    let uploadUnlocked = false;
    let settingsUnlocked = false;
    let settingsDirty = false;

    function pfx(id) { return MODULE_ID + '-' + id; }
    const COL_STORE_CODE = COL.C; // C列（仕入先コード）
    const COL_AB = 27; // AB列（都道府県）
    const COL_SALES_REP = COL.Z; // Z列（営業担当）
    const SHIPPING_AREA_COLS = [COL.J, COL.K, COL.L, COL.M, COL.N, COL.O, COL.P, COL.Q, COL.R, COL.S, COL.T, COL.U, COL.V];
    const DEFAULT_AREA_BY_COL = {
        [COL.J]: 'hokkaido',
        [COL.K]: 'kitaTohoku',
        [COL.L]: 'minamiTohoku',
        [COL.M]: 'kanto',
        [COL.N]: 'shinetsu',
        [COL.O]: 'hokuriku',
        [COL.P]: 'chubu',
        [COL.Q]: 'kansai',
        [COL.R]: 'chugoku',
        [COL.S]: 'shikoku',
        [COL.T]: 'kitaKyushu',
        [COL.U]: 'minamiKyushu',
        [COL.V]: 'okinawa',
    };
    const AREA_LABEL = {
        hokkaido: '北海道',
        kitaTohoku: '北東北',
        minamiTohoku: '南東北',
        kanto: '関東',
        shinetsu: '信越',
        hokuriku: '北陸',
        chubu: '中部',
        kansai: '関西',
        chugoku: '中国',
        shikoku: '四国',
        kitaKyushu: '北九州',
        minamiKyushu: '南九州',
        okinawa: '沖縄',
    };

    function normalizeToken(v) {
        return toStr(v).toLowerCase().replace(/[\s　]/g, '');
    }

    function cloneDefaultSettings() {
        return {
            rebateAron: DEFAULT_SETTINGS.rebateAron,
            rebatePana: DEFAULT_SETTINGS.rebatePana,
            warehouseFee: DEFAULT_SETTINGS.warehouseFee,
            warehouseOutFee: DEFAULT_SETTINGS.warehouseOutFee,
            monthlyRebates: {},
            defaultShippingSmall: DEFAULT_SETTINGS.defaultShippingSmall,
            keywordAron: [...DEFAULT_SETTINGS.keywordAron],
            keywordPana: [...DEFAULT_SETTINGS.keywordPana]
        };
    }

    function escapeHTML(v) {
        return toStr(v)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function toAreaKey(areaText) {
        const s = normalizeToken(areaText);
        if (!s) return '';
        if (s.includes('北海道')) return 'hokkaido';
        if (s.includes('北東北')) return 'kitaTohoku';
        if (s.includes('南東北')) return 'minamiTohoku';
        if (s.includes('関東')) return 'kanto';
        if (s.includes('信越')) return 'shinetsu';
        if (s.includes('北陸')) return 'hokuriku';
        if (s.includes('中部')) return 'chubu';
        if (s.includes('関西')) return 'kansai';
        if (s.includes('中国')) return 'chugoku';
        if (s.includes('四国')) return 'shikoku';
        if (s.includes('北九州')) return 'kitaKyushu';
        if (s.includes('南九州')) return 'minamiKyushu';
        if (s.includes('沖縄')) return 'okinawa';
        if (s.includes('九州')) return 'kitaKyushu'; // ざっくり分類
        if (s.includes('東北')) return 'minamiTohoku'; // ざっくり分類
        return '';
    }

    function prefectureToAreaKey(prefecture) {
        const s = normalizeToken(prefecture);
        if (!s) return '';
        const directArea = toAreaKey(s);
        if (directArea) return directArea;
        if (s.includes('北海道')) return 'hokkaido';

        const m = s.match(/(青森|岩手|秋田|宮城|山形|福島|茨城|栃木|群馬|埼玉|千葉|東京|神奈川|山梨|新潟|長野|富山|石川|福井|岐阜|静岡|愛知|三重|滋賀|京都|大阪|兵庫|奈良|和歌山|鳥取|島根|岡山|広島|山口|徳島|香川|愛媛|高知|福岡|佐賀|長崎|大分|熊本|宮崎|鹿児島|沖縄)/);
        if (!m) return '';
        const p = m[1];
        if (['青森', '岩手', '秋田'].includes(p)) return 'kitaTohoku';
        if (['宮城', '山形', '福島'].includes(p)) return 'minamiTohoku';
        if (['茨城', '栃木', '群馬', '埼玉', '千葉', '東京', '神奈川', '山梨'].includes(p)) return 'kanto';
        if (['新潟', '長野'].includes(p)) return 'shinetsu';
        if (['富山', '石川', '福井'].includes(p)) return 'hokuriku';
        if (['岐阜', '静岡', '愛知', '三重'].includes(p)) return 'chubu';
        if (['滋賀', '京都', '大阪', '兵庫', '奈良', '和歌山'].includes(p)) return 'kansai';
        if (['鳥取', '島根', '岡山', '広島', '山口'].includes(p)) return 'chugoku';
        if (['徳島', '香川', '愛媛', '高知'].includes(p)) return 'shikoku';
        if (['福岡', '佐賀', '長崎', '大分'].includes(p)) return 'kitaKyushu';
        if (['熊本', '宮崎', '鹿児島'].includes(p)) return 'minamiKyushu';
        if (p === '沖縄') return 'okinawa';
        return '';
    }

    function buildShippingAreaColumnMap(rows, headerRow) {
        const candidates = [];
        if (rows[1]) candidates.push(rows[1]); // 要件: J2〜V2
        if (rows[headerRow]) candidates.push(rows[headerRow]);
        if (rows[headerRow - 1]) candidates.push(rows[headerRow - 1]);

        for (const candidate of candidates) {
            const map = {};
            let hit = 0;
            for (const col of SHIPPING_AREA_COLS) {
                const key = toAreaKey(candidate[col]);
                if (!key) continue;
                map[col] = key;
                hit++;
            }
            if (hit >= 5) return map;
        }

        const fallback = {};
        for (const col of SHIPPING_AREA_COLS) fallback[col] = DEFAULT_AREA_BY_COL[col];
        return fallback;
    }

    function isOkinawaPrefecture(prefecture) {
        const s = normalizeToken(prefecture);
        return s.includes('沖縄');
    }

    function resolveShippingCost(shipping, prefecture, settings) {
        const areaCosts = shipping.areaCosts || {};
        let areaKey = prefectureToAreaKey(prefecture);
        let shippingCost = 0;
        let fallback = false;

        if (areaKey && areaCosts[areaKey] > 0) {
            shippingCost = areaCosts[areaKey];
        } else {
            const tried = {};
            const order = [areaKey, 'kanto', 'chubu', 'kansai', 'kitaTohoku', 'minamiTohoku', 'hokkaido', 'shinetsu', 'hokuriku', 'chugoku', 'shikoku', 'kitaKyushu', 'minamiKyushu', 'okinawa'];
            for (const key of order) {
                if (!key || tried[key]) continue;
                tried[key] = true;
                if (areaCosts[key] > 0) {
                    shippingCost = areaCosts[key];
                    areaKey = key;
                    fallback = true;
                    break;
                }
            }
        }

        // 沖縄県は特別条件（サイズ帯100以下は既存ロジック優先）
        if (isOkinawaPrefecture(prefecture) && !(shipping.sizeBand > 0 && shipping.sizeBand <= 100)) {
            shippingCost = 3000;
            areaKey = 'okinawa';
            fallback = false;
        }

        // サイズ帯100以下は設定値を優先
        if (shipping.sizeBand > 0 && shipping.sizeBand <= 100 && settings.defaultShippingSmall > 0) {
            shippingCost = settings.defaultShippingSmall;
        }

        return { shippingCost, areaKey, fallback };
    }

    function log(msg) {
        logLines.push(msg);
        const el = document.getElementById(pfx('log-content'));
        if (el) el.textContent = logLines.join('\n');
        const logBox = document.getElementById(pfx('load-log'));
        if (logBox) logBox.style.display = 'block';
    }

    function resetAnalysisOutputs(statusLabel = '読込済（再分析待ち）') {
        state.results = null;
        state.storeBaseCache = {};
        state.storeViewRuntime = null;
        state.storeDetailRuntime = null;
        state.storeCurrentPage = 1;
        state.storeCurrentPageTotal = 1;

        Object.keys(state.charts).forEach(k => destroyChart(state.charts, k));
        ['overview', 'monthly', 'store', 'store-detail', 'sim', 'details'].forEach(id => {
            const emp = document.getElementById(pfx(id + '-empty'));
            const con = document.getElementById(pfx(id + '-content'));
            if (emp) emp.style.display = '';
            if (con) con.style.display = 'none';
        });
        if (statusLabel) KaientaiM.updateModuleStatus(MODULE_ID, statusLabel, false);
    }

    function readProgressDraftInputs() {
        const snapshot = {};
        const prefix = pfx('progress-');
        document.querySelectorAll(`[id^="${prefix}"]`).forEach(el => {
            if (!(el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement || el instanceof HTMLSelectElement)) return;
            const key = el.id.slice(prefix.length);
            if (!key) return;
            snapshot[key] = el.value;
        });
        return snapshot;
    }

    function applyProgressDraftInputs(snapshot) {
        if (!snapshot || typeof snapshot !== 'object') return;
        const prefix = pfx('progress-');
        Object.keys(snapshot).forEach(key => {
            const el = document.getElementById(prefix + key);
            if (!(el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement || el instanceof HTMLSelectElement)) return;
            el.value = toStr(snapshot[key]);
        });
    }

    function normalizeProgressItem(item, fallbackId) {
        const idNum = Number(item?.id);
        const deadline = toStr(item?.deadline || item?.dueDate);
        return {
            id: Number.isFinite(idNum) && idNum > 0 ? idNum : fallbackId,
            rep: toStr(item?.rep),
            customer: toStr(item?.customer),
            actionPlan: toStr(item?.actionPlan),
            result: toStr(item?.result),
            deadline: /^\d{4}-\d{2}-\d{2}$/.test(deadline) ? deadline : '',
            status: toStr(item?.status) || 'planned',
            updatedAt: toStr(item?.updatedAt) || new Date().toISOString()
        };
    }

    function ensureProgressSeq() {
        let maxId = 0;
        for (const item of state.progressItems) {
            const idNum = Number(item?.id);
            if (Number.isFinite(idNum) && idNum > maxId) maxId = idNum;
        }
        state.progressSeq = Math.max(1, maxId + 1);
    }

    function applyTabLockState() {
        const uploadLockText = document.getElementById(pfx('upload-lock-state'));
        if (uploadLockText) {
            uploadLockText.textContent = uploadUnlocked
                ? '解除済み: データ読込を実行できます'
                : 'ロック中: 読込実行にはパスワードが必要です';
        }
        const settingsLockText = document.getElementById(pfx('settings-lock-state'));
        if (settingsLockText) {
            settingsLockText.textContent = settingsUnlocked
                ? '解除済み: 設定を編集できます'
                : 'ロック中: 設定編集にはパスワードが必要です';
        }
        const settingsLockShell = document.getElementById(pfx('settings-lock-shell'));
        if (settingsLockShell) settingsLockShell.classList.toggle('is-locked', !settingsUnlocked);

        const settingsTab = document.getElementById(pfx('tab-settings'));
        if (settingsTab) {
            settingsTab.querySelectorAll('input, select, textarea, button').forEach(el => {
                const id = el.id || '';
                if (id === pfx('btn-settings-unlock')) return;
                if (id === pfx('btn-recalc')) {
                    el.disabled = !settingsUnlocked || state.salesData.length === 0;
                    return;
                }
                el.disabled = !settingsUnlocked;
            });
        }

        updateSettingsSaveState();
        checkAllLoaded();
    }

    function unlockTabGroup(target) {
        const isUpload = target === 'upload';
        const already = isUpload ? uploadUnlocked : settingsUnlocked;
        if (already) return true;
        const label = isUpload ? 'データ読込' : '設定編集';
        const input = window.prompt(`${label}を解除するパスワードを入力してください`);
        if (input === null) return false;
        if (input !== TAB_UNLOCK_PASSWORD) {
            alert('パスワードが違います');
            return false;
        }
        if (isUpload) uploadUnlocked = true;
        else settingsUnlocked = true;
        applyTabLockState();
        return true;
    }

    function ensureUploadUnlocked() {
        return uploadUnlocked || unlockTabGroup('upload');
    }

    function ensureSettingsUnlocked() {
        return settingsUnlocked || unlockTabGroup('settings');
    }

    function normalizeSalesRow(data) {
        const src = (data && typeof data === 'object') ? data : {};
        return {
            orderNo: toStr(src.orderNo ?? src.order_no ?? src['受注番号']),
            month: toStr(src.month || 'unknown'),
            maker: toStr(src.maker || 'other'),
            makerRaw: toStr(src.makerRaw ?? src.maker_raw),
            salesRep: toStr(src.salesRep ?? src.sales_rep ?? src['営業担当']),
            store: toStr(src.store ?? src.storeName ?? src.customer ?? src['販売店']),
            storeCode: toStr(src.storeCode ?? src.store_code ?? src.supplierCode ?? src.customerCode ?? src.dealerCode ?? src['仕入先コード'] ?? src['得意先コード']),
            prefecture: toStr(src.prefecture ?? src['県名']),
            jan: toStr(src.jan ?? src.JAN),
            name: toStr(src.name ?? src.itemName ?? src['商品名']),
            qty: toNum(src.qty),
            unitPrice: toNum(src.unitPrice ?? src.unit_price),
            totalPrice: toNum(src.totalPrice ?? src.total_price)
        };
    }

    function fillMissingSalesStoreCodes(rows) {
        const storeCodeFreq = {};
        for (const row of rows) {
            if (!row.store || !row.storeCode) continue;
            if (!storeCodeFreq[row.store]) storeCodeFreq[row.store] = {};
            storeCodeFreq[row.store][row.storeCode] = (storeCodeFreq[row.store][row.storeCode] || 0) + 1;
        }
        const storeMainCode = {};
        Object.keys(storeCodeFreq).forEach(store => {
            let bestCode = '';
            let bestCount = 0;
            Object.entries(storeCodeFreq[store]).forEach(([code, count]) => {
                if (count > bestCount) {
                    bestCode = code;
                    bestCount = count;
                }
            });
            if (bestCode) storeMainCode[store] = bestCode;
        });
        return rows.map(row => {
            if (row.storeCode) return row;
            const fallback = storeMainCode[row.store];
            return fallback ? { ...row, storeCode: fallback } : row;
        });
    }

    function detectSalesStoreCodeColumn(rows, headerRow) {
        const scoreCol = (row, col) => {
            const t = normalizeToken(row?.[col]);
            if (!t) return 0;
            if (t.includes('仕入先コード')) return 100;
            if (t.includes('得意先コード')) return 95;
            if (t.includes('販売店コード')) return 90;
            if (t.includes('取引先コード')) return 85;
            if (t.includes('仕入先') && t.includes('コード')) return 80;
            if (t.includes('得意先') && t.includes('コード')) return 78;
            if (t.includes('コード')) return 20;
            return 0;
        };

        const candidates = [headerRow, headerRow - 1, 0].filter(v => v >= 0 && v < rows.length);
        let bestCol = COL_STORE_CODE;
        let bestScore = 0;
        for (const rowIdx of candidates) {
            const row = rows[rowIdx] || [];
            for (let col = 0; col < Math.min(50, row.length); col++) {
                const sc = scoreCol(row, col);
                if (sc > bestScore) {
                    bestScore = sc;
                    bestCol = col;
                }
            }
        }
        return bestCol;
    }

    function buildAutoStatePayload() {
        return {
            schemaVersion: 2,
            savedAt: new Date().toISOString(),
            shippingData: state.shippingData,
            salesData: state.salesData,
            productData: state.productData,
            progressItems: state.progressItems,
            progressSeq: state.progressSeq,
            settings: getSettings(),
            progressDraft: readProgressDraftInputs()
        };
    }

    async function persistCloudStateNow() {
        if (!window.KaientaiCloud || typeof window.KaientaiCloud.saveModuleState !== 'function') return false;
        if (!(window.KaientaiCloud.isReady && window.KaientaiCloud.isReady())) return false;
        await window.KaientaiCloud.saveModuleState(MODULE_ID, buildAutoStatePayload());
        return true;
    }

    function flushCloudStateSaveQueue() {
        if (cloudPersistInFlight) {
            cloudPersistPending = true;
            return;
        }
        cloudPersistInFlight = true;
        (async () => {
            try {
                await persistCloudStateNow();
            } catch (err) {
                console.warn('persistCloudStateNow failed', err);
            } finally {
                cloudPersistInFlight = false;
                if (cloudPersistPending) {
                    cloudPersistPending = false;
                    scheduleCloudStateSave(1200);
                }
            }
        })();
    }

    function scheduleCloudStateSave(delay = 1800) {
        if (cloudPersistTimer) clearTimeout(cloudPersistTimer);
        cloudPersistTimer = setTimeout(() => {
            cloudPersistTimer = null;
            flushCloudStateSaveQueue();
        }, Math.max(200, delay));
    }

    function saveAutoStateNow() {
        try {
            localStorage.setItem(AUTO_STATE_STORAGE_KEY, JSON.stringify(buildAutoStatePayload()));
        } catch (err) {
            console.warn('saveAutoStateNow failed', err);
        }
        scheduleCloudStateSave(1200);
    }

    function scheduleAutoStateSave(delay = 600) {
        if (autoPersistTimer) clearTimeout(autoPersistTimer);
        autoPersistTimer = setTimeout(() => {
            autoPersistTimer = null;
            saveAutoStateNow();
        }, Math.max(0, delay));
    }

    function applyLoadedPayload(payload, sourceLabel) {
        if (!payload || !Array.isArray(payload.shippingData) || !Array.isArray(payload.salesData) || !Array.isArray(payload.productData)) {
            return false;
        }

        state.shippingData = payload.shippingData;
        state.salesData = fillMissingSalesStoreCodes((payload.salesData || []).map(normalizeSalesRow));
        state.productData = payload.productData;
        state.progressItems = Array.isArray(payload.progressItems)
            ? payload.progressItems.map((item, idx) => normalizeProgressItem(item, idx + 1))
            : [];
        state.progressSeq = toNum(payload.progressSeq) || 1;
        ensureProgressSeq();
        resetAnalysisOutputs('データ復元済（再分析待ち）');
        restoreSavedSettings(payload.settings || {});
        applyProgressDraftInputs(payload.progressDraft || {});
        renderProgressFormSelectors();
        renderProgressTable();
        updateUploadCardsByState();
        checkAllLoaded();
        if (sourceLabel) log(sourceLabel);
        if (state.shippingData.length > 0 && state.salesData.length > 0 && state.productData.length > 0) {
            log('--- 自動再分析開始 ---');
            runAnalysis();
            if (currentTab !== 'upload' && currentTab !== 'settings') switchModTab(currentTab || 'overview');
            log('--- 自動再分析完了 ---');
        }
        return true;
    }

    function restoreAutoState() {
        try {
            const raw = localStorage.getItem(AUTO_STATE_STORAGE_KEY);
            if (!raw) return false;
            const payload = JSON.parse(raw);
            return applyLoadedPayload(payload, 'ローカル自動復元完了');
        } catch (err) {
            console.warn('restoreAutoState failed', err);
            return false;
        }
    }

    async function restoreCloudStateIfNeeded() {
        if (state.shippingData.length > 0 || state.salesData.length > 0 || state.productData.length > 0) return false;
        if (!window.KaientaiCloud || typeof window.KaientaiCloud.loadModuleState !== 'function') return false;
        if (!(window.KaientaiCloud.isReady && window.KaientaiCloud.isReady())) return false;
        try {
            const payload = await window.KaientaiCloud.loadModuleState(MODULE_ID);
            if (!payload) return false;
            const ok = applyLoadedPayload(payload, 'クラウド自動復元完了');
            if (ok) saveAutoStateNow();
            return ok;
        } catch (err) {
            console.warn('restoreCloudStateIfNeeded failed', err);
            return false;
        }
    }

    // ── Settings ──
    function getSalesMonthList() {
        const months = [...new Set(state.salesData.map(s => s.month).filter(m => m && m !== 'unknown'))].sort();
        if (months.length > 0) return months;
        const d = new Date();
        return [d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')];
    }

    function normalizeKeywordList(input, fallback) {
        const src = Array.isArray(input) ? input : toStr(input).split(',');
        const values = src.map(v => toStr(v).trim().toLowerCase()).filter(Boolean);
        return values.length > 0 ? values : [...fallback];
    }

    function normalizeMonthlyRebates(raw) {
        const src = raw && typeof raw === 'object' ? raw : {};
        const out = {};
        Object.keys(src).sort().forEach(month => {
            const m = src[month] || {};
            const aron = m.aron || {};
            const pana = m.pana || {};
            const normalized = {
                aron: { achieve: toNum(aron.achieve), car: toNum(aron.car) },
                pana: { achieve: toNum(pana.achieve), car: toNum(pana.car) }
            };
            const sum = normalized.aron.achieve + normalized.aron.car + normalized.pana.achieve + normalized.pana.car;
            if (sum !== 0) out[month] = normalized;
        });
        return out;
    }

    function pickNumber(value, fallback) {
        if (value === '' || value == null) return fallback;
        const n = Number(value);
        return Number.isFinite(n) ? n : fallback;
    }

    function normalizeSettings(raw) {
        const src = raw && typeof raw === 'object' ? raw : {};
        return {
            rebateAron: Math.max(0, pickNumber(src.rebateAron, DEFAULT_SETTINGS.rebateAron)),
            rebatePana: Math.max(0, pickNumber(src.rebatePana, DEFAULT_SETTINGS.rebatePana)),
            warehouseFee: Math.max(0, pickNumber(src.warehouseFee, DEFAULT_SETTINGS.warehouseFee)),
            warehouseOutFee: Math.max(0, pickNumber(src.warehouseOutFee, DEFAULT_SETTINGS.warehouseOutFee)),
            monthlyRebates: normalizeMonthlyRebates(src.monthlyRebates),
            defaultShippingSmall: Math.max(0, pickNumber(src.defaultShippingSmall, DEFAULT_SETTINGS.defaultShippingSmall)),
            keywordAron: normalizeKeywordList(src.keywordAron, DEFAULT_SETTINGS.keywordAron),
            keywordPana: normalizeKeywordList(src.keywordPana, DEFAULT_SETTINGS.keywordPana)
        };
    }

    function serializeSettings(settings) {
        return JSON.stringify(normalizeSettings(settings));
    }

    function markSettingsDirty(flag) {
        settingsDirty = !!flag;
        updateSettingsSaveState();
    }

    function updateSettingsSaveState() {
        const el = document.getElementById(pfx('settings-save-state'));
        if (!el) return;
        el.classList.remove('dirty', 'locked');
        if (!settingsUnlocked) {
            el.classList.add('locked');
            el.textContent = 'ロック中';
            return;
        }
        if (settingsDirty) {
            el.classList.add('dirty');
            el.textContent = '未保存の変更があります';
            return;
        }
        el.textContent = '保存済み設定を使用中';
    }

    function renderMonthlyRebateInputs(seedMonthlyRebates) {
        const container = document.getElementById(pfx('monthly-rebate-body'));
        if (!container) return;

        const oldValues = {};
        container.querySelectorAll('input[data-month][data-maker][data-type]').forEach(input => {
            const k = `${input.dataset.month}|${input.dataset.maker}|${input.dataset.type}`;
            oldValues[k] = toNum(input.value);
        });
        const normalizedSeed = normalizeMonthlyRebates(seedMonthlyRebates);

        const months = getSalesMonthList();
        container.innerHTML = '';
        for (const month of months) {
            const keys = [
                { maker: 'aron', type: 'achieve', label: 'アロン 達成リベート金額 (円)' },
                { maker: 'aron', type: 'car', label: 'アロン 車扱い還元金 (円)' },
                { maker: 'pana', type: 'achieve', label: 'パナ 達成リベート金額 (円)' },
                { maker: 'pana', type: 'car', label: 'パナ 車扱い還元金 (円)' },
            ];
            const fields = keys.map(k => {
                const key = `${month}|${k.maker}|${k.type}`;
                const hasOld = Object.prototype.hasOwnProperty.call(oldValues, key);
                const val = hasOld
                    ? oldValues[key]
                    : toNum(normalizedSeed?.[month]?.[k.maker]?.[k.type]);
                return `<label class="monthly-rebate-field"><span>${k.label}</span><input type="number" class="monthly-rebate-input" data-month="${month}" data-maker="${k.maker}" data-type="${k.type}" value="${val}" step="1000"></label>`;
            }).join('');
            const row = document.createElement('div');
            row.className = 'monthly-rebate-row';
            row.innerHTML = `<div class="monthly-rebate-month">${month}</div><div class="monthly-rebate-grid">${fields}</div>`;
            container.appendChild(row);
        }
    }

    function readMonthlyRebateSettings() {
        const monthlyRebates = {};
        const tbody = document.getElementById(pfx('monthly-rebate-body'));
        if (!tbody) return monthlyRebates;

        tbody.querySelectorAll('input[data-month][data-maker][data-type]').forEach(input => {
            const month = input.dataset.month;
            const maker = input.dataset.maker;
            const type = input.dataset.type;
            if (!month || !maker || !type) return;
            if (!monthlyRebates[month]) {
                monthlyRebates[month] = {
                    aron: { achieve: 0, car: 0 },
                    pana: { achieve: 0, car: 0 }
                };
            }
            if (!monthlyRebates[month][maker]) monthlyRebates[month][maker] = { achieve: 0, car: 0 };
            monthlyRebates[month][maker][type] = toNum(input.value);
        });
        return normalizeMonthlyRebates(monthlyRebates);
    }

    function getMonthlyRebate(settings, month, maker) {
        if (maker !== 'aron' && maker !== 'pana') return { achieve: 0, car: 0, fixed: 0 };
        const m = settings.monthlyRebates?.[month]?.[maker] || {};
        const achieve = toNum(m.achieve);
        const car = toNum(m.car);
        return { achieve, car, fixed: achieve + car };
    }

    function calcMonthlyRebate(entry, settings) {
        const rate = entry.maker === 'aron' ? settings.rebateAron : entry.maker === 'pana' ? settings.rebatePana : 0;
        const variable = entry.sales * rate;
        const fixed = getMonthlyRebate(settings, entry.month, entry.maker);
        return {
            variable,
            achieve: fixed.achieve,
            car: fixed.car,
            fixed: fixed.fixed,
            total: variable + fixed.fixed
        };
    }

    function calcMonthlyMinus(entry, settings, monthSalesTotals) {
        const monthSales = monthSalesTotals?.[entry.month] || 0;
        const warehouseBase = monthSales > 0 ? settings.warehouseFee * (entry.sales / monthSales) : 0;
        const warehouseOut = entry.qty * settings.warehouseOutFee;
        return { warehouseBase, warehouseOut, total: warehouseBase + warehouseOut };
    }

    function readSettingsFromForm() {
        return normalizeSettings({
            rebateAron: toNum(document.getElementById(pfx('rebate-aron'))?.value) / 100,
            rebatePana: toNum(document.getElementById(pfx('rebate-pana'))?.value) / 100,
            warehouseFee: toNum(document.getElementById(pfx('warehouse-fee'))?.value),
            warehouseOutFee: toNum(document.getElementById(pfx('warehouse-out-fee'))?.value),
            monthlyRebates: readMonthlyRebateSettings(),
            defaultShippingSmall: toNum(document.getElementById(pfx('default-shipping-small'))?.value),
            keywordAron: document.getElementById(pfx('keyword-aron'))?.value,
            keywordPana: document.getElementById(pfx('keyword-pana'))?.value
        });
    }

    function applySettingsToForm(settings) {
        const normalized = normalizeSettings(settings);
        setInputValue('rebate-aron', normalized.rebateAron * 100);
        setInputValue('rebate-pana', normalized.rebatePana * 100);
        setInputValue('warehouse-fee', normalized.warehouseFee);
        setInputValue('warehouse-out-fee', normalized.warehouseOutFee);
        setInputValue('default-shipping-small', normalized.defaultShippingSmall);
        setInputValue('keyword-aron', normalized.keywordAron.join(','));
        setInputValue('keyword-pana', normalized.keywordPana.join(','));
        renderMonthlyRebateInputs(normalized.monthlyRebates);
        const wrap = document.getElementById(pfx('monthly-rebate-body'));
        if (!wrap) return;
        wrap.querySelectorAll('input[data-month][data-maker][data-type]').forEach(input => {
            const month = input.dataset.month;
            const maker = input.dataset.maker;
            const type = input.dataset.type;
            input.value = toNum(normalized.monthlyRebates?.[month]?.[maker]?.[type]);
        });
    }

    function getSettings() {
        return normalizeSettings(state.appliedSettings || cloneDefaultSettings());
    }

    function saveSettingsFromForm() {
        if (!ensureSettingsUnlocked()) return false;
        const next = readSettingsFromForm();
        state.appliedSettings = next;
        markSettingsDirty(false);
        scheduleAutoStateSave(120);
        updateSettingsSaveState();
        return true;
    }

    function detectMaker(text) {
        // S列の値やその他テキストからメーカー判定
        const s = getSettings();
        const t = (text || '').toLowerCase().replace(/[\s（）\(\)株]+/g, '');
        for (const kw of s.keywordAron) if (t.includes(kw.replace(/[\s（）\(\)株]+/g, ''))) return 'aron';
        for (const kw of s.keywordPana) if (t.includes(kw.replace(/[\s（）\(\)株]+/g, ''))) return 'pana';
        return 'other';
    }

    // Excelの日付値(シリアル値 or 文字列)から yyyy-MM を抽出
    function extractMonthFromDate(val) {
        if (val == null || val === '') return null;

        // 1) Excelシリアル値（数値）→ JSの日付に変換
        if (typeof val === 'number' && val > 30000 && val < 100000) {
            // Excel日付シリアル: 1900/1/1 = 1, 1900/2/28 の次が 1900/3/1 (Excelバグ互換)
            const epoch = new Date(1899, 11, 30); // 1899-12-30
            const d = new Date(epoch.getTime() + val * 86400000);
            const y = d.getFullYear();
            const m = d.getMonth() + 1;
            if (y >= 2000 && y <= 2099) return y + '-' + String(m).padStart(2, '0');
        }

        const s = String(val).trim();

        // 2) yyyy/mm/dd, yyyy-mm-dd
        let match = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-]\d{1,2}/);
        if (match) return match[1] + '-' + match[2].padStart(2, '0');

        // 3) yyyy年mm月
        match = s.match(/(\d{4})\s*年\s*(\d{1,2})\s*月/);
        if (match) return match[1] + '-' + match[2].padStart(2, '0');

        // 4) mm/dd/yyyy (US形式)
        match = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
        if (match) return match[3] + '-' + match[1].padStart(2, '0');

        // 5) Date型として解析を試みる
        const d = new Date(s);
        if (!isNaN(d.getTime()) && d.getFullYear() >= 2000) {
            return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
        }

        return null;
    }

    function extractMonthFromFileName(fileName) {
        let m;
        m = fileName.match(/(\d{4})\s*年\s*(\d{1,2})\s*月/);
        if (m) return m[1] + '-' + m[2].padStart(2, '0');
        m = fileName.match(/(\d{4})[-\/](\d{1,2})/);
        if (m) return m[1] + '-' + m[2].padStart(2, '0');
        m = fileName.match(/(\d{4})(\d{2})/);
        if (m && parseInt(m[2]) >= 1 && parseInt(m[2]) <= 12) return m[1] + '-' + m[2];
        return null;
    }

    // ── Data Loading ──
    function findHeaderRow(rows, keywords) {
        for (let i = 0; i < Math.min(rows.length, 15); i++) {
            const rowText = (rows[i] || []).map(c => toStr(c).toLowerCase()).join(' ');
            if (keywords.some(kw => rowText.includes(kw))) return i;
        }
        return 0;
    }

    // 複数シートから最適なデータシートを自動検出
    function findBestSheet(parsed, preferNames, minCols, janCol) {
        // 1. シート名が一致するものを優先
        for (const pref of preferNames) {
            const found = parsed.sheetNames.find(n => n.includes(pref));
            if (found) {
                const rows = parsed.sheets[found];
                // 最低限のデータ行があるか確認
                if (rows.length > 5) {
                    log(`  → シート名一致で「${found}」を選択`);
                    return found;
                }
            }
        }
        // 2. 列数が十分でJANっぽいデータがあるシートを探す
        let bestSheet = null, bestScore = 0;
        for (const name of parsed.sheetNames) {
            const rows = parsed.sheets[name];
            if (rows.length < 3) continue;
            let score = 0;
            score += rows.length; // 行数が多いほど良い
            // 数行チェックしてJAN列にデータがあるか
            let janHits = 0;
            for (let i = 1; i < Math.min(rows.length, 20); i++) {
                const row = rows[i] || [];
                if (row.length >= minCols) score += 5;
                const val = toStr(row[janCol]);
                if (val && /^\d{8,13}$/.test(val)) janHits++;
            }
            score += janHits * 10;
            if (score > bestScore) { bestScore = score; bestSheet = name; }
        }
        if (bestSheet) log(`  → データ内容分析で「${bestSheet}」を選択 (スコア=${bestScore})`);
        return bestSheet || parsed.sheetNames[0];
    }

    function loadShipping(parsed) {
        log(`--- 送料マスタ読込開始: ${parsed.fileName} ---`);
        log(`シート数: ${parsed.sheetNames.length} → [${parsed.sheetNames.join(', ')}]`);

        // 「商品」シートを優先的に探す
        const sheetName = findBestSheet(parsed, ['商品'], 10, COL.A);
        const rows = parsed.sheets[sheetName];
        log(`使用シート: 「${sheetName}」 / 総行数: ${rows.length}`);

        // 先頭5行をダンプして構造を可視化
        for (let i = 0; i < Math.min(rows.length, 5); i++) {
            const r = rows[i] || [];
            const preview = `  行${i}: A=[${toStr(r[COL.A])}] B=[${toStr(r[COL.B])}] J=[${toStr(r[COL.J])}] V=[${toStr(r[COL.V])}] (列数=${r.length})`;
            log(preview);
        }

        // ヘッダー行を自動検出
        const headerRow = findHeaderRow(rows, ['jan', 'janコード', '商品', 'コード', '品番']);
        log(`ヘッダー行検出: ${headerRow}行目 → データは${headerRow + 1}行目から`);

        const areaColMap = buildShippingAreaColumnMap(rows, headerRow);
        const areaColsText = SHIPPING_AREA_COLS.map(col => AREA_LABEL[areaColMap[col]] || '-').join(' / ');
        log(`  エリア列(J-V): [${areaColsText}]`);

        state.shippingData = [];
        let skipped = 0;
        for (let i = headerRow + 1; i < rows.length; i++) {
            const row = rows[i] || [];
            const jan = toStr(row[COL.A]);
            if (!jan) { skipped++; continue; }
            const areaCosts = {};
            for (const col of SHIPPING_AREA_COLS) {
                const areaKey = areaColMap[col];
                if (!areaKey) continue;
                const cost = toNum(row[col]);
                if (cost > 0) areaCosts[areaKey] = cost;
            }
            state.shippingData.push({
                jan,
                name: toStr(row[COL.B]),
                sizeBand: toNum(row[COL.I]),
                areaCosts
            });
        }
        log(`送料マスタ: ${state.shippingData.length}件読込 (スキップ: ${skipped}件, ヘッダー後空行含む)`);

        // 読込んだデータのサンプルを表示
        if (state.shippingData.length > 0) {
            const s = state.shippingData[0];
            const keys = Object.keys(s.areaCosts || {});
            const firstKey = keys[0];
            const areaSample = firstKey ? `${AREA_LABEL[firstKey] || firstKey}=${fmtYen(s.areaCosts[firstKey])}` : 'なし';
            log(`  サンプル: JAN=[${s.jan}] 商品名=[${s.name}] サイズ帯=[${s.sizeBand}] エリア送料=[${areaSample}]`);
        }
        if (state.shippingData.length > 1) {
            const s = state.shippingData[1];
            const keys = Object.keys(s.areaCosts || {});
            const firstKey = keys[0];
            const areaSample = firstKey ? `${AREA_LABEL[firstKey] || firstKey}=${fmtYen(s.areaCosts[firstKey])}` : 'なし';
            log(`  サンプル: JAN=[${s.jan}] 商品名=[${s.name}] サイズ帯=[${s.sizeBand}] エリア送料=[${areaSample}]`);
        }

        document.getElementById(pfx('status-shipping')).textContent = `✓ ${state.shippingData.length}件`;
        document.getElementById(pfx('card-shipping')).classList.add('loaded');
        resetAnalysisOutputs('送料更新（再分析待ち）');
        scheduleAutoStateSave();
    }

    function loadSales(parsedList) {
        const normalizeOrderNo = (orderNo) => toStr(orderNo).trim();
        const makerValues = new Set(state.salesData.map(s => s.makerRaw).filter(Boolean));
        const monthValues = new Set(state.salesData.map(s => s.month).filter(Boolean));
        const seenOrderNos = new Set(state.salesData.map(s => normalizeOrderNo(s.orderNo)).filter(Boolean));
        let duplicateOrderCount = 0;
        let addedCount = 0;

        for (const parsed of parsedList) {
            log(`--- 販売実績読込: ${parsed.fileName} ---`);
            const fileMonth = extractMonthFromFileName(parsed.fileName); // フォールバック用
            for (const sheetName of parsed.sheetNames) {
                const rows = parsed.sheets[sheetName];
                log(`  シート[${sheetName}]: ${rows.length}行`);

                // 先頭3行ダンプ（主要列を表示）
                for (let i = 0; i < Math.min(rows.length, 3); i++) {
                    const r = rows[i] || [];
                    log(`    行${i}: A=[${toStr(r[COL.A])}] C=[${toStr(r[COL_STORE_CODE])}] B=[${toStr(r[COL.B])}] D=[${toStr(r[COL.D])}] H=[${toStr(r[COL.H])}] I=[${toStr(r[COL.I])}] K=[${toStr(r[COL.K])}] L=[${toStr(r[COL.L])}] S=[${toStr(r[COL.S])}] Z=[${toStr(r[COL_SALES_REP])}] AB=[${toStr(r[COL_AB])}]`);
                }

                const headerRow = findHeaderRow(rows, ['jan', 'janコード', '商品', 'コード', '品番', '数量', '販売', '受注']);
                log(`  ヘッダー行: ${headerRow}行目`);
                const storeCodeCol = detectSalesStoreCodeColumn(rows, headerRow);
                log(`  仕入先コード列: ${storeCodeCol + 1}列目`);

                let count = 0, dateOk = 0, dateFail = 0;
                for (let i = headerRow + 1; i < rows.length; i++) {
                    const row = rows[i] || [];
                    const jan = toStr(row[COL.H]);
                    if (!jan) continue;
                    const orderNo = normalizeOrderNo(row[COL.A]);

                    // B列の受注日から月を判定（最優先）
                    const rawDate = row[COL.B];
                    let month = extractMonthFromDate(rawDate);
                    if (month) {
                        dateOk++;
                    } else {
                        dateFail++;
                        month = fileMonth || 'unknown'; // B列から取れなければファイル名フォールバック
                    }
                    monthValues.add(month);

                    // S列からメーカー判定
                    const makerRaw = toStr(row[COL.S]);
                    if (makerRaw) makerValues.add(makerRaw);
                    const maker = makerRaw ? detectMaker(makerRaw) : 'other';

                    const name = toStr(row[COL.I]);
                    const store = toStr(row[COL.D]);
                    const storeCode = toStr(row[storeCodeCol]) || toStr(row[COL_STORE_CODE]);
                    const qty = toNum(row[COL.K]);
                    const unitPrice = toNum(row[COL.L]);
                    if (orderNo && seenOrderNos.has(orderNo)) {
                        duplicateOrderCount++;
                        continue;
                    }
                    if (orderNo) seenOrderNos.add(orderNo);
                    const salesRep = toStr(row[COL_SALES_REP]);
                    state.salesData.push({
                        orderNo,
                        month,
                        maker,
                        makerRaw,
                        salesRep,
                        store,
                        storeCode,
                        prefecture: toStr(row[COL_AB]),
                        jan, name,
                        qty,
                        unitPrice,
                        totalPrice: toNum(row[COL.M])
                    });
                    addedCount++;
                    count++;
                }
                log(`  → ${count}件読込 (B列日付OK=${dateOk}, B列日付NG=${dateFail})`);
            }
        }

        const beforeFillCodeCount = state.salesData.filter(s => toStr(s.storeCode)).length;
        state.salesData = fillMissingSalesStoreCodes(state.salesData.map(normalizeSalesRow));
        const afterFillCodeCount = state.salesData.filter(s => toStr(s.storeCode)).length;

        // 診断情報
        log(`検出月一覧: [${[...monthValues].sort().join(', ')}]`);
        log(`S列メーカー表記一覧: [${[...makerValues].join(' / ')}]`);
        const aronCount = state.salesData.filter(s => s.maker === 'aron').length;
        const panaCount = state.salesData.filter(s => s.maker === 'pana').length;
        const otherCount = state.salesData.filter(s => s.maker === 'other').length;
        log(`メーカー判定結果: アロン=${aronCount}件 / パナ=${panaCount}件 / その他=${otherCount}件`);
        log(`仕入先コード: 取込直後=${beforeFillCodeCount}件 / 補完後=${afterFillCodeCount}件 / 未設定=${state.salesData.length - afterFillCodeCount}件`);
        log(`販売実績追加: ${addedCount}件 / 重複受注番号スキップ: ${duplicateOrderCount}件 / 累計: ${state.salesData.length}件`);

        document.getElementById(pfx('status-sales')).textContent = `✓ ${state.salesData.length}件 (+${addedCount}件 / ${parsedList.length}ファイル)`;
        document.getElementById(pfx('card-sales')).classList.add('loaded');
        renderMonthlyRebateInputs(state.appliedSettings?.monthlyRebates || {});
        renderProgressFormSelectors();
        resetAnalysisOutputs('販売実績更新（再分析待ち）');
        scheduleAutoStateSave();
    }

    function loadProduct(parsedList) {
        const files = Array.isArray(parsedList) ? parsedList : [parsedList];
        state.productData = [];
        const productMap = new Map();
        let skipped = 0;
        let overwriteCount = 0;

        for (const parsed of files) {
            log(`--- Product master load: ${parsed.fileName} ---`);
            log(`Sheets: ${parsed.sheetNames.length} / [${parsed.sheetNames.join(', ')}]`);

            const sheetName = findBestSheet(parsed, [], 10, COL.A);
            const rows = parsed.sheets[sheetName];
            log(`Using sheet: ${sheetName} / rows: ${rows.length}`);

            for (let i = 0; i < Math.min(rows.length, 5); i++) {
                const r = rows[i] || [];
                log(`  Row ${i}: A=[${toStr(r[COL.A])}] D=[${toStr(r[COL.D])}] H=[${toStr(r[COL.H])}] M=[${toStr(r[COL.M])}] O=[${toStr(r[COL.O])}]`);
            }

            const headerRow = findHeaderRow(rows, ['jan', 'code', 'item', 'cost', 'price']);
            log(`Header row: ${headerRow}`);

            let fileCount = 0;
            for (let i = headerRow + 1; i < rows.length; i++) {
                const row = rows[i] || [];
                const jan = toStr(row[COL.A]);
                if (!jan) { skipped++; continue; }
                const wc = toNum(row[COL.O]);
                const cost = toNum(row[COL.M]);
                if (productMap.has(jan)) overwriteCount++;
                productMap.set(jan, {
                    jan, name: toStr(row[COL.D]), listPrice: toNum(row[COL.H]),
                    cost, warehouseCost: wc, effectiveCost: wc > 0 ? wc : cost
                });
                fileCount++;
            }
            log(`  ${parsed.fileName}: imported ${fileCount} rows`);
        }

        state.productData = Array.from(productMap.values());
        log(`Product master loaded: ${state.productData.length} rows (files: ${files.length}, skipped: ${skipped}, duplicate JAN overwritten: ${overwriteCount})`);

        if (state.productData.length > 0) {
            const sample = state.productData[0];
            log(`  Sample: JAN=[${sample.jan}] name=[${sample.name}] list=[${sample.listPrice}] cost=[${sample.effectiveCost}]`);
        }

        document.getElementById(pfx('status-product')).textContent = `Loaded: ${state.productData.length} (${files.length} files)`;
        document.getElementById(pfx('card-product')).classList.add('loaded');
        resetAnalysisOutputs('商品マスタ更新（再分析待ち）');
        scheduleAutoStateSave();
    }

    // ── Analysis ──
    function runAnalysis() {
        const settings = getSettings();
        const shippingMap = {};
        for (const s of state.shippingData) shippingMap[s.jan] = s;
        const productMap = {};
        for (const p of state.productData) productMap[p.jan] = p;

        const records = [];
        let matchCount = 0, noShipping = 0, noProduct = 0, excludedCount = 0;
        let noPrefArea = 0, areaFallback = 0, zeroAreaCost = 0;

        for (const sale of state.salesData) {
            const shipping = shippingMap[sale.jan];
            const product = productMap[sale.jan];
            if (!shipping) noShipping++;
            if (!product) noProduct++;
            if (!shipping || !product) { excludedCount++; continue; }
            matchCount++;

            const shipCalc = resolveShippingCost(shipping, sale.prefecture, settings);
            const shippingCost = shipCalc.shippingCost;
            if (!shipCalc.areaKey) noPrefArea++;
            if (shipCalc.fallback) areaFallback++;
            if (shippingCost <= 0) zeroAreaCost++;
            const effectiveCost = product.effectiveCost;
            const listPrice = product.listPrice;
            const salesAmount = sale.unitPrice * sale.qty;
            const totalShipping = sale.qty * shippingCost;
            const totalCost = effectiveCost * sale.qty;
            const grossProfit = salesAmount - totalCost - totalShipping;

            records.push({
                ...sale, shippingCost, shippingArea: shipCalc.areaKey, effectiveCost, listPrice,
                salesAmount, totalShipping, totalCost, grossProfit,
                rateVsList: listPrice > 0 ? sale.unitPrice / listPrice : 0
            });
        }

        log(`マッチング: ${matchCount}件一致(3データ一致) / 除外: ${excludedCount} / 送料未一致: ${noShipping} / 商品マスタ未一致: ${noProduct} / 地域判定不可: ${noPrefArea} / エリア補完: ${areaFallback} / 送料0円: ${zeroAreaCost}`);

        const monthlyAgg = {}, storeAgg = {}, storeSliceAgg = {}, productAgg = {};
        let totalSales = 0, totalCost = 0, totalShipping = 0, totalGross = 0, totalQty = 0;
        let aronSales = 0, panaSales = 0;

        for (const r of records) {
            totalSales += r.salesAmount;
            totalCost += r.totalCost;
            totalShipping += r.totalShipping;
            totalGross += r.grossProfit;
            totalQty += r.qty;
            if (r.maker === 'aron') aronSales += r.salesAmount;
            if (r.maker === 'pana') panaSales += r.salesAmount;

            const mk = r.month + '|' + r.maker;
            if (!monthlyAgg[mk]) monthlyAgg[mk] = { month: r.month, maker: r.maker, sales: 0, cost: 0, shipping: 0, gross: 0, qty: 0 };
            monthlyAgg[mk].sales += r.salesAmount;
            monthlyAgg[mk].cost += r.totalCost;
            monthlyAgg[mk].shipping += r.totalShipping;
            monthlyAgg[mk].gross += r.grossProfit;
            monthlyAgg[mk].qty += r.qty;

            const sk = r.store + '|' + r.maker;
            if (!storeAgg[sk]) storeAgg[sk] = { store: r.store, maker: r.maker, sales: 0, cost: 0, shipping: 0, gross: 0, qty: 0 };
            storeAgg[sk].sales += r.salesAmount;
            storeAgg[sk].cost += r.totalCost;
            storeAgg[sk].shipping += r.totalShipping;
            storeAgg[sk].gross += r.grossProfit;
            storeAgg[sk].qty += r.qty;

            const ssKey = [r.store, r.storeCode || '', r.salesRep || '', r.month, r.maker].join('\t');
            if (!storeSliceAgg[ssKey]) {
                storeSliceAgg[ssKey] = {
                    store: r.store || '(不明)',
                    storeCode: r.storeCode || '',
                    salesRep: r.salesRep || '',
                    month: r.month,
                    maker: r.maker,
                    sales: 0,
                    cost: 0,
                    shipping: 0,
                    gross: 0,
                    qty: 0,
                    aronRateNumerator: 0,
                    aronRateDenominator: 0,
                    panaRateNumerator: 0,
                    panaRateDenominator: 0
                };
            }
            const ss = storeSliceAgg[ssKey];
            ss.sales += r.salesAmount;
            ss.cost += r.totalCost;
            ss.shipping += r.totalShipping;
            ss.gross += r.grossProfit;
            ss.qty += r.qty;
            if (r.listPrice > 0 && r.qty > 0) {
                if (r.maker === 'aron') {
                    ss.aronRateNumerator += r.unitPrice * r.qty;
                    ss.aronRateDenominator += r.listPrice * r.qty;
                } else if (r.maker === 'pana') {
                    ss.panaRateNumerator += r.unitPrice * r.qty;
                    ss.panaRateDenominator += r.listPrice * r.qty;
                }
            }

            if (!productAgg[r.jan]) productAgg[r.jan] = {
                jan: r.jan, name: r.name, maker: r.maker, listPrice: r.listPrice,
                effectiveCost: r.effectiveCost, shippingCost: r.shippingCost,
                sales: 0, cost: 0, shipping: 0, gross: 0, qty: 0, priceSum: 0, priceCount: 0
            };
            productAgg[r.jan].sales += r.salesAmount;
            productAgg[r.jan].cost += r.totalCost;
            productAgg[r.jan].shipping += r.totalShipping;
            productAgg[r.jan].gross += r.grossProfit;
            productAgg[r.jan].qty += r.qty;
            if (r.unitPrice > 0) { productAgg[r.jan].priceSum += r.unitPrice; productAgg[r.jan].priceCount++; }
        }

        const months = [...new Set(records.map(r => r.month))].sort();
        const monthCount = months.length;
        for (const month of months) {
            for (const maker of ['aron', 'pana']) {
                const key = month + '|' + maker;
                if (!monthlyAgg[key] && getMonthlyRebate(settings, month, maker).fixed > 0) {
                    monthlyAgg[key] = { month, maker, sales: 0, cost: 0, shipping: 0, gross: 0, qty: 0 };
                }
            }
        }

        const monthSalesTotals = {};
        for (const e of Object.values(monthlyAgg)) {
            monthSalesTotals[e.month] = (monthSalesTotals[e.month] || 0) + e.sales;
        }

        const rebateByMaker = { aron: 0, pana: 0, other: 0 };
        const minusByMaker = { aron: 0, pana: 0, other: 0 };
        let totalRebate = 0;
        for (const e of Object.values(monthlyAgg)) {
            const rb = calcMonthlyRebate(e, settings).total;
            const minus = calcMonthlyMinus(e, settings, monthSalesTotals).total;
            rebateByMaker[e.maker] = (rebateByMaker[e.maker] || 0) + rb;
            minusByMaker[e.maker] = (minusByMaker[e.maker] || 0) + minus;
            totalRebate += rb;
        }

        const totalWarehouse = settings.warehouseFee * monthCount;
        const totalWarehouseOut = totalQty * settings.warehouseOutFee;
        const totalMinus = totalWarehouse + totalWarehouseOut;
        const realProfit = totalGross + totalRebate - totalMinus;

        state.results = {
            records, monthlyAgg, storeAgg, storeSlices: Object.values(storeSliceAgg), productAgg, months,
            totalSales, totalCost, totalShipping, totalGross, totalQty,
            totalRebate, totalWarehouse, totalWarehouseOut, totalMinus, realProfit,
            aronSales, panaSales, settings, monthSalesTotals, rebateByMaker, minusByMaker
        };
        state.storeBaseCache = { 'all|all': buildStoreBase(records) };
        state.storeViewRuntime = null;
        state.storeCurrentPage = 1;
        state.storeCurrentPageTotal = 1;

        log(`分析完了: 売上 ${fmtYen(totalSales)} / 商品粗利 ${fmtYen(totalGross)} / 実利益 ${fmtYen(realProfit)} / マイナス要件 ${fmtYen(totalMinus)}`);
        KaientaiM.updateModuleStatus(MODULE_ID, '分析済 (' + state.salesData.length + '件)', true);
        return state.results;
    }

    // ── Render: Overview ──
    function renderOverview() {
        const r = state.results; if (!r) return;
        document.getElementById(pfx('overview-empty')).style.display = 'none';
        document.getElementById(pfx('overview-content')).style.display = 'block';

        document.getElementById(pfx('kpi-total-sales')).textContent = fmtYen(r.totalSales);
        document.getElementById(pfx('kpi-total-cost')).textContent = fmtYen(r.totalCost);
        document.getElementById(pfx('kpi-total-shipping')).textContent = fmtYen(r.totalShipping);
        document.getElementById(pfx('kpi-total-gross')).textContent = fmtYen(r.totalGross);
        document.getElementById(pfx('kpi-total-rebate')).textContent = fmtYen(r.totalRebate);
        document.getElementById(pfx('kpi-total-warehouse')).textContent = fmtYen(r.totalMinus);

        const rpEl = document.getElementById(pfx('kpi-real-profit'));
        rpEl.textContent = fmtYen(r.realProfit);
        rpEl.className = 'kpi-value ' + (r.realProfit >= 0 ? 'positive' : 'negative');
        document.getElementById(pfx('kpi-profit-rate')).textContent = r.totalSales > 0 ? fmtPct(r.realProfit / r.totalSales) : '-';

        // Maker bar
        const md = { aron: { sales: 0, gross: 0 }, pana: { sales: 0, gross: 0 }, other: { sales: 0, gross: 0 } };
        for (const rec of r.records) { md[rec.maker].sales += rec.salesAmount; md[rec.maker].gross += rec.grossProfit; }
        md.aron.real = md.aron.gross + (r.rebateByMaker.aron || 0) - (r.minusByMaker.aron || 0);
        md.pana.real = md.pana.gross + (r.rebateByMaker.pana || 0) - (r.minusByMaker.pana || 0);
        md.other.real = md.other.gross + (r.rebateByMaker.other || 0) - (r.minusByMaker.other || 0);

        destroyChart(state.charts, 'maker-bar');
        state.charts['maker-bar'] = new Chart(document.getElementById(pfx('chart-maker-bar')), {
            type: 'bar',
            data: {
                labels: ['アロン化成', 'パナソニック', 'その他'],
                datasets: [
                    { label: '売上', data: [md.aron.sales, md.pana.sales, md.other.sales], backgroundColor: '#42a5f5' },
                    { label: '商品粗利', data: [md.aron.gross, md.pana.gross, md.other.gross], backgroundColor: '#66bb6a' },
                    { label: '実利益', data: [md.aron.real, md.pana.real, md.other.real], backgroundColor: '#ffa726' }
                ]
            },
            options: { responsive: true, plugins: { legend: { position: 'bottom' } }, scales: { y: { ticks: { callback: v => '¥' + fmt(v) } } } }
        });

        // Pie
        destroyChart(state.charts, 'profit-pie');
        state.charts['profit-pie'] = new Chart(document.getElementById(pfx('chart-profit-pie')), {
            type: 'doughnut',
            data: { labels: ['原価', '送料', '粗利'], datasets: [{ data: [r.totalCost, r.totalShipping, Math.max(0, r.totalGross)], backgroundColor: ['#ef5350', '#ff7043', '#66bb6a'] }] },
            options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
        });

        // Annual line
        if (r.months.length > 0) {
            const salesByM = r.months.map(m => { let s = 0; for (const rec of r.records) if (rec.month === m) s += rec.salesAmount; return s; });
            const grossByM = r.months.map(m => { let s = 0; for (const rec of r.records) if (rec.month === m) s += rec.grossProfit; return s; });
            const realByM = r.months.map((m, i) => {
                let rebate = 0, minus = 0;
                for (const e of Object.values(r.monthlyAgg)) {
                    if (e.month !== m) continue;
                    rebate += calcMonthlyRebate(e, r.settings).total;
                    minus += calcMonthlyMinus(e, r.settings, r.monthSalesTotals).total;
                }
                return grossByM[i] + rebate - minus;
            });
            destroyChart(state.charts, 'annual-line');
            state.charts['annual-line'] = new Chart(document.getElementById(pfx('chart-annual')), {
                type: 'line',
                data: {
                    labels: r.months,
                    datasets: [
                        { label: '売上', data: salesByM, borderColor: '#42a5f5', tension: 0.3, fill: false },
                        { label: '商品粗利', data: grossByM, borderColor: '#66bb6a', tension: 0.3, fill: false },
                        { label: '実利益', data: realByM, borderColor: '#ffa726', tension: 0.3, fill: false, borderWidth: 3 }
                    ]
                },
                options: { responsive: true, plugins: { legend: { position: 'bottom' } }, scales: { y: { ticks: { callback: v => '¥' + fmt(v) } } } }
            });
        }
    }

    // ── Helper: 月次行の実利益計算 ──
    function calcMonthlyReal(e, r) {
        const rebate = calcMonthlyRebate(e, r.settings).total;
        const whFee = calcMonthlyMinus(e, r.settings, r.monthSalesTotals).total;
        return { rebate, whFee, realProfit: e.gross + rebate - whFee };
    }

    // ── Render: Monthly ──
    function renderMonthly() {
        const r = state.results; if (!r) return;
        document.getElementById(pfx('monthly-empty')).style.display = 'none';
        document.getElementById(pfx('monthly-content')).style.display = 'block';

        const filter = document.getElementById(pfx('monthly-maker')).value;
        const tbody = document.getElementById(pfx('monthly-tbody'));
        tbody.innerHTML = '';
        const ml = { aron: 'アロン化成', pana: 'パナソニック', other: 'その他' };

        // メーカー別エントリ
        const entries = Object.values(r.monthlyAgg)
            .filter(e => filter === 'all' || e.maker === filter)
            .sort((a, b) => a.month.localeCompare(b.month) || a.maker.localeCompare(b.maker));

        const months = [...new Set(entries.map(e => e.month))].sort();

        // 年間合計用
        const grandTotal = { sales: 0, cost: 0, shipping: 0, gross: 0, rebate: 0, whFee: 0, real: 0 };

        for (const month of months) {
            const monthEntries = entries.filter(e => e.month === month);
            const monthTotal = { sales: 0, cost: 0, shipping: 0, gross: 0, rebate: 0, whFee: 0, real: 0 };

            // 各メーカー行
            for (const e of monthEntries) {
                const { rebate, whFee, realProfit } = calcMonthlyReal(e, r);
                const pr = e.sales > 0 ? realProfit / e.sales : 0;
                const tr = document.createElement('tr');
                tr.innerHTML = `<td>${e.month}</td><td>${ml[e.maker] || e.maker}</td><td>${fmtYen(e.sales)}</td><td>${fmtYen(e.cost)}</td><td>${fmtYen(e.shipping)}</td><td class="${e.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(e.gross)}</td><td>${fmtYen(rebate)}</td><td>${fmtYen(whFee)}</td><td class="${realProfit >= 0 ? 'positive' : 'negative'}">${fmtYen(realProfit)}</td><td>${fmtPct(pr)}</td>`;
                tbody.appendChild(tr);

                monthTotal.sales += e.sales; monthTotal.cost += e.cost; monthTotal.shipping += e.shipping;
                monthTotal.gross += e.gross; monthTotal.rebate += rebate; monthTotal.whFee += whFee; monthTotal.real += realProfit;
            }

            // 月合計行（メーカーが2つ以上ある場合のみ）
            if (filter === 'all' && monthEntries.length > 1) {
                const pr = monthTotal.sales > 0 ? monthTotal.real / monthTotal.sales : 0;
                const tr = document.createElement('tr');
                tr.style.background = '#e8eaf6';
                tr.style.fontWeight = '700';
                tr.innerHTML = `<td>${month}</td><td>【合計】</td><td>${fmtYen(monthTotal.sales)}</td><td>${fmtYen(monthTotal.cost)}</td><td>${fmtYen(monthTotal.shipping)}</td><td class="${monthTotal.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(monthTotal.gross)}</td><td>${fmtYen(monthTotal.rebate)}</td><td>${fmtYen(monthTotal.whFee)}</td><td class="${monthTotal.real >= 0 ? 'positive' : 'negative'}">${fmtYen(monthTotal.real)}</td><td>${fmtPct(pr)}</td>`;
                tbody.appendChild(tr);
            }

            grandTotal.sales += monthTotal.sales; grandTotal.cost += monthTotal.cost;
            grandTotal.shipping += monthTotal.shipping; grandTotal.gross += monthTotal.gross;
            grandTotal.rebate += monthTotal.rebate; grandTotal.whFee += monthTotal.whFee; grandTotal.real += monthTotal.real;
        }

        // 年間合計行
        if (months.length > 1) {
            const pr = grandTotal.sales > 0 ? grandTotal.real / grandTotal.sales : 0;
            const tr = document.createElement('tr');
            tr.style.background = '#fff3e0';
            tr.style.fontWeight = '700';
            tr.style.fontSize = '13px';
            tr.innerHTML = `<td colspan="2">年間合計</td><td>${fmtYen(grandTotal.sales)}</td><td>${fmtYen(grandTotal.cost)}</td><td>${fmtYen(grandTotal.shipping)}</td><td class="${grandTotal.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(grandTotal.gross)}</td><td>${fmtYen(grandTotal.rebate)}</td><td>${fmtYen(grandTotal.whFee)}</td><td class="${grandTotal.real >= 0 ? 'positive' : 'negative'}">${fmtYen(grandTotal.real)}</td><td>${fmtPct(pr)}</td>`;
            tbody.appendChild(tr);
        }

        // Chart
        const makers = filter === 'all' ? ['aron', 'pana', 'other'] : [filter];
        const colors = { aron: '#42a5f5', pana: '#66bb6a', other: '#bdbdbd' };
        const datasets = makers.map(mk => ({
            label: ml[mk] || mk,
            data: months.map(m => {
                const e = entries.find(x => x.month === m && x.maker === mk);
                if (!e) return 0;
                return calcMonthlyReal(e, r).realProfit;
            }),
            backgroundColor: colors[mk],
        }));

        // 合計ラインも追加
        if (filter === 'all') {
            datasets.push({
                label: '合計',
                type: 'line',
                data: months.map(m => {
                    let total = 0;
                    for (const mk of makers) {
                        const e = entries.find(x => x.month === m && x.maker === mk);
                        if (e) total += calcMonthlyReal(e, r).realProfit;
                    }
                    return total;
                }),
                borderColor: '#ff6f00',
                backgroundColor: 'transparent',
                borderWidth: 3,
                tension: 0.3,
                pointRadius: 5,
            });
        }

        destroyChart(state.charts, 'monthly');
        state.charts['monthly'] = new Chart(document.getElementById(pfx('chart-monthly')), {
            type: 'bar', data: { labels: months, datasets },
            options: { responsive: true, plugins: { legend: { position: 'bottom' } }, scales: { y: { ticks: { callback: v => '¥' + fmt(v) } } } }
        });
    }

    function buildStoreBase(records) {
        const repsSet = new Set();
        const recordsByRep = { all: records };
        for (const rec of records) {
            if (!rec.salesRep) continue;
            repsSet.add(rec.salesRep);
            if (!recordsByRep[rec.salesRep]) recordsByRep[rec.salesRep] = [];
            recordsByRep[rec.salesRep].push(rec);
        }
        return {
            reps: [...repsSet].sort((a, b) => a.localeCompare(b, 'ja')),
            recordsByRep,
            entriesByRep: {}
        };
    }

    function getStoreBase(records, makerF, monthF) {
        const key = makerF + '|' + monthF;
        if (state.storeBaseCache[key]) return state.storeBaseCache[key];

        const filtered = [];
        for (const rec of records) {
            if (makerF !== 'all' && rec.maker !== makerF) continue;
            if (monthF !== 'all' && rec.month !== monthF) continue;
            filtered.push(rec);
        }
        const base = buildStoreBase(filtered);
        state.storeBaseCache[key] = base;
        return base;
    }

    function buildStoreEntries(scopedRecords) {
        const sMap = {};
        for (const rec of scopedRecords) {
            const key = rec.store || '(不明)';
            if (!sMap[key]) {
                sMap[key] = {
                    store: key,
                    sales: 0, cost: 0, shipping: 0, gross: 0, qty: 0,
                    reps: new Set(),
                    codes: new Set(),
                    aronRateNumerator: 0, aronRateDenominator: 0,
                    panaRateNumerator: 0, panaRateDenominator: 0
                };
            }
            sMap[key].sales += rec.salesAmount;
            sMap[key].cost += rec.totalCost;
            sMap[key].shipping += rec.totalShipping;
            sMap[key].gross += rec.grossProfit;
            sMap[key].qty += rec.qty;
            if (rec.salesRep) sMap[key].reps.add(rec.salesRep);
            if (rec.storeCode) sMap[key].codes.add(rec.storeCode);
            if (rec.listPrice > 0 && rec.qty > 0) {
                if (rec.maker === 'aron') {
                    sMap[key].aronRateNumerator += rec.unitPrice * rec.qty;
                    sMap[key].aronRateDenominator += rec.listPrice * rec.qty;
                } else if (rec.maker === 'pana') {
                    sMap[key].panaRateNumerator += rec.unitPrice * rec.qty;
                    sMap[key].panaRateDenominator += rec.listPrice * rec.qty;
                }
            }
        }
        return Object.values(sMap).map(e => {
            const repNames = [...e.reps].sort((a, b) => a.localeCompare(b, 'ja'));
            const codes = [...e.codes].sort((a, b) => a.localeCompare(b, 'ja'));
            return {
                store: e.store,
                storeCode: codes.join(' / '),
                salesRep: repNames.join(' / ') || '(未設定)',
                sales: e.sales,
                cost: e.cost,
                shipping: e.shipping,
                gross: e.gross,
                qty: e.qty,
                rate: e.sales > 0 ? e.gross / e.sales : 0,
                aronRate: e.aronRateDenominator > 0 ? e.aronRateNumerator / e.aronRateDenominator : 0,
                panaRate: e.panaRateDenominator > 0 ? e.panaRateNumerator / e.panaRateDenominator : 0
            };
        });
    }

    function sortStoreEntries(entries, sortKey) {
        switch (sortKey) {
            case 'gross-asc': entries.sort((a, b) => a.gross - b.gross); break;
            case 'sales-desc': entries.sort((a, b) => b.sales - a.sales); break;
            case 'sales-asc': entries.sort((a, b) => a.sales - b.sales); break;
            case 'qty-desc': entries.sort((a, b) => b.qty - a.qty); break;
            case 'qty-asc': entries.sort((a, b) => a.qty - b.qty); break;
            case 'rate-desc': entries.sort((a, b) => b.rate - a.rate); break;
            case 'rate-asc': entries.sort((a, b) => a.rate - b.rate); break;
            case 'aron-rate-desc': entries.sort((a, b) => b.aronRate - a.aronRate); break;
            case 'aron-rate-asc': entries.sort((a, b) => a.aronRate - b.aronRate); break;
            case 'pana-rate-desc': entries.sort((a, b) => b.panaRate - a.panaRate); break;
            case 'pana-rate-asc': entries.sort((a, b) => a.panaRate - b.panaRate); break;
            case 'rep-asc': entries.sort((a, b) => a.salesRep.localeCompare(b.salesRep, 'ja')); break;
            case 'rep-desc': entries.sort((a, b) => b.salesRep.localeCompare(a.salesRep, 'ja')); break;
            case 'code-asc': entries.sort((a, b) => a.storeCode.localeCompare(b.storeCode, 'ja')); break;
            case 'code-desc': entries.sort((a, b) => b.storeCode.localeCompare(a.storeCode, 'ja')); break;
            case 'store-asc': entries.sort((a, b) => a.store.localeCompare(b.store, 'ja')); break;
            case 'store-desc': entries.sort((a, b) => b.store.localeCompare(a.store, 'ja')); break;
            case 'gross-desc':
            default:
                entries.sort((a, b) => b.gross - a.gross);
                break;
        }
    }

    function renderStoreSimulation(scopedRecords, entries) {
        const simStoreSel = document.getElementById(pfx('store-sim-store'));
        const simMaker = document.getElementById(pfx('store-sim-maker')).value;
        const simRateChange = toNum(document.getElementById(pfx('store-sim-rate')).value) / 100;
        const simIncreaseQty = Math.max(0, toNum(document.getElementById(pfx('store-sim-qty')).value));
        const simTbody = document.getElementById(pfx('store-sim-tbody'));

        const prevSimStore = simStoreSel.value;
        const simStoreList = [...entries].sort((a, b) => a.store.localeCompare(b.store, 'ja')).map(e => e.store);
        simStoreSel.innerHTML = '';
        if (simStoreList.length === 0) {
            simStoreSel.innerHTML = '<option value="">（データなし）</option>';
        } else {
            for (const s of simStoreList) simStoreSel.innerHTML += `<option value="${s}">${s}</option>`;
        }
        simStoreSel.value = simStoreList.includes(prevSimStore) ? prevSimStore : (simStoreList[0] || '');
        const simStore = simStoreSel.value;

        const fmtQty = (n) => (n == null || isNaN(n)) ? '-' : (Math.round(n * 10) / 10).toLocaleString('ja-JP');
        const signed = (n, formatter) => (n >= 0 ? '+' : '') + formatter(n);

        simTbody.innerHTML = '';
        const storeRecords = scopedRecords.filter(rec => (rec.store || '(不明)') === simStore);
        if (!simStore || storeRecords.length === 0) {
            simTbody.innerHTML = '<tr><td colspan="7">販売店データがありません</td></tr>';
            return;
        }

        const targetRecords = storeRecords.filter(rec => simMaker === 'all' || rec.maker === simMaker);
        const targetQtyBase = targetRecords.reduce((sum, rec) => sum + rec.qty, 0);
        const targetCount = targetRecords.length;

        const before = {
            sales: 0, gross: 0, qty: 0,
            aronNumerator: 0, aronDenominator: 0,
            panaNumerator: 0, panaDenominator: 0
        };
        const after = {
            sales: 0, gross: 0, qty: 0,
            aronNumerator: 0, aronDenominator: 0,
            panaNumerator: 0, panaDenominator: 0
        };

        for (const rec of storeRecords) {
            before.sales += rec.salesAmount;
            before.gross += rec.grossProfit;
            before.qty += rec.qty;
            if (rec.listPrice > 0 && rec.qty > 0) {
                if (rec.maker === 'aron') {
                    before.aronNumerator += rec.unitPrice * rec.qty;
                    before.aronDenominator += rec.listPrice * rec.qty;
                } else if (rec.maker === 'pana') {
                    before.panaNumerator += rec.unitPrice * rec.qty;
                    before.panaDenominator += rec.listPrice * rec.qty;
                }
            }

            const applyChange = simMaker === 'all' || rec.maker === simMaker;
            let addQty = 0;
            if (applyChange && simIncreaseQty > 0) {
                if (targetQtyBase > 0) addQty = simIncreaseQty * (rec.qty / targetQtyBase);
                else if (targetCount > 0) addQty = simIncreaseQty / targetCount;
            }

            const newQty = rec.qty + addQty;
            const newUnitPrice = applyChange ? rec.unitPrice * (1 + simRateChange) : rec.unitPrice;
            const newSales = newQty * newUnitPrice;
            const newCost = newQty * rec.effectiveCost;
            const newShipping = newQty * rec.shippingCost;
            const newGross = newSales - newCost - newShipping;

            after.sales += newSales;
            after.gross += newGross;
            after.qty += newQty;
            if (rec.listPrice > 0 && newQty > 0) {
                if (rec.maker === 'aron') {
                    after.aronNumerator += newUnitPrice * newQty;
                    after.aronDenominator += rec.listPrice * newQty;
                } else if (rec.maker === 'pana') {
                    after.panaNumerator += newUnitPrice * newQty;
                    after.panaDenominator += rec.listPrice * newQty;
                }
            }
        }

        const beforeAronRate = before.aronDenominator > 0 ? before.aronNumerator / before.aronDenominator : 0;
        const beforePanaRate = before.panaDenominator > 0 ? before.panaNumerator / before.panaDenominator : 0;
        const afterAronRate = after.aronDenominator > 0 ? after.aronNumerator / after.aronDenominator : 0;
        const afterPanaRate = after.panaDenominator > 0 ? after.panaNumerator / after.panaDenominator : 0;
        const beforeProfitRate = before.sales > 0 ? before.gross / before.sales : 0;
        const afterProfitRate = after.sales > 0 ? after.gross / after.sales : 0;

        const diff = {
            sales: after.sales - before.sales,
            gross: after.gross - before.gross,
            qty: after.qty - before.qty,
            profitRate: afterProfitRate - beforeProfitRate,
            aronRate: afterAronRate - beforeAronRate,
            panaRate: afterPanaRate - beforePanaRate
        };

        simTbody.innerHTML = [
            `<tr><td>現状</td><td>${fmtYen(before.sales)}</td><td class="${before.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(before.gross)}</td><td>${fmtPct(beforeProfitRate)}</td><td>${fmtQty(before.qty)}</td><td>${beforeAronRate > 0 ? fmtPct(beforeAronRate) : '-'}</td><td>${beforePanaRate > 0 ? fmtPct(beforePanaRate) : '-'}</td></tr>`,
            `<tr><td>変動後</td><td>${fmtYen(after.sales)}</td><td class="${after.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(after.gross)}</td><td>${fmtPct(afterProfitRate)}</td><td>${fmtQty(after.qty)}</td><td>${afterAronRate > 0 ? fmtPct(afterAronRate) : '-'}</td><td>${afterPanaRate > 0 ? fmtPct(afterPanaRate) : '-'}</td></tr>`,
            `<tr style="font-weight:700;background:#fff3e0;"><td>差分</td><td class="${diff.sales >= 0 ? 'positive' : 'negative'}">${signed(diff.sales, fmtYen)}</td><td class="${diff.gross >= 0 ? 'positive' : 'negative'}">${signed(diff.gross, fmtYen)}</td><td class="${diff.profitRate >= 0 ? 'positive' : 'negative'}">${signed(diff.profitRate, fmtPct)}</td><td class="${diff.qty >= 0 ? 'positive' : 'negative'}">${diff.qty >= 0 ? '+' : ''}${fmtQty(diff.qty)}</td><td class="${diff.aronRate >= 0 ? 'positive' : 'negative'}">${signed(diff.aronRate, fmtPct)}</td><td class="${diff.panaRate >= 0 ? 'positive' : 'negative'}">${signed(diff.panaRate, fmtPct)}</td></tr>`
        ].join('');
    }

    function renderStoreSimulationFromCurrent() {
        if (!state.storeViewRuntime) return;
        renderStoreSimulation(state.storeViewRuntime.scopedRecords, state.storeViewRuntime.entries);
    }

    // ── Render: Store ──
    function renderStore() {
        const r = state.results; if (!r) return;
        document.getElementById(pfx('store-empty')).style.display = 'none';
        document.getElementById(pfx('store-content')).style.display = 'block';

        const mSel = document.getElementById(pfx('store-month'));
        const curMonth = mSel.value;
        mSel.innerHTML = '<option value="all">全期間</option>';
        for (const m of r.months) mSel.innerHTML += `<option value="${m}">${m}</option>`;
        mSel.value = curMonth || 'all';

        const makerF = document.getElementById(pfx('store-maker')).value;
        const monthF = mSel.value;
        const sortKey = document.getElementById(pfx('store-sort')).value;
        const repSel = document.getElementById(pfx('store-rep'));
        const prevRep = repSel.value;
        const searchQ = toStr(document.getElementById(pfx('store-search')).value).toLowerCase();
        const limitRaw = document.getElementById(pfx('store-limit')).value;

        const slices = Array.isArray(r.storeSlices) ? r.storeSlices : [];
        const repSet = new Set();
        for (const slice of slices) {
            if (makerF !== 'all' && slice.maker !== makerF) continue;
            if (monthF !== 'all' && slice.month !== monthF) continue;
            if (slice.salesRep) repSet.add(slice.salesRep);
        }
        const reps = [...repSet].sort((a, b) => a.localeCompare(b, 'ja'));
        repSel.innerHTML = '<option value="all">全担当</option>';
        for (const rep of reps) repSel.innerHTML += `<option value="${rep}">${rep}</option>`;
        repSel.value = reps.includes(prevRep) ? prevRep : 'all';
        const repF = repSel.value;

        const agg = {};
        for (const slice of slices) {
            if (makerF !== 'all' && slice.maker !== makerF) continue;
            if (monthF !== 'all' && slice.month !== monthF) continue;
            if (repF !== 'all' && slice.salesRep !== repF) continue;
            const key = slice.store || '(不明)';
            if (!agg[key]) {
                agg[key] = {
                    store: key,
                    sales: 0, cost: 0, shipping: 0, gross: 0, qty: 0,
                    reps: new Set(),
                    codes: new Set(),
                    aronRateNumerator: 0, aronRateDenominator: 0,
                    panaRateNumerator: 0, panaRateDenominator: 0
                };
            }
            const row = agg[key];
            row.sales += slice.sales;
            row.cost += slice.cost;
            row.shipping += slice.shipping;
            row.gross += slice.gross;
            row.qty += slice.qty;
            if (slice.salesRep) row.reps.add(slice.salesRep);
            if (slice.storeCode) row.codes.add(slice.storeCode);
            row.aronRateNumerator += slice.aronRateNumerator;
            row.aronRateDenominator += slice.aronRateDenominator;
            row.panaRateNumerator += slice.panaRateNumerator;
            row.panaRateDenominator += slice.panaRateDenominator;
        }

        let entries = Object.values(agg).map(e => {
            const repNames = [...e.reps].sort((a, b) => a.localeCompare(b, 'ja'));
            const codes = [...e.codes].sort((a, b) => a.localeCompare(b, 'ja'));
            return {
                store: e.store,
                storeCode: codes.join(' / '),
                salesRep: repNames.join(' / ') || '(未設定)',
                sales: e.sales,
                cost: e.cost,
                shipping: e.shipping,
                gross: e.gross,
                qty: e.qty,
                rate: e.sales > 0 ? e.gross / e.sales : 0,
                aronRate: e.aronRateDenominator > 0 ? e.aronRateNumerator / e.aronRateDenominator : 0,
                panaRate: e.panaRateDenominator > 0 ? e.panaRateNumerator / e.panaRateDenominator : 0
            };
        });

        if (searchQ) {
            entries = entries.filter(e => e.store.toLowerCase().includes(searchQ) || e.storeCode.toLowerCase().includes(searchQ));
        }

        sortStoreEntries(entries, sortKey);

        const pageSize = Math.max(1, Math.min(1000, toNum(limitRaw) || 300));
        const totalPages = Math.max(1, Math.ceil(entries.length / pageSize));
        state.storeCurrentPageTotal = totalPages;
        state.storeCurrentPage = Math.min(Math.max(1, state.storeCurrentPage || 1), totalPages);
        const startIndex = (state.storeCurrentPage - 1) * pageSize;
        const displayed = entries.slice(startIndex, startIndex + pageSize);
        const tbody = document.getElementById(pfx('store-tbody'));
        tbody.innerHTML = displayed.map(e =>
            `<tr><td>${e.store}</td><td>${e.storeCode || '-'}</td><td>${e.salesRep}</td><td>${e.aronRate > 0 ? fmtPct(e.aronRate) : '-'}</td><td>${e.panaRate > 0 ? fmtPct(e.panaRate) : '-'}</td><td>${fmtYen(e.sales)}</td><td>${fmtYen(e.cost)}</td><td>${fmtYen(e.shipping)}</td><td class="${e.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(e.gross)}</td><td>${fmtPct(e.rate)}</td><td>${fmt(e.qty)}</td></tr>`
        ).join('');

        const summaryEl = document.getElementById(pfx('store-summary'));
        if (summaryEl) {
            if (entries.length === 0) {
                summaryEl.textContent = '表示: 0件 / 全0件';
            } else {
                const from = startIndex + 1;
                const to = startIndex + displayed.length;
                summaryEl.textContent = `表示: ${fmt(from)}-${fmt(to)}件 / 全${fmt(entries.length)}件`;
            }
        }

        const pagerEl = document.getElementById(pfx('store-pagination'));
        const pageStatusEl = document.getElementById(pfx('store-page-status'));
        const prevBtn = document.getElementById(pfx('store-page-prev'));
        const nextBtn = document.getElementById(pfx('store-page-next'));
        if (pagerEl && pageStatusEl && prevBtn && nextBtn) {
            pagerEl.style.display = totalPages > 1 ? 'flex' : 'none';
            pageStatusEl.textContent = `${fmt(state.storeCurrentPage)} / ${fmt(totalPages)}ページ`;
            prevBtn.disabled = state.storeCurrentPage <= 1;
            nextBtn.disabled = state.storeCurrentPage >= totalPages;
        }

        state.storeViewRuntime = null;

        const top = [...entries].sort((a, b) => b.gross - a.gross).slice(0, 15);
        destroyChart(state.charts, 'store');
        state.charts['store'] = new Chart(document.getElementById(pfx('chart-store')), {
            type: 'bar',
            data: {
                labels: top.map(e => e.store),
                datasets: [
                    { label: '売上', data: top.map(e => e.sales), backgroundColor: '#42a5f5' },
                    { label: '粗利', data: top.map(e => e.gross), backgroundColor: top.map(e => e.gross >= 0 ? '#66bb6a' : '#ef5350') }
                ]
            },
            options: { indexAxis: 'y', responsive: true, plugins: { legend: { position: 'bottom' } }, scales: { x: { ticks: { callback: v => '¥' + fmt(v) } } } }
        });
    }

    // ── Render: Simulation ──
    function renderSimulation() {
        const r = state.results; if (!r) return;
        document.getElementById(pfx('sim-empty')).style.display = 'none';
        document.getElementById(pfx('sim-content')).style.display = 'block';

        let aRS = 0, aRC = 0, pRS = 0, pRC = 0, allRS = 0, allRC = 0;
        for (const rec of r.records) {
            if (rec.listPrice > 0 && rec.unitPrice > 0) {
                const rt = rec.unitPrice / rec.listPrice;
                allRS += rt; allRC++;
                if (rec.maker === 'aron') { aRS += rt; aRC++; }
                if (rec.maker === 'pana') { pRS += rt; pRC++; }
            }
        }
        document.getElementById(pfx('sim-cur-aron')).textContent = fmtPct(aRC > 0 ? aRS / aRC : 0);
        document.getElementById(pfx('sim-cur-pana')).textContent = fmtPct(pRC > 0 ? pRS / pRC : 0);
        document.getElementById(pfx('sim-cur-all')).textContent = fmtPct(allRC > 0 ? allRS / allRC : 0);

        const rateChange = toNum(document.getElementById(pfx('sim-rate')).value) / 100;
        const target = document.getElementById(pfx('sim-target')).value;
        const rebateRateDeltaAron = toNum(document.getElementById(pfx('sim-rebate-aron')).value) / 100;
        const rebateRateDeltaPana = toNum(document.getElementById(pfx('sim-rebate-pana')).value) / 100;
        const fixedDeltaAron = toNum(document.getElementById(pfx('sim-fixed-aron')).value);
        const fixedDeltaPana = toNum(document.getElementById(pfx('sim-fixed-pana')).value);
        document.getElementById(pfx('sim-rate-display')).textContent = (rateChange >= 0 ? '+' : '') + (rateChange * 100).toFixed(1) + '%';

        const monthMakerBase = {};
        Object.values(r.monthlyAgg).forEach(e => {
            monthMakerBase[e.month + '|' + e.maker] = { month: e.month, maker: e.maker, sales: e.sales, gross: e.gross, qty: e.qty };
        });

        const calcScenario = (rate) => {
            let gross = 0;
            const monthMaker = {};
            for (const rec of r.records) {
                const applyChange = target === 'all' || rec.maker === target;
                const newUnitPrice = applyChange ? rec.unitPrice * (1 + rate) : rec.unitPrice;
                const newSales = rec.qty * newUnitPrice;
                const newGross = newSales - rec.totalCost - rec.totalShipping;
                gross += newGross;

                const mk = rec.month + '|' + rec.maker;
                if (!monthMaker[mk]) monthMaker[mk] = { month: rec.month, maker: rec.maker, sales: 0, gross: 0, qty: 0 };
                monthMaker[mk].sales += newSales;
                monthMaker[mk].gross += newGross;
                monthMaker[mk].qty += rec.qty;
            }

            Object.keys(monthMakerBase).forEach(mk => {
                if (!monthMaker[mk]) {
                    const base = monthMakerBase[mk];
                    monthMaker[mk] = { month: base.month, maker: base.maker, sales: 0, gross: 0, qty: 0 };
                }
            });

            const monthSalesTotals = {};
            Object.values(monthMaker).forEach(e => {
                monthSalesTotals[e.month] = (monthSalesTotals[e.month] || 0) + e.sales;
            });

            let rebate = 0;
            let minus = 0;
            for (const e of Object.values(monthMaker)) {
                const baseRate = e.maker === 'aron' ? r.settings.rebateAron : e.maker === 'pana' ? r.settings.rebatePana : 0;
                const rateDelta = e.maker === 'aron' ? rebateRateDeltaAron : e.maker === 'pana' ? rebateRateDeltaPana : 0;
                const variable = e.sales * (baseRate + rateDelta);

                const fixed = getMonthlyRebate(r.settings, e.month, e.maker);
                const fixedDelta = e.maker === 'aron' ? fixedDeltaAron : e.maker === 'pana' ? fixedDeltaPana : 0;
                const fixedTotal = fixed.fixed + fixedDelta;

                rebate += variable + fixedTotal;
                minus += calcMonthlyMinus(e, r.settings, monthSalesTotals).total;
            }
            return { gross, rebate, minus, real: gross + rebate - minus };
        };

        const before = { gross: r.totalGross, rebate: r.totalRebate, minus: r.totalMinus, real: r.realProfit };
        const after = calcScenario(rateChange);
        const diff = after.real - before.real;

        document.getElementById(pfx('sim-before')).textContent = fmtYen(before.real);
        document.getElementById(pfx('sim-after')).textContent = fmtYen(after.real);
        document.getElementById(pfx('sim-after')).className = 'sim-value ' + (after.real >= 0 ? 'positive' : 'negative');
        document.getElementById(pfx('sim-diff')).textContent = (diff >= 0 ? '+' : '') + fmtYen(diff);
        document.getElementById(pfx('sim-diff')).className = 'sim-value ' + (diff >= 0 ? 'positive' : 'negative');

        const breakdownBody = document.getElementById(pfx('sim-breakdown-body'));
        if (breakdownBody) {
            const rows = [
                { label: '商品粗利', before: before.gross, after: after.gross },
                { label: 'リベート', before: before.rebate, after: after.rebate },
                { label: 'マイナス要件', before: before.minus, after: after.minus },
                { label: '実利益', before: before.real, after: after.real }
            ];
            breakdownBody.innerHTML = rows.map(row => {
                const delta = row.after - row.before;
                const cls = delta >= 0 ? 'positive' : 'negative';
                return `<tr><td>${row.label}</td><td>${fmtYen(row.before)}</td><td>${fmtYen(row.after)}</td><td class="${cls}">${delta >= 0 ? '+' : ''}${fmtYen(delta)}</td></tr>`;
            }).join('');
        }

        const steps = [], gv = [];
        for (let pct = -20; pct <= 20; pct += 2) {
            steps.push((pct >= 0 ? '+' : '') + pct + '%');
            gv.push(calcScenario(pct / 100).real);
        }
        destroyChart(state.charts, 'simulation');
        state.charts['simulation'] = new Chart(document.getElementById(pfx('chart-sim')), {
            type: 'line',
            data: { labels: steps, datasets: [{ label: '実利益', data: gv, borderColor: '#ffa726', backgroundColor: 'rgba(255,167,38,0.1)', fill: true, tension: 0.3, borderWidth: 3, pointRadius: 4, pointBackgroundColor: gv.map(v => v >= 0 ? '#66bb6a' : '#ef5350') }] },
            options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { ticks: { callback: v => '¥' + fmt(v) } } } }
        });
    }

    function buildStoreDetailMap() {
        const r = state.results;
        const map = {};
        for (const rec of r.records) {
            const key = rec.store || '(不明)';
            if (!map[key]) {
                map[key] = {
                    store: key,
                    codes: new Set(),
                    reps: new Set(),
                    records: [],
                    sales: 0
                };
            }
            if (rec.storeCode) map[key].codes.add(rec.storeCode);
            if (rec.salesRep) map[key].reps.add(rec.salesRep);
            map[key].records.push(rec);
            map[key].sales += rec.salesAmount;
        }
        return Object.values(map).map(item => ({
            store: item.store,
            codeText: [...item.codes].sort((a, b) => a.localeCompare(b, 'ja')).join(' / '),
            repText: [...item.reps].sort((a, b) => a.localeCompare(b, 'ja')).join(' / '),
            records: item.records,
            sales: item.sales
        })).sort((a, b) => b.sales - a.sales);
    }

    function linearRegression(points) {
        if (!Array.isArray(points) || points.length < 2) return { slope: 0, intercept: 0, r2: 0 };
        let sx = 0, sy = 0, sxx = 0, syy = 0, sxy = 0;
        const n = points.length;
        for (const p of points) {
            sx += p.x; sy += p.y;
            sxx += p.x * p.x;
            syy += p.y * p.y;
            sxy += p.x * p.y;
        }
        const den = (n * sxx - sx * sx);
        if (Math.abs(den) < 1e-9) return { slope: 0, intercept: sy / n, r2: 0 };
        const slope = (n * sxy - sx * sy) / den;
        const intercept = (sy - slope * sx) / n;

        const yMean = sy / n;
        let ssTot = 0, ssRes = 0;
        for (const p of points) {
            const pred = slope * p.x + intercept;
            ssTot += (p.y - yMean) * (p.y - yMean);
            ssRes += (p.y - pred) * (p.y - pred);
        }
        const r2 = ssTot > 1e-9 ? Math.max(0, Math.min(1, 1 - ssRes / ssTot)) : 0;
        return { slope, intercept, r2 };
    }

    function monthStrToIndex(monthStr) {
        const m = toStr(monthStr).match(/^(\d{4})-(\d{1,2})$/);
        if (!m) return null;
        const y = toNum(m[1]);
        const mo = toNum(m[2]);
        if (y <= 0 || mo < 1 || mo > 12) return null;
        return y * 12 + (mo - 1);
    }

    function monthIndexToMonthNo(idx) {
        return ((idx % 12) + 12) % 12 + 1;
    }

    function buildMonthlySeries(records) {
        const map = {};
        for (const rec of records) {
            const month = toStr(rec.month);
            const idx = monthStrToIndex(month);
            if (idx == null) continue;
            if (!map[month]) {
                map[month] = {
                    month,
                    idx,
                    qty: 0,
                    sales: 0,
                    cost: 0,
                    shipping: 0,
                    unitPriceWeighted: 0,
                    listWeighted: 0
                };
            }
            const row = map[month];
            row.qty += rec.qty;
            row.sales += rec.salesAmount;
            row.cost += rec.totalCost;
            row.shipping += rec.totalShipping;
            row.unitPriceWeighted += rec.unitPrice * rec.qty;
            row.listWeighted += rec.listPrice * rec.qty;
        }
        const arr = Object.values(map).sort((a, b) => a.idx - b.idx);
        return arr.map(row => ({
            ...row,
            unitPrice: row.qty > 0 ? row.unitPriceWeighted / row.qty : 0,
            costPerQty: row.qty > 0 ? row.cost / row.qty : 0,
            shippingPerQty: row.qty > 0 ? row.shipping / row.qty : 0,
            listPrice: row.qty > 0 ? row.listWeighted / row.qty : 0
        }));
    }

    function estimateTrendRate(monthlySeries) {
        if (!Array.isArray(monthlySeries) || monthlySeries.length < 2) return { monthlyRate: 0, confidence: 0, points: monthlySeries?.length || 0, r2: 0 };
        const points = monthlySeries.map((row, i) => ({ x: i, y: row.qty }));
        const reg = linearRegression(points);
        const avgQty = monthlySeries.reduce((sum, row) => sum + row.qty, 0) / Math.max(1, monthlySeries.length);
        const monthlyRate = avgQty > 0 ? reg.slope / avgQty : 0;
        const confidence = Math.max(0, Math.min(1, reg.r2 * Math.min(1, monthlySeries.length / 12)));
        return { monthlyRate, confidence, points: monthlySeries.length, r2: reg.r2 };
    }

    function estimateElasticity(monthlySeries, manualFallback) {
        const points = [];
        for (const row of monthlySeries || []) {
            if (row.qty > 0 && row.unitPrice > 0) {
                points.push({ x: Math.log(row.unitPrice), y: Math.log(row.qty) });
            }
        }
        if (points.length < 3) {
            return { value: manualFallback, source: 'manual', confidence: 0, points: points.length, r2: 0 };
        }
        const reg = linearRegression(points);
        let value = reg.slope;
        if (!Number.isFinite(value)) value = manualFallback;
        value = Math.max(-5, Math.min(1, value));
        const confidence = Math.max(0, Math.min(1, reg.r2 * Math.min(1, points.length / 12)));
        return { value, source: 'estimated', confidence, points: points.length, r2: reg.r2 };
    }

    function estimateSeasonality(monthlySeries, horizonMonths) {
        if (!Array.isArray(monthlySeries) || monthlySeries.length < 6) return { factor: 1, source: 'neutral' };
        const avgQty = monthlySeries.reduce((sum, row) => sum + row.qty, 0) / Math.max(1, monthlySeries.length);
        if (avgQty <= 0) return { factor: 1, source: 'neutral' };
        const byMonthNo = {};
        for (const row of monthlySeries) {
            const mo = monthIndexToMonthNo(row.idx);
            if (!byMonthNo[mo]) byMonthNo[mo] = [];
            byMonthNo[mo].push(row.qty / avgQty);
        }
        const meanRatio = {};
        for (const k of Object.keys(byMonthNo)) {
            const arr = byMonthNo[k];
            meanRatio[k] = arr.reduce((sum, v) => sum + v, 0) / Math.max(1, arr.length);
        }
        const lastIdx = monthlySeries[monthlySeries.length - 1].idx;
        const baseMonthNo = monthIndexToMonthNo(lastIdx);
        const targetMonthNo = monthIndexToMonthNo(lastIdx + Math.max(1, horizonMonths));
        const baseRatio = meanRatio[baseMonthNo] || 1;
        const targetRatio = meanRatio[targetMonthNo] || 1;
        const factor = Math.max(0.7, Math.min(1.3, targetRatio / baseRatio));
        return { factor, source: 'estimated' };
    }

    function weightedCostShippingList(records) {
        let qty = 0;
        let unitPriceWeighted = 0;
        let costWeighted = 0;
        let shippingWeighted = 0;
        let aronNum = 0, aronDen = 0, panaNum = 0, panaDen = 0;
        for (const rec of records) {
            if (rec.qty <= 0) continue;
            qty += rec.qty;
            unitPriceWeighted += rec.unitPrice * rec.qty;
            costWeighted += rec.effectiveCost * rec.qty;
            shippingWeighted += rec.shippingCost * rec.qty;
            if (rec.listPrice > 0) {
                if (rec.maker === 'aron') { aronNum += rec.unitPrice * rec.qty; aronDen += rec.listPrice * rec.qty; }
                if (rec.maker === 'pana') { panaNum += rec.unitPrice * rec.qty; panaDen += rec.listPrice * rec.qty; }
            }
        }
        return {
            qty,
            unitPrice: qty > 0 ? unitPriceWeighted / qty : 0,
            costPerQty: qty > 0 ? costWeighted / qty : 0,
            shippingPerQty: qty > 0 ? shippingWeighted / qty : 0,
            aronRate: aronDen > 0 ? aronNum / aronDen : 0,
            panaRate: panaDen > 0 ? panaNum / panaDen : 0
        };
    }

    function renderStoreDetail() {
        const r = state.results; if (!r) return;
        document.getElementById(pfx('store-detail-empty')).style.display = 'none';
        document.getElementById(pfx('store-detail-content')).style.display = 'block';

        const searchInput = document.getElementById(pfx('store-detail-search'));
        const selectEl = document.getElementById(pfx('store-detail-select'));
        const summaryEl = document.getElementById(pfx('store-detail-summary'));
        const tbody = document.getElementById(pfx('store-detail-tbody'));
        const factorBody = document.getElementById(pfx('store-detail-factor-body'));
        const maker = document.getElementById(pfx('store-detail-maker')).value;
        const rateChange = toNum(document.getElementById(pfx('store-detail-rate')).value) / 100;
        const qtyIncreasePerMonth = Math.max(0, toNum(document.getElementById(pfx('store-detail-qty')).value));
        const manualQtyRate = toNum(document.getElementById(pfx('store-detail-manual-rate')).value) / 100;
        const horizonMonths = Math.max(1, Math.round(toNum(document.getElementById(pfx('store-detail-horizon')).value) || 3));
        const useTrend = !!document.getElementById(pfx('store-detail-use-trend')).checked;
        const trendAdjust = toNum(document.getElementById(pfx('store-detail-trend-adjust')).value) / 100;
        const useElasticAuto = !!document.getElementById(pfx('store-detail-use-elastic-auto')).checked;
        const manualElasticity = toNum(document.getElementById(pfx('store-detail-elasticity')).value || -1);
        const useSeasonality = !!document.getElementById(pfx('store-detail-use-seasonality')).checked;
        const windowMonths = Math.max(1, Math.round(toNum(document.getElementById(pfx('store-detail-window')).value) || 6));
        const q = toStr(searchInput.value).toLowerCase();

        const stores = buildStoreDetailMap();
        const filtered = stores.filter(item => {
            if (!q) return true;
            return item.store.toLowerCase().includes(q) || item.codeText.toLowerCase().includes(q);
        });

        const prev = selectEl.value;
        selectEl.innerHTML = filtered.map(item => `<option value="${item.store}">${item.store}${item.codeText ? ` [${item.codeText}]` : ''}</option>`).join('');
        if (filtered.length === 0) {
            selectEl.innerHTML = '<option value="">該当なし</option>';
        }
        selectEl.value = filtered.some(item => item.store === prev) ? prev : (filtered[0]?.store || '');

        const selected = filtered.find(item => item.store === selectEl.value);
        if (!selected) {
            summaryEl.textContent = '販売店が見つかりません';
            tbody.innerHTML = '<tr><td colspan="8">対象データがありません</td></tr>';
            if (factorBody) factorBody.innerHTML = '<tr><td colspan="4">要因データがありません</td></tr>';
            return;
        }

        const targetRecords = selected.records.filter(rec => maker === 'all' || rec.maker === maker);
        if (targetRecords.length === 0) {
            summaryEl.textContent = `仕入先コード: ${selected.codeText || '-'} / 営業担当: ${selected.repText || '(未設定)'} / 対象データ0件`;
            tbody.innerHTML = '<tr><td colspan="8">対象メーカーのデータがありません</td></tr>';
            if (factorBody) factorBody.innerHTML = '<tr><td colspan="4">要因データがありません</td></tr>';
            return;
        }

        const monthlySeries = buildMonthlySeries(targetRecords);
        const recentMonths = monthlySeries.slice(-windowMonths);
        const refSeries = recentMonths.length > 0 ? recentMonths : monthlySeries;
        const baseMonthlyQty = refSeries.reduce((sum, row) => sum + row.qty, 0) / Math.max(1, refSeries.length);

        const trend = estimateTrendRate(monthlySeries);
        const trendRate = useTrend ? (trend.monthlyRate + trendAdjust) : trendAdjust;
        const elasticityEst = estimateElasticity(monthlySeries, manualElasticity);
        const elasticity = useElasticAuto ? elasticityEst.value : manualElasticity;
        const seasonality = useSeasonality ? estimateSeasonality(monthlySeries, horizonMonths).factor : 1;

        const qtyPriceImpact = Math.max(0, 1 + elasticity * rateChange);
        const qtyTrendImpact = Math.max(0, 1 + trendRate * horizonMonths);
        const qtyManualImpact = Math.max(0, 1 + manualQtyRate);
        const combinedMultiplier = Math.max(0, qtyPriceImpact * qtyTrendImpact * qtyManualImpact * seasonality);

        const baseQtyForecast = Math.max(0, baseMonthlyQty * horizonMonths);
        const forecastQty = Math.max(0, baseQtyForecast * combinedMultiplier + qtyIncreasePerMonth * horizonMonths);
        const deltaQty = forecastQty - baseQtyForecast;

        const weightedBase = weightedCostShippingList(targetRecords);
        const baseUnitPrice = weightedBase.unitPrice;
        const newUnitPrice = baseUnitPrice * (1 + rateChange);

        const beforeSales = baseQtyForecast * baseUnitPrice;
        const beforeGross = beforeSales - baseQtyForecast * (weightedBase.costPerQty + weightedBase.shippingPerQty);
        const afterSales = forecastQty * newUnitPrice;
        const afterGross = afterSales - forecastQty * (weightedBase.costPerQty + weightedBase.shippingPerQty);

        const beforeRate = beforeSales > 0 ? beforeGross / beforeSales : 0;
        const afterRate = afterSales > 0 ? afterGross / afterSales : 0;

        const beforeAronRate = weightedBase.aronRate;
        const beforePanaRate = weightedBase.panaRate;
        const afterAronRate = (maker === 'all' || maker === 'aron') ? beforeAronRate * (1 + rateChange) : beforeAronRate;
        const afterPanaRate = (maker === 'all' || maker === 'pana') ? beforePanaRate * (1 + rateChange) : beforePanaRate;

        summaryEl.textContent = `仕入先コード: ${selected.codeText || '-'} / 営業担当: ${selected.repText || '(未設定)'} / 履歴${fmt(monthlySeries.length)}ヶ月・明細${fmt(selected.records.length)}件 / 予測期間: ${fmt(horizonMonths)}ヶ月`;

        const fmtQty = (n) => (n == null || isNaN(n)) ? '-' : (Math.round(n * 10) / 10).toLocaleString('ja-JP');
        const signedYen = (n) => (n >= 0 ? '+' : '') + fmtYen(n);
        const signedPct = (n) => (n >= 0 ? '+' : '') + fmtPct(n);
        tbody.innerHTML = [
            `<tr><td>ベース予測</td><td>${fmtYen(beforeSales)}</td><td class="${beforeGross >= 0 ? 'positive' : 'negative'}">${fmtYen(beforeGross)}</td><td>${fmtPct(beforeRate)}</td><td>${fmtQty(baseQtyForecast)}</td><td>${beforeAronRate > 0 ? fmtPct(beforeAronRate) : '-'}</td><td>${beforePanaRate > 0 ? fmtPct(beforePanaRate) : '-'}</td><td>${maker === 'all' ? '全メーカー' : maker}</td></tr>`,
            `<tr><td>高度シミュ後</td><td>${fmtYen(afterSales)}</td><td class="${afterGross >= 0 ? 'positive' : 'negative'}">${fmtYen(afterGross)}</td><td>${fmtPct(afterRate)}</td><td>${fmtQty(forecastQty)}</td><td>${afterAronRate > 0 ? fmtPct(afterAronRate) : '-'}</td><td>${afterPanaRate > 0 ? fmtPct(afterPanaRate) : '-'}</td><td>${maker === 'all' ? '全メーカー' : maker}</td></tr>`,
            `<tr><td>差分</td><td class="${afterSales - beforeSales >= 0 ? 'positive' : 'negative'}">${signedYen(afterSales - beforeSales)}</td><td class="${afterGross - beforeGross >= 0 ? 'positive' : 'negative'}">${signedYen(afterGross - beforeGross)}</td><td class="${afterRate - beforeRate >= 0 ? 'positive' : 'negative'}">${signedPct(afterRate - beforeRate)}</td><td class="${deltaQty >= 0 ? 'positive' : 'negative'}">${deltaQty >= 0 ? '+' : ''}${fmtQty(deltaQty)}</td><td class="${afterAronRate - beforeAronRate >= 0 ? 'positive' : 'negative'}">${signedPct(afterAronRate - beforeAronRate)}</td><td class="${afterPanaRate - beforePanaRate >= 0 ? 'positive' : 'negative'}">${signedPct(afterPanaRate - beforePanaRate)}</td><td>複合予測</td></tr>`
        ].join('');

        if (factorBody) {
            const confidence = Math.max(0, Math.min(1, (trend.confidence + elasticityEst.confidence) / 2));
            const rows = [
                { label: 'トレンド率(月次)', value: fmtPct(trend.monthlyRate), memo: `推定値 / R²=${(trend.r2 || 0).toFixed(2)} / ${trend.points}点` },
                { label: 'トレンド補正', value: fmtPct(trendAdjust), memo: useTrend ? '自動トレンドに加算' : '手動のみ' },
                { label: '価格弾力性', value: elasticity.toFixed(2), memo: useElasticAuto ? `自動推定 (${elasticityEst.points}点)` : '手動入力' },
                { label: '価格影響倍率', value: (qtyPriceImpact * 100).toFixed(1) + '%', memo: `価格変動 ${(rateChange * 100).toFixed(1)}%` },
                { label: '季節性倍率', value: (seasonality * 100).toFixed(1) + '%', memo: useSeasonality ? '季節性ON' : '季節性OFF' },
                { label: '手動増減率倍率', value: (qtyManualImpact * 100).toFixed(1) + '%', memo: `手動増減率 ${(manualQtyRate * 100).toFixed(1)}%` },
                { label: '複合数量倍率', value: (combinedMultiplier * 100).toFixed(1) + '%', memo: `絶対増加 +${fmtQty(qtyIncreasePerMonth)}個/月` },
                { label: '推定信頼度', value: (confidence * 100).toFixed(0) + '%', memo: 'トレンド×弾力性の統合指標' }
            ];
            factorBody.innerHTML = rows.map(row => `<tr><td>${row.label}</td><td>${row.value}</td><td>${row.memo}</td></tr>`).join('');
        }
    }

    function getProgressStatusLabel(status) {
        const map = { planned: '計画中', doing: '実行中', done: '完了', hold: '保留' };
        return map[status] || status || '-';
    }

    function parseDateYmd(ymd) {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(toStr(ymd))) return null;
        const [y, m, d] = ymd.split('-').map(Number);
        const dt = new Date(y, m - 1, d);
        if (Number.isNaN(dt.getTime())) return null;
        return dt;
    }

    function getDeadlineBadge(item) {
        if (!item || !item.deadline) return { className: '', mark: '', title: '' };
        if (item.status === 'done') return { className: 'done', mark: '', title: '' };
        const due = parseDateYmd(item.deadline);
        if (!due) return { className: '', mark: '', title: '' };
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        due.setHours(0, 0, 0, 0);
        const days = Math.round((due.getTime() - today.getTime()) / 86400000);
        if (days < 0) return { className: 'overdue', mark: '!', title: '期限超過' };
        if (days <= 3) return { className: 'near', mark: '!', title: `期限まで${days}日` };
        return { className: '', mark: '', title: '' };
    }

    function buildProgressRepStoreMap() {
        const map = {};
        const addPair = (repRaw, storeRaw) => {
            const rep = toStr(repRaw);
            const store = toStr(storeRaw);
            if (!rep) return;
            if (!map[rep]) map[rep] = new Set();
            if (store) map[rep].add(store);
        };
        for (const row of state.salesData) addPair(row.salesRep, row.store);
        for (const item of state.progressItems) addPair(item.rep, item.customer);
        return map;
    }

    function renderProgressToggleButtons(container, values, selected, dataKey, emptyText) {
        if (!container) return;
        if (!values.length) {
            container.innerHTML = `<span class="hint">${escapeHTML(emptyText)}</span>`;
            return;
        }
        container.innerHTML = values.map(v => {
            const active = v === selected ? ' active' : '';
            return `<button type="button" class="toggle-chip${active}" data-${dataKey}="${escapeHTML(v)}">${escapeHTML(v)}</button>`;
        }).join('');
    }

    function renderProgressFormSelectors() {
        const repInput = document.getElementById(pfx('progress-rep'));
        const customerInput = document.getElementById(pfx('progress-customer'));
        const repToggle = document.getElementById(pfx('progress-rep-toggle'));
        const customerToggle = document.getElementById(pfx('progress-customer-toggle'));
        if (!repInput || !customerInput) return;

        const map = buildProgressRepStoreMap();
        const reps = Object.keys(map).sort((a, b) => a.localeCompare(b, 'ja'));
        let selectedRep = toStr(repInput.value);
        if (!reps.includes(selectedRep)) selectedRep = '';
        repInput.value = selectedRep;

        const stores = selectedRep
            ? [...(map[selectedRep] || new Set())].sort((a, b) => a.localeCompare(b, 'ja'))
            : [];
        let selectedCustomer = toStr(customerInput.value);
        if (!stores.includes(selectedCustomer)) selectedCustomer = '';
        customerInput.value = selectedCustomer;

        renderProgressToggleButtons(repToggle, reps, selectedRep, 'progress-rep-value', '営業担当が未登録です');
        const storeHint = selectedRep ? 'この担当に紐づく販売店がありません' : '先に営業担当を選択してください';
        renderProgressToggleButtons(customerToggle, stores, selectedCustomer, 'progress-customer-value', storeHint);
    }

    function getProgressFormValues() {
        return {
            rep: toStr(document.getElementById(pfx('progress-rep')).value),
            customer: toStr(document.getElementById(pfx('progress-customer')).value),
            actionPlan: toStr(document.getElementById(pfx('progress-action')).value),
            deadline: toStr(document.getElementById(pfx('progress-deadline')).value),
            result: toStr(document.getElementById(pfx('progress-result')).value),
            status: toStr(document.getElementById(pfx('progress-status')).value || 'planned')
        };
    }

    function clearProgressForm() {
        setInputValue('progress-edit-id', '');
        setInputValue('progress-rep', '');
        setInputValue('progress-customer', '');
        setInputValue('progress-action', '');
        setInputValue('progress-deadline', '');
        setInputValue('progress-result', '');
        setInputValue('progress-status', 'planned');
        renderProgressFormSelectors();
    }

    function renderProgressTable() {
        const body = document.getElementById(pfx('progress-tbody'));
        if (!body) return;
        const repFilterEl = document.getElementById(pfx('progress-filter-rep'));
        const customerFilterEl = document.getElementById(pfx('progress-filter-customer'));
        const repSearch = toStr(document.getElementById(pfx('progress-filter-rep-search'))?.value || '').toLowerCase();
        const customerSearch = toStr(document.getElementById(pfx('progress-filter-customer-search'))?.value || '').toLowerCase();
        const statusFilter = toStr(document.getElementById(pfx('progress-filter-status'))?.value || 'all');
        const search = toStr(document.getElementById(pfx('progress-filter-search'))?.value || '').toLowerCase();

        renderProgressFormSelectors();

        const repStoreMap = buildProgressRepStoreMap();
        const allReps = Object.keys(repStoreMap).sort((a, b) => a.localeCompare(b, 'ja'));
        const reps = repSearch ? allReps.filter(rep => rep.toLowerCase().includes(repSearch)) : allReps;
        if (repFilterEl) {
            const prev = toStr(repFilterEl.value || 'all');
            repFilterEl.innerHTML = `<option value="all">全担当</option>${reps.map(rep => `<option value="${escapeHTML(rep)}">${escapeHTML(rep)}</option>`).join('')}`;
            repFilterEl.value = prev === 'all' || reps.includes(prev) ? prev : 'all';
        }
        const repFilter = toStr(repFilterEl?.value || 'all');

        const storeSet = new Set();
        if (repFilter !== 'all') {
            (repStoreMap[repFilter] || new Set()).forEach(store => { if (store) storeSet.add(store); });
        } else {
            Object.values(repStoreMap).forEach(set => {
                if (!(set instanceof Set)) return;
                set.forEach(store => { if (store) storeSet.add(store); });
            });
        }
        const allCustomers = [...storeSet].sort((a, b) => a.localeCompare(b, 'ja'));
        const customers = customerSearch ? allCustomers.filter(store => store.toLowerCase().includes(customerSearch)) : allCustomers;
        if (customerFilterEl) {
            const prev = toStr(customerFilterEl.value || 'all');
            customerFilterEl.innerHTML = `<option value="all">全販売店</option>${customers.map(store => `<option value="${escapeHTML(store)}">${escapeHTML(store)}</option>`).join('')}`;
            customerFilterEl.value = prev === 'all' || customers.includes(prev) ? prev : 'all';
        }
        const customerFilter = toStr(customerFilterEl?.value || 'all');

        let rows = [...state.progressItems];
        if (repFilter !== 'all') rows = rows.filter(item => item.rep === repFilter);
        if (customerFilter !== 'all') rows = rows.filter(item => item.customer === customerFilter);
        if (statusFilter !== 'all') rows = rows.filter(item => item.status === statusFilter);
        if (search) {
            rows = rows.filter(item => {
                const hay = `${item.rep} ${item.customer} ${item.actionPlan} ${item.result} ${item.deadline}`.toLowerCase();
                return hay.includes(search);
            });
        }
        rows.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));

        body.innerHTML = rows.length === 0
            ? '<tr><td colspan="9">データがありません</td></tr>'
            : rows.map(item => {
                const badge = getDeadlineBadge(item);
                return `
                <tr>
                    <td>${escapeHTML(item.rep || '-')}</td>
                    <td>${escapeHTML(item.customer || '-')}</td>
                    <td>${escapeHTML(item.actionPlan || '-')}</td>
                    <td class="progress-deadline ${badge.className}">${escapeHTML(item.deadline || '-')} ${badge.mark ? `<span class="deadline-flag" title="${escapeHTML(badge.title)}">${badge.mark}</span>` : ''}</td>
                    <td>${escapeHTML(item.result || '-')}</td>
                    <td>${escapeHTML(getProgressStatusLabel(item.status))}</td>
                    <td>${toStr(item.updatedAt).replace('T', ' ').slice(0, 16)}</td>
                    <td><button class="btn-secondary progress-edit-btn" data-progress-edit="${item.id}">編集</button></td>
                    <td><button class="btn-secondary progress-delete-btn" data-progress-delete="${item.id}">削除</button></td>
                </tr>
            `;
            }).join('');
    }

    function saveProgressItem() {
        const editId = toNum(document.getElementById(pfx('progress-edit-id')).value);
        const data = getProgressFormValues();
        if (!data.rep || !data.customer || !data.actionPlan || !data.deadline) {
            alert('営業担当・販売店名・アクションプラン・アクション実行期限は必須です');
            return;
        }

        if (editId > 0) {
            const idx = state.progressItems.findIndex(item => item.id === editId);
            if (idx >= 0) {
                state.progressItems[idx] = normalizeProgressItem({ ...state.progressItems[idx], ...data, updatedAt: new Date().toISOString() }, editId);
            }
        } else {
            const item = normalizeProgressItem({ ...data, id: state.progressSeq, updatedAt: new Date().toISOString() }, state.progressSeq);
            state.progressItems.push(item);
            state.progressSeq += 1;
        }
        clearProgressForm();
        renderProgressTable();
        scheduleAutoStateSave(200);
    }

    function editProgressItemById(id) {
        const item = state.progressItems.find(v => v.id === id);
        if (!item) return;
        setInputValue('progress-edit-id', id);
        setInputValue('progress-rep', item.rep);
        setInputValue('progress-customer', item.customer);
        setInputValue('progress-action', item.actionPlan);
        setInputValue('progress-deadline', item.deadline || '');
        setInputValue('progress-result', item.result);
        setInputValue('progress-status', item.status);
        renderProgressFormSelectors();
    }

    function deleteProgressItemById(id) {
        const before = state.progressItems.length;
        state.progressItems = state.progressItems.filter(item => item.id !== id);
        if (state.progressItems.length === before) return;
        renderProgressTable();
        scheduleAutoStateSave(200);
    }

    // ── Render: Details ──
    function renderDetails() {
        const r = state.results; if (!r) return;
        document.getElementById(pfx('details-empty')).style.display = 'none';
        document.getElementById(pfx('details-content')).style.display = 'block';

        const makerF = document.getElementById(pfx('details-maker')).value;
        const sortKey = document.getElementById(pfx('details-sort')).value;
        const search = document.getElementById(pfx('details-search')).value.toLowerCase();
        const ml = { aron: 'アロン化成', pana: 'パナソニック', other: 'その他' };

        let entries = Object.values(r.productAgg);
        if (makerF !== 'all') entries = entries.filter(e => e.maker === makerF);
        if (search) entries = entries.filter(e => e.jan.toLowerCase().includes(search) || e.name.toLowerCase().includes(search));

        switch (sortKey) {
            case 'profit-desc': entries.sort((a, b) => b.gross - a.gross); break;
            case 'profit-asc': entries.sort((a, b) => a.gross - b.gross); break;
            case 'sales-desc': entries.sort((a, b) => b.sales - a.sales); break;
            case 'qty-desc': entries.sort((a, b) => b.qty - a.qty); break;
        }

        const tbody = document.getElementById(pfx('details-tbody'));
        tbody.innerHTML = '';
        for (const e of entries) {
            const avgP = e.priceCount > 0 ? e.priceSum / e.priceCount : 0;
            const rt = e.listPrice > 0 ? avgP / e.listPrice : 0;
            const pr = e.sales > 0 ? e.gross / e.sales : 0;
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>${e.jan}</td><td>${e.name}</td><td>${ml[e.maker] || e.maker}</td><td>${fmtYen(e.listPrice)}</td><td>${fmtYen(e.effectiveCost)}</td><td>${fmtYen(avgP)}</td><td>${fmtPct(rt)}</td><td>${fmtYen(e.shippingCost)}</td><td>${fmt(e.qty)}</td><td>${fmtYen(e.sales)}</td><td class="${e.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(e.gross)}</td><td>${fmtPct(pr)}</td>`;
            tbody.appendChild(tr);
        }
    }

    // ── Tab Switch (module内部) ──
    function switchModTab(tabId) {
        currentTab = tabId;
        const container = document.getElementById('page-' + MODULE_ID);
        container.querySelectorAll('.mod-tab').forEach(t => t.classList.remove('active'));
        container.querySelectorAll('.mod-nav-btn').forEach(b => b.classList.remove('active'));
        document.getElementById(pfx('tab-' + tabId))?.classList.add('active');
        container.querySelector(`.mod-nav-btn[data-mtab="${tabId}"]`)?.classList.add('active');

        if (state.results) {
            switch (tabId) {
                case 'overview': renderOverview(); break;
                case 'monthly': renderMonthly(); break;
                case 'store':
                    setTimeout(() => {
                        if (currentTab === 'store') renderStore();
                    }, 0);
                    break;
                case 'simulation': renderSimulation(); break;
                case 'store-detail': renderStoreDetail(); break;
                case 'details': renderDetails(); break;
                case 'progress': renderProgressTable(); break;
            }
        }
    }

    function checkAllLoaded() {
        const ok = state.shippingData.length > 0 && state.salesData.length > 0 && state.productData.length > 0;
        const analyzeBtn = document.getElementById(pfx('btn-analyze'));
        if (analyzeBtn) analyzeBtn.disabled = !ok;
        const shippingInput = document.getElementById(pfx('file-shipping'));
        if (shippingInput) shippingInput.disabled = !uploadUnlocked || state.shippingData.length > 0;
        const salesInput = document.getElementById(pfx('file-sales'));
        if (salesInput) salesInput.disabled = !uploadUnlocked;
        const productInput = document.getElementById(pfx('file-product'));
        if (productInput) productInput.disabled = !uploadUnlocked || state.productData.length > 0;
        const clearShippingBtn = document.getElementById(pfx('btn-shipping-clear'));
        if (clearShippingBtn) clearShippingBtn.disabled = !uploadUnlocked || state.shippingData.length === 0;
        const clearProductBtn = document.getElementById(pfx('btn-product-clear'));
        if (clearProductBtn) clearProductBtn.disabled = !uploadUnlocked || state.productData.length === 0;
        const settingsSaveBtn = document.getElementById(pfx('btn-settings-save'));
        if (settingsSaveBtn) settingsSaveBtn.disabled = !settingsUnlocked;
        const recalcBtn = document.getElementById(pfx('btn-recalc'));
        if (recalcBtn) recalcBtn.disabled = !settingsUnlocked || state.salesData.length === 0;
    }

    function setInputValue(id, value) {
        const el = document.getElementById(pfx(id));
        if (!el) return;
        el.value = value;
    }

    function restoreSavedSettings(saved) {
        state.appliedSettings = normalizeSettings(saved || cloneDefaultSettings());
        applySettingsToForm(state.appliedSettings);
        markSettingsDirty(false);
    }

    function updateUploadCardsByState() {
        const map = [
            { card: 'card-shipping', status: 'status-shipping', len: state.shippingData.length },
            { card: 'card-sales', status: 'status-sales', len: state.salesData.length },
            { card: 'card-product', status: 'status-product', len: state.productData.length }
        ];
        for (const item of map) {
            const cardEl = document.getElementById(pfx(item.card));
            const statusEl = document.getElementById(pfx(item.status));
            if (cardEl) cardEl.classList.toggle('loaded', item.len > 0);
            if (statusEl) statusEl.textContent = item.len > 0 ? `✓ ${item.len} rows` : 'No data';
        }
    }

    async function saveCloudState() {
        if (!window.KaientaiCloud || typeof window.KaientaiCloud.saveModuleState !== 'function') {
            alert('Cloud is not available');
            return;
        }
        if (!(window.KaientaiCloud.isReady && window.KaientaiCloud.isReady())) {
            alert('Firebase connection is not ready');
            return;
        }
        if (state.shippingData.length === 0 || state.salesData.length === 0 || state.productData.length === 0) {
            alert('Load all three files before cloud save');
            return;
        }

        const btn = document.getElementById(pfx('btn-cloud-save'));
        const oldText = btn ? btn.textContent : '';
        try {
            if (btn) { btn.disabled = true; btn.textContent = 'Saving...'; }
            const payload = {
                schemaVersion: 1,
                savedAt: new Date().toISOString(),
                shippingData: state.shippingData,
                salesData: state.salesData,
                productData: state.productData,
                progressItems: state.progressItems,
                progressSeq: state.progressSeq,
                progressDraft: readProgressDraftInputs(),
                settings: getSettings()
            };
            const meta = await window.KaientaiCloud.saveModuleState(MODULE_ID, payload);
            log(`Cloud save complete: ${meta.byteLength} bytes (${meta.chunkCount} chunks)`);
            alert('Cloud save complete');
        } catch (err) {
            log('Cloud save error: ' + err.message);
            alert('Cloud save failed: ' + err.message);
        } finally {
            if (btn) { btn.textContent = oldText || 'Cloud Save'; checkAllLoaded(); }
        }
    }

    async function loadCloudState() {
        if (!window.KaientaiCloud || typeof window.KaientaiCloud.loadModuleState !== 'function') {
            alert('Cloud is not available');
            return;
        }
        if (!(window.KaientaiCloud.isReady && window.KaientaiCloud.isReady())) {
            alert('Firebase connection is not ready');
            return;
        }

        const btn = document.getElementById(pfx('btn-cloud-load'));
        const oldText = btn ? btn.textContent : '';
        try {
            if (btn) { btn.disabled = true; btn.textContent = 'Loading...'; }
            const payload = await window.KaientaiCloud.loadModuleState(MODULE_ID);
            if (!payload) {
                alert('No cloud data found');
                return;
            }
            if (!Array.isArray(payload.shippingData) || !Array.isArray(payload.salesData) || !Array.isArray(payload.productData)) {
                throw new Error('Cloud data format is invalid');
            }

            state.shippingData = payload.shippingData;
            state.salesData = fillMissingSalesStoreCodes((payload.salesData || []).map(normalizeSalesRow));
            state.productData = payload.productData;
            state.progressItems = Array.isArray(payload.progressItems)
                ? payload.progressItems.map((item, idx) => normalizeProgressItem(item, idx + 1))
                : [];
            state.progressSeq = toNum(payload.progressSeq) || 1;
            ensureProgressSeq();
            state.results = null;
            state.storeBaseCache = {};
            state.storeViewRuntime = null;
            state.storeDetailRuntime = null;
            state.storeCurrentPage = 1;
            state.storeCurrentPageTotal = 1;
            logLines = [];

            restoreSavedSettings(payload.settings || {});
            applyProgressDraftInputs(payload.progressDraft || {});
            renderProgressFormSelectors();
            renderProgressTable();
            updateUploadCardsByState();
            checkAllLoaded();
            KaientaiM.updateModuleStatus(MODULE_ID, `Data loaded (${state.salesData.length})`, true);

            ['overview', 'monthly', 'store', 'store-detail', 'sim', 'details'].forEach(id => {
                const emp = document.getElementById(pfx(id + '-empty'));
                const con = document.getElementById(pfx(id + '-content'));
                if (emp) emp.style.display = '';
                if (con) con.style.display = 'none';
            });
            Object.keys(state.charts).forEach(k => destroyChart(state.charts, k));
            document.getElementById(pfx('load-log')).style.display = 'none';
            log('Cloud load complete');
            alert('Cloud load complete. Press Analyze.');
        } catch (err) {
            log('Cloud load error: ' + err.message);
            alert('Cloud load failed: ' + err.message);
        } finally {
            if (btn) { btn.disabled = false; btn.textContent = oldText || 'Cloud Load'; }
        }
    }

    function analyze() {
        if (state.shippingData.length === 0 || state.salesData.length === 0 || state.productData.length === 0) {
            alert('3つのデータすべてを読み込んでください。');
            return;
        }
        log('--- 分析開始 ---');
        runAnalysis();
        switchModTab('overview');
        log('--- 分析完了 ---');
    }

    // ── Build HTML ──
    function buildHTML(container) {
        container.innerHTML = `
        <div class="page-header">
            <h2>アロン・パナ分析</h2>
            <p class="page-desc">アロン化成・パナソニック 販売実績分析 — 送料込み実利益算出</p>
        </div>

        <div class="mod-nav">
            <button class="mod-nav-btn active" data-mtab="overview">全体概要</button>
            <button class="mod-nav-btn" data-mtab="monthly">月次分析</button>
            <button class="mod-nav-btn" data-mtab="store">販売店分析</button>
            <button class="mod-nav-btn" data-mtab="simulation">掛け率シミュレーション</button>
            <button class="mod-nav-btn" data-mtab="store-detail">販売店詳細分析</button>
            <button class="mod-nav-btn" data-mtab="details">商品別詳細</button>
            <button class="mod-nav-btn" data-mtab="progress">進捗管理</button>
            <button class="mod-nav-btn" data-mtab="upload">データ読込</button>
            <button class="mod-nav-btn" data-mtab="settings">設定</button>
        </div>

        <!-- Upload -->
        <div class="mod-tab" id="${pfx('tab-upload')}">
            <div class="tab-lock-banner">
                <span id="${pfx('upload-lock-state')}">ロック中: 読込実行にはパスワードが必要です</span>
                <button class="btn-secondary" id="${pfx('btn-upload-unlock')}">読込ロック解除</button>
            </div>
            <div class="upload-grid">
                <div class="upload-card" id="${pfx('card-shipping')}">
                    <div class="upload-icon">&#128666;</div>
                    <h3>送料マスターデータ</h3>
                    <p>A列:JAN / B列:商品名 / I列:サイズ帯 / J〜V列:エリア別送料（W列は未使用）</p>
                    <label class="upload-btn">ファイル選択<input type="file" accept=".xlsx,.xls,.csv" id="${pfx('file-shipping')}" hidden></label>
                    <div class="upload-status" id="${pfx('status-shipping')}">未読込</div>
                    <div class="upload-hint">送料マスタは1セット固定。差し替え時は解除ボタンを使用。</div>
                    <div class="action-bar"><button class="btn-secondary" id="${pfx('btn-shipping-clear')}" disabled>送料データ解除</button></div>
                </div>
                <div class="upload-card" id="${pfx('card-sales')}">
                    <div class="upload-icon">&#128200;</div>
                    <h3>販売実績データ</h3>
                    <p>A列:受注番号 / C列:仕入先コード / B列:受注日 / D列:販売店 / H列:JAN / I列:商品名 / K列:数量 / L列:単価 / M列:合計 / S列:メーカー / Z列:営業担当 / AB列:県名</p>
                    <label class="upload-btn">ファイル選択<input type="file" accept=".xlsx,.xls,.csv" id="${pfx('file-sales')}" hidden multiple></label>
                    <div class="upload-status" id="${pfx('status-sales')}">未読込</div>
                    <div class="upload-hint">※複数月のファイルを同時選択可能</div>
                </div>
                <div class="upload-card" id="${pfx('card-product')}">
                    <div class="upload-icon">&#128230;</div>
                    <h3>商品マスタ</h3>
                    <p>A列:JAN / D列:商品名 / H列:定価 / M列:原価 / O列:倉庫入原価</p>
                    <label class="upload-btn">ファイル選択<input type="file" accept=".xlsx,.xls,.csv" id="${pfx('file-product')}" hidden multiple></label>
                    <div class="upload-status" id="${pfx('status-product')}">未読込</div>
                    <div class="upload-hint">追加取込は不可。追加前に商品マスタをクリア。</div>
                    <div class="action-bar"><button class="btn-secondary" id="${pfx('btn-product-clear')}" disabled>商品マスタクリア</button></div>
                </div>
            </div>
            <div class="action-bar">
                <button class="btn-primary" id="${pfx('btn-analyze')}" disabled>分析開始</button>
            </div>
            <div id="${pfx('load-log')}" class="load-log" style="display:none;">
                <h3>読込ログ</h3>
                <pre id="${pfx('log-content')}"></pre>
            </div>
        </div>

        <!-- Settings -->
        <div class="mod-tab" id="${pfx('tab-settings')}">
            <div class="tab-lock-banner">
                <span id="${pfx('settings-lock-state')}">ロック中: 設定編集にはパスワードが必要です</span>
            </div>
            <div class="settings-lock-shell" id="${pfx('settings-lock-shell')}">
                <div class="settings-lock-overlay">
                    <div class="settings-lock-panel">
                        <p>設定内容の表示・編集にはパスワードが必要です。</p>
                        <button class="btn-primary" id="${pfx('btn-settings-unlock')}">設定ロック解除</button>
                    </div>
                </div>
                <div class="settings-lockable">
                    <div class="settings-grid">
                        <div class="setting-card rebate-setting-card">
                            <h3>リベート設定</h3>
                            <div class="setting-row"><label>アロン化成 リベート率 (%)</label><input type="number" id="${pfx('rebate-aron')}" value="0" step="0.1" min="0" max="100"></div>
                            <div class="setting-row"><label>パナソニック リベート率 (%)</label><input type="number" id="${pfx('rebate-pana')}" value="0" step="0.1" min="0" max="100"></div>
                            <p class="hint">月別固定加算（達成リベート金額・車扱い還元金）を月カード単位で設定できます。</p>
                            <div class="monthly-rebate-list" id="${pfx('monthly-rebate-body')}"></div>
                        </div>
                        <div class="setting-card">
                            <h3>マイナス要件</h3>
                            <div class="setting-row"><label>月額 倉庫引き取り費 (円)</label><input type="number" id="${pfx('warehouse-fee')}" value="0" step="100" min="0"></div>
                            <div class="setting-row"><label>倉庫出し手数料 (円/個)</label><input type="number" id="${pfx('warehouse-out-fee')}" value="50" step="1"></div>
                            <p class="hint">倉庫出し手数料は 総数量 × 単価 で粗利から減算します。</p>
                        </div>
                        <div class="setting-card">
                            <h3>送料ルール</h3>
                            <div class="setting-row"><label>サイズ帯100以下のデフォルト送料 (円)</label><input type="number" id="${pfx('default-shipping-small')}" value="100" step="10" min="0"></div>
                        </div>
                        <div class="setting-card">
                            <h3>メーカー判定キーワード（販売実績 S列から判定）</h3>
                            <div class="setting-row"><label>アロン化成 判定キーワード</label><input type="text" id="${pfx('keyword-aron')}" value="アロン"></div>
                            <div class="setting-row"><label>パナソニック 判定キーワード</label><input type="text" id="${pfx('keyword-pana')}" value="パナソニック,パナ,Panasonic"></div>
                            <p class="hint">※販売実績データのS列の値で判定。(株)等は自動除去。カンマ区切りで複数キーワード指定可。</p>
                        </div>
                    </div>
                    <div class="action-bar">
                        <button class="btn-primary" id="${pfx('btn-settings-save')}">設定保存</button>
                        <button class="btn-secondary" id="${pfx('btn-recalc')}">設定を反映して再計算</button>
                        <span class="settings-save-state" id="${pfx('settings-save-state')}">保存済み設定を使用中</span>
                    </div>
                </div>
            </div>
        </div>

        <!-- Overview -->
        <div class="mod-tab active" id="${pfx('tab-overview')}">
            <div id="${pfx('overview-empty')}" class="empty-state">データを読み込んで分析を実行してください</div>
            <div id="${pfx('overview-content')}" style="display:none;">
                <div class="kpi-grid">
                    <div class="kpi-card"><div class="kpi-label">総売上</div><div class="kpi-value" id="${pfx('kpi-total-sales')}">-</div></div>
                    <div class="kpi-card"><div class="kpi-label">総原価</div><div class="kpi-value" id="${pfx('kpi-total-cost')}">-</div></div>
                    <div class="kpi-card"><div class="kpi-label">総送料</div><div class="kpi-value" id="${pfx('kpi-total-shipping')}">-</div></div>
                    <div class="kpi-card highlight"><div class="kpi-label">商品粗利合計</div><div class="kpi-value" id="${pfx('kpi-total-gross')}">-</div></div>
                    <div class="kpi-card"><div class="kpi-label">リベート合計</div><div class="kpi-value" id="${pfx('kpi-total-rebate')}">-</div></div>
                    <div class="kpi-card"><div class="kpi-label">マイナス要件合計</div><div class="kpi-value" id="${pfx('kpi-total-warehouse')}">-</div></div>
                    <div class="kpi-card accent"><div class="kpi-label">本当の粗利（実利益）</div><div class="kpi-value" id="${pfx('kpi-real-profit')}">-</div></div>
                    <div class="kpi-card"><div class="kpi-label">実利益率</div><div class="kpi-value" id="${pfx('kpi-profit-rate')}">-</div></div>
                </div>
                <div class="chart-row">
                    <div class="chart-box"><h3>メーカー別 売上・粗利</h3><canvas id="${pfx('chart-maker-bar')}"></canvas></div>
                    <div class="chart-box"><h3>利益構成</h3><canvas id="${pfx('chart-profit-pie')}"></canvas></div>
                </div>
                <div class="chart-row"><div class="chart-box full"><h3>年間月次推移</h3><canvas id="${pfx('chart-annual')}"></canvas></div></div>
            </div>
        </div>

        <!-- Monthly -->
        <div class="mod-tab" id="${pfx('tab-monthly')}">
            <div id="${pfx('monthly-empty')}" class="empty-state">データを読み込んで分析を実行してください</div>
            <div id="${pfx('monthly-content')}" style="display:none;">
                <div class="filter-bar"><label>メーカー:</label><select id="${pfx('monthly-maker')}"><option value="all">全て</option><option value="aron">アロン化成</option><option value="pana">パナソニック</option><option value="other">その他</option></select></div>
                <div class="table-wrapper"><table><thead><tr><th>年月</th><th>メーカー</th><th>売上合計</th><th>原価合計</th><th>送料合計</th><th>商品粗利</th><th>リベート</th><th>マイナス要件</th><th>本当の粗利</th><th>実利益率</th></tr></thead><tbody id="${pfx('monthly-tbody')}"></tbody></table></div>
                <div class="chart-row"><div class="chart-box full"><h3>月次推移チャート</h3><canvas id="${pfx('chart-monthly')}"></canvas></div></div>
            </div>
        </div>

        <!-- Store -->
        <div class="mod-tab" id="${pfx('tab-store')}">
            <div id="${pfx('store-empty')}" class="empty-state">データを読み込んで分析を実行してください</div>
            <div id="${pfx('store-content')}" style="display:none;">
                <div class="filter-bar"><label>メーカー:</label><select id="${pfx('store-maker')}"><option value="all">全て</option><option value="aron">アロン化成</option><option value="pana">パナソニック</option></select><label>年月:</label><select id="${pfx('store-month')}"><option value="all">全期間</option></select><label>営業担当:</label><select id="${pfx('store-rep')}"><option value="all">全担当</option></select><label>検索:</label><input type="text" id="${pfx('store-search')}" placeholder="販売店名 / 仕入先コード"><label>並び替え:</label><select id="${pfx('store-sort')}"><option value="gross-desc">粗利(高い順)</option><option value="gross-asc">粗利(低い順)</option><option value="sales-desc">売上(高い順)</option><option value="sales-asc">売上(低い順)</option><option value="qty-desc">数量(多い順)</option><option value="qty-asc">数量(少ない順)</option><option value="rate-desc">粗利率(高い順)</option><option value="rate-asc">粗利率(低い順)</option><option value="aron-rate-desc">アロン掛率(高い順)</option><option value="aron-rate-asc">アロン掛率(低い順)</option><option value="pana-rate-desc">パナ掛率(高い順)</option><option value="pana-rate-asc">パナ掛率(低い順)</option><option value="rep-asc">担当者(昇順)</option><option value="rep-desc">担当者(降順)</option><option value="code-asc">仕入先コード(昇順)</option><option value="code-desc">仕入先コード(降順)</option><option value="store-asc">販売店名(昇順)</option><option value="store-desc">販売店名(降順)</option></select><label>表示件数:</label><select id="${pfx('store-limit')}"><option value="100">100</option><option value="300" selected>300</option><option value="1000">1000</option></select></div>
                <div class="store-meta-row">
                    <div class="hint" id="${pfx('store-summary')}"></div>
                    <div class="store-pagination" id="${pfx('store-pagination')}" style="display:none;">
                        <button type="button" class="btn-secondary" id="${pfx('store-page-prev')}">前へ</button>
                        <span class="store-page-status" id="${pfx('store-page-status')}">1 / 1ページ</span>
                        <button type="button" class="btn-secondary" id="${pfx('store-page-next')}">次へ</button>
                    </div>
                </div>
                <div class="table-wrapper"><table class="store-table"><thead><tr><th>販売店名</th><th>仕入先コード</th><th>営業担当者</th><th>アロン掛率</th><th>パナ掛率</th><th>売上合計</th><th>原価合計</th><th>送料合計</th><th>商品粗利</th><th>粗利率</th><th>数量合計</th></tr></thead><tbody id="${pfx('store-tbody')}"></tbody></table></div>
                <div class="chart-box full"><h3>販売店詳細分析は別タブ</h3><p class="hint">販売店分析を軽量化するため、1件指定の詳細シミュレーションは「販売店詳細分析」タブに分離しています。</p><div class="action-bar"><button class="btn-secondary" id="${pfx('btn-open-store-detail')}">販売店詳細分析を開く</button></div></div>
                <div class="chart-row"><div class="chart-box full"><h3>販売店別 粗利ランキング</h3><canvas id="${pfx('chart-store')}"></canvas></div></div>
            </div>
        </div>

        <!-- Simulation -->
        <div class="mod-tab" id="${pfx('tab-simulation')}">
            <div id="${pfx('sim-empty')}" class="empty-state">データを読み込んで分析を実行してください</div>
            <div id="${pfx('sim-content')}" style="display:none;">
                <div class="sim-controls">
                    <div class="sim-card"><h3>現在の平均掛け率</h3><div class="sim-current"><span>アロン化成: <strong id="${pfx('sim-cur-aron')}">-</strong></span><span>パナソニック: <strong id="${pfx('sim-cur-pana')}">-</strong></span><span>全体: <strong id="${pfx('sim-cur-all')}">-</strong></span></div></div>
                    <div class="sim-card">
                        <h3>掛け率・リベート変動シミュレーション</h3>
                        <div class="setting-row"><label>掛け率変動 (%)</label><input type="range" id="${pfx('sim-rate')}" min="-20" max="20" value="0" step="0.5"><span id="${pfx('sim-rate-display')}">±0%</span></div>
                        <div class="setting-row"><label>対象メーカー</label><select id="${pfx('sim-target')}"><option value="all">全体</option><option value="aron">アロン化成のみ</option><option value="pana">パナソニックのみ</option></select></div>
                        <div class="setting-row"><label>アロン リベート率増減 (pt)</label><input type="number" id="${pfx('sim-rebate-aron')}" value="0" step="0.1"></div>
                        <div class="setting-row"><label>パナ リベート率増減 (pt)</label><input type="number" id="${pfx('sim-rebate-pana')}" value="0" step="0.1"></div>
                        <div class="setting-row"><label>アロン 固定リベート増減 (円/月)</label><input type="number" id="${pfx('sim-fixed-aron')}" value="0" step="1000"></div>
                        <div class="setting-row"><label>パナ 固定リベート増減 (円/月)</label><input type="number" id="${pfx('sim-fixed-pana')}" value="0" step="1000"></div>
                    </div>
                </div>
                <div class="sim-result-grid">
                    <div class="sim-result-card"><div class="sim-label">変動前 実利益</div><div class="sim-value" id="${pfx('sim-before')}">-</div></div>
                    <div class="sim-result-card arrow">&#8594;</div>
                    <div class="sim-result-card"><div class="sim-label">変動後 実利益</div><div class="sim-value" id="${pfx('sim-after')}">-</div></div>
                    <div class="sim-result-card"><div class="sim-label">差額</div><div class="sim-value" id="${pfx('sim-diff')}">-</div></div>
                </div>
                <div class="table-wrapper"><table><thead><tr><th>指標</th><th>変動前</th><th>変動後</th><th>差額</th></tr></thead><tbody id="${pfx('sim-breakdown-body')}"></tbody></table></div>
                <div class="chart-row"><div class="chart-box full"><h3>掛け率 vs 実利益 推移</h3><canvas id="${pfx('chart-sim')}"></canvas></div></div>
            </div>
        </div>

        <!-- Store Detail -->
        <div class="mod-tab" id="${pfx('tab-store-detail')}">
            <div id="${pfx('store-detail-empty')}" class="empty-state">データを読み込んで分析を実行してください</div>
            <div id="${pfx('store-detail-content')}" style="display:none;">
                <div class="chart-box full">
                    <h3>販売店1件の詳細分析（軽量）</h3>
                    <div class="filter-bar">
                        <label>販売店名 / 仕入先コード検索:</label>
                        <input type="text" id="${pfx('store-detail-search')}" placeholder="例: ○○商事 / 12345">
                        <label>候補:</label>
                        <select id="${pfx('store-detail-select')}"></select>
                        <label>対象メーカー:</label>
                        <select id="${pfx('store-detail-maker')}"><option value="all">両メーカー</option><option value="aron">アロン化成のみ</option><option value="pana">パナソニックのみ</option></select>
                        <label>掛率変動(%):</label>
                        <input type="number" id="${pfx('store-detail-rate')}" value="0" step="0.1">
                        <label>予測期間(月):</label>
                        <input type="number" id="${pfx('store-detail-horizon')}" value="3" step="1" min="1" max="24">
                        <label>参照期間(月):</label>
                        <input type="number" id="${pfx('store-detail-window')}" value="6" step="1" min="1" max="24">
                        <label>手動増減率(%):</label>
                        <input type="number" id="${pfx('store-detail-manual-rate')}" value="0" step="0.1">
                        <label>追加数量(個/月):</label>
                        <input type="number" id="${pfx('store-detail-qty')}" value="0" step="1" min="0">
                        <label><input type="checkbox" id="${pfx('store-detail-use-trend')}" checked> トレンド反映</label>
                        <label>トレンド補正(%/月):</label>
                        <input type="number" id="${pfx('store-detail-trend-adjust')}" value="0" step="0.1">
                        <label><input type="checkbox" id="${pfx('store-detail-use-elastic-auto')}" checked> 弾力性自動推定</label>
                        <label>価格弾力性(手動):</label>
                        <input type="number" id="${pfx('store-detail-elasticity')}" value="-1.0" step="0.1">
                        <label><input type="checkbox" id="${pfx('store-detail-use-seasonality')}" checked> 季節性反映</label>
                    </div>
                    <div class="hint" id="${pfx('store-detail-summary')}"></div>
                    <div class="table-wrapper"><table class="store-sim-table"><thead><tr><th>区分</th><th>売上</th><th>粗利</th><th>粗利率</th><th>数量</th><th>アロン掛率</th><th>パナ掛率</th><th>対象</th></tr></thead><tbody id="${pfx('store-detail-tbody')}"></tbody></table></div>
                    <div class="table-wrapper"><table class="store-detail-factor-table"><thead><tr><th>予測要因</th><th>値</th><th>説明</th></tr></thead><tbody id="${pfx('store-detail-factor-body')}"></tbody></table></div>
                </div>
            </div>
        </div>

        <!-- Details -->
        <div class="mod-tab" id="${pfx('tab-details')}">
            <div id="${pfx('details-empty')}" class="empty-state">データを読み込んで分析を実行してください</div>
            <div id="${pfx('details-content')}" style="display:none;">
                <div class="filter-bar">
                    <label>メーカー:</label><select id="${pfx('details-maker')}"><option value="all">全て</option><option value="aron">アロン化成</option><option value="pana">パナソニック</option></select>
                    <label>並び替え:</label><select id="${pfx('details-sort')}"><option value="profit-desc">粗利(高い順)</option><option value="profit-asc">粗利(低い順)</option><option value="sales-desc">売上(高い順)</option><option value="qty-desc">数量(多い順)</option></select>
                    <label>検索:</label><input type="text" id="${pfx('details-search')}" placeholder="JANコードまたは商品名">
                </div>
                <div class="table-wrapper"><table><thead><tr><th>JANコード</th><th>商品名</th><th>メーカー</th><th>定価</th><th>原価</th><th>販売単価(平均)</th><th>掛け率</th><th>送料</th><th>数量合計</th><th>売上合計</th><th>粗利合計</th><th>粗利率</th></tr></thead><tbody id="${pfx('details-tbody')}"></tbody></table></div>
                <div class="action-bar"><button class="btn-secondary" id="${pfx('btn-export')}">CSVエクスポート</button></div>
            </div>
        </div>

        <!-- Progress -->
        <div class="mod-tab" id="${pfx('tab-progress')}">
            <div id="${pfx('progress-empty')}" class="empty-state" style="display:none;">進捗データがありません</div>
            <div id="${pfx('progress-content')}">
                <div class="chart-box full">
                    <h3>進捗管理</h3>
                    <input type="hidden" id="${pfx('progress-edit-id')}" value="">
                    <input type="hidden" id="${pfx('progress-rep')}" value="">
                    <input type="hidden" id="${pfx('progress-customer')}" value="">
                    <div class="progress-form-grid">
                        <div class="progress-form-field">
                            <label>営業担当 <span class="required">*</span></label>
                            <div class="toggle-chip-group" id="${pfx('progress-rep-toggle')}"></div>
                        </div>
                        <div class="progress-form-field">
                            <label>販売店名 <span class="required">*</span></label>
                            <div class="toggle-chip-group" id="${pfx('progress-customer-toggle')}"></div>
                        </div>
                        <div class="progress-form-field">
                            <label>アクション実行期限 <span class="required">*</span></label>
                            <input type="date" id="${pfx('progress-deadline')}">
                        </div>
                        <div class="progress-form-field">
                            <label>ステータス</label>
                            <select id="${pfx('progress-status')}"><option value="planned">計画中</option><option value="doing">実行中</option><option value="done">完了</option><option value="hold">保留</option></select>
                        </div>
                    </div>
                    <div class="setting-row progress-plan-row"><label>アクションプラン <span class="required">*</span></label><textarea id="${pfx('progress-action')}" rows="2" placeholder="例: 値上げ提案の事前打診"></textarea></div>
                    <div class="setting-row progress-plan-row"><label>実行結果</label><textarea id="${pfx('progress-result')}" rows="2" placeholder="例: 次回訪問で詳細見積依頼"></textarea></div>
                    <div class="action-bar">
                        <button class="btn-primary" id="${pfx('progress-save')}">登録 / 更新</button>
                        <button class="btn-secondary" id="${pfx('progress-clear-form')}">入力クリア</button>
                    </div>
                    <div class="filter-bar">
                        <label>営業担当検索</label><input type="text" id="${pfx('progress-filter-rep-search')}" placeholder="例: 田中">
                        <label>営業担当候補</label><select id="${pfx('progress-filter-rep')}"><option value="all">全担当</option></select>
                        <label>販売店検索</label><input type="text" id="${pfx('progress-filter-customer-search')}" placeholder="例: ○○商事 / 12345">
                        <label>販売店候補</label><select id="${pfx('progress-filter-customer')}"><option value="all">全販売店</option></select>
                    </div>
                    <div class="filter-bar">
                        <label>ステータス</label><select id="${pfx('progress-filter-status')}"><option value="all">全て</option><option value="planned">計画中</option><option value="doing">実行中</option><option value="done">完了</option><option value="hold">保留</option></select>
                        <label>自由検索</label><input type="text" id="${pfx('progress-filter-search')}" placeholder="担当/顧客/内容/期限">
                    </div>
                    <div class="table-wrapper"><table class="progress-table"><thead><tr><th>営業担当</th><th>顧客</th><th>アクションプラン</th><th>実行期限</th><th>結果</th><th>状態</th><th>更新日時</th><th>編集</th><th>削除</th></tr></thead><tbody id="${pfx('progress-tbody')}"></tbody></table></div>
                </div>
            </div>
        </div>
        `;
    }

    // ── Bind Events ──
    function bindEvents(container) {
        // Sub-tab nav
        container.querySelectorAll('.mod-nav-btn').forEach(btn => {
            btn.addEventListener('click', () => switchModTab(btn.dataset.mtab));
        });

        const uploadUnlockBtn = document.getElementById(pfx('btn-upload-unlock'));
        if (uploadUnlockBtn) uploadUnlockBtn.addEventListener('click', () => unlockTabGroup('upload'));
        const settingsUnlockBtn = document.getElementById(pfx('btn-settings-unlock'));
        if (settingsUnlockBtn) settingsUnlockBtn.addEventListener('click', () => unlockTabGroup('settings'));

        // File uploads
        document.getElementById(pfx('file-shipping')).addEventListener('change', async (e) => {
            if (!e.target.files.length) return;
            if (!ensureUploadUnlocked()) { e.target.value = ''; return; }
            if (state.shippingData.length > 0) {
                alert('送料マスタは1セット固定です。先に「送料データ解除」を実行してください。');
                e.target.value = '';
                return;
            }
            try { loadShipping(await parseExcel(e.target.files[0])); checkAllLoaded(); e.target.value = ''; }
            catch (err) { log('送料マスタ読込エラー: ' + err.message); alert('送料マスタの読込に失敗しました'); }
        });
        document.getElementById(pfx('file-sales')).addEventListener('change', async (e) => {
            if (!e.target.files.length) return;
            if (!ensureUploadUnlocked()) { e.target.value = ''; return; }
            try {
                const list = [];
                for (const f of Array.from(e.target.files)) list.push(await parseExcel(f));
                loadSales(list); checkAllLoaded();
                e.target.value = '';
            } catch (err) { log('販売実績読込エラー: ' + err.message); alert('販売実績の読込に失敗しました'); }
        });
        document.getElementById(pfx('file-product')).addEventListener('change', async (e) => {
            if (!e.target.files.length) return;
            if (!ensureUploadUnlocked()) { e.target.value = ''; return; }
            if (state.productData.length > 0) {
                alert('商品マスタ追加はできません。先に「商品マスタクリア」を実行してください。');
                e.target.value = '';
                return;
            }
            try {
                const list = [];
                for (const f of Array.from(e.target.files)) list.push(await parseExcel(f));
                loadProduct(list); checkAllLoaded();
                e.target.value = '';
            }
            catch (err) { log('商品マスタ読込エラー: ' + err.message); alert('商品マスタの読込に失敗しました'); }
        });

        document.getElementById(pfx('btn-analyze')).addEventListener('click', analyze);
        document.getElementById(pfx('btn-shipping-clear')).addEventListener('click', () => {
            if (!ensureUploadUnlocked()) return;
            if (state.shippingData.length === 0) return;
            state.shippingData = [];
            resetAnalysisOutputs('送料解除（再分析待ち）');
            updateUploadCardsByState();
            checkAllLoaded();
            const shippingInput = document.getElementById(pfx('file-shipping'));
            if (shippingInput) shippingInput.value = '';
            scheduleAutoStateSave();
        });
        document.getElementById(pfx('btn-product-clear')).addEventListener('click', () => {
            if (!ensureUploadUnlocked()) return;
            if (state.productData.length === 0) return;
            state.productData = [];
            resetAnalysisOutputs('商品マスタクリア（再分析待ち）');
            updateUploadCardsByState();
            checkAllLoaded();
            const productInput = document.getElementById(pfx('file-product'));
            if (productInput) productInput.value = '';
            scheduleAutoStateSave();
        });
        document.getElementById(pfx('btn-recalc')).addEventListener('click', () => {
            if (!ensureSettingsUnlocked()) return;
            if (state.salesData.length === 0) { alert('データを先に読み込んでください。'); return; }
            if (settingsDirty) {
                alert('未保存の設定があります。先に「設定保存」を押してください。');
                return;
            }
            analyze();
        });
        const settingsSaveBtn = document.getElementById(pfx('btn-settings-save'));
        if (settingsSaveBtn) {
            settingsSaveBtn.addEventListener('click', () => {
                if (!saveSettingsFromForm()) return;
                alert('設定を保存しました');
            });
        }

        const settingsTabEl = document.getElementById(pfx('tab-settings'));
        if (settingsTabEl) {
            const onSettingsChanged = (e) => {
                const target = e.target;
                if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement || target instanceof HTMLTextAreaElement)) return;
                const draft = readSettingsFromForm();
                markSettingsDirty(serializeSettings(draft) !== serializeSettings(state.appliedSettings));
            };
            settingsTabEl.addEventListener('input', onSettingsChanged);
            settingsTabEl.addEventListener('change', onSettingsChanged);
        }
        const onProgressChanged = (e) => {
            const target = e.target;
            if (!(target instanceof HTMLElement)) return;
            if (!target.id || !target.id.startsWith(pfx('progress-'))) return;
            if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement || target instanceof HTMLTextAreaElement)) return;
            scheduleAutoStateSave(500);
        };
        container.addEventListener('input', onProgressChanged);
        container.addEventListener('change', onProgressChanged);

        // Filters
        document.getElementById(pfx('monthly-maker')).addEventListener('change', renderMonthly);
        const resetStorePageAndRender = () => { state.storeCurrentPage = 1; renderStore(); };
        document.getElementById(pfx('store-maker')).addEventListener('change', resetStorePageAndRender);
        document.getElementById(pfx('store-month')).addEventListener('change', resetStorePageAndRender);
        document.getElementById(pfx('store-rep')).addEventListener('change', resetStorePageAndRender);
        document.getElementById(pfx('store-search')).addEventListener('input', resetStorePageAndRender);
        document.getElementById(pfx('store-sort')).addEventListener('change', resetStorePageAndRender);
        document.getElementById(pfx('store-limit')).addEventListener('change', resetStorePageAndRender);
        document.getElementById(pfx('store-page-prev')).addEventListener('click', () => {
            if (state.storeCurrentPage <= 1) return;
            state.storeCurrentPage -= 1;
            renderStore();
        });
        document.getElementById(pfx('store-page-next')).addEventListener('click', () => {
            if (state.storeCurrentPage >= state.storeCurrentPageTotal) return;
            state.storeCurrentPage += 1;
            renderStore();
        });
        const openStoreDetailBtn = document.getElementById(pfx('btn-open-store-detail'));
        if (openStoreDetailBtn) openStoreDetailBtn.addEventListener('click', () => switchModTab('store-detail'));
        const oldStoreSimStore = document.getElementById(pfx('store-sim-store'));
        if (oldStoreSimStore) oldStoreSimStore.addEventListener('change', renderStoreSimulationFromCurrent);
        const oldStoreSimMaker = document.getElementById(pfx('store-sim-maker'));
        if (oldStoreSimMaker) oldStoreSimMaker.addEventListener('change', renderStoreSimulationFromCurrent);
        const oldStoreSimRate = document.getElementById(pfx('store-sim-rate'));
        if (oldStoreSimRate) oldStoreSimRate.addEventListener('input', renderStoreSimulationFromCurrent);
        const oldStoreSimQty = document.getElementById(pfx('store-sim-qty'));
        if (oldStoreSimQty) oldStoreSimQty.addEventListener('input', renderStoreSimulationFromCurrent);
        document.getElementById(pfx('details-maker')).addEventListener('change', renderDetails);
        document.getElementById(pfx('details-sort')).addEventListener('change', renderDetails);
        document.getElementById(pfx('details-search')).addEventListener('input', renderDetails);
        document.getElementById(pfx('sim-rate')).addEventListener('input', renderSimulation);
        document.getElementById(pfx('sim-target')).addEventListener('change', renderSimulation);
        document.getElementById(pfx('sim-rebate-aron')).addEventListener('input', renderSimulation);
        document.getElementById(pfx('sim-rebate-pana')).addEventListener('input', renderSimulation);
        document.getElementById(pfx('sim-fixed-aron')).addEventListener('input', renderSimulation);
        document.getElementById(pfx('sim-fixed-pana')).addEventListener('input', renderSimulation);

        document.getElementById(pfx('store-detail-search')).addEventListener('input', renderStoreDetail);
        document.getElementById(pfx('store-detail-select')).addEventListener('change', renderStoreDetail);
        document.getElementById(pfx('store-detail-maker')).addEventListener('change', renderStoreDetail);
        document.getElementById(pfx('store-detail-rate')).addEventListener('input', renderStoreDetail);
        document.getElementById(pfx('store-detail-horizon')).addEventListener('input', renderStoreDetail);
        document.getElementById(pfx('store-detail-window')).addEventListener('input', renderStoreDetail);
        document.getElementById(pfx('store-detail-manual-rate')).addEventListener('input', renderStoreDetail);
        document.getElementById(pfx('store-detail-qty')).addEventListener('input', renderStoreDetail);
        document.getElementById(pfx('store-detail-use-trend')).addEventListener('change', renderStoreDetail);
        document.getElementById(pfx('store-detail-trend-adjust')).addEventListener('input', renderStoreDetail);
        document.getElementById(pfx('store-detail-use-elastic-auto')).addEventListener('change', renderStoreDetail);
        document.getElementById(pfx('store-detail-elasticity')).addEventListener('input', renderStoreDetail);
        document.getElementById(pfx('store-detail-use-seasonality')).addEventListener('change', renderStoreDetail);

        document.getElementById(pfx('progress-save')).addEventListener('click', saveProgressItem);
        document.getElementById(pfx('progress-clear-form')).addEventListener('click', () => {
            clearProgressForm();
            scheduleAutoStateSave(200);
        });
        document.getElementById(pfx('progress-filter-rep-search')).addEventListener('input', renderProgressTable);
        document.getElementById(pfx('progress-filter-rep')).addEventListener('change', renderProgressTable);
        document.getElementById(pfx('progress-filter-customer-search')).addEventListener('input', renderProgressTable);
        document.getElementById(pfx('progress-filter-customer')).addEventListener('change', renderProgressTable);
        document.getElementById(pfx('progress-filter-status')).addEventListener('change', renderProgressTable);
        document.getElementById(pfx('progress-filter-search')).addEventListener('input', renderProgressTable);
        document.getElementById(pfx('progress-rep-toggle')).addEventListener('click', (e) => {
            const target = e.target;
            if (!(target instanceof HTMLElement)) return;
            const rep = toStr(target.getAttribute('data-progress-rep-value'));
            if (!rep) return;
            setInputValue('progress-rep', rep);
            setInputValue('progress-customer', '');
            renderProgressFormSelectors();
            scheduleAutoStateSave(250);
        });
        document.getElementById(pfx('progress-customer-toggle')).addEventListener('click', (e) => {
            const target = e.target;
            if (!(target instanceof HTMLElement)) return;
            const customer = toStr(target.getAttribute('data-progress-customer-value'));
            if (!customer) return;
            setInputValue('progress-customer', customer);
            renderProgressFormSelectors();
            scheduleAutoStateSave(250);
        });
        document.getElementById(pfx('progress-tbody')).addEventListener('click', (e) => {
            const target = e.target;
            if (!(target instanceof HTMLElement)) return;
            const editId = toNum(target.getAttribute('data-progress-edit'));
            const deleteId = toNum(target.getAttribute('data-progress-delete'));
            if (editId > 0) {
                editProgressItemById(editId);
            } else if (deleteId > 0) {
                deleteProgressItemById(deleteId);
            }
        });

        // Export
        document.getElementById(pfx('btn-export')).addEventListener('click', () => {
            const r = state.results; if (!r) return;
            const ml = { aron: 'アロン化成', pana: 'パナソニック', other: 'その他' };
            const header = ['JANコード', '商品名', 'メーカー', '定価', '原価', '送料', '数量合計', '売上合計', '粗利合計'];
            const rows = Object.values(r.productAgg).map(e => [e.jan, e.name, ml[e.maker] || e.maker, e.listPrice, e.effectiveCost, e.shippingCost, e.qty, e.sales, e.gross]);
            exportCSV(header, rows, 'aron_pana_export.csv');
        });
    }

    // ── Register Module ──
    KaientaiM.registerModule({
        id: MODULE_ID,
        title: 'アロン・パナ分析',
        icon: '&#128202;',
        description: 'アロン化成・パナソニック販売実績の送料込み実利益分析。月次・販売店別・掛け率シミュレーション対応。',
        color: '#1565c0',
        init(container) {
            buildHTML(container);
            bindEvents(container);
            restoreSavedSettings(cloneDefaultSettings());
            renderProgressFormSelectors();
            renderProgressTable();
            applyTabLockState();
            const restoredLocal = restoreAutoState();
            if (!restoredLocal) restoreCloudStateIfNeeded();
            scheduleCloudStateSave(2500);
            checkAllLoaded();
        },
        onShow() {
            if (currentTab === 'progress') {
                switchModTab('progress');
                return;
            }
            if (state.results && currentTab !== 'upload' && currentTab !== 'settings') {
                switchModTab(currentTab);
            }
        }
    });

})();
