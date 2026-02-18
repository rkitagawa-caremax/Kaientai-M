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

    // ── Module-local state ──
    const state = {
        shippingData: [],
        salesData: [],
        productData: [],
        results: null,
        charts: {},
        storeBaseCache: {},
        storeViewRuntime: null,
        storeCurrentPage: 1,
        storeCurrentPageTotal: 1
    };

    let logLines = [];
    let currentTab = 'upload';
    const AUTO_STATE_STORAGE_KEY = 'kaientai-aron-pana-autostate-v1';
    let autoPersistTimer = null;
    let cloudPersistTimer = null;
    let cloudPersistInFlight = false;
    let cloudPersistPending = false;

    function pfx(id) { return MODULE_ID + '-' + id; }
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
        state.storeCurrentPage = 1;
        state.storeCurrentPageTotal = 1;

        Object.keys(state.charts).forEach(k => destroyChart(state.charts, k));
        ['overview', 'monthly', 'store', 'sim', 'details'].forEach(id => {
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

    function buildAutoStatePayload() {
        return {
            schemaVersion: 2,
            savedAt: new Date().toISOString(),
            shippingData: state.shippingData,
            salesData: state.salesData,
            productData: state.productData,
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
        state.salesData = payload.salesData;
        state.productData = payload.productData;
        resetAnalysisOutputs('データ復元済（再分析待ち）');
        restoreSavedSettings(payload.settings || {});
        applyProgressDraftInputs(payload.progressDraft || {});
        updateUploadCardsByState();
        checkAllLoaded();
        if (sourceLabel) log(sourceLabel);
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

    function renderMonthlyRebateInputs() {
        const container = document.getElementById(pfx('monthly-rebate-body'));
        if (!container) return;

        const oldValues = {};
        container.querySelectorAll('input[data-month][data-maker][data-type]').forEach(input => {
            const k = `${input.dataset.month}|${input.dataset.maker}|${input.dataset.type}`;
            oldValues[k] = toNum(input.value);
        });

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
                const val = oldValues[key] ?? 0;
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
        return monthlyRebates;
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

    function getSettings() {
        return {
            rebateAron: toNum(document.getElementById(pfx('rebate-aron'))?.value) / 100,
            rebatePana: toNum(document.getElementById(pfx('rebate-pana'))?.value) / 100,
            warehouseFee: toNum(document.getElementById(pfx('warehouse-fee'))?.value),
            warehouseOutFee: toNum(document.getElementById(pfx('warehouse-out-fee'))?.value || 50),
            monthlyRebates: readMonthlyRebateSettings(),
            defaultShippingSmall: toNum(document.getElementById(pfx('default-shipping-small'))?.value),
            keywordAron: (document.getElementById(pfx('keyword-aron'))?.value || 'アロン').split(',').map(s => s.trim().toLowerCase()).filter(Boolean),
            keywordPana: (document.getElementById(pfx('keyword-pana'))?.value || 'パナソニック,パナ,Panasonic').split(',').map(s => s.trim().toLowerCase()).filter(Boolean),
        };
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
                    log(`    行${i}: A=[${toStr(r[COL.A])}] B=[${toStr(r[COL.B])}] D=[${toStr(r[COL.D])}] H=[${toStr(r[COL.H])}] I=[${toStr(r[COL.I])}] K=[${toStr(r[COL.K])}] L=[${toStr(r[COL.L])}] S=[${toStr(r[COL.S])}] Z=[${toStr(r[COL_SALES_REP])}] AB=[${toStr(r[COL_AB])}]`);
                }

                const headerRow = findHeaderRow(rows, ['jan', 'janコード', '商品', 'コード', '品番', '数量', '販売', '受注']);
                log(`  ヘッダー行: ${headerRow}行目`);

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

        // 診断情報
        log(`検出月一覧: [${[...monthValues].sort().join(', ')}]`);
        log(`S列メーカー表記一覧: [${[...makerValues].join(' / ')}]`);
        const aronCount = state.salesData.filter(s => s.maker === 'aron').length;
        const panaCount = state.salesData.filter(s => s.maker === 'pana').length;
        const otherCount = state.salesData.filter(s => s.maker === 'other').length;
        log(`メーカー判定結果: アロン=${aronCount}件 / パナ=${panaCount}件 / その他=${otherCount}件`);
        log(`販売実績追加: ${addedCount}件 / 重複受注番号スキップ: ${duplicateOrderCount}件 / 累計: ${state.salesData.length}件`);

        document.getElementById(pfx('status-sales')).textContent = `✓ ${state.salesData.length}件 (+${addedCount}件 / ${parsedList.length}ファイル)`;
        document.getElementById(pfx('card-sales')).classList.add('loaded');
        renderMonthlyRebateInputs();
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

        const monthlyAgg = {}, storeAgg = {}, productAgg = {};
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
            records, monthlyAgg, storeAgg, productAgg, months,
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
            return {
                store: e.store,
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
        const limitRaw = document.getElementById(pfx('store-limit')).value;

        const base = getStoreBase(r.records, makerF, monthF);
        repSel.innerHTML = '<option value="all">全担当</option>';
        for (const rep of base.reps) repSel.innerHTML += `<option value="${rep}">${rep}</option>`;
        repSel.value = base.reps.includes(prevRep) ? prevRep : 'all';
        const repF = repSel.value;

        if (!base.entriesByRep[repF]) {
            const scopedRecords = base.recordsByRep[repF] || [];
            base.entriesByRep[repF] = buildStoreEntries(scopedRecords);
        }

        const entries = [...base.entriesByRep[repF]];
        sortStoreEntries(entries, sortKey);

        const isAll = limitRaw === 'all';
        const pageSize = isAll ? Math.max(1, entries.length || 1) : Math.max(1, toNum(limitRaw) || 300);
        const totalPages = isAll ? 1 : Math.max(1, Math.ceil(entries.length / pageSize));
        state.storeCurrentPageTotal = totalPages;
        state.storeCurrentPage = isAll ? 1 : Math.min(Math.max(1, state.storeCurrentPage || 1), totalPages);
        const startIndex = isAll ? 0 : (state.storeCurrentPage - 1) * pageSize;
        const displayed = entries.slice(startIndex, startIndex + pageSize);
        const tbody = document.getElementById(pfx('store-tbody'));
        tbody.innerHTML = displayed.map(e =>
            `<tr><td>${e.store}</td><td>${e.salesRep}</td><td>${e.aronRate > 0 ? fmtPct(e.aronRate) : '-'}</td><td>${e.panaRate > 0 ? fmtPct(e.panaRate) : '-'}</td><td>${fmtYen(e.sales)}</td><td>${fmtYen(e.cost)}</td><td>${fmtYen(e.shipping)}</td><td class="${e.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(e.gross)}</td><td>${fmtPct(e.rate)}</td><td>${fmt(e.qty)}</td></tr>`
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

        const scopedRecords = base.recordsByRep[repF] || [];
        state.storeViewRuntime = { scopedRecords, entries };
        renderStoreSimulation(scopedRecords, entries);

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
        document.getElementById(pfx('sim-rate-display')).textContent = (rateChange >= 0 ? '+' : '') + (rateChange * 100).toFixed(1) + '%';

        let beforeG = 0, afterG = 0;
        for (const rec of r.records) {
            beforeG += rec.grossProfit;
            if (target === 'all' || rec.maker === target) {
                afterG += rec.unitPrice * (1 + rateChange) * rec.qty - rec.totalCost - rec.totalShipping;
            } else {
                afterG += rec.grossProfit;
            }
        }
        const diff = afterG - beforeG;
        document.getElementById(pfx('sim-before')).textContent = fmtYen(beforeG);
        document.getElementById(pfx('sim-after')).textContent = fmtYen(afterG);
        document.getElementById(pfx('sim-after')).className = 'sim-value ' + (afterG >= 0 ? 'positive' : 'negative');
        document.getElementById(pfx('sim-diff')).textContent = (diff >= 0 ? '+' : '') + fmtYen(diff);
        document.getElementById(pfx('sim-diff')).className = 'sim-value ' + (diff >= 0 ? 'positive' : 'negative');

        const steps = [], gv = [];
        for (let pct = -20; pct <= 20; pct += 2) {
            steps.push((pct >= 0 ? '+' : '') + pct + '%');
            let g = 0;
            for (const rec of r.records) {
                if (target === 'all' || rec.maker === target) g += rec.unitPrice * (1 + pct / 100) * rec.qty - rec.totalCost - rec.totalShipping;
                else g += rec.grossProfit;
            }
            gv.push(g);
        }
        destroyChart(state.charts, 'simulation');
        state.charts['simulation'] = new Chart(document.getElementById(pfx('chart-sim')), {
            type: 'line',
            data: { labels: steps, datasets: [{ label: '粗利', data: gv, borderColor: '#ffa726', backgroundColor: 'rgba(255,167,38,0.1)', fill: true, tension: 0.3, borderWidth: 3, pointRadius: 4, pointBackgroundColor: gv.map(v => v >= 0 ? '#66bb6a' : '#ef5350') }] },
            options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { ticks: { callback: v => '¥' + fmt(v) } } } }
        });
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
                case 'details': renderDetails(); break;
            }
        }
    }

    function checkAllLoaded() {
        const ok = state.shippingData.length > 0 && state.salesData.length > 0 && state.productData.length > 0;
        const analyzeBtn = document.getElementById(pfx('btn-analyze'));
        if (analyzeBtn) analyzeBtn.disabled = !ok;
        const shippingInput = document.getElementById(pfx('file-shipping'));
        if (shippingInput) shippingInput.disabled = state.shippingData.length > 0;
        const productInput = document.getElementById(pfx('file-product'));
        if (productInput) productInput.disabled = state.productData.length > 0;
        const clearShippingBtn = document.getElementById(pfx('btn-shipping-clear'));
        if (clearShippingBtn) clearShippingBtn.disabled = state.shippingData.length === 0;
        const clearProductBtn = document.getElementById(pfx('btn-product-clear'));
        if (clearProductBtn) clearProductBtn.disabled = state.productData.length === 0;
    }

    function setInputValue(id, value) {
        const el = document.getElementById(pfx(id));
        if (!el) return;
        el.value = value;
    }

    function restoreSavedSettings(saved) {
        if (!saved || typeof saved !== 'object') {
            renderMonthlyRebateInputs();
            return;
        }

        setInputValue('rebate-aron', toNum(saved.rebateAron) * 100);
        setInputValue('rebate-pana', toNum(saved.rebatePana) * 100);
        setInputValue('warehouse-fee', toNum(saved.warehouseFee));
        setInputValue('warehouse-out-fee', toNum(saved.warehouseOutFee) || 50);
        setInputValue('default-shipping-small', toNum(saved.defaultShippingSmall));

        const kwAron = Array.isArray(saved.keywordAron) ? saved.keywordAron.join(',') : '';
        const kwPana = Array.isArray(saved.keywordPana) ? saved.keywordPana.join(',') : '';
        if (kwAron) setInputValue('keyword-aron', kwAron);
        if (kwPana) setInputValue('keyword-pana', kwPana);

        renderMonthlyRebateInputs();
        const wrap = document.getElementById(pfx('monthly-rebate-body'));
        if (!wrap) return;
        wrap.querySelectorAll('input[data-month][data-maker][data-type]').forEach(input => {
            const month = input.dataset.month;
            const maker = input.dataset.maker;
            const type = input.dataset.type;
            input.value = toNum(saved.monthlyRebates?.[month]?.[maker]?.[type]);
        });
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
            state.salesData = payload.salesData;
            state.productData = payload.productData;
            state.results = null;
            state.storeBaseCache = {};
            state.storeViewRuntime = null;
            state.storeCurrentPage = 1;
            state.storeCurrentPageTotal = 1;
            logLines = [];

            restoreSavedSettings(payload.settings || {});
            updateUploadCardsByState();
            checkAllLoaded();
            KaientaiM.updateModuleStatus(MODULE_ID, `Data loaded (${state.salesData.length})`, true);

            ['overview', 'monthly', 'store', 'sim', 'details'].forEach(id => {
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
            <button class="mod-nav-btn active" data-mtab="upload">データ読込</button>
            <button class="mod-nav-btn" data-mtab="settings">設定</button>
            <button class="mod-nav-btn" data-mtab="overview">全体概要</button>
            <button class="mod-nav-btn" data-mtab="monthly">月次分析</button>
            <button class="mod-nav-btn" data-mtab="store">販売店分析</button>
            <button class="mod-nav-btn" data-mtab="simulation">掛け率シミュレーション</button>
            <button class="mod-nav-btn" data-mtab="details">商品別詳細</button>
        </div>

        <!-- Upload -->
        <div class="mod-tab active" id="${pfx('tab-upload')}">
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
                    <p>A列:受注番号 / B列:受注日 / D列:販売店 / H列:JAN / I列:商品名 / K列:数量 / L列:単価 / M列:合計 / S列:メーカー / Z列:営業担当 / AB列:県名</p>
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
                <button class="btn-primary" id="${pfx('btn-recalc')}">再計算</button>
            </div>
        </div>

        <!-- Overview -->
        <div class="mod-tab" id="${pfx('tab-overview')}">
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
                <div class="filter-bar"><label>メーカー:</label><select id="${pfx('store-maker')}"><option value="all">全て</option><option value="aron">アロン化成</option><option value="pana">パナソニック</option></select><label>年月:</label><select id="${pfx('store-month')}"><option value="all">全期間</option></select><label>営業担当:</label><select id="${pfx('store-rep')}"><option value="all">全担当</option></select><label>並び替え:</label><select id="${pfx('store-sort')}"><option value="gross-desc">粗利(高い順)</option><option value="gross-asc">粗利(低い順)</option><option value="sales-desc">売上(高い順)</option><option value="sales-asc">売上(低い順)</option><option value="qty-desc">数量(多い順)</option><option value="qty-asc">数量(少ない順)</option><option value="rate-desc">粗利率(高い順)</option><option value="rate-asc">粗利率(低い順)</option><option value="aron-rate-desc">アロン掛率(高い順)</option><option value="aron-rate-asc">アロン掛率(低い順)</option><option value="pana-rate-desc">パナ掛率(高い順)</option><option value="pana-rate-asc">パナ掛率(低い順)</option><option value="rep-asc">担当者(昇順)</option><option value="rep-desc">担当者(降順)</option><option value="store-asc">販売店名(昇順)</option><option value="store-desc">販売店名(降順)</option></select><label>表示件数:</label><select id="${pfx('store-limit')}"><option value="300">300</option><option value="1000">1000</option><option value="all">全件</option></select></div>
                <div class="store-meta-row">
                    <div class="hint" id="${pfx('store-summary')}"></div>
                    <div class="store-pagination" id="${pfx('store-pagination')}" style="display:none;">
                        <button type="button" class="btn-secondary" id="${pfx('store-page-prev')}">前へ</button>
                        <span class="store-page-status" id="${pfx('store-page-status')}">1 / 1ページ</span>
                        <button type="button" class="btn-secondary" id="${pfx('store-page-next')}">次へ</button>
                    </div>
                </div>
                <div class="table-wrapper"><table><thead><tr><th>販売店名</th><th>営業担当者</th><th>アロン掛率</th><th>パナ掛率</th><th>売上合計</th><th>原価合計</th><th>送料合計</th><th>商品粗利</th><th>粗利率</th><th>数量合計</th></tr></thead><tbody id="${pfx('store-tbody')}"></tbody></table></div>
                <div class="chart-box full">
                    <h3>販売店別 掛率シミュレーション</h3>
                    <div class="filter-bar"><label>販売店:</label><select id="${pfx('store-sim-store')}"></select><label>対象メーカー:</label><select id="${pfx('store-sim-maker')}"><option value="all">両メーカー</option><option value="aron">アロン化成のみ</option><option value="pana">パナソニックのみ</option></select><label>掛率変動(%):</label><input type="number" id="${pfx('store-sim-rate')}" value="0" step="0.1"><label>予想販売増加数(個):</label><input type="number" id="${pfx('store-sim-qty')}" value="0" step="1" min="0"></div>
                    <div class="table-wrapper"><table class="store-sim-table"><thead><tr><th>区分</th><th>売上</th><th>粗利</th><th>粗利率</th><th>数量</th><th>アロン掛率</th><th>パナ掛率</th></tr></thead><tbody id="${pfx('store-sim-tbody')}"></tbody></table></div>
                </div>
                <div class="chart-row"><div class="chart-box full"><h3>販売店別 粗利ランキング</h3><canvas id="${pfx('chart-store')}"></canvas></div></div>
            </div>
        </div>

        <!-- Simulation -->
        <div class="mod-tab" id="${pfx('tab-simulation')}">
            <div id="${pfx('sim-empty')}" class="empty-state">データを読み込んで分析を実行してください</div>
            <div id="${pfx('sim-content')}" style="display:none;">
                <div class="sim-controls">
                    <div class="sim-card"><h3>現在の平均掛け率</h3><div class="sim-current"><span>アロン化成: <strong id="${pfx('sim-cur-aron')}">-</strong></span><span>パナソニック: <strong id="${pfx('sim-cur-pana')}">-</strong></span><span>全体: <strong id="${pfx('sim-cur-all')}">-</strong></span></div></div>
                    <div class="sim-card"><h3>掛け率変動シミュレーション</h3><div class="setting-row"><label>掛け率変動 (%)</label><input type="range" id="${pfx('sim-rate')}" min="-20" max="20" value="0" step="0.5"><span id="${pfx('sim-rate-display')}">±0%</span></div><div class="setting-row"><label>対象メーカー</label><select id="${pfx('sim-target')}"><option value="all">全体</option><option value="aron">アロン化成のみ</option><option value="pana">パナソニックのみ</option></select></div></div>
                </div>
                <div class="sim-result-grid">
                    <div class="sim-result-card"><div class="sim-label">変動前 粗利</div><div class="sim-value" id="${pfx('sim-before')}">-</div></div>
                    <div class="sim-result-card arrow">&#8594;</div>
                    <div class="sim-result-card"><div class="sim-label">変動後 粗利</div><div class="sim-value" id="${pfx('sim-after')}">-</div></div>
                    <div class="sim-result-card"><div class="sim-label">差額</div><div class="sim-value" id="${pfx('sim-diff')}">-</div></div>
                </div>
                <div class="chart-row"><div class="chart-box full"><h3>掛け率 vs 粗利 推移</h3><canvas id="${pfx('chart-sim')}"></canvas></div></div>
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
        `;
    }

    // ── Bind Events ──
    function bindEvents(container) {
        // Sub-tab nav
        container.querySelectorAll('.mod-nav-btn').forEach(btn => {
            btn.addEventListener('click', () => switchModTab(btn.dataset.mtab));
        });

        // File uploads
        document.getElementById(pfx('file-shipping')).addEventListener('change', async (e) => {
            if (!e.target.files.length) return;
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
            try {
                const list = [];
                for (const f of Array.from(e.target.files)) list.push(await parseExcel(f));
                loadSales(list); checkAllLoaded();
                e.target.value = '';
            } catch (err) { log('販売実績読込エラー: ' + err.message); alert('販売実績の読込に失敗しました'); }
        });
        document.getElementById(pfx('file-product')).addEventListener('change', async (e) => {
            if (!e.target.files.length) return;
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
            if (state.salesData.length === 0) { alert('データを先に読み込んでください。'); return; }
            analyze();
        });

        const settingsTabEl = document.getElementById(pfx('tab-settings'));
        if (settingsTabEl) {
            const onSettingsChanged = (e) => {
                const target = e.target;
                if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement || target instanceof HTMLTextAreaElement)) return;
                scheduleAutoStateSave(500);
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
        document.getElementById(pfx('store-sim-store')).addEventListener('change', renderStoreSimulationFromCurrent);
        document.getElementById(pfx('store-sim-maker')).addEventListener('change', renderStoreSimulationFromCurrent);
        document.getElementById(pfx('store-sim-rate')).addEventListener('input', renderStoreSimulationFromCurrent);
        document.getElementById(pfx('store-sim-qty')).addEventListener('input', renderStoreSimulationFromCurrent);
        document.getElementById(pfx('details-maker')).addEventListener('change', renderDetails);
        document.getElementById(pfx('details-sort')).addEventListener('change', renderDetails);
        document.getElementById(pfx('details-search')).addEventListener('input', renderDetails);
        document.getElementById(pfx('sim-rate')).addEventListener('input', renderSimulation);
        document.getElementById(pfx('sim-target')).addEventListener('change', renderSimulation);

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
            renderMonthlyRebateInputs();
            const restoredLocal = restoreAutoState();
            if (!restoredLocal) restoreCloudStateIfNeeded();
            checkAllLoaded();
        },
        onShow() {
            if (state.results && currentTab !== 'upload' && currentTab !== 'settings') {
                switchModTab(currentTab);
            }
        }
    });

})();
