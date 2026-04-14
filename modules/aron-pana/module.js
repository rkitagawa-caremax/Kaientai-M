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
        storeSortedCache: {},
        storeViewRuntime: null,
        storeCurrentPage: 1,
        storeCurrentPageTotal: 1,
        storeHeavyRenderToken: 0,
        storeHeavyRenderTimer: null,
        storeAdvancedEnabled: false,
        detailsCurrentPage: 1,
        detailsCurrentPageTotal: 1,
        storeDetailIndex: null,
        storeDetailCurrentPage: 1,
        storeDetailCurrentPageTotal: 1,
        progressItems: [],
        progressEditingId: '',
        progressCurrentPage: 1,
        progressCurrentPageTotal: 1,
        authVerified: false,
        uploadUnlocked: false,
        settingsUnlocked: false
    };

    let logLines = [];
    let currentTab = 'upload';
    const PROGRESS_STORAGE_KEY = 'kaientai-aron-pana-progress-v1';
    const AUTO_STATE_STORAGE_KEY = 'kaientai-aron-pana-autostate-v1';
    const AUTO_STATE_META_KEY = AUTO_STATE_STORAGE_KEY + '-meta';
    const AUTO_STATE_CHUNK_PREFIX = AUTO_STATE_STORAGE_KEY + '-chunk-';
    const AUTO_STATE_CHUNK_SIZE = 120000;
    let authHydrated = false;
    let autoPersistTimer = null;
    let cloudPersistTimer = null;
    let cloudPersistInFlight = false;
    let cloudPersistPending = false;
    let cloudPersistRetryMs = 1000;
    let cloudRestoreRetryTimer = null;
    let cloudRestoreRetryCount = 0;
    const CLOUD_RESTORE_MAX_RETRY = 20;
    const CLOUD_RESTORE_RETRY_MS = 1500;

    function pfx(id) { return MODULE_ID + '-' + id; }
    const COL_AB = 27; // AB列（都道府県）
    const COL_SUPPLIER_CODE = COL.C; // C列（得意先コード）
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

    function normalizeJanCode(v) {
        const s = toStr(v)
            .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
            .replace(/[^0-9]/g, '');
        return s;
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
        const isOkinawaOrder = isOkinawaPrefecture(prefecture);
        const isSmallSize = shipping.sizeBand > 0 && shipping.sizeBand <= 100;
        let areaKey = prefectureToAreaKey(prefecture);
        if (isOkinawaOrder && !isSmallSize) areaKey = 'kansai';
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

        // 沖縄県は特別条件:
        // サイズ帯100以下は500円固定、それ以外は関西エリア条件を適用
        if (isOkinawaOrder && isSmallSize) {
            shippingCost = 500;
            areaKey = 'okinawa';
            fallback = false;
        } else if (isSmallSize && settings.defaultShippingSmall > 0) {
            // 沖縄以外のサイズ帯100以下は設定値を優先
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

    // ── Settings ──
    function getSalesMonthList() {
        const months = [...new Set(state.salesData.map(s => s.month).filter(m => m && m !== 'unknown'))].sort();
        if (months.length > 0) return months;
        const d = new Date();
        return [d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')];
    }

    function isValidMonthToken(month) {
        return /^\d{4}-\d{2}$/.test(toStr(month));
    }

    function computeDataPeriodFromRecords(records) {
        const months = [...new Set((records || []).map(r => toStr(r.month)).filter(isValidMonthToken))].sort();
        if (months.length === 0) return null;
        return {
            from: months[0],
            to: months[months.length - 1],
            monthCount: months.length
        };
    }

    function readMultiSelectValues(selectEl) {
        if (!selectEl) return [];
        return Array.from(selectEl.options).filter(opt => opt.selected).map(opt => opt.value).filter(Boolean);
    }

    function makeSelectionKey(values) {
        const list = [...new Set((values || []).map(v => toStr(v)).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ja'));
        return list.length > 0 ? list.join('|') : 'all';
    }

    function setMultiSelectOptions(selectEl, values, prevSelected) {
        if (!selectEl) return [];
        const selectedSet = new Set((prevSelected || []).map(v => toStr(v)));
        selectEl.innerHTML = (values || []).map(v => `<option value="${escHtml(v)}">${escHtml(v)}</option>`).join('');
        const applied = [];
        Array.from(selectEl.options).forEach(opt => {
            if (selectedSet.has(opt.value)) {
                opt.selected = true;
                applied.push(opt.value);
            }
        });
        return applied;
    }

    function enableSimpleMultiSelect(selectEl) {
        if (!selectEl || selectEl.dataset.simpleMultiBound === '1') return;
        selectEl.dataset.simpleMultiBound = '1';
        selectEl.addEventListener('mousedown', (event) => {
            if (!(event.target instanceof HTMLOptionElement)) return;
            event.preventDefault();
            event.target.selected = !event.target.selected;
            selectEl.dispatchEvent(new Event('change', { bubbles: true }));
        });
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
                return `<label class="monthly-rebate-field"><span>${escHtml(k.label)}</span><input type="number" class="monthly-rebate-input" data-month="${escHtml(month)}" data-maker="${escHtml(k.maker)}" data-type="${escHtml(k.type)}" value="${Number(val) || 0}" step="1000"></label>`;
            }).join('');
            const row = document.createElement('div');
            row.className = 'monthly-rebate-row';
            row.innerHTML = `<div class="monthly-rebate-month">${escHtml(month)}</div><div class="monthly-rebate-grid">${fields}</div>`;
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

        const toMonthToken = (d) => {
            const y = d.getFullYear();
            const m = d.getMonth() + 1;
            return (y >= 2000 && y <= 2099) ? (y + '-' + String(m).padStart(2, '0')) : null;
        };
        const tryExcelSerial = (num) => {
            if (!(num > 30000 && num < 100000)) return null;
            const epoch = new Date(1899, 11, 30); // Excel serial epoch
            return toMonthToken(new Date(epoch.getTime() + num * 86400000));
        };

        // 1) Dateオブジェクト
        if (val instanceof Date && !isNaN(val.getTime())) {
            return toMonthToken(val);
        }

        // 2) 数値 / 数値文字列のExcelシリアル
        if (typeof val === 'number') {
            const month = tryExcelSerial(val);
            if (month) return month;
        }

        const s = String(val).trim();
        if (!s) return null;
        const normalized = s
            .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 65248))
            .replace(/[．。]/g, '.')
            .replace(/[／]/g, '/')
            .replace(/[－―ー]/g, '-');

        const numeric = parseFloat(normalized.replace(/,/g, ''));
        if (!Number.isNaN(numeric) && /^\d+(\.\d+)?$/.test(normalized.replace(/,/g, ''))) {
            const month = tryExcelSerial(numeric);
            if (month) return month;
        }

        // 3) yyyy/mm/dd, yyyy-mm-dd, yyyy.mm.dd, yyyy/mm
        let match = normalized.match(/(\d{4})[\/\-.](\d{1,2})(?:[\/\-.]\d{1,2})?/);
        if (match) {
            const y = toNum(match[1]);
            const m = toNum(match[2]);
            if (y >= 2000 && y <= 2099 && m >= 1 && m <= 12) return `${match[1]}-${String(m).padStart(2, '0')}`;
        }

        // 4) yyyy年mm月
        match = normalized.match(/(\d{4})\s*年\s*(\d{1,2})\s*月/);
        if (match) {
            const m = toNum(match[2]);
            if (m >= 1 && m <= 12) return `${match[1]}-${String(m).padStart(2, '0')}`;
        }

        // 5) mm/dd/yyyy, mm-dd-yyyy, mm.dd.yyyy
        match = normalized.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
        if (match) {
            const y = toNum(match[3]);
            const m = toNum(match[1]);
            if (y >= 2000 && y <= 2099 && m >= 1 && m <= 12) return `${match[3]}-${String(m).padStart(2, '0')}`;
        }

        // 6) yyyymmdd / yyyymm
        match = normalized.match(/^(\d{4})(\d{2})(\d{2})?$/);
        if (match) {
            const y = toNum(match[1]);
            const m = toNum(match[2]);
            if (y >= 2000 && y <= 2099 && m >= 1 && m <= 12) return `${match[1]}-${String(m).padStart(2, '0')}`;
        }

        // 7) Date解析の最終手段
        const d = new Date(normalized);
        if (!isNaN(d.getTime())) return toMonthToken(d);

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

    function findProductHeaderRow(rows) {
        const limit = Math.min(rows.length, 20);
        for (let i = 0; i < limit; i++) {
            const rowText = (rows[i] || []).map(c => toStr(c).toLowerCase()).join(' ');
            const hasJan = rowText.includes('jan') || rowText.includes('商品コード') || rowText.includes('商品ｺｰﾄﾞ') || rowText.includes('品番') || rowText.includes('code');
            const hasName = rowText.includes('商品名') || rowText.includes('品名') || rowText.includes('item');
            const hasCost = rowText.includes('原価') || rowText.includes('cost') || rowText.includes('仕切');
            const hasPrice = rowText.includes('定価') || rowText.includes('上代') || rowText.includes('price') || rowText.includes('売価');
            if (hasJan && hasName && (hasCost || hasPrice)) return i;
        }
        return -1;
    }

    function findBestProductSheet(parsed) {
        let bestSheet = '';
        let bestHeaderRow = -1;
        let bestScore = -1;
        for (const name of parsed.sheetNames) {
            const rows = parsed.sheets[name];
            if (!rows || rows.length < 3) continue;
            const headerRow = findProductHeaderRow(rows);
            if (headerRow < 0) continue;

            const normalizedName = normalizeToken(name);
            let score = 0;
            if (normalizedName.includes('商品') || normalizedName.includes('マスタ')) score += 300;
            if (normalizedName.includes('品番')) score += 120;
            score += Math.min(rows.length, 1500);

            let janHits = 0;
            for (let i = headerRow + 1; i < Math.min(rows.length, headerRow + 120); i++) {
                const row = rows[i] || [];
                const jan = normalizeJanCode(row[COL.A]);
                if (/^\d{8,13}$/.test(jan)) janHits++;
            }
            score += janHits * 20;

            if (score > bestScore) {
                bestScore = score;
                bestSheet = name;
                bestHeaderRow = headerRow;
            }
        }
        if (bestSheet) log(`  -> product sheet selected: "${bestSheet}" (score=${bestScore}, header=${bestHeaderRow})`);
        return { sheetName: bestSheet, headerRow: bestHeaderRow };
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
            const jan = normalizeJanCode(row[COL.A]);
            if (!/^\d{8,13}$/.test(jan)) { skipped++; continue; }
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
        saveAutoStateNow();
    }

    function loadSales(parsedList) {
        const normalizeOrderNo = (orderNo) => toStr(orderNo).trim();
        const buildSalesRowSignature = (row) => ([
            normalizeOrderNo(row.orderNo),
            toStr(row.month),
            toStr(row.store),
            toStr(row.supplierCode),
            toStr(row.salesRep),
            toStr(row.jan),
            toStr(row.name),
            String(toNum(row.qty)),
            String(toNum(row.unitPrice)),
            String(toNum(row.totalPrice)),
            toStr(row.prefecture),
            toStr(row.makerRaw)
        ]).join('|');
        const makerValues = new Set(state.salesData.map(s => s.makerRaw).filter(Boolean));
        const monthValues = new Set(state.salesData.map(s => s.month).filter(Boolean));
        const seenRowSignatures = new Set(state.salesData.map(buildSalesRowSignature));
        let duplicateRowCount = 0;
        let addedCount = 0;

        for (const parsed of parsedList) {
            log(`--- 販売実績読込: ${parsed.fileName} ---`);
            const fileMonth = extractMonthFromFileName(parsed.fileName);
            const fileMonthStats = {};
            for (const sheetName of parsed.sheetNames) {
                const rows = parsed.sheets[sheetName];
                log(`  シート[${sheetName}]: ${rows.length}行`);

                for (let i = 0; i < Math.min(rows.length, 3); i++) {
                    const r = rows[i] || [];
                    log(`    行${i}: A=[${toStr(r[COL.A])}] B=[${toStr(r[COL.B])}] C=[${toStr(r[COL_SUPPLIER_CODE])}] D=[${toStr(r[COL.D])}] H=[${toStr(r[COL.H])}] I=[${toStr(r[COL.I])}] K=[${toStr(r[COL.K])}] L=[${toStr(r[COL.L])}] S=[${toStr(r[COL.S])}] Z=[${toStr(r[COL_SALES_REP])}] AB=[${toStr(r[COL_AB])}]`);
                }

                const headerRow = findHeaderRow(rows, ['jan', '商品', 'コード', '数量', '受注']);
                log(`  ヘッダー行: ${headerRow}`);

                let count = 0, dateOk = 0, dateFail = 0;
                let inferredUnitPriceCount = 0, inferredTotalPriceCount = 0, inferredQtyCount = 0;
                for (let i = headerRow + 1; i < rows.length; i++) {
                    const row = rows[i] || [];
                    const jan = normalizeJanCode(row[COL.H]);
                    if (!/^\d{8,13}$/.test(jan)) continue;
                    const orderNo = normalizeOrderNo(row[COL.A]);

                    const rawDate = row[COL.B];
                    let month = extractMonthFromDate(rawDate);
                    if (month) {
                        dateOk++;
                    } else {
                        dateFail++;
                        month = fileMonth || 'unknown';
                    }
                    monthValues.add(month);
                    fileMonthStats[month] = (fileMonthStats[month] || 0) + 1;

                    const makerRaw = toStr(row[COL.S]);
                    if (makerRaw) makerValues.add(makerRaw);
                    const maker = makerRaw ? detectMaker(makerRaw) : 'other';

                    const name = toStr(row[COL.I]);
                    const store = toStr(row[COL.D]);
                    let qty = toNum(row[COL.K]);
                    let unitPrice = toNum(row[COL.L]);
                    let totalPrice = toNum(row[COL.M]);

                    const qtyAbs = Math.abs(qty);
                    const unitAbs = Math.abs(unitPrice);
                    const totalAbs = Math.abs(totalPrice);
                    if (unitAbs <= 0 && qtyAbs > 0 && totalAbs > 0) {
                        unitPrice = totalPrice / qty;
                        inferredUnitPriceCount++;
                    }
                    if (totalAbs <= 0 && qtyAbs > 0 && Math.abs(unitPrice) > 0) {
                        totalPrice = unitPrice * qty;
                        inferredTotalPriceCount++;
                    }
                    if (qtyAbs <= 0 && Math.abs(unitPrice) > 0 && Math.abs(totalPrice) > 0) {
                        qty = totalPrice / unitPrice;
                        inferredQtyCount++;
                    }
                    const salesRep = toStr(row[COL_SALES_REP]);
                    const supplierCode = toStr(row[COL_SUPPLIER_CODE]);
                    const saleRow = {
                        orderNo,
                        month,
                        maker,
                        makerRaw,
                        salesRep,
                        supplierCode,
                        store,
                        prefecture: toStr(row[COL_AB]),
                        jan,
                        name,
                        qty,
                        unitPrice,
                        totalPrice
                    };
                    const rowSignature = buildSalesRowSignature(saleRow);
                    if (seenRowSignatures.has(rowSignature)) {
                        duplicateRowCount++;
                        continue;
                    }
                    seenRowSignatures.add(rowSignature);
                    state.salesData.push(saleRow);
                    addedCount++;
                    count++;
                }
                log(`  ${count}件読込 (B列日付OK=${dateOk}, B列日付NG=${dateFail})`);
            }
            const monthBreakdown = Object.keys(fileMonthStats).sort().map(m => `${m}:${fileMonthStats[m]}`).join(', ');
            if (monthBreakdown) log(`  月判定内訳: ${monthBreakdown}`);
        }

        log(`月一覧: [${[...monthValues].sort().join(', ')}]`);
        log(`S列メーカー生値: [${[...makerValues].join(' / ')}]`);
        const aronCount = state.salesData.filter(s => s.maker === 'aron').length;
        const panaCount = state.salesData.filter(s => s.maker === 'pana').length;
        const otherCount = state.salesData.filter(s => s.maker === 'other').length;
        log(`メーカー判定件数: アロン=${aronCount}件 / パナ=${panaCount}件 / その他=${otherCount}件`);
        log(`販売実績追加: ${addedCount}件 / 行重複スキップ: ${duplicateRowCount}件 / 累計: ${state.salesData.length}件`);

        document.getElementById(pfx('status-sales')).textContent = `計: ${state.salesData.length}件 (+${addedCount}件 / ${parsedList.length}ファイル)`;
        document.getElementById(pfx('card-sales')).classList.add('loaded');
        renderMonthlyRebateInputs();
        saveAutoStateNow();
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

            const picked = findBestProductSheet(parsed);
            const sheetName = picked.sheetName;
            const headerRow = picked.headerRow;
            if (!sheetName || headerRow < 0) {
                log(`  ! ${parsed.fileName}: 商品マスタ形式のヘッダーを検出できないためスキップ`);
                continue;
            }
            const rows = parsed.sheets[sheetName];
            log(`Using sheet: ${sheetName} / rows: ${rows.length}`);

            for (let i = 0; i < Math.min(rows.length, 5); i++) {
                const r = rows[i] || [];
                log(`  Row ${i}: A=[${toStr(r[COL.A])}] D=[${toStr(r[COL.D])}] H=[${toStr(r[COL.H])}] M=[${toStr(r[COL.M])}] O=[${toStr(r[COL.O])}]`);
            }
            log(`Header row: ${headerRow}`);

            let fileCount = 0;
            for (let i = headerRow + 1; i < rows.length; i++) {
                const row = rows[i] || [];
                const jan = normalizeJanCode(row[COL.A]);
                if (!/^\d{8,13}$/.test(jan)) { skipped++; continue; }
                const name = toStr(row[COL.D]);
                const listPrice = toNum(row[COL.H]);
                const wc = toNum(row[COL.O]);
                const cost = toNum(row[COL.M]);
                if (!name && listPrice <= 0 && wc <= 0 && cost <= 0) { skipped++; continue; }
                if (productMap.has(jan)) overwriteCount++;
                productMap.set(jan, {
                    jan, name, listPrice,
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
        saveAutoStateNow();
    }

    // ── Analysis ──
    function runAnalysis() {
        const settings = getSettings();
        const shippingMap = {};
        for (const s of state.shippingData) shippingMap[s.jan] = s;
        const productMap = {};
        for (const p of state.productData) productMap[p.jan] = p;

        // Build price references to recover rows where both unit/total are zero.
        // Priority: same JAN average unit price, then maker-level average list-rate.
        const janPriceStats = {};
        const makerRateStats = {};
        for (const sale of state.salesData) {
            const qtyAbs = Math.abs(toNum(sale.qty));
            const unitAbs = Math.abs(toNum(sale.unitPrice));
            if (qtyAbs <= 0 || unitAbs <= 0) continue;

            if (!janPriceStats[sale.jan]) janPriceStats[sale.jan] = { amount: 0, qty: 0 };
            janPriceStats[sale.jan].amount += unitAbs * qtyAbs;
            janPriceStats[sale.jan].qty += qtyAbs;

            const product = productMap[sale.jan];
            if (!product || product.listPrice <= 0) continue;
            const rate = unitAbs / product.listPrice;
            if (!Number.isFinite(rate) || rate <= 0) continue;
            if (!makerRateStats[sale.maker]) makerRateStats[sale.maker] = { weightedRate: 0, qty: 0 };
            makerRateStats[sale.maker].weightedRate += rate * qtyAbs;
            makerRateStats[sale.maker].qty += qtyAbs;
        }
        const janPriceRef = {};
        Object.keys(janPriceStats).forEach(jan => {
            const s = janPriceStats[jan];
            janPriceRef[jan] = s.qty > 0 ? s.amount / s.qty : 0;
        });
        const makerRateRef = {};
        Object.keys(makerRateStats).forEach(maker => {
            const s = makerRateStats[maker];
            makerRateRef[maker] = s.qty > 0 ? s.weightedRate / s.qty : 0;
        });

        const records = [];
        let matchCount = 0, noShipping = 0, noProduct = 0, excludedCount = 0;
        let noPrefArea = 0, areaFallback = 0, zeroAreaCost = 0;
        let inferredByJan = 0, inferredByMakerRate = 0, missingPriceExcluded = 0;
        let excludedByManualRule = 0;
        const noProductJanSamples = new Set();
        const excludedStoreCode = '405';
        const excludedStoreNameToken = normalizeToken('(株)ケアマックスコーポレーション　大阪第2倉庫');

        for (const sale of state.salesData) {
            const supplierCode = toStr(sale.supplierCode).replace(/^0+/, '');
            const storeToken = normalizeToken(sale.store);
            if (supplierCode === excludedStoreCode && (storeToken === excludedStoreNameToken || storeToken.includes(excludedStoreNameToken))) {
                excludedByManualRule++;
                excludedCount++;
                continue;
            }
            const shipping = shippingMap[sale.jan];
            const product = productMap[sale.jan];
            const isProductValid = !!(product && (toStr(product.name) || toNum(product.listPrice) > 0 || toNum(product.effectiveCost) > 0));
            if (!shipping) noShipping++;
            if (!isProductValid) {
                noProduct++;
                if (noProductJanSamples.size < 20) noProductJanSamples.add(sale.jan);
            }
            if (!shipping || !isProductValid) { excludedCount++; continue; }
            matchCount++;

            const shipCalc = resolveShippingCost(shipping, sale.prefecture, settings);
            const shippingCost = shipCalc.shippingCost;
            if (!shipCalc.areaKey) noPrefArea++;
            if (shipCalc.fallback) areaFallback++;
            if (shippingCost <= 0) zeroAreaCost++;
            const effectiveCost = product.effectiveCost;
            const listPrice = product.listPrice;
            let qty = toNum(sale.qty);
            let unitPrice = toNum(sale.unitPrice);
            let totalPrice = toNum(sale.totalPrice);

            if (Math.abs(unitPrice) <= 0 && Math.abs(qty) > 0 && Math.abs(totalPrice) > 0) {
                unitPrice = totalPrice / qty;
            }
            if (Math.abs(totalPrice) <= 0 && Math.abs(qty) > 0 && Math.abs(unitPrice) > 0) {
                totalPrice = unitPrice * qty;
            }
            if (Math.abs(unitPrice) <= 0 && Math.abs(totalPrice) <= 0 && Math.abs(qty) > 0) {
                const janRef = janPriceRef[sale.jan] || 0;
                if (janRef > 0) {
                    unitPrice = janRef;
                    totalPrice = unitPrice * qty;
                    inferredByJan++;
                } else {
                    const makerRef = makerRateRef[sale.maker] || 0;
                    if (makerRef > 0 && listPrice > 0) {
                        unitPrice = listPrice * makerRef;
                        totalPrice = unitPrice * qty;
                        inferredByMakerRate++;
                    } else {
                        missingPriceExcluded++;
                        excludedCount++;
                        continue;
                    }
                }
            }

            const salesAmount = Math.abs(totalPrice) > 0 ? totalPrice : (unitPrice * qty);
            const totalShipping = qty * shippingCost;
            const totalCost = effectiveCost * qty;
            const grossProfit = salesAmount - totalCost - totalShipping;

            records.push({
                ...sale, qty, unitPrice, totalPrice, shippingCost, shippingArea: shipCalc.areaKey, effectiveCost, listPrice,
                salesAmount, totalShipping, totalCost, grossProfit,
                rateVsList: listPrice > 0 ? unitPrice / listPrice : 0
            });
        }

        log(`マッチング: ${matchCount}件一致(3データ一致) / 除外: ${excludedCount} / 手動除外(コード405・ケアマックス大阪第2倉庫): ${excludedByManualRule} / 送料未一致: ${noShipping} / 商品マスタ未一致: ${noProduct} / 単価不明除外: ${missingPriceExcluded} / 単価補完(JAN平均): ${inferredByJan} / 単価補完(メーカー平均掛率): ${inferredByMakerRate} / 地域判定不可: ${noPrefArea} / エリア補完: ${areaFallback} / 送料0円: ${zeroAreaCost}`);
        if (noProductJanSamples.size > 0) {
            log(`商品マスタ未一致JAN例: ${[...noProductJanSamples].join(', ')}`);
        }

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

        const dataPeriod = computeDataPeriodFromRecords(state.salesData);
        state.results = {
            records, monthlyAgg, storeAgg, productAgg, months,
            totalSales, totalCost, totalShipping, totalGross, totalQty,
            totalRebate, totalWarehouse, totalWarehouseOut, totalMinus, realProfit,
            aronSales, panaSales, settings, monthSalesTotals, rebateByMaker, minusByMaker,
            dataPeriod
        };
        state.storeBaseCache = { 'all|all': buildStoreBase(records) };
        state.storeSortedCache = {};
        state.storeViewRuntime = null;
        state.storeAdvancedEnabled = false;
        state.storeCurrentPage = 1;
        state.storeCurrentPageTotal = 1;
        state.detailsCurrentPage = 1;
        state.detailsCurrentPageTotal = 1;
        state.storeDetailIndex = buildStoreDetailIndex(records);
        state.storeDetailCurrentPage = 1;
        state.storeDetailCurrentPageTotal = 1;
        state.progressCurrentPage = 1;
        state.progressCurrentPageTotal = 1;
        state.progressEditingId = '';
        state.storeHeavyRenderToken = 0;
        if (state.storeHeavyRenderTimer) {
            clearTimeout(state.storeHeavyRenderTimer);
            state.storeHeavyRenderTimer = null;
        }
        setStoreAdvancedMode(false);

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
        const periodEl = document.getElementById(pfx('overview-period'));
        if (periodEl) {
            if (r.dataPeriod) {
                periodEl.textContent = `データ対象期間: ${r.dataPeriod.from} ～ ${r.dataPeriod.to} (${fmt(r.dataPeriod.monthCount)}か月)`;
            } else {
                periodEl.textContent = 'データ対象期間: 年月情報なし';
            }
        }

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
                tr.innerHTML = `<td>${escHtml(e.month)}</td><td>${escHtml(ml[e.maker] || e.maker)}</td><td>${fmtYen(e.sales)}</td><td>${fmtYen(e.cost)}</td><td>${fmtYen(e.shipping)}</td><td class="${e.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(e.gross)}</td><td>${fmtYen(rebate)}</td><td>${fmtYen(whFee)}</td><td class="${realProfit >= 0 ? 'positive' : 'negative'}">${fmtYen(realProfit)}</td><td>${fmtPct(pr)}</td>`;
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
                tr.innerHTML = `<td>${escHtml(month)}</td><td>【合計】</td><td>${fmtYen(monthTotal.sales)}</td><td>${fmtYen(monthTotal.cost)}</td><td>${fmtYen(monthTotal.shipping)}</td><td class="${monthTotal.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(monthTotal.gross)}</td><td>${fmtYen(monthTotal.rebate)}</td><td>${fmtYen(monthTotal.whFee)}</td><td class="${monthTotal.real >= 0 ? 'positive' : 'negative'}">${fmtYen(monthTotal.real)}</td><td>${fmtPct(pr)}</td>`;
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

    function buildStoreDetailIndex(records) {
        const map = {};
        for (const rec of records) {
            const store = rec.store || '(未設定)';
            if (!map[store]) {
                map[store] = {
                    records: [],
                    months: new Set(),
                    reps: new Set(),
                    supplierCodes: new Set()
                };
            }
            map[store].records.push(rec);
            if (rec.month) map[store].months.add(rec.month);
            if (rec.salesRep) map[store].reps.add(rec.salesRep);
            if (rec.supplierCode) map[store].supplierCodes.add(rec.supplierCode);
        }
        return {
            stores: Object.keys(map).sort((a, b) => a.localeCompare(b, 'ja')),
            map
        };
    }

    function getStoreBase(records, makerF, monthValues) {
        const monthSet = new Set((monthValues || []).filter(Boolean));
        const monthKey = makeSelectionKey([...monthSet]);
        const key = makerF + '|' + monthKey;
        if (state.storeBaseCache[key]) return state.storeBaseCache[key];

        const filtered = [];
        for (const rec of records) {
            if (makerF !== 'all' && rec.maker !== makerF) continue;
            if (monthSet.size > 0 && !monthSet.has(rec.month)) continue;
            filtered.push(rec);
        }
        const base = buildStoreBase(filtered);
        state.storeBaseCache[key] = base;
        return base;
    }

    function selectStoreRecordsByReps(base, repValues) {
        if (!base) return [];
        const reps = (repValues || []).filter(Boolean);
        if (reps.length === 0) return base.recordsByRep.all || [];
        const out = [];
        for (const rep of reps) {
            const rows = base.recordsByRep[rep];
            if (Array.isArray(rows)) out.push(...rows);
        }
        return out;
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
                // 掛率は卸売価格÷定価のみ（送料・リベートは含めない）
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

    function renderStoreChart(entries) {
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

    function renderStoreSimulation(runtime, rebuildStoreList) {
        const simStoreSel = document.getElementById(pfx('store-sim-store'));
        const simMaker = document.getElementById(pfx('store-sim-maker')).value;
        const simRateChange = toNum(document.getElementById(pfx('store-sim-rate')).value) / 100;
        const simIncreaseQty = Math.max(0, toNum(document.getElementById(pfx('store-sim-qty')).value));
        const simTbody = document.getElementById(pfx('store-sim-tbody'));

        const fmtQty = (n) => (n == null || isNaN(n)) ? '-' : (Math.round(n * 10) / 10).toLocaleString('ja-JP');
        const signed = (n, formatter) => (n >= 0 ? '+' : '') + formatter(n);

        if (!runtime) {
            simTbody.innerHTML = '<tr><td colspan="7">データがありません</td></tr>';
            return;
        }

        const storeRecordsMap = runtime.storeRecordsMap || {};
        const simStoreList = runtime.simStoreList || [];
        if (rebuildStoreList) {
            const prevSimStore = simStoreSel.value;
            simStoreSel.innerHTML = '';
            if (simStoreList.length === 0) {
                simStoreSel.innerHTML = '<option value="">販売店なし</option>';
            } else {
                for (const sName of simStoreList) simStoreSel.innerHTML += `<option value="${escHtml(sName)}">${escHtml(sName)}</option>`;
            }
            simStoreSel.value = simStoreList.includes(prevSimStore) ? prevSimStore : (simStoreList[0] || '');
        } else if (!simStoreSel.value && simStoreList.length > 0) {
            simStoreSel.value = simStoreList[0];
        }
        const simStore = simStoreSel.value;

        simTbody.innerHTML = '';
        const storeRecords = storeRecordsMap[simStore] || [];
        if (!simStore || storeRecords.length === 0) {
            simTbody.innerHTML = '<tr><td colspan="7">対象データがありません</td></tr>';
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
                // 掛率は卸売価格÷定価のみ（送料・リベートは含めない）
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
                // 掛率は卸売価格÷定価のみ（送料・リベートは含めない）
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
            `<tr><td>変更後</td><td>${fmtYen(after.sales)}</td><td class="${after.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(after.gross)}</td><td>${fmtPct(afterProfitRate)}</td><td>${fmtQty(after.qty)}</td><td>${afterAronRate > 0 ? fmtPct(afterAronRate) : '-'}</td><td>${afterPanaRate > 0 ? fmtPct(afterPanaRate) : '-'}</td></tr>`,
            `<tr style="font-weight:700;background:#fff3e0;"><td>差分</td><td class="${diff.sales >= 0 ? 'positive' : 'negative'}">${signed(diff.sales, fmtYen)}</td><td class="${diff.gross >= 0 ? 'positive' : 'negative'}">${signed(diff.gross, fmtYen)}</td><td class="${diff.profitRate >= 0 ? 'positive' : 'negative'}">${signed(diff.profitRate, fmtPct)}</td><td class="${diff.qty >= 0 ? 'positive' : 'negative'}">${diff.qty >= 0 ? '+' : ''}${fmtQty(diff.qty)}</td><td class="${diff.aronRate >= 0 ? 'positive' : 'negative'}">${signed(diff.aronRate, fmtPct)}</td><td class="${diff.panaRate >= 0 ? 'positive' : 'negative'}">${signed(diff.panaRate, fmtPct)}</td></tr>`
        ].join('');
    }

    function renderStoreSimulationFromCurrent() {
        if (!state.storeAdvancedEnabled || !state.storeViewRuntime) return;
        renderStoreSimulation(state.storeViewRuntime, false);
    }

    function cancelStoreHeavyRender() {
        state.storeHeavyRenderToken += 1;
        if (state.storeHeavyRenderTimer) {
            clearTimeout(state.storeHeavyRenderTimer);
            state.storeHeavyRenderTimer = null;
        }
    }

    function setStoreAdvancedMode(enabled, message) {
        state.storeAdvancedEnabled = !!enabled;

        const panel = document.getElementById(pfx('store-advanced-panel'));
        const statusEl = document.getElementById(pfx('store-advanced-status'));
        const enableBtn = document.getElementById(pfx('btn-store-advanced-enable'));
        const disableBtn = document.getElementById(pfx('btn-store-advanced-disable'));
        if (panel) panel.style.display = state.storeAdvancedEnabled ? '' : 'none';
        if (enableBtn) enableBtn.disabled = state.storeAdvancedEnabled;
        if (disableBtn) disableBtn.disabled = !state.storeAdvancedEnabled;

        if (!state.storeAdvancedEnabled) {
            cancelStoreHeavyRender();
            state.storeViewRuntime = null;
            destroyChart(state.charts, 'store');
            const simTbody = document.getElementById(pfx('store-sim-tbody'));
            if (simTbody) simTbody.innerHTML = '<tr><td colspan="7">軽量モード中です。詳細分析を読み込むと表示します。</td></tr>';
        }

        if (statusEl) {
            statusEl.textContent = message || (
                state.storeAdvancedEnabled
                    ? '詳細分析を有効化しました。フィルタ条件ごとに再計算します。'
                    : '軽量モード中です。詳細分析（シミュレーション/ランキング）は停止しています。'
            );
        }
    }

    // ── Render: Store ──
    function renderStore(keepFilterOptions) {
        const r = state.results; if (!r) return;
        document.getElementById(pfx('store-empty')).style.display = 'none';
        document.getElementById(pfx('store-content')).style.display = 'block';

        const mSel = document.getElementById(pfx('store-month'));
        const repSel = document.getElementById(pfx('store-rep'));
        enableSimpleMultiSelect(mSel);
        enableSimpleMultiSelect(repSel);

        const prevMonths = readMultiSelectValues(mSel);
        if (!keepFilterOptions) {
            setMultiSelectOptions(mSel, r.months, prevMonths);
        }

        const makerF = document.getElementById(pfx('store-maker')).value;
        const monthValues = readMultiSelectValues(mSel);
        const monthKey = makeSelectionKey(monthValues);
        const sortKey = document.getElementById(pfx('store-sort')).value;
        const prevReps = readMultiSelectValues(repSel);
        const limitRaw = document.getElementById(pfx('store-limit')).value;

        const base = getStoreBase(r.records, makerF, monthValues);
        if (!keepFilterOptions) {
            setMultiSelectOptions(repSel, base.reps, prevReps);
        }
        const repValues = readMultiSelectValues(repSel);
        const repKey = makeSelectionKey(repValues);
        const dataKey = makerF + '|' + monthKey + '|' + repKey;

        if (!base.entriesByRep[repKey]) {
            const scopedRecordsByRep = selectStoreRecordsByReps(base, repValues);
            base.entriesByRep[repKey] = buildStoreEntries(scopedRecordsByRep);
        }

        const sortedKey = dataKey + '|' + sortKey;
        let entries = state.storeSortedCache[sortedKey];
        if (!entries) {
            entries = [...base.entriesByRep[repKey]];
            sortStoreEntries(entries, sortKey);
            state.storeSortedCache[sortedKey] = entries;
        }

        const isAll = limitRaw === 'all';
        const pageSize = isAll ? Math.max(1, entries.length || 1) : Math.max(1, toNum(limitRaw) || 300);
        const totalPages = isAll ? 1 : Math.max(1, Math.ceil(entries.length / pageSize));
        state.storeCurrentPageTotal = totalPages;
        state.storeCurrentPage = isAll ? 1 : Math.min(Math.max(1, state.storeCurrentPage || 1), totalPages);
        const startIndex = isAll ? 0 : (state.storeCurrentPage - 1) * pageSize;
        const displayed = entries.slice(startIndex, startIndex + pageSize);
        const tbody = document.getElementById(pfx('store-tbody'));
        tbody.innerHTML = displayed.map(e =>
            `<tr><td>${escHtml(e.store)}</td><td>${escHtml(e.salesRep)}</td><td>${e.aronRate > 0 ? fmtPct(e.aronRate) : '-'}</td><td>${e.panaRate > 0 ? fmtPct(e.panaRate) : '-'}</td><td>${fmtYen(e.sales)}</td><td>${fmtYen(e.cost)}</td><td>${fmtYen(e.shipping)}</td><td class="${e.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(e.gross)}</td><td>${fmtPct(e.rate)}</td><td>${fmt(e.qty)}</td></tr>`
        ).join('');

        const summaryEl = document.getElementById(pfx('store-summary'));
        if (summaryEl) {
            const monthLabel = monthValues.length > 0 ? monthValues.join(', ') : '全期間';
            const repLabel = repValues.length > 0 ? repValues.join(', ') : '全担当';
            if (entries.length === 0) {
                summaryEl.textContent = `表示: 0件 / 全0件 / 年月: ${monthLabel} / 営業担当: ${repLabel}`;
            } else {
                const from = startIndex + 1;
                const to = startIndex + displayed.length;
                summaryEl.textContent = `表示: ${fmt(from)}-${fmt(to)}件 / 全${fmt(entries.length)}件 / 年月: ${monthLabel} / 営業担当: ${repLabel}`;
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

        if (!state.storeAdvancedEnabled) return;

        const scopedRecords = selectStoreRecordsByReps(base, repValues);
        const statusEl = document.getElementById(pfx('store-advanced-status'));
        const heavyChanged = !state.storeViewRuntime || state.storeViewRuntime.key !== dataKey;
        if (!heavyChanged) {
            if (statusEl) statusEl.textContent = '詳細分析を表示中です。';
            return;
        }

        const storeRecordsMap = {};
        for (const rec of scopedRecords) {
            const key = rec.store || '(未設定)';
            if (!storeRecordsMap[key]) storeRecordsMap[key] = [];
            storeRecordsMap[key].push(rec);
        }

        state.storeViewRuntime = {
            key: dataKey,
            scopedRecords,
            storeRecordsMap,
            simStoreList: Object.keys(storeRecordsMap).sort((a, b) => a.localeCompare(b, 'ja'))
        };

        const simTbody = document.getElementById(pfx('store-sim-tbody'));
        if (simTbody) simTbody.innerHTML = '<tr><td colspan="7">詳細分析を読み込み中...</td></tr>';
        if (statusEl) statusEl.textContent = '詳細分析を読み込み中です...';

        cancelStoreHeavyRender();
        const token = state.storeHeavyRenderToken;
        state.storeHeavyRenderTimer = setTimeout(() => {
            if (currentTab !== 'store' || token !== state.storeHeavyRenderToken || !state.storeAdvancedEnabled) {
                state.storeHeavyRenderTimer = null;
                return;
            }
            renderStoreSimulation(state.storeViewRuntime, true);
            renderStoreChart(entries);
            state.storeHeavyRenderTimer = null;
            const currentStatusEl = document.getElementById(pfx('store-advanced-status'));
            if (currentStatusEl) currentStatusEl.textContent = '詳細分析を表示中です。';
        }, 0);
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

        const settings = r.settings || getSettings();
        const months = Array.isArray(r.months) && r.months.length > 0
            ? r.months
            : [...new Set(r.records.map(x => x.month).filter(Boolean))].sort();
        const simRebateTbody = document.getElementById(pfx('sim-rebate-tbody'));
        const simRebateSummary = document.getElementById(pfx('sim-rebate-summary'));

        const evalScenario = (pct) => {
            const monthlyAgg = {};
            for (const rec of r.records) {
                const applyChange = target === 'all' || rec.maker === target;
                const newUnitPrice = applyChange ? rec.unitPrice * (1 + pct) : rec.unitPrice;
                const sales = newUnitPrice * rec.qty;
                const cost = rec.totalCost;
                const shipping = rec.totalShipping;
                const gross = sales - cost - shipping;
                const key = rec.month + '|' + rec.maker;
                if (!monthlyAgg[key]) {
                    monthlyAgg[key] = {
                        month: rec.month,
                        maker: rec.maker,
                        sales: 0,
                        cost: 0,
                        shipping: 0,
                        gross: 0,
                        qty: 0
                    };
                }
                monthlyAgg[key].sales += sales;
                monthlyAgg[key].cost += cost;
                monthlyAgg[key].shipping += shipping;
                monthlyAgg[key].gross += gross;
                monthlyAgg[key].qty += rec.qty;
            }

            for (const month of months) {
                for (const maker of ['aron', 'pana']) {
                    const key = month + '|' + maker;
                    if (!monthlyAgg[key] && getMonthlyRebate(settings, month, maker).fixed > 0) {
                        monthlyAgg[key] = { month, maker, sales: 0, cost: 0, shipping: 0, gross: 0, qty: 0 };
                    }
                }
            }

            const monthSalesTotals = {};
            let totalGross = 0;
            for (const e of Object.values(monthlyAgg)) {
                monthSalesTotals[e.month] = (monthSalesTotals[e.month] || 0) + e.sales;
                totalGross += e.gross;
            }

            const makerSeed = { aron: 0, pana: 0, other: 0 };
            const rebateVariableByMaker = { ...makerSeed };
            const rebateFixedByMaker = { ...makerSeed };
            const rebateTotalByMaker = { ...makerSeed };
            let totalRebate = 0;
            let totalMinus = 0;
            for (const e of Object.values(monthlyAgg)) {
                const maker = (e.maker === 'aron' || e.maker === 'pana') ? e.maker : 'other';
                const rb = calcMonthlyRebate(e, settings);
                totalRebate += rb.total;
                rebateVariableByMaker[maker] += rb.variable;
                rebateFixedByMaker[maker] += rb.fixed;
                rebateTotalByMaker[maker] += rb.total;
                totalMinus += calcMonthlyMinus(e, settings, monthSalesTotals).total;
            }
            const real = totalGross + totalRebate - totalMinus;
            return {
                real,
                rebate: {
                    total: totalRebate,
                    variableByMaker: rebateVariableByMaker,
                    fixedByMaker: rebateFixedByMaker,
                    totalByMaker: rebateTotalByMaker
                }
            };
        };

        const before = evalScenario(0);
        const after = evalScenario(rateChange);
        const diff = after.real - before.real;
        document.getElementById(pfx('sim-before')).textContent = fmtYen(before.real);
        document.getElementById(pfx('sim-after')).textContent = fmtYen(after.real);
        document.getElementById(pfx('sim-after')).className = 'sim-value ' + (after.real >= 0 ? 'positive' : 'negative');
        document.getElementById(pfx('sim-diff')).textContent = (diff >= 0 ? '+' : '') + fmtYen(diff);
        document.getElementById(pfx('sim-diff')).className = 'sim-value ' + (diff >= 0 ? 'positive' : 'negative');

        const makerDiff = (afterMap, beforeMap) => ({
            aron: toNum(afterMap.aron) - toNum(beforeMap.aron),
            pana: toNum(afterMap.pana) - toNum(beforeMap.pana),
            other: toNum(afterMap.other) - toNum(beforeMap.other)
        });
        const diffVariable = makerDiff(after.rebate.variableByMaker, before.rebate.variableByMaker);
        const diffFixed = makerDiff(after.rebate.fixedByMaker, before.rebate.fixedByMaker);
        const diffTotal = makerDiff(after.rebate.totalByMaker, before.rebate.totalByMaker);
        const row = (label, total, byMaker, isDiff) => {
            const tClass = isDiff ? (total >= 0 ? 'positive' : 'negative') : '';
            const aClass = isDiff ? (byMaker.aron >= 0 ? 'positive' : 'negative') : '';
            const pClass = isDiff ? (byMaker.pana >= 0 ? 'positive' : 'negative') : '';
            const oClass = isDiff ? (byMaker.other >= 0 ? 'positive' : 'negative') : '';
            const f = (v) => isDiff ? ((v >= 0 ? '+' : '') + fmtYen(v)) : fmtYen(v);
            return `<tr><td>${label}</td><td class="${tClass}">${f(total)}</td><td class="${aClass}">${f(byMaker.aron)}</td><td class="${pClass}">${f(byMaker.pana)}</td><td class="${oClass}">${f(byMaker.other)}</td></tr>`;
        };
        if (simRebateTbody) {
            simRebateTbody.innerHTML = [
                row('変動前(率)', before.rebate.variableByMaker.aron + before.rebate.variableByMaker.pana + before.rebate.variableByMaker.other, before.rebate.variableByMaker, false),
                row('変動前(固定)', before.rebate.fixedByMaker.aron + before.rebate.fixedByMaker.pana + before.rebate.fixedByMaker.other, before.rebate.fixedByMaker, false),
                row('変動前(合計)', before.rebate.total, before.rebate.totalByMaker, false),
                row('変動後(率)', after.rebate.variableByMaker.aron + after.rebate.variableByMaker.pana + after.rebate.variableByMaker.other, after.rebate.variableByMaker, false),
                row('変動後(固定)', after.rebate.fixedByMaker.aron + after.rebate.fixedByMaker.pana + after.rebate.fixedByMaker.other, after.rebate.fixedByMaker, false),
                row('変動後(合計)', after.rebate.total, after.rebate.totalByMaker, false),
                row('差分(率)', (after.rebate.variableByMaker.aron + after.rebate.variableByMaker.pana + after.rebate.variableByMaker.other) - (before.rebate.variableByMaker.aron + before.rebate.variableByMaker.pana + before.rebate.variableByMaker.other), diffVariable, true),
                row('差分(固定)', (after.rebate.fixedByMaker.aron + after.rebate.fixedByMaker.pana + after.rebate.fixedByMaker.other) - (before.rebate.fixedByMaker.aron + before.rebate.fixedByMaker.pana + before.rebate.fixedByMaker.other), diffFixed, true),
                row('差分(合計)', after.rebate.total - before.rebate.total, diffTotal, true)
            ].join('');
        }
        if (simRebateSummary) {
            const makerLabel = target === 'all' ? '全体' : (target === 'aron' ? 'アロン化成のみ' : 'パナソニックのみ');
            const rebateDiff = after.rebate.total - before.rebate.total;
            simRebateSummary.textContent = `対象: ${makerLabel} / リベート差分: ${(rebateDiff >= 0 ? '+' : '') + fmtYen(rebateDiff)} / 実利益差分: ${(diff >= 0 ? '+' : '') + fmtYen(diff)} / リベートは設定タブの率・固定値を自動適用`;
        }

        const steps = [], gv = [];
        for (let pct = -20; pct <= 20; pct += 2) {
            steps.push((pct >= 0 ? '+' : '') + pct + '%');
            gv.push(evalScenario(pct / 100).real);
        }
        destroyChart(state.charts, 'simulation');
        state.charts['simulation'] = new Chart(document.getElementById(pfx('chart-sim')), {
            type: 'line',
            data: { labels: steps, datasets: [{ label: '実利益', data: gv, borderColor: '#ffa726', backgroundColor: 'rgba(255,167,38,0.1)', fill: true, tension: 0.3, borderWidth: 3, pointRadius: 4, pointBackgroundColor: gv.map(v => v >= 0 ? '#66bb6a' : '#ef5350') }] },
            options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { ticks: { callback: v => '¥' + fmt(v) } } } }
        });
    }

    function clamp(n, min, max) {
        return Math.max(min, Math.min(max, n));
    }

    function calcTrendRate(records, makerFilter) {
        const monthMap = {};
        for (const rec of records) {
            if (makerFilter !== 'all' && rec.maker !== makerFilter) continue;
            if (!rec.month || rec.month === 'unknown') continue;
            if (!monthMap[rec.month]) monthMap[rec.month] = { qty: 0 };
            monthMap[rec.month].qty += rec.qty;
        }
        const months = Object.keys(monthMap).sort();
        if (months.length < 2) return { trendRate: 0, monthsUsed: months.length };

        const growth = [];
        for (let i = 1; i < months.length; i++) {
            const prev = monthMap[months[i - 1]].qty;
            const curr = monthMap[months[i]].qty;
            if (prev <= 0 || curr <= 0) continue;
            growth.push((curr - prev) / prev);
        }
        if (growth.length === 0) return { trendRate: 0, monthsUsed: months.length };

        const recent = growth.slice(-6).sort((a, b) => a - b);
        const trimmed = recent.length >= 5 ? recent.slice(1, recent.length - 1) : recent;
        const avg = trimmed.reduce((sum, x) => sum + x, 0) / Math.max(1, trimmed.length);
        return { trendRate: clamp(avg, -0.35, 0.35), monthsUsed: months.length };
    }

    function calcPriceElasticity(records, makerFilter) {
        const skuMap = {};
        for (const rec of records) {
            if (makerFilter !== 'all' && rec.maker !== makerFilter) continue;
            if (!rec.jan || !rec.month || rec.month === 'unknown') continue;
            if (rec.unitPrice <= 0 || rec.qty <= 0) continue;
            const key = rec.maker + '|' + rec.jan;
            if (!skuMap[key]) skuMap[key] = { qty: 0, byMonth: {} };
            const sku = skuMap[key];
            sku.qty += rec.qty;
            if (!sku.byMonth[rec.month]) sku.byMonth[rec.month] = { qty: 0, priceQty: 0 };
            sku.byMonth[rec.month].qty += rec.qty;
            sku.byMonth[rec.month].priceQty += rec.unitPrice * rec.qty;
        }

        let weighted = 0;
        let totalWeight = 0;
        let skuUsed = 0;
        let pointCount = 0;
        for (const sku of Object.values(skuMap)) {
            const points = Object.values(sku.byMonth)
                .filter(p => p.qty > 0 && p.priceQty > 0)
                .map(p => ({ p: p.priceQty / p.qty, q: p.qty }))
                .filter(x => x.p > 0 && x.q > 0);
            if (points.length < 3) continue;
            const xs = points.map(v => Math.log(v.p));
            const ys = points.map(v => Math.log(v.q));
            const n = xs.length;
            const avgX = xs.reduce((sum, x) => sum + x, 0) / n;
            const avgY = ys.reduce((sum, y) => sum + y, 0) / n;
            let cov = 0, varX = 0;
            for (let i = 0; i < n; i++) {
                cov += (xs[i] - avgX) * (ys[i] - avgY);
                varX += (xs[i] - avgX) * (xs[i] - avgX);
            }
            if (varX <= 1e-9) continue;
            const slope = clamp(cov / varX, -3.0, 0.5);
            const w = Math.sqrt(Math.max(1, sku.qty));
            weighted += slope * w;
            totalWeight += w;
            skuUsed++;
            pointCount += n;
        }

        if (totalWeight <= 0) {
            return { elasticity: -1.1, skuUsed: 0, pointCount: 0, confidence: '低' };
        }
        const elasticity = clamp(weighted / totalWeight, -2.5, 0.2);
        const confidence = skuUsed >= 15 ? '高' : skuUsed >= 5 ? '中' : '低';
        return { elasticity, skuUsed, pointCount, confidence };
    }

    function calcSeasonalityFactors(records, makerFilter) {
        const qtyByMonthPart = {};
        for (const rec of records) {
            if (makerFilter !== 'all' && rec.maker !== makerFilter) continue;
            const month = toStr(rec.month);
            if (!isValidMonthToken(month)) continue;
            const monthPart = month.slice(5, 7);
            qtyByMonthPart[monthPart] = (qtyByMonthPart[monthPart] || 0) + rec.qty;
        }

        const monthParts = Object.keys(qtyByMonthPart).sort();
        if (monthParts.length === 0) {
            return { factors: {}, monthsUsed: 0, confidence: '低' };
        }

        const avg = monthParts.reduce((sum, mm) => sum + qtyByMonthPart[mm], 0) / Math.max(1, monthParts.length);
        if (avg <= 0) {
            return { factors: {}, monthsUsed: monthParts.length, confidence: '低' };
        }

        const factors = {};
        for (const mm of monthParts) {
            factors[mm] = clamp(qtyByMonthPart[mm] / avg, 0.6, 1.4);
        }
        const confidence = monthParts.length >= 10 ? '高' : monthParts.length >= 6 ? '中' : '低';
        return { factors, monthsUsed: monthParts.length, confidence };
    }

    function simulateStoreMetrics(records, makerFilter, simRateChange, simIncreaseQty, trendRate, elasticity, simContext, seasonalityFactors) {
        const settings = simContext?.settings || getSettings();
        const scopeSalesByMakerMonth = simContext?.scopeSalesByMakerMonth || {};
        const scopeSalesByMonth = simContext?.scopeSalesByMonth || {};
        const factors = seasonalityFactors && typeof seasonalityFactors === 'object'
            ? seasonalityFactors
            : {};
        const readSeasonality = (month) => {
            const token = toStr(month);
            if (!isValidMonthToken(token)) return 1;
            return clamp(toNum(factors[token.slice(5, 7)]) || 1, 0.6, 1.4);
        };

        const before = {
            sales: 0, gross: 0, qty: 0,
            aronNumerator: 0, aronDenominator: 0,
            panaNumerator: 0, panaDenominator: 0,
            rebate: 0, minus: 0
        };
        const after = {
            sales: 0, gross: 0, qty: 0,
            aronNumerator: 0, aronDenominator: 0,
            panaNumerator: 0, panaDenominator: 0,
            rebate: 0, minus: 0
        };

        const beforeMonthly = {};
        const afterMonthly = {};
        const deltaMakerMonthSales = {};
        const deltaMonthSales = {};

        const projectedTargetBase = [];
        let targetProjectedTotal = 0;
        let targetCount = 0;
        for (const rec of records) {
            const apply = makerFilter === 'all' || rec.maker === makerFilter;
            if (!apply) continue;
            targetCount++;
            const priceFactor = Math.max(0.05, 1 + simRateChange);
            const elasticityFactor = Math.pow(priceFactor, elasticity);
            const seasonalityFactor = readSeasonality(rec.month);
            const qtyBase = rec.qty * (1 + trendRate) * clamp(elasticityFactor, 0.2, 3.0) * seasonalityFactor;
            projectedTargetBase.push({ rec, qtyBase });
            targetProjectedTotal += qtyBase;
        }

        for (const rec of records) {
            const month = rec.month || 'unknown';
            const maker = rec.maker || 'other';
            const key = month + '|' + maker;

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
            if (!beforeMonthly[key]) {
                beforeMonthly[key] = { month, maker, sales: 0, cost: 0, shipping: 0, gross: 0, qty: 0 };
            }
            beforeMonthly[key].sales += rec.salesAmount;
            beforeMonthly[key].cost += rec.totalCost;
            beforeMonthly[key].shipping += rec.totalShipping;
            beforeMonthly[key].gross += rec.grossProfit;
            beforeMonthly[key].qty += rec.qty;

            const applyChange = makerFilter === 'all' || rec.maker === makerFilter;
            const priceFactor = Math.max(0.05, 1 + simRateChange);
            const elasticityFactor = applyChange ? Math.pow(priceFactor, elasticity) : 1;
            const seasonalityFactor = applyChange ? readSeasonality(rec.month) : 1;
            let newQty = rec.qty * (1 + trendRate) * clamp(elasticityFactor, 0.2, 3.0) * seasonalityFactor;
            if (applyChange && simIncreaseQty > 0) {
                if (targetProjectedTotal > 0) newQty += simIncreaseQty * (newQty / targetProjectedTotal);
                else if (targetCount > 0) newQty += simIncreaseQty / targetCount;
            }
            const newUnitPrice = applyChange ? rec.unitPrice * (1 + simRateChange) : rec.unitPrice;
            const newSales = newQty * newUnitPrice;
            const newCost = newQty * rec.effectiveCost;
            const newShipping = newQty * rec.shippingCost;
            const newGross = newSales - newCost - newShipping;

            after.sales += newSales;
            after.gross += newGross;
            after.qty += newQty;
            if (!afterMonthly[key]) {
                afterMonthly[key] = { month, maker, sales: 0, cost: 0, shipping: 0, gross: 0, qty: 0 };
            }
            afterMonthly[key].sales += newSales;
            afterMonthly[key].cost += newCost;
            afterMonthly[key].shipping += newShipping;
            afterMonthly[key].gross += newGross;
            afterMonthly[key].qty += newQty;
            deltaMakerMonthSales[key] = (deltaMakerMonthSales[key] || 0) + (newSales - rec.salesAmount);
            deltaMonthSales[month] = (deltaMonthSales[month] || 0) + (newSales - rec.salesAmount);

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

        const calcRebateMinus = (monthlyMap, isAfter) => {
            const out = { rebate: 0, minus: 0 };
            for (const e of Object.values(monthlyMap)) {
                const key = e.month + '|' + e.maker;
                const baseMakerSales = toNum(scopeSalesByMakerMonth[key]);
                const baseMonthSales = toNum(scopeSalesByMonth[e.month]);
                const makerSalesTotal = Math.max(0, baseMakerSales + (isAfter ? (deltaMakerMonthSales[key] || 0) : 0));
                const monthSalesTotal = Math.max(0, baseMonthSales + (isAfter ? (deltaMonthSales[e.month] || 0) : 0));

                const variable = calcMonthlyRebate(e, settings).variable;
                const fixedTotal = getMonthlyRebate(settings, e.month, e.maker).fixed;
                const fixedShare = makerSalesTotal > 0 ? fixedTotal * (e.sales / makerSalesTotal) : 0;
                const rebate = variable + fixedShare;
                const warehouseBase = monthSalesTotal > 0 ? settings.warehouseFee * (e.sales / monthSalesTotal) : 0;
                const warehouseOut = e.qty * settings.warehouseOutFee;
                out.rebate += rebate;
                out.minus += warehouseBase + warehouseOut;
            }
            return out;
        };

        const beforeAdjust = calcRebateMinus(beforeMonthly, false);
        const afterAdjust = calcRebateMinus(afterMonthly, true);
        before.rebate = beforeAdjust.rebate;
        before.minus = beforeAdjust.minus;
        after.rebate = afterAdjust.rebate;
        after.minus = afterAdjust.minus;
        before.gross = before.gross + before.rebate - before.minus;
        after.gross = after.gross + after.rebate - after.minus;

        const beforeAronRate = before.aronDenominator > 0 ? before.aronNumerator / before.aronDenominator : 0;
        const beforePanaRate = before.panaDenominator > 0 ? before.panaNumerator / before.panaDenominator : 0;
        const afterAronRate = after.aronDenominator > 0 ? after.aronNumerator / after.aronDenominator : 0;
        const afterPanaRate = after.panaDenominator > 0 ? after.panaNumerator / after.panaDenominator : 0;
        const beforeProfitRate = before.sales > 0 ? before.gross / before.sales : 0;
        const afterProfitRate = after.sales > 0 ? after.gross / after.sales : 0;

        return {
            before: {
                sales: before.sales, gross: before.gross, qty: before.qty,
                profitRate: beforeProfitRate, aronRate: beforeAronRate, panaRate: beforePanaRate
            },
            after: {
                sales: after.sales, gross: after.gross, qty: after.qty,
                profitRate: afterProfitRate, aronRate: afterAronRate, panaRate: afterPanaRate
            }
        };
    }

    function renderStoreDetail(keepPage) {
        const r = state.results; if (!r) return;
        document.getElementById(pfx('store-detail-empty')).style.display = 'none';
        document.getElementById(pfx('store-detail-content')).style.display = 'block';

        if (!state.storeDetailIndex) state.storeDetailIndex = buildStoreDetailIndex(r.records);
        const index = state.storeDetailIndex;
        const ml = { aron: 'アロン化成', pana: 'パナソニック', other: 'その他' };

        const storeSearchInput = document.getElementById(pfx('store-detail-store-search'));
        const storeSel = document.getElementById(pfx('store-detail-store'));
        const makerSel = document.getElementById(pfx('store-detail-maker'));
        const monthSel = document.getElementById(pfx('store-detail-month'));
        const repSel = document.getElementById(pfx('store-detail-rep'));
        const simMakerSel = document.getElementById(pfx('store-detail-sim-maker'));
        const simRateInput = document.getElementById(pfx('store-detail-rate'));
        const simQtyInput = document.getElementById(pfx('store-detail-qty'));
        const useTrendInput = document.getElementById(pfx('store-detail-use-trend'));
        const trendAdjustInput = document.getElementById(pfx('store-detail-trend-adjust'));
        const useElasticAutoInput = document.getElementById(pfx('store-detail-use-elastic-auto'));
        const elasticityInput = document.getElementById(pfx('store-detail-elasticity'));
        const useSeasonalityInput = document.getElementById(pfx('store-detail-use-seasonality'));
        const productSearchInput = document.getElementById(pfx('store-detail-product-search'));
        const productLimitSel = document.getElementById(pfx('store-detail-limit'));
        const summaryEl = document.getElementById(pfx('store-detail-summary'));
        const modelSummaryEl = document.getElementById(pfx('store-detail-model-summary'));
        const simTbody = document.getElementById(pfx('store-detail-sim-tbody'));
        const productsTbody = document.getElementById(pfx('store-detail-products-tbody'));
        const productSummaryEl = document.getElementById(pfx('store-detail-product-summary'));
        const pagerEl = document.getElementById(pfx('store-detail-pagination'));
        const pageStatusEl = document.getElementById(pfx('store-detail-page-status'));
        const prevBtn = document.getElementById(pfx('store-detail-page-prev'));
        const nextBtn = document.getElementById(pfx('store-detail-page-next'));

        const updateAdvancedInputState = () => {
            if (trendAdjustInput) trendAdjustInput.disabled = !(useTrendInput && useTrendInput.checked);
            if (elasticityInput) elasticityInput.disabled = !!(useElasticAutoInput && useElasticAutoInput.checked);
        };
        updateAdvancedInputState();

        const searchToken = normalizeToken(storeSearchInput?.value || '');
        const filteredStores = index.stores.filter(storeName => {
            if (!searchToken) return true;
            if (normalizeToken(storeName).includes(searchToken)) return true;
            const info = index.map[storeName];
            if (!info) return false;
            for (const code of info.supplierCodes || []) {
                if (normalizeToken(code).includes(searchToken)) return true;
            }
            return false;
        });

        const storeVersion = `${searchToken}|${filteredStores.length}|${filteredStores[0] || ''}|${filteredStores[filteredStores.length - 1] || ''}`;
        if (storeSel.dataset.version !== storeVersion) {
            const prevStore = storeSel.value;
            if (filteredStores.length === 0) {
                storeSel.innerHTML = '<option value="">該当なし</option>';
            } else {
                storeSel.innerHTML = filteredStores.map(storeName => {
                    const info = index.map[storeName];
                    const codeLabel = info && info.supplierCodes.size > 0
                        ? ` [${[...info.supplierCodes].sort().join(',')}]`
                        : '';
                    return `<option value="${escHtml(storeName)}">${escHtml(storeName)}${escHtml(codeLabel)}</option>`;
                }).join('');
            }
            storeSel.value = filteredStores.includes(prevStore) ? prevStore : (filteredStores[0] || '');
            storeSel.dataset.version = storeVersion;
            state.storeDetailCurrentPage = 1;
        }

        const selectedStore = storeSel.value;
        const storeInfo = index.map[selectedStore];
        if (!storeInfo) {
            if (summaryEl) summaryEl.textContent = '対象の販売店データがありません。';
            if (modelSummaryEl) modelSummaryEl.textContent = '';
            if (simTbody) simTbody.innerHTML = '<tr><td colspan="7">データがありません</td></tr>';
            if (productsTbody) productsTbody.innerHTML = '<tr><td colspan="9">データがありません</td></tr>';
            if (productSummaryEl) productSummaryEl.textContent = '表示: 0件 / 全0件';
            if (pagerEl) pagerEl.style.display = 'none';
            return;
        }

        const months = [...storeInfo.months].filter(Boolean).sort();
        const reps = [...storeInfo.reps].filter(Boolean).sort((a, b) => a.localeCompare(b, 'ja'));
        const monthVersion = `${selectedStore}|${months.join('|')}`;
        const repVersion = `${selectedStore}|${reps.join('|')}`;

        if (monthSel.dataset.version !== monthVersion) {
            const prevMonth = monthSel.value;
            monthSel.innerHTML = '<option value="all">全期間</option>';
            for (const m of months) monthSel.innerHTML += `<option value="${escHtml(m)}">${escHtml(m)}</option>`;
            monthSel.value = months.includes(prevMonth) ? prevMonth : 'all';
            monthSel.dataset.version = monthVersion;
            if (!keepPage) state.storeDetailCurrentPage = 1;
        }
        if (repSel.dataset.version !== repVersion) {
            const prevRep = repSel.value;
            repSel.innerHTML = '<option value="all">全担当</option>';
            for (const rep of reps) repSel.innerHTML += `<option value="${escHtml(rep)}">${escHtml(rep)}</option>`;
            repSel.value = reps.includes(prevRep) ? prevRep : 'all';
            repSel.dataset.version = repVersion;
            if (!keepPage) state.storeDetailCurrentPage = 1;
        }

        const makerF = makerSel.value;
        const monthF = monthSel.value;
        const repF = repSel.value;
        const simMaker = simMakerSel.value;
        const simRateChange = toNum(simRateInput.value) / 100;
        const simIncreaseQty = Math.max(0, toNum(simQtyInput.value));
        const useTrend = !!(useTrendInput && useTrendInput.checked);
        const trendAdjustRate = toNum(trendAdjustInput?.value) / 100;
        const useElasticAuto = !!(useElasticAutoInput && useElasticAutoInput.checked);
        const manualElasticity = toNum(elasticityInput?.value);
        const useSeasonality = !!(useSeasonalityInput && useSeasonalityInput.checked);

        const scopeRecords = r.records.filter(rec => {
            if (makerF !== 'all' && rec.maker !== makerF) return false;
            if (monthF !== 'all' && rec.month !== monthF) return false;
            if (repF !== 'all' && rec.salesRep !== repF) return false;
            return true;
        });

        const scopeSalesByMakerMonth = {};
        const scopeSalesByMonth = {};
        for (const rec of scopeRecords) {
            const month = rec.month || 'unknown';
            const maker = rec.maker || 'other';
            const key = month + '|' + maker;
            scopeSalesByMakerMonth[key] = (scopeSalesByMakerMonth[key] || 0) + rec.salesAmount;
            scopeSalesByMonth[month] = (scopeSalesByMonth[month] || 0) + rec.salesAmount;
        }
        const simContext = {
            settings: r.settings || getSettings(),
            scopeSalesByMakerMonth,
            scopeSalesByMonth
        };

        const filteredRecords = storeInfo.records.filter(rec => {
            if (makerF !== 'all' && rec.maker !== makerF) return false;
            if (monthF !== 'all' && rec.month !== monthF) return false;
            if (repF !== 'all' && rec.salesRep !== repF) return false;
            return true;
        });

        const fmtQty = (n) => (n == null || isNaN(n)) ? '-' : (Math.round(n * 10) / 10).toLocaleString('ja-JP');
        const signed = (n, formatter) => (n >= 0 ? '+' : '') + formatter(n);
        const makeDiff = (after, before) => ({
            sales: after.sales - before.sales,
            gross: after.gross - before.gross,
            qty: after.qty - before.qty,
            profitRate: after.profitRate - before.profitRate,
            aronRate: after.aronRate - before.aronRate,
            panaRate: after.panaRate - before.panaRate
        });

        if (filteredRecords.length === 0) {
            if (summaryEl) summaryEl.textContent = `対象販売店: ${selectedStore} / 条件に一致するデータがありません。`;
            if (modelSummaryEl) modelSummaryEl.textContent = '';
            if (simTbody) simTbody.innerHTML = '<tr><td colspan="7">データがありません</td></tr>';
            if (productsTbody) productsTbody.innerHTML = '<tr><td colspan="9">データがありません</td></tr>';
            if (productSummaryEl) productSummaryEl.textContent = '表示: 0件 / 全0件';
            if (pagerEl) pagerEl.style.display = 'none';
            return;
        }

        const baseScenario = simulateStoreMetrics(filteredRecords, simMaker, simRateChange, simIncreaseQty, 0, 0, simContext, {});
        const trend = calcTrendRate(filteredRecords, simMaker);
        const elasticity = calcPriceElasticity(filteredRecords, simMaker);
        const seasonality = calcSeasonalityFactors(filteredRecords, simMaker);
        const appliedTrendRate = useTrend ? clamp(trend.trendRate + trendAdjustRate, -0.5, 0.5) : 0;
        const appliedElasticity = useElasticAuto ? elasticity.elasticity : manualElasticity;
        const appliedSeasonality = useSeasonality ? seasonality.factors : {};
        const advancedScenario = simulateStoreMetrics(
            filteredRecords,
            simMaker,
            simRateChange,
            simIncreaseQty,
            appliedTrendRate,
            appliedElasticity,
            simContext,
            appliedSeasonality
        );
        const baseDiff = makeDiff(baseScenario.after, baseScenario.before);
        const advancedDiff = makeDiff(advancedScenario.after, baseScenario.before);

        const supplierCodes = [...storeInfo.supplierCodes].sort();
        const supplierLabel = supplierCodes.length > 0 ? supplierCodes.join(',') : '-';
        if (summaryEl) {
            summaryEl.textContent =
                `対象販売店: ${selectedStore} / 得意先コード: ${supplierLabel} / 明細 ${fmt(filteredRecords.length)}件 / 現状売上 ${fmtYen(baseScenario.before.sales)} / 現状実利益 ${fmtYen(baseScenario.before.gross)} / 現状実利益率 ${fmtPct(baseScenario.before.profitRate)}`;
        }
        if (modelSummaryEl) {
            const trendLabel = useTrend
                ? `ON(${fmtPct(appliedTrendRate)}: 推定${fmtPct(trend.trendRate)} + 補正${fmtPct(trendAdjustRate)})`
                : `OFF(推定${fmtPct(trend.trendRate)})`;
            const elasticityLabel = useElasticAuto
                ? `自動ON(${elasticity.elasticity.toFixed(2)} / SKU ${fmt(elasticity.skuUsed)} / 信頼度 ${elasticity.confidence})`
                : `手動(${appliedElasticity.toFixed(2)})`;
            const seasonalityLabel = useSeasonality
                ? `ON(月カテゴリ${fmt(seasonality.monthsUsed)} / 信頼度 ${seasonality.confidence})`
                : 'OFF';
            modelSummaryEl.textContent =
                `高度シミュレーター適用: トレンド ${trendLabel} / 価格弾力性 ${elasticityLabel} / 季節性 ${seasonalityLabel}`;
        }
        if (simTbody) {
            simTbody.innerHTML = [
                `<tr><td>現状</td><td>${fmtYen(baseScenario.before.sales)}</td><td class="${baseScenario.before.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(baseScenario.before.gross)}</td><td>${fmtPct(baseScenario.before.profitRate)}</td><td>${fmtQty(baseScenario.before.qty)}</td><td>${baseScenario.before.aronRate > 0 ? fmtPct(baseScenario.before.aronRate) : '-'}</td><td>${baseScenario.before.panaRate > 0 ? fmtPct(baseScenario.before.panaRate) : '-'}</td></tr>`,
                `<tr><td>実績ベース変更後</td><td>${fmtYen(baseScenario.after.sales)}</td><td class="${baseScenario.after.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(baseScenario.after.gross)}</td><td>${fmtPct(baseScenario.after.profitRate)}</td><td>${fmtQty(baseScenario.after.qty)}</td><td>${baseScenario.after.aronRate > 0 ? fmtPct(baseScenario.after.aronRate) : '-'}</td><td>${baseScenario.after.panaRate > 0 ? fmtPct(baseScenario.after.panaRate) : '-'}</td></tr>`,
                `<tr><td>高度シミュレーション</td><td>${fmtYen(advancedScenario.after.sales)}</td><td class="${advancedScenario.after.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(advancedScenario.after.gross)}</td><td>${fmtPct(advancedScenario.after.profitRate)}</td><td>${fmtQty(advancedScenario.after.qty)}</td><td>${advancedScenario.after.aronRate > 0 ? fmtPct(advancedScenario.after.aronRate) : '-'}</td><td>${advancedScenario.after.panaRate > 0 ? fmtPct(advancedScenario.after.panaRate) : '-'}</td></tr>`,
                `<tr style="font-weight:700;background:#e3f2fd;"><td>高度差分(対現状)</td><td class="${advancedDiff.sales >= 0 ? 'positive' : 'negative'}">${signed(advancedDiff.sales, fmtYen)}</td><td class="${advancedDiff.gross >= 0 ? 'positive' : 'negative'}">${signed(advancedDiff.gross, fmtYen)}</td><td class="${advancedDiff.profitRate >= 0 ? 'positive' : 'negative'}">${signed(advancedDiff.profitRate, fmtPct)}</td><td class="${advancedDiff.qty >= 0 ? 'positive' : 'negative'}">${advancedDiff.qty >= 0 ? '+' : ''}${fmtQty(advancedDiff.qty)}</td><td class="${advancedDiff.aronRate >= 0 ? 'positive' : 'negative'}">${signed(advancedDiff.aronRate, fmtPct)}</td><td class="${advancedDiff.panaRate >= 0 ? 'positive' : 'negative'}">${signed(advancedDiff.panaRate, fmtPct)}</td></tr>`,
                `<tr style="font-weight:700;background:#fff3e0;"><td>実績差分(対現状)</td><td class="${baseDiff.sales >= 0 ? 'positive' : 'negative'}">${signed(baseDiff.sales, fmtYen)}</td><td class="${baseDiff.gross >= 0 ? 'positive' : 'negative'}">${signed(baseDiff.gross, fmtYen)}</td><td class="${baseDiff.profitRate >= 0 ? 'positive' : 'negative'}">${signed(baseDiff.profitRate, fmtPct)}</td><td class="${baseDiff.qty >= 0 ? 'positive' : 'negative'}">${baseDiff.qty >= 0 ? '+' : ''}${fmtQty(baseDiff.qty)}</td><td class="${baseDiff.aronRate >= 0 ? 'positive' : 'negative'}">${signed(baseDiff.aronRate, fmtPct)}</td><td class="${baseDiff.panaRate >= 0 ? 'positive' : 'negative'}">${signed(baseDiff.panaRate, fmtPct)}</td></tr>`
            ].join('');
        }

        const productMap = {};
        for (const rec of filteredRecords) {
            const key = rec.jan || '(未設定)';
            if (!productMap[key]) {
                productMap[key] = {
                    jan: rec.jan,
                    name: rec.name,
                    maker: rec.maker,
                    qty: 0,
                    sales: 0,
                    cost: 0,
                    shipping: 0,
                    gross: 0
                };
            }
            const e = productMap[key];
            e.qty += rec.qty;
            e.sales += rec.salesAmount;
            e.cost += rec.totalCost;
            e.shipping += rec.totalShipping;
            e.gross += rec.grossProfit;
        }

        const search = (productSearchInput.value || '').toLowerCase().trim();
        let products = Object.values(productMap);
        if (search) {
            products = products.filter(e =>
                (e.jan || '').toLowerCase().includes(search) ||
                (e.name || '').toLowerCase().includes(search)
            );
        }
        products.sort((a, b) => b.gross - a.gross || b.sales - a.sales || a.jan.localeCompare(b.jan, 'ja'));

        const limitRaw = productLimitSel.value;
        const isAll = limitRaw === 'all';
        const pageSize = isAll ? Math.max(1, products.length || 1) : Math.max(1, toNum(limitRaw) || 200);
        const totalPages = isAll ? 1 : Math.max(1, Math.ceil(products.length / pageSize));
        state.storeDetailCurrentPageTotal = totalPages;
        state.storeDetailCurrentPage = isAll ? 1 : Math.min(Math.max(1, state.storeDetailCurrentPage || 1), totalPages);
        const startIndex = isAll ? 0 : (state.storeDetailCurrentPage - 1) * pageSize;
        const displayed = products.slice(startIndex, startIndex + pageSize);

        if (productsTbody) {
            productsTbody.innerHTML = displayed.map(e => {
                const pr = e.sales > 0 ? e.gross / e.sales : 0;
                return `<tr><td>${escHtml(e.jan || '-')}</td><td>${escHtml(e.name || '-')}</td><td>${escHtml(ml[e.maker] || e.maker)}</td><td>${fmt(e.qty)}</td><td>${fmtYen(e.sales)}</td><td>${fmtYen(e.cost)}</td><td>${fmtYen(e.shipping)}</td><td class="${e.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(e.gross)}</td><td>${fmtPct(pr)}</td></tr>`;
            }).join('');
        }
        if (productSummaryEl) {
            if (products.length === 0) {
                productSummaryEl.textContent = '表示: 0件 / 全0件';
            } else {
                const from = startIndex + 1;
                const to = startIndex + displayed.length;
                productSummaryEl.textContent = `表示: ${fmt(from)}-${fmt(to)}件 / 全${fmt(products.length)}件`;
            }
        }
        if (pagerEl && pageStatusEl && prevBtn && nextBtn) {
            pagerEl.style.display = totalPages > 1 ? 'flex' : 'none';
            pageStatusEl.textContent = `${fmt(state.storeDetailCurrentPage)} / ${fmt(totalPages)}ページ`;
            prevBtn.disabled = state.storeDetailCurrentPage <= 1;
            nextBtn.disabled = state.storeDetailCurrentPage >= totalPages;
        }
    }

    function escHtml(v) {
        return String(v ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function padChunkId(index) {
        return String(index).padStart(4, '0');
    }

    function splitToChunks(text, size) {
        const chunks = [];
        for (let i = 0; i < text.length; i += size) {
            chunks.push(text.slice(i, i + size));
        }
        return chunks.length > 0 ? chunks : [''];
    }

    function writeLocalAutoState(serialized) {
        let prevCount = 0;
        try {
            const prevMetaRaw = localStorage.getItem(AUTO_STATE_META_KEY);
            if (prevMetaRaw) {
                const prevMeta = JSON.parse(prevMetaRaw);
                prevCount = toNum(prevMeta?.chunkCount);
            }
        } catch (err) {
            prevCount = 0;
        }

        const chunks = splitToChunks(serialized, AUTO_STATE_CHUNK_SIZE);
        localStorage.setItem(AUTO_STATE_META_KEY, JSON.stringify({
            chunkCount: chunks.length,
            savedAt: new Date().toISOString()
        }));
        for (let i = 0; i < chunks.length; i++) {
            localStorage.setItem(AUTO_STATE_CHUNK_PREFIX + padChunkId(i), chunks[i]);
        }
        for (let i = chunks.length; i < prevCount; i++) {
            localStorage.removeItem(AUTO_STATE_CHUNK_PREFIX + padChunkId(i));
        }
        // legacy key cleanup
        localStorage.removeItem(AUTO_STATE_STORAGE_KEY);
    }

    function readLocalAutoState() {
        const metaRaw = localStorage.getItem(AUTO_STATE_META_KEY);
        if (metaRaw) {
            const meta = JSON.parse(metaRaw);
            const chunkCount = toNum(meta?.chunkCount);
            if (chunkCount > 0) {
                let serialized = '';
                for (let i = 0; i < chunkCount; i++) {
                    serialized += localStorage.getItem(AUTO_STATE_CHUNK_PREFIX + padChunkId(i)) || '';
                }
                if (serialized) return serialized;
            }
        }
        return localStorage.getItem(AUTO_STATE_STORAGE_KEY) || '';
    }

    function buildAutoStatePayload() {
        return {
            schemaVersion: 1,
            savedAt: new Date().toISOString(),
            shippingData: state.shippingData,
            salesData: state.salesData,
            productData: state.productData,
            settings: getSettings(),
            progressItems: state.progressItems
        };
    }

    function saveAutoStateNow() {
        const payload = buildAutoStatePayload();
        try {
            writeLocalAutoState(JSON.stringify(payload));
        } catch (err) {
            console.warn('saveAutoStateNow failed', err);
        }
        scheduleCloudStateSave(0);
    }

    async function persistCloudStateNow() {
        if (!window.KaientaiCloud || typeof window.KaientaiCloud.saveModuleState !== 'function') {
            throw new Error('クラウドAPIを利用できません');
        }
        if (!(window.KaientaiCloud.isReady && window.KaientaiCloud.isReady())) {
            throw new Error('クラウド保存の準備ができていません');
        }
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
                cloudPersistRetryMs = 1000;
            } catch (err) {
                console.warn('persistCloudStateNow failed', err);
                cloudPersistPending = true;
                cloudPersistRetryMs = Math.min(30000, Math.max(1000, cloudPersistRetryMs * 2));
            } finally {
                cloudPersistInFlight = false;
                if (cloudPersistPending) {
                    cloudPersistPending = false;
                    scheduleCloudStateSave(cloudPersistRetryMs);
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

    function scheduleAutoStateSave(delay = 600) {
        if (autoPersistTimer) clearTimeout(autoPersistTimer);
        autoPersistTimer = setTimeout(() => {
            autoPersistTimer = null;
            saveAutoStateNow();
        }, Math.max(0, delay));
    }

    function clearAutoState() {
        if (autoPersistTimer) {
            clearTimeout(autoPersistTimer);
            autoPersistTimer = null;
        }
        if (cloudPersistTimer) {
            clearTimeout(cloudPersistTimer);
            cloudPersistTimer = null;
        }
        try {
            localStorage.removeItem(AUTO_STATE_STORAGE_KEY);
            const metaRaw = localStorage.getItem(AUTO_STATE_META_KEY);
            if (metaRaw) {
                const meta = JSON.parse(metaRaw);
                const chunkCount = toNum(meta?.chunkCount);
                for (let i = 0; i < chunkCount; i++) {
                    localStorage.removeItem(AUTO_STATE_CHUNK_PREFIX + padChunkId(i));
                }
            }
            localStorage.removeItem(AUTO_STATE_META_KEY);
            scheduleCloudStateSave(300);
        } catch (err) {
            console.warn('clearAutoState failed', err);
        }
    }

    function applyLoadedPayload(payload, sourceLabel) {
        if (!payload || !Array.isArray(payload.shippingData) || !Array.isArray(payload.salesData) || !Array.isArray(payload.productData)) {
            return false;
        }

        state.shippingData = payload.shippingData;
        state.salesData = payload.salesData;
        state.productData = payload.productData;
        if (Array.isArray(payload.progressItems)) {
            state.progressItems = payload.progressItems.filter(x => x && typeof x === 'object');
            saveProgressItems();
        }

        state.results = null;
        state.storeBaseCache = {};
        state.storeSortedCache = {};
        state.storeViewRuntime = null;
        state.storeAdvancedEnabled = false;
        state.storeCurrentPage = 1;
        state.storeCurrentPageTotal = 1;
        state.detailsCurrentPage = 1;
        state.detailsCurrentPageTotal = 1;
        state.storeDetailIndex = null;
        state.storeDetailCurrentPage = 1;
        state.storeDetailCurrentPageTotal = 1;
        state.progressCurrentPage = 1;
        state.progressCurrentPageTotal = 1;
        state.progressEditingId = '';
        state.storeHeavyRenderToken = 0;
        if (state.storeHeavyRenderTimer) {
            clearTimeout(state.storeHeavyRenderTimer);
            state.storeHeavyRenderTimer = null;
        }
        logLines = [];
        setStoreAdvancedMode(false);

        restoreSavedSettings(payload.settings || {});
        setProgressForm(null);
        updateUploadCardsByState();
        renderProgress();
        checkAllLoaded();
        KaientaiM.updateModuleStatus(MODULE_ID, `データ読込済み (${state.salesData.length})`, true);

        ['overview', 'monthly', 'store', 'sim', 'store-detail', 'progress', 'details'].forEach(id => {
            const emp = document.getElementById(pfx(id + '-empty'));
            const con = document.getElementById(pfx(id + '-content'));
            if (emp) emp.style.display = '';
            if (con) con.style.display = 'none';
        });
        Object.keys(state.charts).forEach(k => destroyChart(state.charts, k));
        const logEl = document.getElementById(pfx('load-log'));
        if (logEl) logEl.style.display = 'none';

        const canAnalyze = state.shippingData.length > 0 && state.salesData.length > 0 && state.productData.length > 0;
        if (canAnalyze) {
            runAnalysis();
            switchModTab('overview');
        }

        if (sourceLabel) {
            log(`${sourceLabel}: shipping:${state.shippingData.length} sales:${state.salesData.length} product:${state.productData.length}`);
        }
        return true;
    }

    function restoreAutoState() {
        try {
            const raw = readLocalAutoState();
            if (!raw) return false;
            const payload = JSON.parse(raw);
            return applyLoadedPayload(payload, '自動復元が完了しました');
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
            const applied = applyLoadedPayload(payload, 'クラウド自動復元が完了しました');
            if (applied) saveAutoStateNow();
            return applied;
        } catch (err) {
            console.warn('restoreCloudStateIfNeeded failed', err);
            return false;
        }
    }

    function clearCloudRestoreRetry() {
        if (cloudRestoreRetryTimer) {
            clearTimeout(cloudRestoreRetryTimer);
            cloudRestoreRetryTimer = null;
        }
    }

    async function startCloudRestoreRetry() {
        clearCloudRestoreRetry();
        cloudRestoreRetryCount = 0;

        const attempt = async () => {
            const restored = await restoreCloudStateIfNeeded();
            if (restored) {
                clearCloudRestoreRetry();
                return;
            }
            cloudRestoreRetryCount++;
            if (cloudRestoreRetryCount >= CLOUD_RESTORE_MAX_RETRY) return;
            cloudRestoreRetryTimer = setTimeout(attempt, CLOUD_RESTORE_RETRY_MS);
        };

        await attempt();
    }

    function loadProgressItems() {
        try {
            const raw = localStorage.getItem(PROGRESS_STORAGE_KEY);
            if (!raw) { state.progressItems = []; return; }
            const parsed = JSON.parse(raw);
            if (!Array.isArray(parsed)) { state.progressItems = []; return; }
            state.progressItems = parsed.filter(x => x && typeof x === 'object').map(x => ({
                id: x.id || ('p-' + Date.now() + '-' + Math.random().toString(36).slice(2, 8)),
                salesRep: toStr(x.salesRep),
                store: toStr(x.store),
                supplierCode: toStr(x.supplierCode),
                status: toStr(x.status) || '未着手',
                plan: toStr(x.plan),
                nextDate: toStr(x.nextDate),
                expectedImpact: toNum(x.expectedImpact),
                actualImpact: toNum(x.actualImpact),
                resultMemo: toStr(x.resultMemo),
                createdAt: toStr(x.createdAt) || new Date().toISOString(),
                updatedAt: toStr(x.updatedAt) || new Date().toISOString()
            }));
        } catch (err) {
            console.warn('loadProgressItems failed', err);
            state.progressItems = [];
        }
    }

    function saveProgressItems() {
        try {
            localStorage.setItem(PROGRESS_STORAGE_KEY, JSON.stringify(state.progressItems));
        } catch (err) {
            console.warn('saveProgressItems failed', err);
        }
    }

    function getProgressMasterData() {
        const source = state.results?.records?.length ? state.results.records : state.salesData;
        const reps = new Set();
        const storeMap = {};
        for (const rec of source || []) {
            const rep = toStr(rec.salesRep);
            if (rep) reps.add(rep);
            const store = toStr(rec.store) || '(未設定)';
            if (!storeMap[store]) storeMap[store] = new Set();
            const code = toStr(rec.supplierCode);
            if (code) storeMap[store].add(code);
        }
        for (const item of state.progressItems) {
            if (item.salesRep) reps.add(item.salesRep);
            const store = item.store || '(未設定)';
            if (!storeMap[store]) storeMap[store] = new Set();
            if (item.supplierCode) storeMap[store].add(item.supplierCode);
        }
        const stores = Object.keys(storeMap).sort((a, b) => a.localeCompare(b, 'ja')).map(name => ({
            name,
            supplierCodes: [...storeMap[name]].sort()
        }));
        return {
            reps: [...reps].sort((a, b) => a.localeCompare(b, 'ja')),
            stores
        };
    }

    function setProgressForm(item) {
        const editing = item || null;
        state.progressEditingId = editing ? editing.id : '';
        document.getElementById(pfx('progress-form-rep')).value = editing?.salesRep || '';
        document.getElementById(pfx('progress-form-store')).value = editing?.store || '';
        document.getElementById(pfx('progress-form-supplier-code')).value = editing?.supplierCode || '';
        document.getElementById(pfx('progress-form-status')).value = editing?.status || '未着手';
        document.getElementById(pfx('progress-form-next-date')).value = editing?.nextDate || '';
        document.getElementById(pfx('progress-form-plan')).value = editing?.plan || '';
        document.getElementById(pfx('progress-form-expected')).value = editing ? String(editing.expectedImpact || 0) : '0';
        document.getElementById(pfx('progress-form-actual')).value = editing ? String(editing.actualImpact || 0) : '0';
        document.getElementById(pfx('progress-form-result')).value = editing?.resultMemo || '';
        const saveBtn = document.getElementById(pfx('progress-btn-save'));
        if (saveBtn) saveBtn.textContent = editing ? '更新' : '登録';
    }

    function renderProgress(keepPage) {
        document.getElementById(pfx('progress-empty')).style.display = 'none';
        document.getElementById(pfx('progress-content')).style.display = 'block';
        const signed = (n, formatter) => (n >= 0 ? '+' : '') + formatter(n);

        const repFilter = document.getElementById(pfx('progress-rep-filter'));
        const statusFilter = document.getElementById(pfx('progress-status-filter'));
        const storeSearchInput = document.getElementById(pfx('progress-store-search'));
        const limitSel = document.getElementById(pfx('progress-limit'));
        const summaryEl = document.getElementById(pfx('progress-summary'));
        const tbody = document.getElementById(pfx('progress-tbody'));
        const pagerEl = document.getElementById(pfx('progress-pagination'));
        const pageStatusEl = document.getElementById(pfx('progress-page-status'));
        const prevBtn = document.getElementById(pfx('progress-page-prev'));
        const nextBtn = document.getElementById(pfx('progress-page-next'));
        const repFormSel = document.getElementById(pfx('progress-form-rep'));
        const storeListEl = document.getElementById(pfx('progress-store-list'));

        const master = getProgressMasterData();
        const repVersion = master.reps.join('|');
        if (repFilter.dataset.version !== repVersion) {
            const prev = repFilter.value;
            repFilter.innerHTML = '<option value="all">全担当</option>' + master.reps.map(rep => `<option value="${escHtml(rep)}">${escHtml(rep)}</option>`).join('');
            repFilter.value = master.reps.includes(prev) ? prev : 'all';
            repFilter.dataset.version = repVersion;
            if (!keepPage) state.progressCurrentPage = 1;
        }
        if (repFormSel.dataset.version !== repVersion) {
            const prev = repFormSel.value;
            repFormSel.innerHTML = '<option value="">担当を選択</option>' + master.reps.map(rep => `<option value="${escHtml(rep)}">${escHtml(rep)}</option>`).join('');
            repFormSel.value = master.reps.includes(prev) ? prev : '';
            repFormSel.dataset.version = repVersion;
        }

        const storeVersion = master.stores.map(s => `${s.name}:${s.supplierCodes.join('/')}`).join('|');
        if (storeListEl.dataset.version !== storeVersion) {
            storeListEl.innerHTML = master.stores.map(s => {
                const label = s.supplierCodes.length > 0 ? `得意先コード:${s.supplierCodes.join(',')}` : '得意先コードなし';
                return `<option value="${escHtml(s.name)}" label="${escHtml(label)}"></option>`;
            }).join('');
            storeListEl.dataset.version = storeVersion;
        }

        const repF = repFilter.value;
        const statusF = statusFilter.value;
        const storeToken = normalizeToken(storeSearchInput.value);
        let items = state.progressItems.filter(item => {
            if (repF !== 'all' && item.salesRep !== repF) return false;
            if (statusF !== 'all' && item.status !== statusF) return false;
            if (storeToken) {
                const hay = normalizeToken(`${item.store} ${item.supplierCode} ${item.plan} ${item.resultMemo}`);
                if (!hay.includes(storeToken)) return false;
            }
            return true;
        });
        items.sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));

        const limitRaw = limitSel.value;
        const isAll = limitRaw === 'all';
        const pageSize = isAll ? Math.max(1, items.length || 1) : Math.max(1, toNum(limitRaw) || 50);
        const totalPages = isAll ? 1 : Math.max(1, Math.ceil(items.length / pageSize));
        state.progressCurrentPageTotal = totalPages;
        state.progressCurrentPage = isAll ? 1 : Math.min(Math.max(1, state.progressCurrentPage || 1), totalPages);
        const startIndex = isAll ? 0 : (state.progressCurrentPage - 1) * pageSize;
        const displayed = items.slice(startIndex, startIndex + pageSize);

        const statusClass = (status) => {
            if (status === '完了') return 'positive';
            if (status === '中断') return 'negative';
            return '';
        };
        tbody.innerHTML = displayed.map(item => {
            const expectedText = item.expectedImpact ? fmtYen(item.expectedImpact) : '-';
            const actualText = item.actualImpact ? fmtYen(item.actualImpact) : '-';
            const diff = item.actualImpact - item.expectedImpact;
            const diffText = (item.actualImpact || item.expectedImpact) ? signed(diff, fmtYen) : '-';
            return `<tr>
                <td>${escHtml(item.updatedAt.slice(0, 10))}</td>
                <td>${escHtml(item.salesRep || '-')}</td>
                <td>${escHtml(item.store || '-')}</td>
                <td>${escHtml(item.supplierCode || '-')}</td>
                <td class="${statusClass(item.status)}">${escHtml(item.status)}</td>
                <td>${escHtml(item.nextDate || '-')}</td>
                <td>${escHtml(item.plan || '-')}</td>
                <td>${expectedText}</td>
                <td>${actualText}</td>
                <td class="${diff >= 0 ? 'positive' : 'negative'}">${diffText}</td>
                <td>${escHtml(item.resultMemo || '-')}</td>
                <td><button type="button" class="btn-secondary progress-edit-btn" data-id="${escHtml(item.id)}">編集</button><button type="button" class="btn-secondary progress-delete-btn" data-id="${escHtml(item.id)}">削除</button></td>
            </tr>`;
        }).join('');

        if (summaryEl) {
            const done = items.filter(x => x.status === '完了').length;
            const inProgress = items.filter(x => x.status === '進行中').length;
            const pending = items.filter(x => x.status === '未着手').length;
            summaryEl.textContent = `表示: ${fmt(displayed.length)}件 / 全${fmt(items.length)}件（完了:${fmt(done)} 進行中:${fmt(inProgress)} 未着手:${fmt(pending)}）`;
        }
        if (pagerEl && pageStatusEl && prevBtn && nextBtn) {
            pagerEl.style.display = totalPages > 1 ? 'flex' : 'none';
            pageStatusEl.textContent = `${fmt(state.progressCurrentPage)} / ${fmt(totalPages)}ページ`;
            prevBtn.disabled = state.progressCurrentPage <= 1;
            nextBtn.disabled = state.progressCurrentPage >= totalPages;
        }

        if (state.progressEditingId) {
            const editing = state.progressItems.find(x => x.id === state.progressEditingId);
            if (!editing) setProgressForm(null);
        }
    }

    function saveProgressFromForm() {
        const salesRep = toStr(document.getElementById(pfx('progress-form-rep')).value);
        const store = toStr(document.getElementById(pfx('progress-form-store')).value);
        const supplierCode = toStr(document.getElementById(pfx('progress-form-supplier-code')).value);
        const status = toStr(document.getElementById(pfx('progress-form-status')).value) || '未着手';
        const nextDate = toStr(document.getElementById(pfx('progress-form-next-date')).value);
        const plan = toStr(document.getElementById(pfx('progress-form-plan')).value);
        const expectedImpact = toNum(document.getElementById(pfx('progress-form-expected')).value);
        const actualImpact = toNum(document.getElementById(pfx('progress-form-actual')).value);
        const resultMemo = toStr(document.getElementById(pfx('progress-form-result')).value);

        if (!salesRep) { alert('営業担当を選択してください。'); return; }
        if (!store) { alert('販売店を入力してください。'); return; }
        if (!plan) { alert('アクションプランを入力してください。'); return; }

        const now = new Date().toISOString();
        if (state.progressEditingId) {
            const idx = state.progressItems.findIndex(x => x.id === state.progressEditingId);
            if (idx >= 0) {
                const prev = state.progressItems[idx];
                state.progressItems[idx] = {
                    ...prev,
                    salesRep, store, supplierCode, status, nextDate, plan,
                    expectedImpact, actualImpact, resultMemo,
                    updatedAt: now
                };
            }
        } else {
            state.progressItems.push({
                id: 'p-' + Date.now() + '-' + Math.random().toString(36).slice(2, 8),
                salesRep, store, supplierCode, status, nextDate, plan,
                expectedImpact, actualImpact, resultMemo,
                createdAt: now,
                updatedAt: now
            });
        }
        saveProgressItems();
        scheduleAutoStateSave();
        setProgressForm(null);
        state.progressCurrentPage = 1;
        renderProgress();
    }

    function deleteProgressItem(id) {
        const idx = state.progressItems.findIndex(x => x.id === id);
        if (idx < 0) return;
        state.progressItems.splice(idx, 1);
        if (state.progressEditingId === id) setProgressForm(null);
        saveProgressItems();
        scheduleAutoStateSave();
        renderProgress(true);
    }

    function buildProductExecutiveSummary(entry) {
        const avgPrice = entry.qty !== 0 ? entry.sales / entry.qty : 0;
        const listRate = entry.listPrice > 0 ? avgPrice / entry.listPrice : 0;
        const grossRate = entry.sales > 0 ? entry.gross / entry.sales : 0;
        const shippingRatio = entry.sales > 0 ? entry.shipping / entry.sales : 0;
        const avgShippingPerUnit = entry.qty > 0 ? entry.shipping / entry.qty : entry.shippingCost;
        const unitBalance = avgPrice - entry.effectiveCost - avgShippingPerUnit;
        const reasons = [];

        if (unitBalance < 0) {
            reasons.push(`逆ざや（1個あたり${fmtYen(unitBalance)}）`);
        }
        if (listRate > 0 && listRate < 0.7) {
            reasons.push(`定価比${fmtPct(listRate)}まで値引き`);
        }
        if (shippingRatio >= 0.15) {
            reasons.push(`送料比率${fmtPct(shippingRatio)}が高い`);
        }
        if (entry.qty >= 50 && entry.gross < 0) {
            reasons.push(`販売数量${fmt(entry.qty)}個で損失が累積`);
        }
        const costToPrice = avgPrice > 0 ? entry.effectiveCost / avgPrice : 0;
        if (costToPrice >= 0.9) {
            reasons.push(`原価率${fmtPct(costToPrice)}で高止まり`);
        }

        if (entry.gross < 0) {
            const detail = reasons.length > 0 ? reasons.slice(0, 3).join(' / ') : '主因の特定に十分な比較データなし';
            return `赤字${fmtYen(entry.gross)}。${detail}。`;
        }
        if (reasons.length > 0) {
            return `黒字${fmtYen(entry.gross)}だが注意: ${reasons[0]}。`;
        }
        return `黒字${fmtYen(entry.gross)}。粗利率${fmtPct(grossRate)}で安定。`;
    }

    // ── Render: Details ──
    function renderDetails() {
        const r = state.results; if (!r) return;
        document.getElementById(pfx('details-empty')).style.display = 'none';
        document.getElementById(pfx('details-content')).style.display = 'block';

        const makerF = document.getElementById(pfx('details-maker')).value;
        const sortKey = document.getElementById(pfx('details-sort')).value;
        const limitRaw = document.getElementById(pfx('details-limit')).value;
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

        const isAll = limitRaw === 'all';
        const pageSize = isAll ? Math.max(1, entries.length || 1) : Math.max(1, toNum(limitRaw) || 300);
        const totalPages = isAll ? 1 : Math.max(1, Math.ceil(entries.length / pageSize));
        state.detailsCurrentPageTotal = totalPages;
        state.detailsCurrentPage = isAll ? 1 : Math.min(Math.max(1, state.detailsCurrentPage || 1), totalPages);
        const startIndex = isAll ? 0 : (state.detailsCurrentPage - 1) * pageSize;
        const displayed = entries.slice(startIndex, startIndex + pageSize);

        const aiSummaryEl = document.getElementById(pfx('details-ai-summary'));
        if (aiSummaryEl) {
            const worst = entries.filter(e => e.gross < 0).sort((a, b) => a.gross - b.gross).slice(0, 3);
            if (worst.length === 0) {
                aiSummaryEl.textContent = 'AIエグゼクティブサマリー: 現在、赤字商品はありません。';
            } else {
                const lines = worst.map(e => {
                    const summary = buildProductExecutiveSummary(e);
                    return `・${e.jan || '-'} ${e.name || ''}: ${summary}`;
                });
                aiSummaryEl.textContent = `AIエグゼクティブサマリー(赤字上位): ${lines.join(' ')}`;
            }
        }

        const tbody = document.getElementById(pfx('details-tbody'));
        tbody.innerHTML = displayed.map(e => {
            const avgP = e.qty !== 0 ? e.sales / e.qty : 0;
            const rt = e.listPrice > 0 ? avgP / e.listPrice : 0;
            const pr = e.sales > 0 ? e.gross / e.sales : 0;
            const summary = buildProductExecutiveSummary(e);
            return `<tr><td>${escHtml(e.jan)}</td><td>${escHtml(e.name)}</td><td>${escHtml(ml[e.maker] || e.maker)}</td><td>${fmtYen(e.listPrice)}</td><td>${fmtYen(e.effectiveCost)}</td><td>${fmtYen(avgP)}</td><td>${fmtPct(rt)}</td><td>${fmtYen(e.shippingCost)}</td><td>${fmt(e.qty)}</td><td>${fmtYen(e.sales)}</td><td class="${e.gross >= 0 ? 'positive' : 'negative'}">${fmtYen(e.gross)}</td><td>${fmtPct(pr)}</td><td class="details-exec">${escHtml(summary)}</td></tr>`;
        }).join('');

        const summaryEl = document.getElementById(pfx('details-summary'));
        if (summaryEl) {
            if (entries.length === 0) {
                summaryEl.textContent = '表示: 0件 / 全0件';
            } else {
                const from = startIndex + 1;
                const to = startIndex + displayed.length;
                summaryEl.textContent = `表示: ${fmt(from)}-${fmt(to)}件 / 全${fmt(entries.length)}件`;
            }
        }

        const pagerEl = document.getElementById(pfx('details-pagination'));
        const pageStatusEl = document.getElementById(pfx('details-page-status'));
        const prevBtn = document.getElementById(pfx('details-page-prev'));
        const nextBtn = document.getElementById(pfx('details-page-next'));
        if (pagerEl && pageStatusEl && prevBtn && nextBtn) {
            pagerEl.style.display = totalPages > 1 ? 'flex' : 'none';
            pageStatusEl.textContent = `${fmt(state.detailsCurrentPage)} / ${fmt(totalPages)}ページ`;
            prevBtn.disabled = state.detailsCurrentPage <= 1;
            nextBtn.disabled = state.detailsCurrentPage >= totalPages;
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

        if (tabId === 'progress') {
            renderProgress();
            return;
        }

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
                case 'store-detail':
                    setTimeout(() => {
                        if (currentTab === 'store-detail') renderStoreDetail();
                    }, 0);
                    break;
                case 'progress':
                    renderProgress();
                    break;
                case 'details': renderDetails(); break;
            }
        }
    }

    function getAuthClient() {
        if (!window.KaientaiAuth || typeof window.KaientaiAuth.getStatus !== 'function') return null;
        return window.KaientaiAuth;
    }

    function resetSettingsToDefaults() {
        setInputValue('rebate-aron', 0);
        setInputValue('rebate-pana', 0);
        setInputValue('warehouse-fee', 0);
        setInputValue('warehouse-out-fee', 50);
        setInputValue('default-shipping-small', 100);
        setInputValue('keyword-aron', 'アロン');
        setInputValue('keyword-pana', 'パナソニック,パナ,Panasonic');
        renderMonthlyRebateInputs();
    }

    function clearAuthorizedRuntimeState() {
        clearCloudRestoreRetry();
        if (autoPersistTimer) {
            clearTimeout(autoPersistTimer);
            autoPersistTimer = null;
        }
        if (cloudPersistTimer) {
            clearTimeout(cloudPersistTimer);
            cloudPersistTimer = null;
        }
        cloudPersistPending = false;
        cloudPersistRetryMs = 1000;

        state.shippingData = [];
        state.salesData = [];
        state.productData = [];
        state.progressItems = [];
        logLines = [];
        currentTab = 'upload';

        resetAnalysisOutputsForDataChange();
        resetSettingsToDefaults();

        ['file-shipping', 'file-sales', 'file-product'].forEach(id => {
            const input = document.getElementById(pfx(id));
            if (input) input.value = '';
        });

        const logEl = document.getElementById(pfx('load-log'));
        if (logEl) logEl.style.display = 'none';

        updateUploadCardsByState();
        setProgressForm(null);
        renderProgress();
        switchModTab('upload');
        KaientaiM.updateModuleStatus(MODULE_ID, '認証待ち', false);
    }

    function hydrateAuthorizedSession() {
        if (authHydrated) return;
        authHydrated = true;

        loadProgressItems();
        const restoredLocal = restoreAutoState();
        if (!restoredLocal) startCloudRestoreRetry();
        setProgressForm(null);
        renderProgress();
    }

    function syncGoogleAuthState(authStatus) {
        const authorized = !!authStatus?.authorized;
        const wasAuthorized = state.authVerified;

        state.authVerified = authorized;
        state.uploadUnlocked = authorized;
        state.settingsUnlocked = authorized;
        applyAuthLocks();

        if (authorized) {
            hydrateAuthorizedSession();
            KaientaiM.updateModuleStatus(
                MODULE_ID,
                state.salesData.length > 0 ? `データ読込済み (${state.salesData.length})` : '未設定',
                state.salesData.length > 0
            );
            return;
        }

        authHydrated = false;
        if (wasAuthorized) {
            clearAuthorizedRuntimeState();
        } else {
            KaientaiM.updateModuleStatus(MODULE_ID, '認証待ち', false);
        }
    }

    function bindGoogleAuth() {
        const auth = getAuthClient();
        if (!auth || typeof auth.onAuthStateChanged !== 'function') {
            KaientaiM.updateModuleStatus(MODULE_ID, '認証エラー', false);
            return;
        }
        auth.onAuthStateChanged(syncGoogleAuthState);
    }

    function requestEditUnlock(area) {
        if (state.authVerified) {
            state.uploadUnlocked = true;
            state.settingsUnlocked = true;
            applyAuthLocks();
            return true;
        }

        const auth = getAuthClient();
        if (!auth || typeof auth.signInWithGoogle !== 'function') {
            alert('Google認証の初期化に失敗しました。');
            return false;
        }

        auth.signInWithGoogle().catch(err => {
            const code = String(err?.code || '');
            if (code === 'auth/popup-closed-by-user' || code === 'auth/cancelled-popup-request') return;
            if (err?.message) alert(err.message);
        });
        return false;
    }

    function applyAuthLocks() {
        const uploadLocked = !state.uploadUnlocked;
        const settingsLocked = !state.settingsUnlocked;

        const uploadStatus = document.getElementById(pfx('upload-lock-status'));
        if (uploadStatus) {
            uploadStatus.textContent = uploadLocked
                ? 'ロック中: 読込/分析には承認済みGoogleログインが必要です'
                : '認証済み: データ読込を実行できます';
            uploadStatus.className = 'auth-lock-status ' + (uploadLocked ? 'locked' : 'unlocked');
        }
        const settingsStatus = document.getElementById(pfx('settings-lock-status'));
        if (settingsStatus) {
            settingsStatus.textContent = settingsLocked
                ? 'ロック中: 設定編集には承認済みGoogleログインが必要です'
                : '認証済み: 設定を編集できます';
            settingsStatus.className = 'auth-lock-status ' + (settingsLocked ? 'locked' : 'unlocked');
        }

        const uploadUnlockBtn = document.getElementById(pfx('btn-upload-unlock'));
        if (uploadUnlockBtn) {
            uploadUnlockBtn.textContent = uploadLocked ? 'Googleでログイン' : '認証済み';
            uploadUnlockBtn.disabled = !uploadLocked;
        }
        const settingsUnlockBtn = document.getElementById(pfx('btn-settings-unlock'));
        if (settingsUnlockBtn) {
            settingsUnlockBtn.textContent = settingsLocked ? 'Googleでログイン' : '認証済み';
            settingsUnlockBtn.disabled = !settingsLocked;
        }

        const uploadTargetIds = ['file-shipping', 'file-sales', 'file-product', 'btn-shipping-clear', 'btn-sales-clear', 'btn-product-clear'];
        for (const id of uploadTargetIds) {
            const el = document.getElementById(pfx(id));
            if (el) el.disabled = uploadLocked;
        }
        const uploadTab = document.getElementById(pfx('tab-upload'));
        if (uploadTab) {
            uploadTab.querySelectorAll('.upload-btn').forEach(el => {
                el.classList.toggle('locked', uploadLocked);
            });
        }

        const settingsTab = document.getElementById(pfx('tab-settings'));
        if (settingsTab) {
            settingsTab.querySelectorAll('input, select, textarea').forEach(el => {
                el.disabled = settingsLocked;
            });
        }

        checkAllLoaded();
    }

    function checkAllLoaded() {
        const ok = state.shippingData.length > 0 && state.salesData.length > 0 && state.productData.length > 0;
        const analyzeBtn = document.getElementById(pfx('btn-analyze'));
        if (analyzeBtn) {
            analyzeBtn.dataset.forceDisabled = ok ? '0' : '1';
            analyzeBtn.disabled = !ok || !state.uploadUnlocked;
        }
        const clearShippingBtn = document.getElementById(pfx('btn-shipping-clear'));
        if (clearShippingBtn) clearShippingBtn.disabled = !state.uploadUnlocked || state.shippingData.length === 0;
        const clearSalesBtn = document.getElementById(pfx('btn-sales-clear'));
        if (clearSalesBtn) clearSalesBtn.disabled = !state.uploadUnlocked || state.salesData.length === 0;
        const clearProductBtn = document.getElementById(pfx('btn-product-clear'));
        if (clearProductBtn) clearProductBtn.disabled = !state.uploadUnlocked || state.productData.length === 0;
    }

    function resetAnalysisOutputsForDataChange() {
        state.results = null;
        state.storeBaseCache = {};
        state.storeSortedCache = {};
        state.storeViewRuntime = null;
        state.storeAdvancedEnabled = false;
        state.storeCurrentPage = 1;
        state.storeCurrentPageTotal = 1;
        state.detailsCurrentPage = 1;
        state.detailsCurrentPageTotal = 1;
        state.storeDetailIndex = null;
        state.storeDetailCurrentPage = 1;
        state.storeDetailCurrentPageTotal = 1;
        state.progressCurrentPage = 1;
        state.progressCurrentPageTotal = 1;
        state.progressEditingId = '';
        state.storeHeavyRenderToken = 0;
        if (state.storeHeavyRenderTimer) {
            clearTimeout(state.storeHeavyRenderTimer);
            state.storeHeavyRenderTimer = null;
        }
        ['overview', 'monthly', 'store', 'sim', 'store-detail', 'progress', 'details'].forEach(id => {
            const emp = document.getElementById(pfx(id + '-empty'));
            const con = document.getElementById(pfx(id + '-content'));
            if (emp) emp.style.display = '';
            if (con) con.style.display = 'none';
        });
        Object.keys(state.charts).forEach(k => destroyChart(state.charts, k));
        setStoreAdvancedMode(false);
        KaientaiM.updateModuleStatus(MODULE_ID, '未設定', false);
    }

    function setInputValue(id, value) {
        const el = document.getElementById(pfx(id));
        if (!el) return;
        el.value = value;
    }

    function restoreSavedSettings(saved) {
        if (!saved || typeof saved !== 'object') {
            renderMonthlyRebateInputs();
            applyAuthLocks();
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
        applyAuthLocks();
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
            if (statusEl) statusEl.textContent = item.len > 0 ? `✓ ${item.len}件` : '未読込';
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
            <button class="mod-nav-btn" data-mtab="overview">全体概要</button>
            <button class="mod-nav-btn" data-mtab="monthly">月次分析</button>
            <button class="mod-nav-btn" data-mtab="store">販売店分析</button>
            <button class="mod-nav-btn" data-mtab="simulation">掛け率シミュレーション</button>
            <button class="mod-nav-btn" data-mtab="store-detail">販売店詳細分析</button>
            <button class="mod-nav-btn" data-mtab="progress">進捗管理</button>
            <button class="mod-nav-btn" data-mtab="details">商品別詳細</button>
            <button class="mod-nav-btn active" data-mtab="upload">データ読込</button>
            <button class="mod-nav-btn" data-mtab="settings">設定</button>
        </div>

        <!-- Upload -->
        <div class="mod-tab active" id="${pfx('tab-upload')}">
            <div class="auth-lock-bar">
                <span class="auth-lock-status locked" id="${pfx('upload-lock-status')}">ロック中: 読込/分析には承認済みGoogleログインが必要です</span>
                <button type="button" class="btn-secondary" id="${pfx('btn-upload-unlock')}">Googleでログイン</button>
            </div>
            <div class="upload-grid">
                <div class="upload-card" id="${pfx('card-shipping')}">
                    <div class="upload-icon">&#128666;</div>
                    <h3>送料マスターデータ</h3>
                    <p>A列:JAN / B列:商品名 / I列:サイズ帯 / J〜V列:エリア別送料（W列は未使用）</p>
                    <label class="upload-btn">ファイル選択<input type="file" accept=".xlsx,.xls,.csv" id="${pfx('file-shipping')}" hidden></label>
                    <div class="upload-status" id="${pfx('status-shipping')}">未読込</div>
                    <div class="action-bar"><button class="btn-secondary" id="${pfx('btn-shipping-clear')}">送料データをクリア</button></div>
                </div>
                <div class="upload-card" id="${pfx('card-sales')}">
                    <div class="upload-icon">&#128200;</div>
                    <h3>販売実績データ</h3>
                    <p>A列:受注番号 / B列:受注日 / D列:販売店 / H列:JAN / I列:商品名 / K列:数量 / L列:単価 / M列:合計 / S列:メーカー / Z列:営業担当 / AB列:県名</p>
                    <label class="upload-btn">ファイル選択<input type="file" accept=".xlsx,.xls,.csv" id="${pfx('file-sales')}" hidden multiple></label>
                    <div class="upload-status" id="${pfx('status-sales')}">未読込</div>
                    <div class="upload-hint">※複数月のファイルを同時選択可能</div>
                    <div class="action-bar"><button class="btn-secondary" id="${pfx('btn-sales-clear')}">販売実績データをクリア</button></div>
                </div>
                <div class="upload-card" id="${pfx('card-product')}">
                    <div class="upload-icon">&#128230;</div>
                    <h3>商品マスタ</h3>
                    <p>A列:JAN / D列:商品名 / H列:定価 / M列:原価 / O列:倉庫入原価</p>
                    <label class="upload-btn">ファイル選択<input type="file" accept=".xlsx,.xls,.csv" id="${pfx('file-product')}" hidden multiple></label>
                    <div class="upload-status" id="${pfx('status-product')}">未読込</div>
                    <div class="action-bar"><button class="btn-secondary" id="${pfx('btn-product-clear')}">商品マスタをクリア</button></div>
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
            <div class="auth-lock-bar">
                <span class="auth-lock-status locked" id="${pfx('settings-lock-status')}">ロック中: 設定編集には承認済みGoogleログインが必要です</span>
                <button type="button" class="btn-secondary" id="${pfx('btn-settings-unlock')}">Googleでログイン</button>
            </div>
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
                <div class="store-meta-row"><div class="hint" id="${pfx('overview-period')}">データ対象期間: -</div></div>
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
                <div class="filter-bar">
                    <label>メーカー:</label><select id="${pfx('store-maker')}"><option value="all">全て</option><option value="aron">アロン化成</option><option value="pana">パナソニック</option></select>
                    <label>年月(複数選択):</label><select id="${pfx('store-month')}" multiple size="5"></select>
                    <label>営業担当(複数選択):</label><select id="${pfx('store-rep')}" multiple size="5"></select>
                    <label>並び替え:</label><select id="${pfx('store-sort')}"><option value="gross-desc">粗利(高い順)</option><option value="gross-asc">粗利(低い順)</option><option value="sales-desc">売上(高い順)</option><option value="sales-asc">売上(低い順)</option><option value="qty-desc">数量(多い順)</option><option value="qty-asc">数量(少ない順)</option><option value="rate-desc">粗利率(高い順)</option><option value="rate-asc">粗利率(低い順)</option><option value="aron-rate-desc">アロン掛率(高い順)</option><option value="aron-rate-asc">アロン掛率(低い順)</option><option value="pana-rate-desc">パナ掛率(高い順)</option><option value="pana-rate-asc">パナ掛率(低い順)</option><option value="rep-asc">担当者(昇順)</option><option value="rep-desc">担当者(降順)</option><option value="store-asc">販売店名(昇順)</option><option value="store-desc">販売店名(降順)</option></select>
                    <label>表示件数:</label><select id="${pfx('store-limit')}"><option value="300">300</option><option value="1000">1000</option><option value="all">全件</option></select>
                </div>
                <div class="hint">※年月・営業担当はクリックで複数選択できます。未選択時は全件対象です。</div>
                <div class="store-meta-row">
                    <div class="hint" id="${pfx('store-summary')}"></div>
                    <div class="store-pagination" id="${pfx('store-pagination')}" style="display:none;">
                        <button type="button" class="btn-secondary" id="${pfx('store-page-prev')}">前へ</button>
                        <span class="store-page-status" id="${pfx('store-page-status')}">1 / 1ページ</span>
                        <button type="button" class="btn-secondary" id="${pfx('store-page-next')}">次へ</button>
                    </div>
                </div>
                <div class="table-wrapper"><table><thead><tr><th>販売店名</th><th>営業担当者</th><th>アロン掛率</th><th>パナ掛率</th><th>売上合計</th><th>原価合計</th><th>送料合計</th><th>商品粗利</th><th>粗利率</th><th>数量合計</th></tr></thead><tbody id="${pfx('store-tbody')}"></tbody></table></div>
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
                <div class="hint">※掛け率は「卸売価格 ÷ 定価」で計算（送料・リベートは含めません）。リベートは設定タブの値を自動適用します。</div>
                <div class="sim-result-grid">
                    <div class="sim-result-card"><div class="sim-label">変動前 実利益</div><div class="sim-value" id="${pfx('sim-before')}">-</div></div>
                    <div class="sim-result-card arrow">&#8594;</div>
                    <div class="sim-result-card"><div class="sim-label">変動後 実利益</div><div class="sim-value" id="${pfx('sim-after')}">-</div></div>
                    <div class="sim-result-card"><div class="sim-label">差額</div><div class="sim-value" id="${pfx('sim-diff')}">-</div></div>
                </div>
                <div class="store-meta-row"><div class="hint" id="${pfx('sim-rebate-summary')}"></div></div>
                <div class="table-wrapper"><table><thead><tr><th>区分</th><th>総リベート</th><th>アロン</th><th>パナ</th><th>その他</th></tr></thead><tbody id="${pfx('sim-rebate-tbody')}"></tbody></table></div>
                <div class="chart-row"><div class="chart-box full"><h3>掛け率 vs 実利益 推移</h3><canvas id="${pfx('chart-sim')}"></canvas></div></div>
            </div>
        </div>

        <!-- Store Detail -->
        <div class="mod-tab" id="${pfx('tab-store-detail')}">
            <div id="${pfx('store-detail-empty')}" class="empty-state">データを読み込んで分析を実行してください</div>
            <div id="${pfx('store-detail-content')}" style="display:none;">
                <div class="filter-bar">
                    <label>販売店検索:</label><input type="text" id="${pfx('store-detail-store-search')}" placeholder="販売店名 / 得意先コード">
                    <label>販売店:</label><select id="${pfx('store-detail-store')}"></select>
                    <label>集計メーカー:</label><select id="${pfx('store-detail-maker')}"><option value="all">全て</option><option value="aron">アロン化成</option><option value="pana">パナソニック</option></select>
                    <label>年月:</label><select id="${pfx('store-detail-month')}"><option value="all">全期間</option></select>
                    <label>営業担当:</label><select id="${pfx('store-detail-rep')}"><option value="all">全担当</option></select>
                </div>
                <div class="filter-bar">
                    <label>シミュレーション対象:</label><select id="${pfx('store-detail-sim-maker')}"><option value="all">両メーカー</option><option value="aron">アロン化成のみ</option><option value="pana">パナソニックのみ</option></select>
                    <label>掛率変動(%):</label><input type="number" id="${pfx('store-detail-rate')}" value="0" step="0.1">
                    <label>予想販売増加数(個):</label><input type="number" id="${pfx('store-detail-qty')}" value="0" step="1" min="0">
                </div>
                <div class="advanced-sim-box">
                    <h3>高度シミュレーター</h3>
                    <div class="filter-bar advanced-sim-controls">
                        <label><input type="checkbox" id="${pfx('store-detail-use-trend')}">トレンド反映</label>
                        <label>トレンド補正(%/月):</label><input type="number" id="${pfx('store-detail-trend-adjust')}" value="0" step="0.1">
                        <label><input type="checkbox" id="${pfx('store-detail-use-elastic-auto')}">弾力性自動推定</label>
                        <label>価格弾力性(手動):</label><input type="number" id="${pfx('store-detail-elasticity')}" value="0" step="0.1">
                        <label><input type="checkbox" id="${pfx('store-detail-use-seasonality')}">季節性反映</label>
                    </div>
                    <div class="advanced-sim-formula">
                        <p><strong>計算式</strong> 数量' = 数量 × (1 + トレンド率 + 補正率) × 価格係数 × 季節係数 + 追加数量配賦</p>
                        <p>価格係数 = (1 + 掛率変動率) ^ 価格弾力性</p>
                        <p>トレンド率: 月次数量の直近6期間の成長率平均（外れ値トリム）</p>
                        <p>価格弾力性(自動): SKU別に ln(数量) と ln(価格) の回帰傾きの加重平均</p>
                        <p>季節係数: 月別平均数量 / 全月平均数量（0.6〜1.4に制限）</p>
                    </div>
                </div>
                <div class="store-meta-row"><div class="hint" id="${pfx('store-detail-summary')}"></div></div>
                <div class="store-meta-row"><div class="hint" id="${pfx('store-detail-model-summary')}"></div></div>
                <div class="table-wrapper"><table class="store-sim-table"><thead><tr><th>区分</th><th>売上</th><th>実利益</th><th>実利益率</th><th>数量</th><th>アロン掛率</th><th>パナ掛率</th></tr></thead><tbody id="${pfx('store-detail-sim-tbody')}"><tr><td colspan="7">販売店を選択してください</td></tr></tbody></table></div>
                <div class="filter-bar">
                    <label>商品検索:</label><input type="text" id="${pfx('store-detail-product-search')}" placeholder="JANコードまたは商品名">
                    <label>表示件数:</label><select id="${pfx('store-detail-limit')}"><option value="200">200</option><option value="500">500</option><option value="all">全件</option></select>
                </div>
                <div class="store-meta-row">
                    <div class="hint" id="${pfx('store-detail-product-summary')}"></div>
                    <div class="store-pagination" id="${pfx('store-detail-pagination')}" style="display:none;">
                        <button type="button" class="btn-secondary" id="${pfx('store-detail-page-prev')}">前へ</button>
                        <span class="store-page-status" id="${pfx('store-detail-page-status')}">1 / 1ページ</span>
                        <button type="button" class="btn-secondary" id="${pfx('store-detail-page-next')}">次へ</button>
                    </div>
                </div>
                <div class="table-wrapper"><table><thead><tr><th>JANコード</th><th>商品名</th><th>メーカー</th><th>数量合計</th><th>売上合計</th><th>原価合計</th><th>送料合計</th><th>粗利合計</th><th>粗利率</th></tr></thead><tbody id="${pfx('store-detail-products-tbody')}"></tbody></table></div>
            </div>
        </div>

        <!-- Progress -->
        <div class="mod-tab" id="${pfx('tab-progress')}">
            <div id="${pfx('progress-empty')}" class="empty-state">進捗管理データを作成してください</div>
            <div id="${pfx('progress-content')}" style="display:none;">
                <div class="filter-bar">
                    <label>営業担当:</label><select id="${pfx('progress-rep-filter')}"><option value="all">全担当</option></select>
                    <label>ステータス:</label><select id="${pfx('progress-status-filter')}"><option value="all">全て</option><option value="未着手">未着手</option><option value="進行中">進行中</option><option value="完了">完了</option><option value="中断">中断</option></select>
                    <label>検索:</label><input type="text" id="${pfx('progress-store-search')}" placeholder="販売店名 / 得意先コード / メモ">
                    <label>表示件数:</label><select id="${pfx('progress-limit')}"><option value="50">50</option><option value="200">200</option><option value="all">全件</option></select>
                </div>
                <div class="store-meta-row">
                    <div class="hint" id="${pfx('progress-summary')}"></div>
                    <div class="store-pagination" id="${pfx('progress-pagination')}" style="display:none;">
                        <button type="button" class="btn-secondary" id="${pfx('progress-page-prev')}">前へ</button>
                        <span class="store-page-status" id="${pfx('progress-page-status')}">1 / 1ページ</span>
                        <button type="button" class="btn-secondary" id="${pfx('progress-page-next')}">次へ</button>
                    </div>
                </div>
                <div class="table-wrapper"><table><thead><tr><th>更新日</th><th>営業担当</th><th>販売店</th><th>得意先コード</th><th>状態</th><th>次アクション日</th><th>アクションプラン</th><th>想定効果</th><th>実績効果</th><th>差分</th><th>結果メモ</th><th>操作</th></tr></thead><tbody id="${pfx('progress-tbody')}"></tbody></table></div>
                <div class="setting-card">
                    <h3>アクション登録 / 更新</h3>
                    <div class="filter-bar">
                        <label>営業担当:</label><select id="${pfx('progress-form-rep')}"><option value="">担当を選択</option></select>
                        <label>販売店:</label><input type="text" id="${pfx('progress-form-store')}" list="${pfx('progress-store-list')}" placeholder="販売店名">
                        <datalist id="${pfx('progress-store-list')}"></datalist>
                        <label>得意先コード:</label><input type="text" id="${pfx('progress-form-supplier-code')}" placeholder="例: 12345">
                        <label>状態:</label><select id="${pfx('progress-form-status')}"><option value="未着手">未着手</option><option value="進行中">進行中</option><option value="完了">完了</option><option value="中断">中断</option></select>
                        <label>次アクション日:</label><input type="date" id="${pfx('progress-form-next-date')}">
                    </div>
                    <div class="filter-bar">
                        <label>アクションプラン:</label><input type="text" id="${pfx('progress-form-plan')}" placeholder="例: 掛率交渉とセット提案">
                        <label>想定効果(円):</label><input type="number" id="${pfx('progress-form-expected')}" value="0" step="1000">
                        <label>実績効果(円):</label><input type="number" id="${pfx('progress-form-actual')}" value="0" step="1000">
                    </div>
                    <div class="filter-bar">
                        <label>結果メモ:</label><input type="text" id="${pfx('progress-form-result')}" placeholder="実施結果・課題・次回打ち手">
                    </div>
                    <div class="action-bar">
                        <button class="btn-primary" id="${pfx('progress-btn-save')}">登録</button>
                        <button class="btn-secondary" id="${pfx('progress-btn-clear')}">入力クリア</button>
                    </div>
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
                    <label>表示件数:</label><select id="${pfx('details-limit')}"><option value="300">300</option><option value="1000">1000</option><option value="all">全件</option></select>
                    <label>検索:</label><input type="text" id="${pfx('details-search')}" placeholder="JANコードまたは商品名">
                </div>
                <div class="store-meta-row">
                    <div class="hint" id="${pfx('details-summary')}"></div>
                    <div class="store-pagination" id="${pfx('details-pagination')}" style="display:none;">
                        <button type="button" class="btn-secondary" id="${pfx('details-page-prev')}">前へ</button>
                        <span class="store-page-status" id="${pfx('details-page-status')}">1 / 1ページ</span>
                        <button type="button" class="btn-secondary" id="${pfx('details-page-next')}">次へ</button>
                    </div>
                </div>
                <div class="store-meta-row"><div class="hint" id="${pfx('details-ai-summary')}"></div></div>
                <div class="table-wrapper"><table class="details-table"><thead><tr><th>JANコード</th><th>商品名</th><th>メーカー</th><th>定価</th><th>原価</th><th>販売単価(平均)</th><th>掛け率</th><th>送料</th><th>数量合計</th><th>売上合計</th><th>粗利合計</th><th>粗利率</th><th>AIエグゼクティブサマリー</th></tr></thead><tbody id="${pfx('details-tbody')}"></tbody></table></div>
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
        document.getElementById(pfx('btn-upload-unlock')).addEventListener('click', () => requestEditUnlock('upload'));
        document.getElementById(pfx('btn-settings-unlock')).addEventListener('click', () => requestEditUnlock('settings'));

        // File uploads
        document.getElementById(pfx('file-shipping')).addEventListener('change', async (e) => {
            if (!e.target.files.length) return;
            if (!state.uploadUnlocked && !requestEditUnlock('upload')) { e.target.value = ''; return; }
            try { loadShipping(await parseExcel(e.target.files[0])); checkAllLoaded(); }
            catch (err) { log('送料マスタ読込エラー: ' + err.message); alert('送料マスタの読込に失敗しました'); }
        });
        document.getElementById(pfx('file-sales')).addEventListener('change', async (e) => {
            if (!e.target.files.length) return;
            if (!state.uploadUnlocked && !requestEditUnlock('upload')) { e.target.value = ''; return; }
            try {
                const list = [];
                for (const f of Array.from(e.target.files)) list.push(await parseExcel(f));
                loadSales(list); checkAllLoaded();
            } catch (err) { log('販売実績読込エラー: ' + err.message); alert('販売実績の読込に失敗しました'); }
        });
        document.getElementById(pfx('file-product')).addEventListener('change', async (e) => {
            if (!e.target.files.length) return;
            if (!state.uploadUnlocked && !requestEditUnlock('upload')) { e.target.value = ''; return; }
            try {
                const list = [];
                for (const f of Array.from(e.target.files)) list.push(await parseExcel(f));
                loadProduct(list); checkAllLoaded();
            }
            catch (err) { log('商品マスタ読込エラー: ' + err.message); alert('商品マスタの読込に失敗しました'); }
        });

        document.getElementById(pfx('btn-analyze')).addEventListener('click', () => {
            if (!state.uploadUnlocked && !requestEditUnlock('upload')) return;
            analyze();
        });
        const settingsTabEl = document.getElementById(pfx('tab-settings'));
        if (settingsTabEl) {
            const persistOnInput = (e) => {
                const target = e.target;
                if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement || target instanceof HTMLTextAreaElement)) return;
                scheduleAutoStateSave(200);
            };
            const persistOnChange = (e) => {
                const target = e.target;
                if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement || target instanceof HTMLTextAreaElement)) return;
                saveAutoStateNow();
            };
            settingsTabEl.addEventListener('input', persistOnInput);
            settingsTabEl.addEventListener('change', persistOnChange);
        }
        document.getElementById(pfx('btn-recalc')).addEventListener('click', () => {
            if (state.salesData.length === 0) { alert('データを先に読み込んでください。'); return; }
            analyze();
        });
        document.getElementById(pfx('btn-shipping-clear')).addEventListener('click', () => {
            if (!state.uploadUnlocked && !requestEditUnlock('upload')) return;
            if (state.shippingData.length === 0) return;
            state.shippingData = [];
            state.storeDetailIndex = null;
            const input = document.getElementById(pfx('file-shipping'));
            if (input) input.value = '';
            resetAnalysisOutputsForDataChange();
            updateUploadCardsByState();
            checkAllLoaded();
            saveAutoStateNow();
        });
        document.getElementById(pfx('btn-sales-clear')).addEventListener('click', () => {
            if (!state.uploadUnlocked && !requestEditUnlock('upload')) return;
            if (state.salesData.length === 0) return;
            state.salesData = [];
            state.storeDetailIndex = null;
            const input = document.getElementById(pfx('file-sales'));
            if (input) input.value = '';
            renderMonthlyRebateInputs();
            resetAnalysisOutputsForDataChange();
            updateUploadCardsByState();
            checkAllLoaded();
            saveAutoStateNow();
        });
        document.getElementById(pfx('btn-product-clear')).addEventListener('click', () => {
            if (!state.uploadUnlocked && !requestEditUnlock('upload')) return;
            if (state.productData.length === 0) return;
            state.productData = [];
            state.storeDetailIndex = null;
            const input = document.getElementById(pfx('file-product'));
            if (input) input.value = '';
            resetAnalysisOutputsForDataChange();
            updateUploadCardsByState();
            checkAllLoaded();
            saveAutoStateNow();
        });

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
            renderStore(true);
        });
        document.getElementById(pfx('store-page-next')).addEventListener('click', () => {
            if (state.storeCurrentPage >= state.storeCurrentPageTotal) return;
            state.storeCurrentPage += 1;
            renderStore(true);
        });
        const resetStoreDetailPageAndRender = () => { state.storeDetailCurrentPage = 1; renderStoreDetail(); };
        document.getElementById(pfx('store-detail-store-search')).addEventListener('input', resetStoreDetailPageAndRender);
        document.getElementById(pfx('store-detail-store')).addEventListener('change', resetStoreDetailPageAndRender);
        document.getElementById(pfx('store-detail-maker')).addEventListener('change', resetStoreDetailPageAndRender);
        document.getElementById(pfx('store-detail-month')).addEventListener('change', resetStoreDetailPageAndRender);
        document.getElementById(pfx('store-detail-rep')).addEventListener('change', resetStoreDetailPageAndRender);
        document.getElementById(pfx('store-detail-limit')).addEventListener('change', resetStoreDetailPageAndRender);
        document.getElementById(pfx('store-detail-product-search')).addEventListener('input', resetStoreDetailPageAndRender);
        document.getElementById(pfx('store-detail-sim-maker')).addEventListener('change', () => renderStoreDetail(true));
        document.getElementById(pfx('store-detail-rate')).addEventListener('input', () => renderStoreDetail(true));
        document.getElementById(pfx('store-detail-qty')).addEventListener('input', () => renderStoreDetail(true));
        document.getElementById(pfx('store-detail-use-trend')).addEventListener('change', () => renderStoreDetail(true));
        document.getElementById(pfx('store-detail-trend-adjust')).addEventListener('input', () => renderStoreDetail(true));
        document.getElementById(pfx('store-detail-use-elastic-auto')).addEventListener('change', () => renderStoreDetail(true));
        document.getElementById(pfx('store-detail-elasticity')).addEventListener('input', () => renderStoreDetail(true));
        document.getElementById(pfx('store-detail-use-seasonality')).addEventListener('change', () => renderStoreDetail(true));
        document.getElementById(pfx('store-detail-page-prev')).addEventListener('click', () => {
            if (state.storeDetailCurrentPage <= 1) return;
            state.storeDetailCurrentPage -= 1;
            renderStoreDetail(true);
        });
        document.getElementById(pfx('store-detail-page-next')).addEventListener('click', () => {
            if (state.storeDetailCurrentPage >= state.storeDetailCurrentPageTotal) return;
            state.storeDetailCurrentPage += 1;
            renderStoreDetail(true);
        });
        const resetDetailsPageAndRender = () => { state.detailsCurrentPage = 1; renderDetails(); };
        document.getElementById(pfx('details-maker')).addEventListener('change', resetDetailsPageAndRender);
        document.getElementById(pfx('details-sort')).addEventListener('change', resetDetailsPageAndRender);
        document.getElementById(pfx('details-limit')).addEventListener('change', resetDetailsPageAndRender);
        document.getElementById(pfx('details-search')).addEventListener('input', resetDetailsPageAndRender);
        document.getElementById(pfx('details-page-prev')).addEventListener('click', () => {
            if (state.detailsCurrentPage <= 1) return;
            state.detailsCurrentPage -= 1;
            renderDetails();
        });
        document.getElementById(pfx('details-page-next')).addEventListener('click', () => {
            if (state.detailsCurrentPage >= state.detailsCurrentPageTotal) return;
            state.detailsCurrentPage += 1;
            renderDetails();
        });

        const resetProgressPageAndRender = () => { state.progressCurrentPage = 1; renderProgress(); };
        const syncProgressSupplierCode = () => {
            const storeInput = document.getElementById(pfx('progress-form-store'));
            const codeInput = document.getElementById(pfx('progress-form-supplier-code'));
            const store = toStr(storeInput.value);
            if (!store) return;
            const master = getProgressMasterData();
            const found = master.stores.find(s => s.name === store);
            if (!found || found.supplierCodes.length === 0) return;
            if (!toStr(codeInput.value)) codeInput.value = found.supplierCodes[0];
        };
        document.getElementById(pfx('progress-rep-filter')).addEventListener('change', resetProgressPageAndRender);
        document.getElementById(pfx('progress-status-filter')).addEventListener('change', resetProgressPageAndRender);
        document.getElementById(pfx('progress-store-search')).addEventListener('input', resetProgressPageAndRender);
        document.getElementById(pfx('progress-limit')).addEventListener('change', resetProgressPageAndRender);
        document.getElementById(pfx('progress-page-prev')).addEventListener('click', () => {
            if (state.progressCurrentPage <= 1) return;
            state.progressCurrentPage -= 1;
            renderProgress(true);
        });
        document.getElementById(pfx('progress-page-next')).addEventListener('click', () => {
            if (state.progressCurrentPage >= state.progressCurrentPageTotal) return;
            state.progressCurrentPage += 1;
            renderProgress(true);
        });
        document.getElementById(pfx('progress-form-store')).addEventListener('change', syncProgressSupplierCode);
        document.getElementById(pfx('progress-btn-save')).addEventListener('click', saveProgressFromForm);
        document.getElementById(pfx('progress-btn-clear')).addEventListener('click', () => setProgressForm(null));
        document.getElementById(pfx('progress-tbody')).addEventListener('click', (e) => {
            const target = e.target;
            if (!(target instanceof HTMLElement)) return;
            const id = target.dataset.id;
            if (!id) return;
            if (target.classList.contains('progress-edit-btn')) {
                const item = state.progressItems.find(x => x.id === id);
                if (item) setProgressForm(item);
                return;
            }
            if (target.classList.contains('progress-delete-btn')) {
                if (confirm('この進捗を削除しますか？')) deleteProgressItem(id);
            }
        });

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
        applyAuthLocks();
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
            applyAuthLocks();
            setProgressForm(null);
            renderProgress();
            setStoreAdvancedMode(false);
            KaientaiM.updateModuleStatus(MODULE_ID, '認証待ち', false);
            bindGoogleAuth();
        },
        onShow() {
            if (currentTab === 'progress') {
                switchModTab(currentTab);
                return;
            }
            if (state.results && currentTab !== 'upload' && currentTab !== 'settings') {
                switchModTab(currentTab);
            }
        }
    });

})();


