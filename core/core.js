// ============================================================
// Kaientai-M  —  Core Module System
// ============================================================
// 各分析モジュールはこのAPIを使って自分自身を登録する。
// 新しいモジュールは KaientaiM.registerModule({...}) を呼ぶだけ。
// ============================================================

window.KaientaiM = (function () {
    'use strict';

    const modules = [];
    let currentPage = 'home';

    // ── Shared Utilities (全モジュール共通) ──
    const util = {
        $(sel) { return document.querySelector(sel); },
        $$(sel) { return document.querySelectorAll(sel); },
        fmt(n) {
            if (n == null || isNaN(n)) return '-';
            return Math.round(n).toLocaleString('ja-JP');
        },
        fmtYen(n) {
            if (n == null || isNaN(n)) return '-';
            return '¥' + util.fmt(n);
        },
        fmtPct(n) {
            if (n == null || isNaN(n)) return '-';
            return (n * 100).toFixed(1) + '%';
        },
        toNum(v) {
            if (v == null) return 0;
            const n = typeof v === 'string' ? parseFloat(v.replace(/[,¥￥\s]/g, '')) : Number(v);
            return isNaN(n) ? 0 : n;
        },
        toStr(v) {
            if (v == null || v === '') return '';
            // 数値型JANコード対策: 科学的記数法を整数文字列に変換
            if (typeof v === 'number') {
                if (Number.isInteger(v)) return String(v);
                // 小数点がある場合も整数化を試みる（Excelの誤差対策）
                const rounded = Math.round(v);
                if (Math.abs(v - rounded) < 0.001) return String(rounded);
                return String(v);
            }
            return String(v).trim();
        },
        COL: { A:0,B:1,C:2,D:3,E:4,F:5,G:6,H:7,I:8,J:9,K:10,L:11,M:12,N:13,O:14,P:15,Q:16,R:17,S:18,T:19,U:20,V:21,W:22,X:23,Y:24,Z:25 },

        parseExcel(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    try {
                        const wb = XLSX.read(e.target.result, { type: 'array' });
                        const sheets = {};
                        wb.SheetNames.forEach(name => {
                            sheets[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
                        });
                        resolve({ sheets, sheetNames: wb.SheetNames, fileName: file.name });
                    } catch (err) { reject(err); }
                };
                reader.onerror = reject;
                reader.readAsArrayBuffer(file);
            });
        },

        exportCSV(header, rows, filename) {
            const bom = '\uFEFF';
            const csv = bom + [header, ...rows].map(r =>
                r.map(v => '"' + String(v).replace(/"/g, '""') + '"').join(',')
            ).join('\n');
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            a.click();
            URL.revokeObjectURL(url);
        },

        destroyChart(chartStore, id) {
            if (chartStore[id]) { chartStore[id].destroy(); delete chartStore[id]; }
        },

        createEl(tag, attrs, innerHTML) {
            const el = document.createElement(tag);
            if (attrs) Object.entries(attrs).forEach(([k, v]) => el.setAttribute(k, v));
            if (innerHTML) el.innerHTML = innerHTML;
            return el;
        }
    };

    // ── Module Registration ──
    function registerModule(config) {
        // config: { id, title, icon, description, color, init(containerEl, util) }
        modules.push(config);
        addModuleCard(config);
        addSidebarItem(config);
        addModulePage(config);
    }

    function addModuleCard(cfg) {
        const grid = document.getElementById('module-grid');
        const card = document.createElement('div');
        card.className = 'module-card';
        card.style.borderTopColor = cfg.color || '#1a237e';
        card.innerHTML = `
            <div class="module-card-icon" style="color:${cfg.color || '#1a237e'}">${cfg.icon || '&#128202;'}</div>
            <h3 class="module-card-title">${cfg.title}</h3>
            <p class="module-card-desc">${cfg.description || ''}</p>
            <div class="module-card-status" id="mod-status-${cfg.id}">未設定</div>
        `;
        card.addEventListener('click', () => navigateTo(cfg.id));
        grid.appendChild(card);
    }

    function addSidebarItem(cfg) {
        const nav = document.getElementById('sidebar-nav');
        const btn = document.createElement('button');
        btn.className = 'sidebar-btn';
        btn.dataset.page = cfg.id;
        btn.innerHTML = `<span class="sidebar-icon">${cfg.icon || '&#128202;'}</span><span>${cfg.title}</span>`;
        btn.addEventListener('click', () => navigateTo(cfg.id));
        nav.appendChild(btn);
    }

    function addModulePage(cfg) {
        const main = document.getElementById('main');
        const page = document.createElement('div');
        page.className = 'page';
        page.id = 'page-' + cfg.id;
        main.appendChild(page);

        // Initialize module — pass container and utilities
        if (typeof cfg.init === 'function') {
            cfg.init(page, util);
        }
    }

    // ── Navigation ──
    function navigateTo(pageId) {
        currentPage = pageId;
        document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
        document.querySelectorAll('.sidebar-btn').forEach(b => b.classList.remove('active'));

        const page = document.getElementById('page-' + pageId);
        if (page) page.classList.add('active');

        const btn = document.querySelector(`.sidebar-btn[data-page="${pageId}"]`);
        if (btn) btn.classList.add('active');

        // Fire module onShow if exists
        const mod = modules.find(m => m.id === pageId);
        if (mod && typeof mod.onShow === 'function') mod.onShow();
    }

    function updateModuleStatus(moduleId, text, ok) {
        const el = document.getElementById('mod-status-' + moduleId);
        if (el) {
            el.textContent = text;
            el.className = 'module-card-status ' + (ok ? 'ok' : '');
        }
    }

    // Home button
    document.querySelector('.sidebar-btn[data-page="home"]').addEventListener('click', () => navigateTo('home'));

    return {
        registerModule,
        navigateTo,
        updateModuleStatus,
        util,
        getModules() { return modules; }
    };
})();
