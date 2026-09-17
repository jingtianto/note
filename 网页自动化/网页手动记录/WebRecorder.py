"""
web_recorder.py (Playwright 版 v1.2) - 支持多标签页/新窗口的完整录制器

用法：
    python web_recorder.py https://www.example.com

输出：
    actions.jsonl        - 结构化记录
    actions_readable.txt - AI 友好版
    snapshots/           - DOM 快照（按内容哈希去重）

特点：
    - 无需 webdriver，直接用系统 Edge
    - context.add_init_script 自动注入所有新页面和 iframe
    - launch_persistent_context 保留登录态
"""
import sys
import os
import json
import time
import hashlib
import signal
import threading
import traceback
from datetime import datetime

# 中文 Windows 的 cmd 默认 GBK 代码页：直接 print emoji 会 UnicodeEncodeError 直接崩。
# 这里只把"编不出来的字符"降级（中文照常显示），输出文件仍是 UTF-8 不受影响。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors='replace')
    except Exception:
        pass

def _wait_enter(msg):
    """双击运行时窗口不能一闪而过：让用户看完提示再关。"""
    try:
        input(msg)
    except Exception:
        pass

def _on_uncaught(etype, value, tb):
    traceback.print_exception(etype, value, tb)
    print("\n❌ 程序异常结束。已采集的数据在 runs 目录下的 actions.partial.jsonl 里。")
    _wait_enter("按回车关闭窗口…")

sys.excepthook = _on_uncaught

try:
    from playwright.sync_api import sync_playwright, Error as PlaywrightError
except ImportError as e:
    print("❌ 缺少 playwright，无法启动浏览器。")
    print("   请先在本机执行：pip install playwright")
    print(f"   （原始错误：{e}）")
    _wait_enter("按回车关闭窗口…")
    sys.exit(1)

# Windows：Ctrl+Break 也走正常收尾流程（程序化/脚本化停止时用得上，Ctrl+C 在部分场景送不进去）
def _on_break(signum, frame):
    raise KeyboardInterrupt()
try:
    signal.signal(signal.SIGBREAK, _on_break)
except Exception:
    pass

# ======================== 配置 ========================
# 不给网址就开空白页，由用户在浏览器里自己输（双击运行时的默认行为）
TARGET_URL = sys.argv[1].strip() if len(sys.argv) > 1 else ""

# 每次运行一个独立文件夹：runs/年月日时分秒/
BASE_DIR = os.path.dirname(os.path.abspath(__file__))   # 始终放脚本旁边，换工作目录也不会散落
RUNS_DIR = os.path.join(BASE_DIR, "runs")
RUN_START = datetime.now()
RUN_ID = RUN_START.strftime("%Y%m%d%H%M%S")
RUN_DIR = os.path.join(RUNS_DIR, RUN_ID)
_suffix = 1
while os.path.exists(RUN_DIR):        # 同一秒内重复启动也不会互相覆盖
    _suffix += 1
    RUN_DIR = os.path.join(RUNS_DIR, f"{RUN_ID}_{_suffix}")

SNAPSHOT_NAME = "snapshots"          # 快照子目录名（相对本次运行文件夹）
SNAPSHOT_DIR = os.path.join(RUN_DIR, SNAPSHOT_NAME)
TABLES_NAME = "tables"               # 表格结构化 dump 子目录
TABLES_DIR = os.path.join(RUN_DIR, TABLES_NAME)
OUTPUT_JSONL = os.path.join(RUN_DIR, "actions.jsonl")
OUTPUT_TXT = os.path.join(RUN_DIR, "actions_readable.txt")
OUTPUT_INDEX = os.path.join(RUN_DIR, "AI_INDEX.md")   # 给 AI 的入口说明（先读这个）
PARTIAL_JSONL = os.path.join(RUN_DIR, "actions.partial.jsonl")   # 增量备份，异常中断时靠它救数据
USER_DATA_DIR = os.path.join(BASE_DIR, "edge_profile")
POLL_INTERVAL = 0.5
DEBUG = True
MAX_ACTIONS_PER_POLL = 5
WINDOW_ASSOC_WINDOW = 8.0
DRAIN_SECONDS = 4.0      # 收尾时最多再采集这么久，不再静默干等
CLOSE_TIMEOUT = 15.0     # 关闭浏览器最长等待，超时强制退出进程

try:
    os.makedirs(SNAPSHOT_DIR, exist_ok=True)   # 顺带建好 runs/<时间戳>/snapshots
    os.makedirs(TABLES_DIR, exist_ok=True)
except OSError as e:
    print(f"❌ 无法在脚本目录下创建输出文件夹：{e}")
    print(f"   脚本位置：{BASE_DIR}")
    print("   请把 WebRecorder.py 放到你有写权限的目录（例如 D:\\ 下的自建文件夹），"
          "不要放在 C:\\Program Files 这类受保护目录。")
    sys.exit(1)
# 可选：设置 WEBREC_DEBUG_PORT 会把这个端口开给 Chrome DevTools 协议，
# 方便外部工具附着到录制中的浏览器（排查/自动化测试）
_extra_args = []
_dbg_port = os.environ.get("WEBREC_DEBUG_PORT")
if _dbg_port:
    _extra_args.append(f"--remote-debugging-port={_dbg_port}")
    print(f"🔧 已开启调试端口：{_dbg_port}")

# ======================== 快照管理 ========================
snapshot_map = {}

def save_snapshot(html):
    if not html or not isinstance(html, str):
        return None
    h = hashlib.sha1(html.encode('utf-8', errors='ignore')).hexdigest()[:16]
    if h in snapshot_map:
        return snapshot_map[h]
    fname = f"snap_{h}.html"
    path = os.path.join(SNAPSHOT_DIR, fname)
    try:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(html)
    except Exception:
        return None
    snapshot_map[h] = fname
    return fname

# ======================== 表格 dump 落盘 ========================
table_map = {}   # hash -> True，同一种表格状态只存一份

def save_table_payload(it):
    """把动作里携带的表格状态（含"操作后"状态）去重落盘，动作里只留文件引用。"""
    h = it.get('table_hash')
    full = it.pop('table_full', None)
    if h:
        if full is not None and h not in table_map:
            table_map[h] = True
            write_table_files(h, full)
        it['table_file'] = f"{TABLES_NAME}/table_{h}.json" if h in table_map else None
    after = it.get('table_after')
    if isinstance(after, dict):
        ah = after.get('hash')
        afull = after.pop('full', None)
        if ah and afull is not None and ah not in table_map:
            table_map[ah] = True
            write_table_files(ah, afull)
        if ah:
            after['file'] = f"{TABLES_NAME}/table_{ah}.json" if ah in table_map else None

def write_table_files(h, full):
    cols = [(c.get('text') or '').strip() for c in (full.get('columns') or [])]
    rows = full.get('rows') or []
    try:
        with open(os.path.join(TABLES_DIR, f"table_{h}.json"), 'w', encoding='utf-8') as f:
            json.dump(full, f, ensure_ascii=False, indent=1)
    except OSError as e:
        print(f"⚠️  表格 JSON 写入失败：{e}")
    try:
        # 额外给一份 TSV：直接看表、也方便丢给 AI
        with open(os.path.join(TABLES_DIR, f"table_{h}.tsv"), 'w', encoding='utf-8') as f:
            f.write('\t'.join(['row_key'] + cols) + '\n')
            for r in rows:
                cells = [str(c).replace('\t', ' ') for c in (r.get('cells') or [])]
                f.write('\t'.join([str(r.get('key', ''))] + cells) + '\n')
    except OSError as e:
        print(f"⚠️  表格 TSV 写入失败：{e}")

# ======================== 注入 JS ========================
MONITOR_JS = r"""
(function() {
    if (window._recorderInitialized) return;
    window._recorderInitialized = true;

    function isDynamicClass(cls) {
        if (!cls) return false;
        return /^(css-|sc-|jsx-|emotion-)/i.test(cls)
            || /\d{4,}/.test(cls)
            || /^[a-z]{1,3}-[a-z0-9]{8,}$/i.test(cls)
            || /^[a-f0-9]{6,}$/i.test(cls);
    }
    function getClassNameString(el) {
        if (!el || !el.className) return '';
        const cls = el.className;
        if (typeof cls === 'string') return cls;
        if (cls.baseVal) return cls.baseVal;
        return '';
    }
    function getStableClasses(el) {
        const cls = getClassNameString(el);
        if (!cls) return [];
        return cls.trim().split(/\s+/).filter(c => c && !isDynamicClass(c));
    }
    function getDynamicClasses(el) {
        const cls = getClassNameString(el);
        if (!cls) return [];
        return cls.trim().split(/\s+/).filter(c => c && isDynamicClass(c));
    }
    function sanitizeValue(el, value) {
        if (value == null || value === '') return value;
        const type = (el.type || '').toLowerCase();
        if (type === 'password') return '***';
        const nameId = ((el.name || '') + (el.id || '') + (el.getAttribute && (el.getAttribute('autocomplete') || '') || '')).toLowerCase();
        if (/token|secret|api[_-]?key|password|pwd|credit|card|cvv|ssn|auth/.test(nameId)) return '***';
        return value;
    }
    function safeCssEscape(s) {
        if (window.CSS && CSS.escape) return CSS.escape(s);
        return String(s).replace(/([^\w-])/g, '\\$1');
    }
    function isInShadow(el) {
        try { return el.getRootNode() instanceof ShadowRoot; } catch (e) { return false; }
    }

    function getXPath(el) {
        if (!el || el.nodeType !== 1) return '';
        if (el.id && !isInShadow(el)) return '//*[@id="' + el.id + '"]';
        const parts = [];
        let current = el, guard = 0;
        while (current && current.nodeType === 1 && guard++ < 100) {
            if (current.id && !isInShadow(current)) {
                const prefix = '//*[@id="' + current.id + '"]';
                return parts.length ? prefix + '/' + parts.join('/') : prefix;
            }
            const root = current.getRootNode && current.getRootNode();
            if (root && root instanceof ShadowRoot) {
                let idx = 1, sib = current;
                while (sib.previousElementSibling) {
                    sib = sib.previousElementSibling;
                    if (sib.tagName === current.tagName) idx++;
                }
                parts.unshift('shadow::' + current.tagName.toLowerCase() + (idx > 1 ? '[' + idx + ']' : ''));
                current = root.host;
                continue;
            }
            let idx = 0, hasSame = false;
            const siblings = current.parentNode ? current.parentNode.childNodes : [];
            for (const sib of siblings) {
                if (sib.nodeType === 1 && sib.tagName === current.tagName) {
                    hasSame = true; idx++;
                    if (sib === current) break;
                }
            }
            const tag = current.tagName.toLowerCase();
            parts.unshift(hasSame ? tag + '[' + idx + ']' : tag);
            if (current === document.documentElement) break;
            current = current.parentElement;
        }
        return '/' + parts.join('/');
    }

    function getCssPath(el) {
        if (!el || el.nodeType !== 1) return '';
        const parts = [];
        let current = el, guard = 0;
        while (current && current.nodeType === 1 && guard++ < 100) {
            if (current.id && !isInShadow(current)) {
                parts.unshift('#' + safeCssEscape(current.id));
                break;
            }
            const root = current.getRootNode && current.getRootNode();
            if (root && root instanceof ShadowRoot) {
                parts.unshift('shadow::' + current.tagName.toLowerCase());
                current = root.host;
                continue;
            }
            let sel = current.tagName.toLowerCase();
            const stableCls = getStableClasses(current);
            if (stableCls.length) sel += '.' + stableCls.map(safeCssEscape).join('.');
            let nth = 1, sib = current;
            while (sib.previousElementSibling) {
                sib = sib.previousElementSibling;
                if (sib.tagName === current.tagName) nth++;
            }
            if (nth > 1) sel += ':nth-of-type(' + nth + ')';
            parts.unshift(sel);
            if (current === document.body) break;
            current = current.parentElement;
        }
        return parts.join(' > ');
    }

    const KEEP_ATTRS = ['id','class','name','type','value','href','src','role',
                        'data-testid','data-test','data-cy','aria-label',
                        'placeholder','title','alt','for','autocomplete'];
    function getElementInfo(el) {
        if (!el || el.nodeType !== 1) return { tagName: 'UNKNOWN' };
        const rect = el.getBoundingClientRect();
        const attrs = {};
        for (const attr of el.attributes || []) {
            if (KEEP_ATTRS.includes(attr.name) || attr.name.startsWith('data-') || attr.name.startsWith('aria-')) {
                attrs[attr.name] = attr.value;
            }
        }
        if (attrs.value !== undefined) attrs.value = sanitizeValue(el, attrs.value);
        return {
            tagName: el.tagName,
            id: el.id || '',
            className_stable: getStableClasses(el),
            className_dynamic: getDynamicClasses(el),
            text: ((el.innerText || el.textContent || '').trim()).slice(0, 200),
            value: sanitizeValue(el, el.value || ''),
            xpath: getXPath(el),
            css: getCssPath(el),
            attributes: attrs,
            rect: { x: Math.round(rect.x), y: Math.round(rect.y), w: Math.round(rect.width), h: Math.round(rect.height) },
            in_shadow_dom: isInShadow(el)
        };
    }

    function getParentChain(el, depth) {
        depth = depth || 3;
        const chain = [];
        let cur = el.parentElement, i = 0;
        while (cur && i < depth) {
            chain.push({
                tag: cur.tagName, id: cur.id || '',
                class: getStableClasses(cur).join(' '),
                text: ((cur.innerText || '').trim()).slice(0, 80)
            });
            cur = cur.parentElement; i++;
        }
        return chain;
    }
    function getContext(el) {
        let dialog = null, table = null, row = null, list = null;
        try { dialog = el.closest('[role="dialog"],[role="alertdialog"],.modal,.ant-modal,.el-dialog,.MuiDialog-root'); } catch(e) {}
        try { table = el.closest('table'); } catch(e) {}
        try { row = el.closest('tr'); } catch(e) {}
        try { list = el.closest('ul,ol,[role="list"],[role="listbox"]'); } catch(e) {}
        let rowIndex = null;
        if (row && row.parentElement) rowIndex = Array.from(row.parentElement.children).indexOf(row);
        let listIndex = null;
        if (list) {
            const items = Array.from(list.children);
            let item = el;
            try { item = el.closest('li,[role="listitem"],[role="option"]') || el; } catch(e) {}
            listIndex = items.indexOf(item);
        }
        return {
            in_dialog: !!dialog,
            dialog_selector: dialog ? (
                dialog.id ? '#' + dialog.id :
                getStableClasses(dialog).length ? '.' + getStableClasses(dialog).join('.') : '[role="dialog"]'
            ) : null,
            in_table: !!table, row_index: rowIndex,
            in_list: !!list, list_index: listIndex,
            parent_chain: getParentChain(el, 3)
        };
    }

    function captureDom() {
        try {
            // 大页面（典型：几千行的表格）整页克隆会明显卡顿，超限就放弃快照，
            // 表格内容另有结构化 dump 兜底，不会丢信息
            const n = document.getElementsByTagName('*').length;
            if (n > SNAPSHOT_MAX_ELEMENTS) return null;
            const clone = document.documentElement.cloneNode(true);
            clone.querySelectorAll('script,style,noscript,link[rel="stylesheet"],svg defs').forEach(n => n.remove());
            const all = clone.querySelectorAll('*');
            for (const n of all) {
                if (!n.attributes) continue;
                const toRemove = [];
                for (const attr of n.attributes) if (attr.name.startsWith('on')) toRemove.push(attr.name);
                for (const name of toRemove) n.removeAttribute(name);
            }
            return clone.outerHTML;
        } catch (e) { return '<error>' + e.message + '</error>'; }
    }
    function captureStableDom(callback, quietMs, maxWaitMs) {
        quietMs = quietMs || 400; maxWaitMs = maxWaitMs || 2500;
        let timer = null, done = false;
        const finish = () => {
            if (done) return;
            done = true;
            try { observer.disconnect(); } catch (e) {}
            try { callback(captureDom()); } catch (e) { callback('<error>' + e.message + '</error>'); }
        };
        let observer;
        try {
            observer = new MutationObserver(() => {
                if (timer) clearTimeout(timer);
                timer = setTimeout(finish, quietMs);
            });
            observer.observe(document, { subtree: true, childList: true, attributes: true, characterData: true });
        } catch (e) { finish(); return; }
        setTimeout(finish, maxWaitMs);
    }

    function backupPending() {
        try {
            const pending = (window._recordedActions || []).filter(a => !a._sent);
            if (!pending.length) {
                sessionStorage.removeItem('_recorder_pending');
                return;
            }
            let slim = pending.map(a => {
                const c = {};
                for (const k in a) {
                    if (k === 'after_snapshot') continue;
                    c[k] = a[k];
                }
                if (a._ready && a.after_snapshot) c.after_snapshot = a.after_snapshot;
                return c;
            });
            try {
                sessionStorage.setItem('_recorder_pending', JSON.stringify(slim));
                return;
            } catch (e) {}
            slim = pending.map(a => {
                const c = {};
                for (const k in a) {
                    if (k === 'before_snapshot' || k === 'after_snapshot' || k === 'table_full') continue;
                    c[k] = a[k];
                }
                return c;
            });
            sessionStorage.setItem('_recorder_pending', JSON.stringify(slim));
        } catch (e) {}
    }
    setInterval(backupPending, 1000);   // 原来 300ms 太频，大表格下 stringify 开销明显
    window.addEventListener('beforeunload', backupPending);
    window.addEventListener('pagehide', backupPending);

    if (!window._recordedActions) window._recordedActions = [];
    try {
        const slim = JSON.parse(sessionStorage.getItem('_recorder_pending') || '[]');
        if (slim.length) {
            slim.forEach(a => { a._recovered = true; a._sent = false; a._ready = true; });
            window._recordedActions = slim.concat(window._recordedActions);
            sessionStorage.removeItem('_recorder_pending');
        }
    } catch (e) {}

    // ==================== 表格结构化采集 ====================
    // 目标：不管原生 table 还是 antd / Element / AG Grid 这类 div 表格，
    // 都要把「哪张表、哪些列、多少行、点的是哪一行哪一列的哪个单元格、
    // 那一行的完整内容、总数/页码、滚动容器」全部记下来，才能喂给 AI 用。
    const TABLE_MAX_ROWS = 300;
    const TABLE_MAX_COLS = 60;
    const CELL_MAX = 120;
    const SNAPSHOT_MAX_ELEMENTS = 6000;

    const ROW_SEL = 'tr,[role="row"],.ant-table-row,.el-table__row,.ag-row,[data-row-key],[data-index]';
    const CELL_SEL = 'td,th,[role="gridcell"],[role="columnheader"],.ant-table-cell,.el-table__cell,.ag-cell';
    const CELL_QUERY = CELL_SEL + ',[class*="cell"],.th,.td';
    const HEADER_SEL = 'thead th,[role="columnheader"],.ant-table-thead th,.el-table__header th,.ag-header-cell,[class*="thead"]';
    const TABLE_BOX_SEL = 'table,[role="grid"],[role="treegrid"],[role="table"],.ant-table,.el-table,.ag-root-wrapper,.ag-root,.vxe-table,.ivu-table';

    // 很多自研 div 表格直接用 .th / .td / xxx-cell 这类类名，不是 <th>/<td> 标签：
    // 只认标签会把它们全漏掉（列名识别不出来、行列定位也没有），这里补上类名判断
    const CELL_TOKENS = ['th', 'td', 'cell', 'column', 'columnheader', 'gridcell', 'headercell'];
    function isCellEl(el) {
        if (!el || el.nodeType !== 1) return false;
        try { if (el.matches(CELL_SEL)) return true; } catch (e) {}
        const cls = ('' + (el.className || '')).trim();
        if (!cls) return false;
        for (const t of cls.split(/\s+/)) {
            const s = t.toLowerCase();
            if (CELL_TOKENS.includes(s)) return true;
            if (/^(cell|td|th|column)[-_]/.test(s)) return true;
            if (/[-_](cell|td|th|column)$/.test(s)) return true;
        }
        return false;
    }
    function isHeadish(el) {
        const cls = ('' + (el && el.className || '')).toLowerCase();
        return /(^|[-_ ])(head|thead|header)([-_ ]|$)/.test(cls);
    }
    function closestRowLike(el) {
        let cur = el, guard = 0;
        while (cur && cur.nodeType === 1 && guard++ < 15) {
            try { if (cur.matches(ROW_SEL)) return cur; } catch (e) {}
            cur = cur.parentElement;
        }
        return null;
    }
    function closestCellLike(el) {
        let cur = el, guard = 0;
        while (cur && cur.nodeType === 1 && guard++ < 12) {
            if (isCellEl(cur)) return cur;
            cur = cur.parentElement;
        }
        return null;
    }

    // 表头筛选面板、下拉、弹窗往往被渲染在 body 下（不在表格 DOM 里），
    // 认出来才能把里面的输入/勾选关联到对应的列
    function floatingPanelOf(el) {
        if (!el || el.nodeType !== 1) return null;
        let cur = el, last = null, guard = 0;
        while (cur && cur.nodeType === 1 && guard++ < 12) {
            if (cur.parentElement === document.body || cur.parentNode === document) { last = cur; break; }
            cur = cur.parentElement;
        }
        if (!last) return null;
        let floating = false;
        try {
            const cs = getComputedStyle(last);
            if (cs.position === 'fixed' || cs.position === 'absolute' || cs.position === 'sticky') floating = true;
        } catch (e) {}
        const hint = ('' + (last.className || '') + ' ' + (last.id || '')).toLowerCase();
        if (/(panel|popover|popup|dropdown|popper|modal|dialog|drawer|overlay|filter|menu|tooltip|select)/.test(hint)) floating = true;
        if (!floating) return null;
        // 单个面板元素用 innerText 不心疼（能拿到可读的分隔），文本太长才退化
        let text = '';
        try { text = ((last.innerText || last.textContent) || '').trim().replace(/\s+/g, ' ').slice(0, 140); } catch (e) {}
        return { id: last.id || '', cls: ('' + (last.className || '')).slice(0, 60), text: text };
    }

    function cellText(el) {
        // 用 textContent 而不是 innerText：innerText 会触发排版计算，
        // 表格 dump 要遍历几千个单元格，用 innerText 能把页面卡死
        const t = ((el && el.textContent) || '').trim().replace(/\s+/g, ' ');
        return t.slice(0, CELL_MAX);
    }

    function ownRows(root) {
        let all = [];
        try { all = Array.from(root.querySelectorAll(ROW_SEL)); } catch (e) { return []; }
        if (!all.length) return [];
        const set = new Set(all);
        return all.filter(r => {
            let p = r.parentElement, guard = 0;
            while (p && guard++ < 40) {
                if (set.has(p)) return false;      // 嵌套的行，丢掉外层
                p = p.parentElement;
            }
            return true;
        });
    }

    function cellElsOf(row) {
        let cs = [];
        try { cs = Array.from(row.children).filter(isCellEl); } catch (e) { cs = []; }
        if (!cs.length) {
            try { cs = Array.from(row.querySelectorAll(CELL_QUERY)).filter(isCellEl); } catch (e) { return []; }
            const set = new Set(cs);
            cs = cs.filter(c => {
                let p = c.parentElement, g = 0;
                while (p && p !== row && g++ < 12) {
                    if (set.has(p)) return false;
                    p = p.parentElement;
                }
                return true;
            });
        }
        return cs;
    }

    function tableRootOf(el, deep) {
        if (!el) return null;
        // 1) 组件库特征容器
        try {
            const t = el.closest(TABLE_BOX_SEL);
            if (t && ownRows(t).length) return t;
        } catch (e) {}
        // 2) 从元素往上找「最近的行容器」——不能取最外层，否则会退化成整个页面
        let cur = el, rowBox = null, guard = 0;
        while (cur && cur.nodeType === 1 && cur !== document.body && guard++ < 15) {
            if (ownRows(cur).length >= 2) { rowBox = cur; break; }
            cur = cur.parentElement;
        }
        // 3) 滚动场景：目标可能是滚动容器本身，往下找
        if (!rowBox && deep && el.querySelector) {
            for (const sel of ['table', '[role="grid"]', '.ant-table', '.el-table', '.ag-root-wrapper', '.ag-root', '[class*="table"]']) {
                try {
                    const t = el.querySelector(sel);
                    if (t && ownRows(t).length) { rowBox = t; break; }
                } catch (e) {}
            }
        }
        if (!rowBox) return null;
        // 4) 行容器自己就带表头就直接用它，否则往上扩到「带表头或总数」的最近祖先
        try { if (headerCellEls(rowBox).length > 0) return rowBox; } catch (e) {}
        let best = rowBox, r = rowBox.parentElement, g2 = 0;
        while (r && r.nodeType === 1 && r !== document.body && g2++ < 6) {
            let hasHead = false, hasTotal = false;
            try { hasHead = headerCellEls(r).length > 0; } catch (e) {}
            try { hasTotal = totalTextOf(r) !== ''; } catch (e) {}
            if (hasHead || hasTotal) { best = r; break; }
            r = r.parentElement;
        }
        return best;
    }

    function scrollBoxOf(el) {
        const doc = document.scrollingElement || document.documentElement;
        if (!el || el.nodeType !== 1) return doc;
        const usable = x => {
            try { return x.clientHeight > 80 && x.scrollHeight > x.clientHeight + 20; } catch (e) { return false; }
        };
        if (usable(el)) return el;
        let cur = el.parentElement, guard = 0;
        while (cur && guard++ < 40) {
            if (usable(cur)) return cur;
            cur = cur.parentElement;
        }
        // 表格自己的滚动区往往在内部（.el-table__body-wrapper / .ant-table-body / .ag-body-viewport）
        try {
            for (const sel of ['.el-table__body-wrapper', '.el-scrollbar__wrap', '.ant-table-body',
                               '.ant-table-content', '.ag-body-viewport', '.vxe-table--body-wrapper',
                               '[class*="body-wrapper"]', '[style*="overflow"]']) {
                const d = el.querySelector(sel);
                if (d && usable(d)) return d;
            }
            // 候选项限定在类名像"滚动区"的元素上，避免对大表格做全量 getComputedStyle
            const cand = el.querySelectorAll('[class*="body"],[class*="wrap"],[class*="scroll"],[class*="viewport"],[class*="content"]');
            let n = 0;
            for (const d of cand) {
                if (n++ > 200) break;
                if (usable(d)) return d;
            }
        } catch (e) {}
        return doc;
    }

    function headRowEl(root) {
        // 自定义 div 表格的"表头"往往就是一个类名带 head 的普通 div 行
        try {
            for (const c of Array.from(root.querySelectorAll('div,tr,ul,li,section'))) {
                if (!isHeadish(c)) continue;
                if (cellElsOf(c).length >= 2) return c;
            }
        } catch (e) {}
        const rs = ownRows(root);
        if (rs.length && isHeadish(rs[0])) return rs[0];
        return null;
    }

    const _headCache = new WeakMap();
    function headerCellEls(root) {
        const c = _headCache.get(root);
        if (c && Date.now() - c.at < 500) return c.cells;
        let out = [];
        let hs = [];
        try { hs = Array.from(root.querySelectorAll(HEADER_SEL)); } catch (e) { hs = []; }
        if (hs.length) {
            const set = new Set(hs);
            hs = hs.filter(h => {
                let p = h.parentElement, g = 0;
                while (p && g++ < 40) {
                    if (set.has(p)) return false;
                    p = p.parentElement;
                }
                return true;
            });
            if (hs.some(h => cellText(h))) out = hs.slice(0, TABLE_MAX_COLS);
        }
        if (!out.length) {
            const hr = headRowEl(root);
            if (hr) out = cellElsOf(hr).slice(0, TABLE_MAX_COLS);
        }
        _headCache.set(root, { at: Date.now(), cells: out });
        return out;
    }

    function headerCellText(el) {
        // 表头单元格里往往还有排序箭头、筛选按钮、图标，要剔掉才能得到干净列名
        try {
            const c = el.cloneNode(true);
            c.querySelectorAll('button,svg,i,input,[class*="sort"],[class*="btn"],[class*="arrow"],[class*="icon"],[class*="filter"]')
                .forEach(n => n.remove());
            const t = ((c.textContent) || '').trim().replace(/\s+/g, ' ');
            if (t) return t.slice(0, CELL_MAX);
        } catch (e) {}
        return cellText(el).replace(/[\u21c5\u25b2\u25bc\u2191\u2193]/g, '').replace(/\s*筛选\s*/g, ' ').trim();
    }

    function headerTexts(root) {
        return headerCellEls(root).map((h, i) => {
            let w = 0;
            try { w = Math.round(h.getBoundingClientRect().width); } catch (e) {}
            return { i: i, text: headerCellText(h), w: w };
        });
    }

    function rowKeyOf(row, cells) {
        const a = row.attributes;
        if (a) {
            for (const k of ['data-row-key', 'data-key', 'data-id', 'data-index', 'id', 'aria-rowindex']) {
                if (a[k] && a[k].value) return k + '=' + String(a[k].value).slice(0, 60);
            }
        }
        for (const t of cells) if (t) return 'text:' + t.slice(0, 40);
        return '';
    }

    const _totalCache = new WeakMap();
    function totalTextOf(root) {
        const c = _totalCache.get(root);
        if (c && Date.now() - c.at < 500) return c.text;
        let text = '';
        let scope = root;
        try { scope = root.closest('.ant-table-wrapper,.el-table,.ant-table,.ag-root-wrapper,.vxe-table') || root.parentElement || root; } catch (e) {}
        const sels = ['.ant-pagination-total-text', '.el-pagination__total', '.vxe-pager--total',
                      '[class*="pagination"] [class*="total"]', '[class*="total"]', '[class*="count"]'];
        const scopes = [scope, scope && scope.parentElement, scope && scope.parentElement && scope.parentElement];
        for (const sel of sels) {
            for (const sc of scopes) {
                try {
                    const n = sc && sc.querySelector(sel);
                    if (!n) continue;
                    const t = ((n.textContent) || '').trim().replace(/\s+/g, ' ');
                    if (/\d/.test(t)) { text = t.slice(0, 60); break; }
                } catch (e) {}
            }
            if (text) break;
        }
        if (!text) {
            // 兜底："共 40 条"这种文本可能挂在别处；只在候选元素里找且限数量，
            // 否则在大表格上会把主线程堵死
            try {
                for (const sc of scopes) {
                    if (!sc) continue;
                    const cand = sc.querySelectorAll('[class*="total"],[class*="count"],[class*="pager"],[class*="sum"],span,b');
                    let n = 0;
                    for (const el of cand) {
                        if (n++ > 300) break;
                        if (el.childElementCount > 8) continue;
                        const t = ((el.textContent) || '').trim().replace(/\s+/g, ' ');
                        if (t.length > 2 && t.length < 80 && /(共|总共|合计|total)\s*[:：]?\s*\d+/i.test(t)) {
                            text = t.slice(0, 60); break;
                        }
                    }
                    if (text) break;
                }
            } catch (e) {}
        }
        _totalCache.set(root, { at: Date.now(), text: text });
        return text;
    }

    function pageTextOf(root) {
        const sels = ['.ant-pagination-item-active', '.el-pager li.is-active', '.el-pager li.active', '.vxe-pager--num-btn.active'];
        for (const sel of sels) {
            try {
                const n = root.closest('div') && (root.closest('.ant-table-wrapper,.el-table,.vxe-table,.ant-table') || root).querySelector(sel);
                if (n && (n.innerText || '').trim()) return (n.innerText || '').trim().slice(0, 12);
            } catch (e) {}
        }
        return '';
    }

    function looksVirtual(root) {
        try {
            if (root.querySelector('.ag-center-cols-container,.ag-body-viewport')) return true;
            if (root.querySelector('[style*="translateY"],[style*="translate3d"]')) return true;
            const r = ownRows(root)[0];
            if (r) {
                const t = getComputedStyle(r).transform;
                if (t && t !== 'none' && t.indexOf('matrix') === 0) return true;
                const pos = getComputedStyle(r).position;
                if (pos === 'absolute') return true;
            }
        } catch (e) {}
        return false;
    }

    function buildTable(root, targetEl) {
        const rows = ownRows(root);
        const cols = headerTexts(root);
        let targetRowEl = null, targetColEl = null, targetRowIdx = -1, targetColIdx = -1;
        let inHeader = false;
        if (targetEl) {
            targetRowEl = closestRowLike(targetEl);
            if (targetRowEl && !rows.includes(targetRowEl)) {
                // 目标行没被识别进来（例如它自己被当成嵌套行），直接补上
                rows.unshift(targetRowEl);
            }
            if (targetRowEl) {
                targetRowIdx = rows.indexOf(targetRowEl);
                const celEls = cellElsOf(targetRowEl);
                targetColEl = closestCellLike(targetEl);
                if (targetColEl) targetColIdx = celEls.indexOf(targetColEl);
                if (targetColIdx < 0) targetColIdx = celEls.findIndex(c => c.contains(targetEl));
            } else {
                // 不在数据行里：可能是点了表头（排序 / 筛选按钮），把列信息也记下来
                const hc = closestCellLike(targetEl);
                if (hc) {
                    const hcells = headerCellEls(root);
                    const hi = hcells.indexOf(hc);
                    if (hi >= 0) { inHeader = true; targetColIdx = hi; targetColEl = hc; }
                }
            }
        }

        const dump = [];
        let hidden = 0;
        rows.forEach(r => {
            const cells = cellElsOf(r).map(cellText);
            const item = { key: rowKeyOf(r, cells), cells: cells, target: r === targetRowEl };
            let vis = true;
            try {
                const cs = getComputedStyle(r);
                vis = cs.display !== 'none' && cs.visibility !== 'hidden' && r.getClientRects().length > 0;
            } catch (e) {}
            if (!vis) { item.hidden = true; hidden++; }
            dump.push(item);
        });
        let truncated = 0;
        if (rows.length > TABLE_MAX_ROWS) {
            truncated = rows.length;
            const keep = dump.filter(x => x.target).concat(dump.slice(0, TABLE_MAX_ROWS - 1));
            dump.length = 0;
            keep.forEach(x => dump.push(x));
        }

        let box = null;
        try {
            const s = scrollBoxOf(root);
            box = {
                selector: getCssPath(s),
                scroll_top: Math.round(s.scrollTop || 0),
                scroll_height: s.scrollHeight || 0,
                client_height: s.clientHeight || 0,
                at_bottom: (s.scrollTop + s.clientHeight) >= (s.scrollHeight - 40)
            };
        } catch (e) {}

        let target = null;
        if (targetRowEl) {
            target = {
                in_header: false,
                row_index_in_dom: targetRowIdx,
                row_key: rowKeyOf(targetRowEl, cellElsOf(targetRowEl).map(cellText)),
                row_text: cellText(targetRowEl).slice(0, 300),
                col_index: targetColIdx,
                col_name: (cols[targetColIdx] || {}).text || '',
                cell_text: targetColEl ? cellText(targetColEl) : '',
                tag: targetEl ? targetEl.tagName : ''
            };
        } else if (inHeader) {
            target = {
                in_header: true,
                col_index: targetColIdx,
                col_name: (cols[targetColIdx] || {}).text || '',
                cell_text: targetColEl ? cellText(targetColEl) : '',
                tag: targetEl ? targetEl.tagName : ''
            };
        }

        return {
            kind: (root.tagName === 'TABLE' || root.getAttribute('role') === 'grid' || root.getAttribute('role') === 'table') ? 'html' : 'div',
            selector: getCssPath(root),
            virtual: looksVirtual(root),
            count_in_dom: rows.length,
            hidden_rows: hidden,
            truncated_rows: truncated,
            total_text: totalTextOf(root),
            page_text: pageTextOf(root),
            scroll_box: box,
            columns: cols,
            rows: dump,
            target: target
        };
    }

    function fastHash(s) {
        let h1 = 0x811c9dc5, h2 = 0x1000193;
        for (let i = 0; i < s.length; i++) {
            const c = s.charCodeAt(i);
            h1 = ((h1 ^ c) * 16777619) >>> 0;
            h2 = (h2 + c * (i + 1)) >>> 0;
        }
        return h1.toString(16) + h2.toString(16);
    }

    const _tableHashes = new Map();
    function tableHashOf(info) {
        // 只用列名，不用列宽：滚动条出现/消失会让宽度变几像素，
        // 那是布局抖动，不是表格内容变化
        return fastHash(JSON.stringify([
            (info.columns || []).map(c => c.text), info.rows,
            info.count_in_dom, info.total_text, info.page_text
        ]));
    }
    function tableMetaOf(info) {
        return {
            kind: info.kind, selector: info.selector, virtual: info.virtual,
            rows_in_dom: info.count_in_dom, rows_hidden: info.hidden_rows,
            rows_truncated: info.truncated_rows,
            total_text: info.total_text, page_text: info.page_text,
            scroll_box: info.scroll_box
        };
    }

    function tableRefOfRoot(root, targetEl) {
        if (!root || !root.isConnected) return null;
        let info;
        try { info = buildTable(root, targetEl || null); } catch (e) { return null; }
        const h = tableHashOf(info);
        // 留一份当前状态，作为随后"操作后"对比的基准
        try { _lastDump.set(root, { columns: info.columns, rows: info.rows }); } catch (e) {}
        const ref = { table_hash: h, table_meta: tableMetaOf(info), table_target: info.target };
        if (_tableHashes.get(root) !== h) {
            _tableHashes.set(root, h);
            ref.table_full = { columns: info.columns, rows: info.rows };
        }
        return ref;
    }

    function tableRef(el, deep) {
        const root = tableRootOf(el, deep);
        if (!root) return null;
        // 见过的表就持续盯着它：即使它不是组件库标准类名，也能捕获滚动加载/轮询刷新
        try { if (!_knownTables.has(root)) _knownTables.add(root); } catch (e) {}
        const ref = tableRefOfRoot(root, deep ? null : el);
        return ref ? { ref: ref, root: root } : null;
    }

    // ============ 操作前后对比：AI 靠它判断"我为什么点这里/输入这个" ============
    const _lastDump = new Map();   // root -> {columns, rows}
    function rowIdent(r, i) { return r.key || ('#idx' + i); }
    function diffRows(before, after) {
        const bKeys = before.map(rowIdent);
        const aKeys = after.map(rowIdent);
        const bMap = new Map(before.map((r, i) => [bKeys[i], r]));
        const aMap = new Map(after.map((r, i) => [aKeys[i], r]));
        let added = 0, removed = 0, changed = 0;
        const samples = [];
        for (let i = 0; i < before.length; i++) {
            const r = before[i], r2 = aMap.get(bKeys[i]);
            if (!r2) { removed++; continue; }
            const c1 = r.cells || [], c2 = r2.cells || [];
            if (JSON.stringify(c1) !== JSON.stringify(c2)) {
                changed++;
                if (samples.length < 5) {
                    for (let j = 0; j < Math.max(c1.length, c2.length); j++) {
                        if (c1[j] !== c2[j]) {
                            samples.push({ row: bKeys[i], col: j, from: c1[j], to: c2[j] });
                            break;
                        }
                    }
                }
            }
        }
        for (let i = 0; i < after.length; i++) if (!bMap.has(aKeys[i])) added++;
        const reordered = (added === 0 && removed === 0 &&
                           bKeys.join('\u0001') !== aKeys.join('\u0001'));
        // 整页换了一批行（排序 / 翻页 / 筛选）跟"多几行少几行"是两回事，得区分开
        const pageReplaced = (before.length > 0 && after.length > 0 &&
                              added === after.length && removed === before.length);
        return { added: added, removed: removed, changed: changed,
                 reordered: reordered, page_replaced: pageReplaced,
                 rows_before: before.length, rows_after: after.length,
                 samples: samples };
    }

    function noteReady(a) {
        // 一个动作可能要等两件事：HTML 快照 + 表格前后对比；都完了才能让 Python 取走
        try {
            if (!a) return;
            a._pending = (a._pending || 1) - 1;
            if (a._pending <= 0) { delete a._pending; a._ready = true; }
        } catch (e) {}
    }

    const _settlePending = new Map();   // root -> {idx, timer}
    function settleTable(root, idx) {
        // 上一步的"操作后"还没结算，这次动作就紧接着发生了：
        // 把变化归到本次动作，避免把后面几步的效果算到上一步头上
        const prev = _settlePending.get(root);
        if (prev) {
            clearTimeout(prev.timer);
            const pa = window._recordedActions[prev.idx];
            if (pa && !pa._settled) {
                pa._settled = true;
                pa.table_diff = { merged_into_next: true };
                noteReady(pa);
            }
        }
        // 等页面把这次操作的结果渲染完，再采一次表格状态
        const timer = setTimeout(() => {
            const cur = _settlePending.get(root);
            if (cur && cur.idx === idx) _settlePending.delete(root);
            const a = window._recordedActions[idx];
            try {
                if (!a || a._settled) return;
                a._settled = true;
                if (!root || !root.isConnected) return;
                const info = buildTable(root, null);
                const h = tableHashOf(info);
                a.table_after_hash = h;
                if (h === a.table_hash) {
                    a.table_diff = { unchanged: true, added: 0, removed: 0, changed: 0 };
                } else {
                    const last = _lastDump.get(root);
                    if (last) a.table_diff = diffRows(last.rows, info.rows);
                    const out = { hash: h, meta: tableMetaOf(info),
                                  trigger: _scrollSinceAction ? 'scroll' : 'action' };
                    if (_tableHashes.get(root) !== h) {
                        _tableHashes.set(root, h);
                        out.full = { columns: info.columns, rows: info.rows };
                    }
                    a.table_after = out;
                    // 这次变化已经归到本动作上了，别让后面的轮询刷新被错标成"滚动加载"
                    _scrollSinceAction = false;
                }
                _lastDump.set(root, { columns: info.columns, rows: info.rows });
            } catch (e) {} finally {
                noteReady(a);
            }
        }, 700);
        _settlePending.set(root, { idx: idx, timer: timer });
    }

    function pushRaw(obj) {
        try { window._recordedActions.push(obj); } catch (e) {}
        return obj;
    }

    let _lastActionTs = 0;
    let _lastHeaderCol = null;     // 最近点过的表头列（用来把筛选面板里的动作关联到列）

    // 表格内容自己变了（滚到底加载 / 定时轮询刷新）也要留痕，否则 AI 看不到数据是怎么变多的
    const _knownTables = new Set();
    function _hasRegisteredAncestor(el) {
        let p = el.parentElement, g = 0;
        while (p && g++ < 20) {
            if (_knownTables.has(p)) return true;
            p = p.parentElement;
        }
        return false;
    }
    function registerTables() {
        // 不能只认组件库类名：自研 div 表格（.table-box / .grid 之类）也要收进来
        const sels = [TABLE_BOX_SEL, '[class*="table"]', '[class*="grid"]', '[class*="datagrid"]', '[class*="listview"]'];
        let n = 0;
        for (const sel of sels) {
            let list = [];
            try { list = Array.from(document.querySelectorAll(sel)); } catch (e) { continue; }
            for (const t of list) {
                if (n++ > 100) return;
                if (_knownTables.has(t) || _hasRegisteredAncestor(t)) continue;
                try { if (ownRows(t).length >= 2) _knownTables.add(t); } catch (e) {}
            }
        }
    }
    setInterval(registerTables, 2000);

    let _tableWatchTimer = null;
    let _lastScrollTs = 0;
    let _scrollSinceAction = false;   // 滚动之后、尚未出现其它动作 → 这次表格变化就是滚动加载引起的
    function onTableMutated() {
        if (_tableWatchTimer) return;
        _tableWatchTimer = setTimeout(() => {
            _tableWatchTimer = null;
            if (Date.now() - _lastActionTs < 600) return;   // 刚有用户动作，那个动作自己已经带上表格状态了
            _knownTables.forEach(root => {
                if (!root.isConnected) { _knownTables.delete(root); return; }
                let info;
                try { info = buildTable(root, null); } catch (e) { return; }
                const h = tableHashOf(info);
                const prev = _tableHashes.get(root);
                if (prev === undefined) { _tableHashes.set(root, h); return; }
                if (prev === h) return;
                _tableHashes.set(root, h);
                _lastDump.set(root, { columns: info.columns, rows: info.rows });
                pushRaw({
                    type: 'table_update',
                    trigger: _scrollSinceAction ? 'scroll' : 'auto',
                    timestamp: new Date().toISOString(),
                    url: location.href, title: document.title,
                    scroll: { x: window.scrollX, y: window.scrollY },
                    table_hash: h, table_meta: tableMetaOf(info),
                    table_full: { columns: info.columns, rows: info.rows },
                    _sent: false, _ready: true, _recovered: false
                });
                _scrollSinceAction = false;   // 一次滚动加载消费一个标记
            });
        }, 800);
    }
    try {
        new MutationObserver(onTableMutated).observe(document, { subtree: true, childList: true, characterData: true });
    } catch (e) {}
    registerTables();

    function recordAction(type, el, extra, opts) {
        extra = extra || {};
        opts = opts || {};
        _lastActionTs = Date.now();
        if (type === 'scroll') _scrollSinceAction = true;
        else if (type !== 'table_update') _scrollSinceAction = false;
        const wantSnap = opts.snap !== false;
        const beforeSnap = wantSnap ? captureDom() : null;
        const action = {
            type: type,
            element: getElementInfo(el),
            context: getContext(el),
            timestamp: new Date().toISOString(),
            url: location.href,
            title: document.title,
            scroll: { x: window.scrollX, y: window.scrollY },
            before_snapshot: beforeSnap,
            after_snapshot: null,
            dom_changed: null,
            _sent: false,
            _ready: false,
            _recovered: false
        };
        let tRefRoot = null;
        try {
            const tr = tableRef(el, !!opts.deep);
            if (tr) {
                tRefRoot = tr.root;               // DOM 元素不能进动作体，否则序列化会炸
                for (const k in tr.ref) action[k] = tr.ref[k];
            }
        } catch (e) {}
        try {
            const tgt = action.table_target;
            if (tgt && tgt.in_header) {
                _lastHeaderCol = { col_index: tgt.col_index, col_name: tgt.col_name,
                                   at: Date.now(), root: tRefRoot };
            } else if (tgt && !tgt.in_header && type === 'click') {
                _lastHeaderCol = null;      // 点了数据行，说明已经离开表头操作了
            }
            if (!action.table_hash) {
                // 不在表格里，但可能在一个浮层（表头筛选面板 / 下拉 / 弹窗）上
                const fp = floatingPanelOf(el);
                if (fp) {
                    action.floating_panel = true;
                    action.panel_id = fp.id;
                    action.panel_class = fp.cls;
                    action.panel_text = fp.text;
                    if (_lastHeaderCol && Date.now() - _lastHeaderCol.at < 60000) {
                        action.panel_of_column = {
                            col_index: _lastHeaderCol.col_index,
                            col_name: _lastHeaderCol.col_name
                        };
                        // 这个面板属于那张表：把表格状态也挂上，前后对比才能连到真正触发变化的那一步
                        const own = _lastHeaderCol.root;
                        const ref2 = own && own.isConnected ? tableRefOfRoot(own, null) : null;
                        if (ref2) {
                            tRefRoot = own;
                            for (const k in ref2) action[k] = ref2[k];
                        }
                    }
                }
            }
        } catch (e) {}
        for (const k in extra) action[k] = extra[k];
        const idx = window._recordedActions.push(action) - 1;
        // _ready 必须等"快照 + 表格前后对比"都完成，否则轮询会提前取走，对比数据就丢了
        let pending = 0;
        if (wantSnap) pending++;
        if (tRefRoot) pending++;
        action._pending = pending;
        if (tRefRoot) settleTable(tRefRoot, idx);
        if (!pending) {
            action._ready = true;
            return;
        }
        if (wantSnap) {
            captureStableDom((afterSnap) => {
                try {
                    const a = window._recordedActions[idx];
                    if (!a) return;
                    a.after_snapshot = afterSnap;
                    if (beforeSnap === null || afterSnap === null) {
                        a.dom_changed = null;
                        a.snapshot_skipped = true;
                        try { a.dom_element_count = document.getElementsByTagName('*').length; } catch (e) {}
                    } else {
                        a.dom_changed = afterSnap !== beforeSnap;
                    }
                    noteReady(a);
                } catch (e) {}
            }, 400, 2500);
        }
    }

    document.addEventListener('click', e => recordAction('click', e.target, { button: e.button }), true);
    document.addEventListener('dblclick', e => recordAction('dblclick', e.target, {}), true);
    document.addEventListener('contextmenu', e => recordAction('contextmenu', e.target, {}), true);
    document.addEventListener('change', e => {
        const el = e.target;
        if (el.tagName === 'INPUT' && el.type === 'file') {
            const files = Array.from(el.files || []).map(f => f.name);
            recordAction('file_upload', el, { files: files });
        } else if (['INPUT','TEXTAREA','SELECT'].includes(el.tagName)) {
            const ty = (el.type || '').toLowerCase();
            const isCheck = (ty === 'checkbox' || ty === 'radio');
            recordAction('change', el, {
                // 勾选框的 el.value 恒为 'on'，没信息量，换成可读的选中状态
                value: isCheck ? (el.checked ? '选中' : '取消选中') : sanitizeValue(el, el.value),
                checked: el.checked
            });
        }
    }, true);
    document.addEventListener('submit', e => recordAction('submit', e.target, {}), true);
    document.addEventListener('keydown', e => {
        if (['Enter','Tab','Escape'].includes(e.key)) recordAction('keydown', e.target, { key: e.key });
    }, true);
    document.addEventListener('dragstart', e => recordAction('dragstart', e.target, {}), true);
    document.addEventListener('drop', e => recordAction('drop', e.target, {}), true);

    // ---- 滚动：无限滚动 / 懒加载表格靠它触发加载，必须留痕 ----
    const _scrollBoxes = new Set();          // 已知的滚动容器
    const _scrollPos = new Map();            // 上次观察到的位置（用来拿到"滚动前"的位置）
    const _scrollPending = new Map();        // el -> 本轮位移的起点/终点
    const _scrollTimers = new Map();
    function noteScrollPos(el) {
        try { _scrollPos.set(el, { top: el.scrollTop, left: el.scrollLeft }); } catch (e) {}
    }
    function collectScrollBoxes() {
        const se = document.scrollingElement || document.documentElement;
        _scrollBoxes.add(se); noteScrollPos(se);
        try {
            document.querySelectorAll('[class*="body"],[class*="wrap"],[class*="scroll"],[class*="viewport"],[class*="table"],[class*="grid"],[style*="overflow"]').forEach(el => {
                try {
                    if (el.clientHeight > 40 && el.scrollHeight > el.clientHeight + 10) {
                        _scrollBoxes.add(el);
                        noteScrollPos(el);
                    }
                } catch (e) {}
            });
        } catch (e) {}
    }
    setInterval(collectScrollBoxes, 1200);
    collectScrollBoxes();

    document.addEventListener('scroll', e => {
        const t = e.target;
        const el = (t && t.nodeType === 1) ? t : (document.scrollingElement || document.documentElement);
        if (!el) return;
        _lastScrollTs = Date.now();      // 立刻标记，别等去抖结束（表格加载往往在滚动结束后 300ms 内就发生）
        _scrollSinceAction = true;       // 这一次滚动之后发生的表格变化都算"滚动加载"
        // 拖滚动条 / PageDown / 程序设置滚动位置都是"一次跳到位"：
        // 第一个 scroll 事件时位置已经变了，所以起算点必须用上一次观察到的位置，
        // 否则会算出位移 0 而把整次滚动丢掉
        const prev = _scrollPos.get(el);
        let pend = _scrollPending.get(el);
        if (!pend) {
            pend = prev ? { fromTop: prev.top, fromLeft: prev.left }
                        : { fromTop: null, fromLeft: null };
            _scrollPending.set(el, pend);
        }
        pend.lastTop = el.scrollTop;
        pend.lastLeft = el.scrollLeft;
        const bottomNow = (el.scrollTop + el.clientHeight) >= (el.scrollHeight - 40);
        // 只要本轮里"曾经"到底就算触发了加载（加载出新数据后位置就不在底部了）
        pend.hit_bottom = pend.hit_bottom || bottomNow;
        pend.end_at_bottom = bottomNow;
        pend.height = el.scrollHeight;
        pend.clientHeight = el.clientHeight;
        noteScrollPos(el);
        _scrollBoxes.add(el);
        clearTimeout(_scrollTimers.get(el));
        _scrollTimers.set(el, setTimeout(() => {
            try {
                const p = _scrollPending.get(el);
                _scrollPending.delete(el);
                if (!p) return;
                const unknown = (p.fromTop === null);
                const fromTop = unknown ? p.lastTop : p.fromTop;
                const fromLeft = unknown ? p.lastLeft : p.fromLeft;
                const dy = unknown ? null : Math.round(p.lastTop - fromTop);
                const dx = unknown ? null : Math.round(p.lastLeft - fromLeft);
                const movedY = (dy === null) ? true : Math.abs(dy) >= 40;
                const movedX = (dx === null) ? true : Math.abs(dx) >= 40;
                if (!movedY && !movedX) return;   // 抖动，不记
                // 起点未知且最后停在 0/0：多半是点击前的 scrollIntoView 或布局抖动，没信息量
                if (unknown && p.lastTop === 0 && p.lastLeft === 0 && !p.hit_bottom) return;
                _lastScrollTs = Date.now();
                const ddy = p.lastTop - fromTop, ddx = p.lastLeft - fromLeft;
                recordAction('scroll', el, {
                    delta_x: dx, delta_y: dy,
                    from_unknown: unknown,
                    axis: Math.abs(ddy) >= Math.abs(ddx) ? 'y' : 'x',
                    dir: unknown ? null : (ddy > 0 ? 'down' : (ddy < 0 ? 'up' : (ddx > 0 ? 'right' : 'left'))),
                    from_top: Math.round(fromTop),
                    from_left: Math.round(fromLeft),
                    scroll_top: Math.round(p.lastTop),
                    scroll_left: Math.round(p.lastLeft),
                    scroll_height: p.height,
                    client_height: p.clientHeight,
                    at_bottom: !!p.hit_bottom,
                    end_at_bottom: !!p.end_at_bottom
                }, { snap: false, deep: true });
            } catch (e2) {}
        }, 450));
    }, true);

    // ---- 边输边筛：change 只在回车/失焦才触发，很多人输了字直接点结果，必须补 input ----
    const _inputTimers = new Map();
    function isTextInput(el) {
        if (!el || el.nodeType !== 1) return false;
        if (el.tagName === 'TEXTAREA') return true;
        if (el.tagName !== 'INPUT') return false;
        const ty = (el.type || 'text').toLowerCase();
        return ['text', 'search', 'password', 'email', 'number', 'tel', 'url'].includes(ty);
    }
    document.addEventListener('input', e => {
        const el = e.target;
        if (!isTextInput(el)) return;
        clearTimeout(_inputTimers.get(el));
        _inputTimers.set(el, setTimeout(() => {
            _inputTimers.delete(el);
            recordAction('input', el, { value: sanitizeValue(el, el.value), checked: el.checked }, { snap: false });
        }, 700));
    }, true);
    document.addEventListener('compositionend', e => {
        const el = e.target;
        if (!isTextInput(el)) return;
        clearTimeout(_inputTimers.get(el));
        _inputTimers.delete(el);
        // 中文输入法：这时 value 才是最终汉字，不然会记成拼音
        recordAction('input', el, { value: sanitizeValue(el, el.value), ime: true }, { snap: false });
    }, true);

    let lastUrl = location.href;
    function checkUrlChange() {
        if (location.href !== lastUrl) {
            window._recordedActions.push({
                type: 'navigation', from: lastUrl, to: location.href,
                timestamp: new Date().toISOString(), _sent: false, _ready: true
            });
            lastUrl = location.href;
        }
    }
    try { new MutationObserver(checkUrlChange).observe(document, { subtree: true, childList: true }); } catch (e) {}
    window.addEventListener('popstate', checkUrlChange);
    window.addEventListener('hashchange', checkUrlChange);

    const _nativeWindowOpen = window.open;
    let _pageCurrentOpen = window.open;

    function _recorderWrappedOpen(...args) {
        try {
            window._recordedActions.push({
                type: 'window_open',
                args: args.map(String),
                timestamp: new Date().toISOString(),
                _sent: false, _ready: true
            });
        } catch (e) {}

        let target = _pageCurrentOpen;
        if (typeof target !== 'function') target = _nativeWindowOpen;
        if (typeof target !== 'function') return null;

        try {
            return target.apply(window, args);
        } catch (e) {
            return null;
        }
    }
    window.open = _recorderWrappedOpen;

    setInterval(function() {
        if (window.open !== _recorderWrappedOpen) {
            _pageCurrentOpen = window.open;
            window.open = _recorderWrappedOpen;
        }
    }, 1000);

    console.log('[Recorder] 监听器已注入:', location.href);
})();
"""

# ======================== 采集 JS（IIFE） ========================
COLLECT_JS = """
(() => {
    var MAX_N = %d;
    function collectNew(win) {
        var arr = [];
        try {
            var list = win._recordedActions || [];
            for (var i = 0; i < list.length; i++) {
                if (arr.length >= MAX_N) return arr;
                var a = list[i];
                if (!a._sent && a._ready) {
                    a._sent = true;
                    var copied = {};
                    for (var k in a) copied[k] = a[k];
                    arr.push(copied);
                    a.before_snapshot = null;
                    a.after_snapshot = null;
                }
            }
        } catch (e) {}
        return arr;
    }
    return collectNew(window);
})()
""" % MAX_ACTIONS_PER_POLL

# ======================== 工具函数 ========================
def parse_iso(ts):
    try:
        return datetime.fromisoformat(ts.replace('Z', '+00:00')).timestamp()
    except Exception:
        return None

# ======================== 全局状态 ========================
all_actions = []
page_labels = {}    # id(page) -> "tabN"
page_meta = {}      # id(page) -> {...}

def get_page_label(pg):
    key = id(pg)
    if key not in page_labels:
        page_labels[key] = f"tab{len(page_labels)+1}"
    return page_labels[key]

def get_page_meta(pg):
    return page_meta.get(id(pg))

def set_page_meta(pg, meta):
    page_meta[id(pg)] = meta

# ======================== 增量落盘 / 看门狗 ========================
_partial_file = None
_partial_idx = 0

def open_partial():
    """每采集一批就追加写盘：收尾阶段出任何事都不会丢掉已采集的记录。"""
    global _partial_file, _partial_idx
    try:
        _partial_file = open(PARTIAL_JSONL, 'w', encoding='utf-8')
        _partial_idx = 0
        print(f"🛟 增量备份：{PARTIAL_JSONL}")
    except OSError as e:
        print(f"⚠️  无法创建增量备份文件：{e}")

def flush_partial():
    global _partial_idx
    if _partial_file is None:
        return
    try:
        for a in all_actions[_partial_idx:]:
            _partial_file.write(json.dumps(a, ensure_ascii=False) + '\n')
        _partial_idx = len(all_actions)
        _partial_file.flush()
    except Exception as e:
        print(f"⚠️  增量写入失败：{e}")

_watchdog = {'armed': False}

def start_watchdog(seconds, message):
    """兜底：关闭浏览器等操作可能永久阻塞，超时直接强杀进程。"""
    _watchdog['armed'] = True

    def _run():
        time.sleep(seconds)
        if not _watchdog['armed']:
            return          # 已经顺利关完了，别再把进程杀掉
        try:
            sys.stdout.flush()
        except Exception:
            pass
        print(f"\n⚠️  {message}（超过 {int(seconds)}s），强制退出。")
        os._exit(0)
    threading.Thread(target=_run, daemon=True).start()

# ======================== 启动 Playwright ========================
pw_driver = sync_playwright().start()

context = None
try:
    context = pw_driver.chromium.launch_persistent_context(
        user_data_dir=USER_DATA_DIR,
        channel="msedge",
        headless=False,
        viewport=None,
        args=["--start-maximized"] + _extra_args
    )
except Exception as e:
    print(f"⚠️  msedge 通道启动失败：{e}")
    print("   尝试改用 chromium...")
    try:
        context = pw_driver.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=False,
            viewport=None,
            args=_extra_args or None
        )
    except Exception as e2:
        print(f"❌ 启动浏览器失败：{e2}")
        try:
            pw_driver.stop()
        except Exception:
            pass
        sys.exit(1)

context.add_init_script(MONITOR_JS)

def on_new_page(pg):
    label = get_page_label(pg)
    try:
        url = pg.url or ""
    except Exception:
        url = ""
    set_page_meta(pg, {
        "first_seen": time.time(),
        "first_url": url,
        "label": label
    })
    print(f"   [新标签页] {label}")

context.on("page", on_new_page)

if context.pages:
    page = context.pages[0]
else:
    page = context.new_page()

if get_page_meta(page) is None:
    get_page_label(page)
    set_page_meta(page, {
        "first_seen": time.time(),
        "first_url": "",
        "label": page_labels[id(page)]
    })

if TARGET_URL:
    try:
        page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=30000)
    except Exception as e:
        print(f"⚠️  导航失败: {e}")
else:
    print("ℹ️  没有指定网址：浏览器已打开空白页，请在浏览器里自己输入网址开始操作。")
    print("   （下次也可以带网址启动：python WebRecorder.py \"https://...\"）")

time.sleep(1.5)

try:
    page.evaluate(MONITOR_JS)
except Exception as e:
    print(f"⚠️  首次注入失败: {e}")

meta = get_page_meta(page)
if meta is not None:
    try:
        meta["first_url"] = page.url
    except Exception:
        pass

print(f"✅ 已开始录制：{TARGET_URL or '（等你手动输入网址）'}")
print(f"[OUTPUT] 本次输出目录：{RUN_DIR}")
print(f"📁 Profile 目录：{USER_DATA_DIR}（登录态会保留）")
print("👉 在浏览器里正常操作（含新标签页）。")
print("⌨️  结束方式：回到【这个窗口】输入 c 再按回车（也可以按 Ctrl+C）。收尾完成后窗口不会自动关。")

# ======================== 结束信号 ========================
_stop = threading.Event()       # 收到结束指令
_exit_now = threading.Event()   # 收尾完毕后用户要退出
STOP_WORDS = ('c', 'q', 'x', 'stop', 'done', 'exit', '结束', '完成', '停止')

def _watch_console():
    """后台读控制台输入：录制中输入 c 回车 = 结束；收尾后再输入任意内容 = 退出。
    这样双击运行的用户不必依赖 Ctrl+C。"""
    try:
        for line in sys.stdin:
            w = (line or '').strip().lower()
            if _stop.is_set():
                _exit_now.set()
                return
            if w in STOP_WORDS:
                print(f"\n⏹️  收到结束指令（{line.strip()}），开始收尾...")
                _stop.set()
                continue
            if w:
                print("   （未识别的指令：输入 c 再按回车 = 结束录制）")
    except Exception:
        pass

threading.Thread(target=_watch_console, daemon=True).start()

# ======================== 采集函数 ========================
def collect_all():
    total = 0
    try:
        pages = list(context.pages)
    except Exception:
        return 0
    for pg in pages:
        try:
            if pg.is_closed():
                continue
            if get_page_meta(pg) is None:
                label = get_page_label(pg)
                try:
                    url = pg.url or ""
                except Exception:
                    url = ""
                set_page_meta(pg, {
                    "first_seen": time.time(),
                    "first_url": url,
                    "label": label
                })
            else:
                m = get_page_meta(pg)
                if not m.get("first_url"):
                    try:
                        m["first_url"] = pg.url or ""
                    except Exception:
                        pass

            items = pg.evaluate(COLLECT_JS) or []
            label = get_page_label(pg)
            for it in items:
                it['_window'] = label
                bf = save_snapshot(it.get('before_snapshot'))
                af = save_snapshot(it.get('after_snapshot'))
                it['before_snapshot_file'] = bf
                it['after_snapshot_file'] = af
                it.pop('before_snapshot', None)
                it.pop('after_snapshot', None)
                save_table_payload(it)
            all_actions.extend(items)
            total += len(items)
        except PlaywrightError:
            pass
        except Exception:
            pass
    return total

# ======================== 主循环 ========================
open_partial()

try:
    while True:
        if _stop.is_set():
            break
        total_this_cycle = 0
        while True:
            n = collect_all()
            total_this_cycle += n
            if n < MAX_ACTIONS_PER_POLL:
                break
        flush_partial()
        if DEBUG and total_this_cycle:
            print(f"   [捕获] +{total_this_cycle} 条，累计 {len(all_actions)}")
        time.sleep(POLL_INTERVAL)
except KeyboardInterrupt:
    pass
print(f"\n⏹️  收尾中...（最多再等 {int(DRAIN_SECONDS)} 秒，请勿重复按 Ctrl+C）")

# 收尾采集：整体限时；万一开始卡住，再按一次 Ctrl+C 会跳到下面的保存步骤而不是直接退出
try:
    _deadline = time.time() + DRAIN_SECONDS
    while time.time() < _deadline:
        n = collect_all()
        flush_partial()
        if n < MAX_ACTIONS_PER_POLL:
            break
        time.sleep(0.1)
except KeyboardInterrupt:
    print("   ⚠️  收尾被再次中断，立刻保存已采集的数据")
except Exception as e:
    print(f"   ⚠️  收尾采集出错：{e}")

# 从这一刻起忽略 Ctrl+C：数据已经齐了，不能因为再按一下而丢掉
signal.signal(signal.SIGINT, signal.SIG_IGN)
print("   （保存/关闭阶段已忽略 Ctrl+C，稍等即可自动退出）")

flush_partial()
print(f"💾 准备写入 {len(all_actions)} 条动作...")

# ======================== 关联新标签页来源 ========================
for key, meta in page_meta.items():
    ts_open = meta["first_seen"]
    best = None
    best_dt = WINDOW_ASSOC_WINDOW
    for a in all_actions:
        if a.get('type') not in ('click', 'window_open', 'submit'):
            continue
        at = parse_iso(a.get('timestamp', ''))
        if at is None:
            continue
        dt = ts_open - at
        if 0 <= dt < best_dt:
            best_dt = dt
            best = a
    if best is not None:
        meta["opened_by_step"] = best.get('step')

# ======================== 输出 JSONL ========================
with open(OUTPUT_JSONL, 'w', encoding='utf-8') as f:
    for i, action in enumerate(all_actions, 1):
        action['step'] = i
        action.pop('_sent', None)
        action.pop('_ready', None)
        action.pop('_pending', None)
        action.pop('_settled', None)
        action['window_label'] = action.pop('_window', 'tab1')
        for key, meta in page_meta.items():
            if meta['label'] == action['window_label']:
                if meta.get('opened_by_step'):
                    action['window_opened_by_step'] = meta['opened_by_step']
                break
        f.write(json.dumps(action, ensure_ascii=False) + '\n')

print(f"📄 JSONL 已保存：{OUTPUT_JSONL}（{len(all_actions)} 条）")
print(f"📁 快照：{SNAPSHOT_DIR}/（{len(snapshot_map)} 个去重文件）")
print(f"📄 表格快照：{TABLES_DIR}/（{len(table_map)} 种表格状态）")
print(f"🗂️  本次运行目录：{RUN_DIR}")

def _short(s, n=42):
    s = ' '.join(str(s or '').split())
    return (s[:n] + '…') if len(s) > n else s

def _target_str(a):
    el = a.get('element') or {}
    if a.get('type') in ('navigation', 'window_open'):
        return str(a.get('to') or a.get('args') or '')
    if not el.get('tagName'):
        return ''
    if a.get('type') == 'scroll':
        return f"滚动容器 <{el.get('tagName')}> {a.get('dir') or '方向未知'}"
    tag = el.get('tagName', '')
    ident = el.get('id') or next((c for c in (el.get('className_stable') or []) if len(c) >= 5), '')
    text = _short((el.get('text') or '').split('\n')[0], 18)
    return ' '.join(x for x in (f"<{tag}>", ident, text) if x)

def _pos_str(a):
    t = a.get('table_target')
    if not t:
        return ''
    col = t.get('col_name') or f"列#{t.get('col_index')}"
    if t.get('in_header'):
        return f"表头「{col}」"
    return f"行#{t.get('row_index_in_dom')}「{col}」{t.get('row_key') or ''}"

def _diff_str(a):
    d = a.get('table_diff')
    if not d:
        return ''
    if d.get('merged_into_next'):
        return '变化并入下一步'
    if d.get('unchanged'):
        return '表格无变化'
    bits = []
    if d.get('page_replaced'):
        bits.append('整页换行')
    else:
        if d.get('added'):
            bits.append(f"+{d['added']}行")
        if d.get('removed'):
            bits.append(f"-{d['removed']}行")
    if d.get('changed'):
        bits.append(f"{d['changed']}行改")
    if d.get('reordered'):
        bits.append('重排')
    if d.get('rows_before') is not None:
        bits.append(f"{d.get('rows_before')}→{d.get('rows_after')}行")
    return '、'.join(bits)

def write_ai_index():
    """写 AI_INDEX.md：把“怎么读这个目录”讲清楚，减少 AI 猜错/漏读。"""
    from collections import Counter
    types = Counter(a.get('type', '?') for a in all_actions)
    inter = [a for a in all_actions if a.get('type') in ('click', 'change', 'input', 'keydown')]
    with_tgt = [a for a in inter if a.get('table_target')]
    panel = [a for a in all_actions if a.get('floating_panel')]
    scrolls = [a for a in all_actions if a.get('type') == 'scroll']
    skipped = [a for a in all_actions if a.get('snapshot_skipped')]
    diffs = [a for a in all_actions
             if a.get('table_diff') and not a['table_diff'].get('merged_into_next')]

    states = {}
    for i, a in enumerate(all_actions, 1):
        for h in (a.get('table_hash'), (a.get('table_after') or {}).get('hash')):
            if h and h not in states:
                states[h] = {'first': i, 'rows': None, 'cols': None}
    for h in list(states):
        try:
            info = json.load(open(os.path.join(TABLES_DIR, f"table_{h}.json"), encoding='utf-8'))
            states[h]['rows'] = len(info.get('rows') or [])
            states[h]['cols'] = len(info.get('columns') or [])
        except Exception:
            pass

    L = []
    L.append("# 浏览器操作录制 — AI 阅读入口\n")
    L.append(f"> 录制目录：`{os.path.basename(RUN_DIR)}`　起始地址：{TARGET_URL or '（启动时未指定，用户手动输入）'}")
    L.append(f"> 录制时间：{RUN_START:%Y-%m-%d %H:%M:%S} ~ {datetime.now():%Y-%m-%d %H:%M:%S}")
    L.append("> **请先读本文件**，再按需打开下表列出的附属文件。\n")

    L.append("## 1. 概要\n")
    L.append(f"- 动作 {len(all_actions)} 条：" + '、'.join(f"{k}×{v}" for k, v in types.most_common()))
    L.append(f"- 涉及标签页 {len(page_labels)} 个；表格状态快照 {len(states)} 种；页面快照 {len(snapshot_map)} 张")
    L.append(f"- 交互动作能对上表格行列的：{len(with_tgt)}/{len(inter)}" +
             (f"；带表格前后对比的动作 {len(diffs)} 条" if diffs else ''))
    L.append(f"- 浮层（筛选面板/弹窗）内动作 {len(panel)} 条；滚动动作 {len(scrolls)} 条\n")

    L.append("## 2. 文件清单与用途\n")
    L.append("| 文件 | 内容 | 什么时候读 |")
    L.append("|---|---|---|")
    L.append("| `actions_readable.txt` | 逐步可读记录（目标元素/行列定位/表格上下文/前后对比/浮层关联） | **先读这个** |")
    L.append("| `actions.jsonl` | 同样内容，机读版，字段最全 | 需要精确字段或程序化处理时 |")
    L.append("| `tables/table_*.json` / `.tsv` | 表格在某一时刻的**完整内容**（列名 + 全部行 + 行标识） | 想知道某步前后表格的实际数据时（动作里的 `table_file` / `table_after.file` 指向它） |")
    L.append("| `snapshots/snap_*.html` | 操作前后的整页 HTML（已去脚本/样式） | 需核对页面结构或元素层级时 |")
    L.append("| `replay_<时间>/` | 用 replay.py 重放过的日志与截图 | 重放排查时 |")
    L.append("\n> 表格内容按**内容哈希去重**：同一状态只存一份，多条动作可能引用同一个 `table_*.json`。若只需要重点内容，读 `actions_readable.txt` 已足够；需要表格全量数据时再打开 `tables/`。\n")

    L.append("## 3. 步骤总览（一行一步）\n")
    L.append("行号/列号为页面 DOM 中的 0 起序号；带「」的是列名。\n")
    L.append("| # | 类型 | 页 | 目标 | 值 | 表格位置 | 表格变化 |")
    L.append("|---|---|---|---|---|---|---|")
    for i, a in enumerate(all_actions, 1):
        val = a.get('value')
        if val is None:
            val = f"按键 {a.get('key')}" if a.get('key') else ''
        L.append(f"| {i} | {a.get('type')} | {a.get('window_label') or ''} | {_short(_target_str(a), 34)} | "
                 f"{_short(val, 22)} | {_short(_pos_str(a), 30)} | {_diff_str(a)} |")
    L.append('')

    if states:
        L.append("## 4. 表格状态索引\n")
        L.append("| 快照文件 | 行数 | 列数 | 首次出现 |")
        L.append("|---|---|---|---|")
        for h, s in sorted(states.items(), key=lambda kv: kv[1]['first']):
            L.append(f"| `tables/table_{h}.json`（同名 .tsv 可直接看表） | {s['rows']} | {s['cols']} | 第 {s['first']} 步 |")
        L.append('')

    L.append("## 5. 已知信息缺口（请不要假设这里有）\n")
    L.append("- **没有“用户意图/目的”字段**：需要从页面标题、URL、动作序列与前后对比来推断“为什么这么点”")
    L.append("- **密码/敏感字段被脱敏为 `***`**（设计如此）；重放不依赖录制里的密码，登录态靠持久化 profile")
    L.append("- **只记最终值**：输入记的是去抖后/提交时的值，不记逐字过程；中文输入法记的是最终汉字")
    L.append("- **文字划选与拖拽过程未记录**（只记拖拽起止）")
    if skipped:
        L.append(f"- **{len(skipped)} 条动作跳过了整页 HTML 快照**（页面元素数超过 6000，避免卡页面；表格结构化内容仍然完整）")
    else:
        L.append("- 本次没有跳过 HTML 快照")
    L.append("- **虚拟滚动表格只能看到 DOM 内可见的那些行**（浏览器自己也只有这些），滚动过程中每个状态会各存一份")
    L.append("- 同一时刻只记一张表；侧边栏/弹窗里的表若被操作也会各记一份\n")

    with open(OUTPUT_INDEX, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')
    print(f"🧭 AI 阅读入口（发给 AI 时先看这个）：{OUTPUT_INDEX}")
    return OUTPUT_INDEX

def write_table_section(f, action):
    """把动作携带的表格上下文 / 浮层面板信息写成 AI 能直接读明白的一段。"""
    meta = action.get('table_meta')
    if action.get('floating_panel'):
        f.write("\n### 浮层面板（表头筛选/下拉/弹窗）\n\n")
        f.write(f"- 面板内容: `{action.get('panel_text', '')}`\n")
        pid, pcls = action.get('panel_id') or '', action.get('panel_class') or ''
        if pid or pcls:
            f.write(f"- 面板元素: id=`{pid}` class=`{pcls}`\n")
        col = action.get('panel_of_column')
        if col:
            f.write(f"- 关联列: 第 {col.get('col_index')} 列「{col.get('col_name')}」"
                    f"（依据：之前点过该列的表头筛选按钮）\n")
        else:
            f.write("- 关联列: 未知（无法确定这个面板属于哪一列）\n")
    if not meta:
        return
    tgt = action.get('table_target') or {}
    f.write("\n### 表格上下文\n\n")
    kind = '原生 HTML 表格' if meta.get('kind') == 'html' else 'div 表格'
    f.write(f"- 表格: `{meta.get('selector', '')}`（{kind}"
            f"{'，虚拟滚动' if meta.get('virtual') else ''}）\n")
    line = f"- 行数（DOM 内）: {meta.get('rows_in_dom')}"
    if meta.get('rows_hidden'):
        line += f"（其中 {meta['rows_hidden']} 行被隐藏/筛掉）"
    if meta.get('rows_truncated'):
        line += f"（原 {meta['rows_truncated']} 行，已截断保存）"
    f.write(line + "\n")
    if meta.get('total_text'):
        f.write(f"- 总条数: {meta['total_text']}\n")
    if meta.get('page_text'):
        f.write(f"- 当前页: {meta['page_text']}\n")
    sb = meta.get('scroll_box') or {}
    if sb.get('selector'):
        f.write(f"- 滚动容器: `{sb['selector']}`（scrollTop {sb.get('scroll_top')}/"
                f"{sb.get('scroll_height')}，可视高 {sb.get('client_height')}）\n")
    if tgt:
        f.write(f"- 目标位置: 第 {tgt.get('row_index_in_dom')} 行、第 {tgt.get('col_index')} 列"
                f"「{tgt.get('col_name')}」里的 <{tgt.get('tag')}>，单元格文本=`{tgt.get('cell_text')}`\n")
        f.write(f"- 行标识（重放时用来确认是同一行）: `{tgt.get('row_key')}`\n")
        f.write(f"- 整行内容: `{tgt.get('row_text')}`\n")
    if action.get('table_file'):
        f.write(f"- 表格快照（动作发生时）: `{action['table_file']}`（同目录同名 .tsv 可直接看表）\n")

    after = action.get('table_after') or {}
    diff = action.get('table_diff') or {}
    if after or diff:
        f.write("\n**这次操作让表格发生了什么**\n\n")
        if diff.get('unchanged'):
            f.write("- 表格内容没有变化\n")
        elif diff.get('merged_into_next'):
            f.write("- 紧接着又发生了下一步操作，本次变化归到那一步（避免误归因）\n")
        else:
            parts = []
            if diff.get('page_replaced'):
                parts.append("整页数据被替换（排序 / 翻页 / 筛选）")
            else:
                if diff.get('added'):
                    parts.append(f"新增 {diff['added']} 行")
                if diff.get('removed'):
                    parts.append(f"减少 {diff['removed']} 行")
            if diff.get('changed'):
                parts.append(f"{diff['changed']} 行内容改变")
            if diff.get('reordered'):
                parts.append("行顺序变化（排序/重排）")
            f.write(f"- 变化摘要: {'、'.join(parts) if parts else '无'}")
            if diff.get('rows_before') is not None:
                f.write(f"（表格行数 {diff.get('rows_before')} → {diff.get('rows_after')}）")
            f.write("\n")
            trig = after.get('trigger')
            if trig == 'scroll':
                f.write("- 变化来源: 滚动触发的加载（懒加载/无限滚动）\n")
            elif trig == 'action':
                f.write("- 变化来源: 本次操作自身触发\n")
            for s in (diff.get('samples') or []):
                f.write(f"  - 行 `{s.get('row')}` 第 {s.get('col')} 列: "
                        f"`{s.get('from', '')}` → `{s.get('to', '')}`\n")
        if after.get('hash'):
            afile = after.get('file') or ''
            tail = f"（{afile}）" if afile else ''
            f.write(f"- 操作后表格快照: `table_{after['hash']}.json`{tail}\n")

# ======================== 输出 AI 友好 TXT ========================
with open(OUTPUT_TXT, 'w', encoding='utf-8') as f:
    f.write("# 用户手动操作记录（Playwright 版）\n\n")
    f.write(f"- 起始 URL: {TARGET_URL or '（启动时未指定，由用户手动输入）'}\n")
    f.write(f"- 录制时间: {RUN_START:%Y-%m-%d %H:%M:%S} ~ {datetime.now():%Y-%m-%d %H:%M:%S}\n")
    f.write(f"- 运行目录: `{RUN_DIR}`\n")
    f.write(f"- 总步数: {len(all_actions)}\n")
    f.write(f"- 涉及标签页: {len(page_labels)} 个\n")
    f.write(f"- 快照目录: `{SNAPSHOT_NAME}/`\n")
    f.write(f"- 表格快照目录: `{TABLES_NAME}/`（{len(table_map)} 种表格状态，.tsv 与 .json 同名）\n\n")

    if page_labels:
        f.write("## 标签页概览\n\n")
        for key, label in page_labels.items():
            meta = page_meta.get(key, {})
            f.write(f"- **{label}**: 首次出现于 {meta.get('first_url', '(未知)')}")
            if meta.get('opened_by_step'):
                f.write(f"（由步骤 {meta['opened_by_step']} 触发打开）")
            f.write("\n")
        f.write("\n---\n\n")

    for i, action in enumerate(all_actions, 1):
        atype = action.get('type', '?')
        el = action.get('element') or {}
        ctx = action.get('context') or {}
        win = action.get('window_label', '?')

        f.write(f"## 步骤 {i}: {atype}  [{win}]\n\n")
        f.write(f"- 时间: {action.get('timestamp', '')}\n")
        f.write(f"- URL: {action.get('url', '')}\n")
        f.write(f"- 标题: {action.get('title', '')}\n")
        f.write(f"- 滚动: ({action.get('scroll', {}).get('x', 0)}, {action.get('scroll', {}).get('y', 0)})\n")
        if action.get('_recovered'):
            f.write(f"- ⚠️ 该步骤从导航前备份恢复\n")
        if action.get('window_opened_by_step'):
            f.write(f"- 本标签页由步骤 {action['window_opened_by_step']} 打开\n")

        if atype == 'navigation':
            f.write(f"- 跳转: `{action.get('from')}` → `{action.get('to')}`\n\n---\n\n")
            continue
        if atype == 'window_open':
            f.write(f"- 打开新窗口: {action.get('args')}\n\n---\n\n")
            continue
        if atype == 'scroll':
            if action.get('from_unknown'):
                move = "位移未知（首次滚动就跳到位，未能取到起点）"
            else:
                move = f"位移 Δx={action.get('delta_x')} Δy={action.get('delta_y')}"
            f.write(f"- 滚动: 方向 {action.get('dir') or '未知'}，{move}，"
                    f"位置 {action.get('scroll_top')}/{action.get('scroll_height')}"
                    f"（可视高 {action.get('client_height')}）"
                    f"{'，曾到底→可能触发加载' if action.get('at_bottom') else ''}\n")
            write_table_section(f, action)
            f.write("\n---\n\n")
            continue
        if atype == 'table_update':
            trig = '滚动加载' if action.get('trigger') == 'scroll' else '页面自动刷新（轮询/推送）'
            f.write(f"- 表格内容自行变化（{trig}）：数据变了，用户在此时并未操作\n")
            write_table_section(f, action)
            f.write("\n---\n\n")
            continue

        if el.get('tagName'):
            f.write(f"\n### 目标元素\n\n")
            f.write(f"- 标签: `<{el.get('tagName')}>`\n")
            if el.get('id'): f.write(f"- ID: `{el['id']}`\n")
            if el.get('className_stable'): f.write(f"- 稳定 class: `{' '.join(el['className_stable'])}`\n")
            if el.get('className_dynamic'): f.write(f"- 动态 class（勿用）: `{' '.join(el['className_dynamic'])}`\n")
            if el.get('text'): f.write(f"- 文本: `{el['text'][:120]}`\n")
            if el.get('value'): f.write(f"- 值: `{el['value']}`\n")
            if el.get('in_shadow_dom'): f.write(f"- ⚠️ 位于 Shadow DOM 内\n")
            f.write(f"- XPath: `{el.get('xpath', '')}`\n")
            f.write(f"- CSS: `{el.get('css', '')}`\n")
            if el.get('attributes'):
                attrs = ', '.join(f"{k}={v}" for k, v in el['attributes'].items())
                f.write(f"- 属性: {attrs}\n")
            if el.get('rect'):
                r = el['rect']
                f.write(f"- 位置: ({r['x']}, {r['y']}) 尺寸: {r['w']}x{r['h']}\n")

        f.write(f"\n### 上下文\n\n")
        f.write(f"- 在弹窗内: {'是' if ctx.get('in_dialog') else '否'}")
        if ctx.get('dialog_selector'): f.write(f"（{ctx['dialog_selector']}）")
        f.write("\n")
        f.write(f"- 在表格内: {'是' if ctx.get('in_table') else '否'}")
        if ctx.get('row_index') is not None: f.write(f"（第 {ctx['row_index']} 行）")
        f.write("\n")
        f.write(f"- 在列表内: {'是' if ctx.get('in_list') else '否'}")
        if ctx.get('list_index') is not None: f.write(f"（第 {ctx['list_index']} 项）")
        f.write("\n")

        if ctx.get('parent_chain'):
            f.write(f"\n**父链（由近到远）:**\n")
            for anc in ctx['parent_chain']:
                f.write(f"  - `<{anc['tag']}` id=`{anc['id']}` class=`{anc['class']}`> `{anc.get('text', '')[:60]}`\n")

        if 'value' in action and atype in ('change', 'input'):
            f.write(f"\n### 输入内容\n\n- 值: `{action['value']}`"
                    f"{'（中文输入法完成后的最终值）' if action.get('ime') else ''}\n")
        if 'key' in action:
            f.write(f"\n### 按键\n\n- Key: `{action['key']}`\n")
        if 'files' in action:
            f.write(f"\n### 上传文件\n\n- 文件: {action['files']}\n")
        if 'button' in action and atype == 'click':
            f.write(f"\n- 鼠标按钮: {action['button']}\n")

        write_table_section(f, action)

        f.write(f"\n### DOM 变化\n\n")
        if action.get('snapshot_skipped'):
            f.write(f"- 快照已跳过：页面元素数 {action.get('dom_element_count')} 超过阈值"
                    f"（大表格页面整页克隆会卡顿），表格内容见上方「表格上下文」\n")
        else:
            f.write(f"- 有变化: {'是' if action.get('dom_changed') else '否'}\n")
        if action.get('before_snapshot_file'):
            f.write(f"- 操作前快照: `{SNAPSHOT_NAME}/{action['before_snapshot_file']}`\n")
        if action.get('after_snapshot_file'):
            f.write(f"- 操作后快照: `{SNAPSHOT_NAME}/{action['after_snapshot_file']}`\n")

        f.write("\n---\n\n")

print(f"📝 可读 TXT 已保存：{OUTPUT_TXT}")

# ======================== 录制覆盖率自检 ========================
# 换到真实网站时先看这段：比例异常就说明页面结构没被识别出来
def _ratio(a, b):
    return f"{a}/{b}" + (f"（{a * 100 // b}%）" if b else '')

def print_coverage():
    from collections import Counter
    types = Counter(a.get('type', '?') for a in all_actions)
    inter = [a for a in all_actions if a.get('type') in ('click', 'change', 'input', 'keydown')]
    with_tgt = [a for a in inter if a.get('table_target')]
    panel = [a for a in all_actions if a.get('floating_panel')]
    panel_col = [a for a in panel if a.get('panel_of_column')]
    scrolls = [a for a in all_actions if a.get('type') == 'scroll']
    skipped = [a for a in all_actions if a.get('snapshot_skipped')]
    hashes = {a.get('table_hash') for a in all_actions if a.get('table_hash')}
    hashes |= {(a.get('table_after') or {}).get('hash') for a in all_actions
               if isinstance(a.get('table_after'), dict)}
    hashes.discard(None)
    diffs = [a for a in all_actions
             if a.get('table_diff') and not a['table_diff'].get('merged_into_next')]
    print("\n📊 录制覆盖率自检（换新网站时先看这里）")
    print(f"   - 动作 {len(all_actions)} 条：" + '、'.join(f"{k}×{v}" for k, v in types.most_common()))
    print(f"   - 交互动作能对上表格行列的：{_ratio(len(with_tgt), len(inter))}")
    print(f"   - 表格状态快照：{len(hashes)} 种（{TABLES_NAME}/，含操作前后）")
    print(f"   - 带前后对比的动作：{len(diffs)} 条")
    print(f"   - 浮层（筛选面板/弹窗）内动作：{len(panel)} 条，其中关联到列 {len(panel_col)} 条")
    print(f"   - 滚动动作：{len(scrolls)} 条（其中曾到底 {sum(1 for a in scrolls if a.get('at_bottom'))} 条）")
    if skipped:
        print(f"   - 因页面元素过多跳过快照：{len(skipped)} 条（表格内容不受影响）")
    if inter and len(with_tgt) * 2 < len(inter):
        print("   ⚠️  带行列定位的比例偏低：多半是表格结构没被识别——")
        print("      把页面网址 + 表格最外层那几行 HTML 发我，我按实际结构补识别规则")

print_coverage()
try:
    write_ai_index()
except Exception as e:
    print(f"⚠️  AI 索引生成失败（不影响其它产物）：{e}")

# ======================== 退出 ========================
if _partial_file is not None:
    try:
        _partial_file.close()
    except Exception:
        pass
try:
    os.remove(PARTIAL_JSONL)   # 正式文件写好了，增量备份就没用了
    print(f"🧹 已清理增量备份：{PARTIAL_JSONL}")
except OSError:
    pass

start_watchdog(CLOSE_TIMEOUT, "关闭浏览器无响应")

print("🔒 正在关闭浏览器...")
try:
    context.close()
    pw_driver.stop()
    _watchdog['armed'] = False          # 关完了，别让看门狗再杀进程
except KeyboardInterrupt:
    print("   ⚠️  关闭被打断，记录已保存")
    _watchdog['armed'] = False
except Exception as e:
    print(f"⚠️  退出时出错（记录已保存）：{e}")
    _watchdog['armed'] = False

# 保存已完成：恢复 Ctrl+C（它的作用变成"退出本程序"）
try:
    signal.signal(signal.SIGINT, signal.default_int_handler)
except Exception:
    pass

print("")
print("=" * 68)
print("✅ 已完成：本次录制全部写完，浏览器也已关闭。")
print(f"   产物目录：{RUN_DIR}")
print("   建议先把这个目录整个给 AI 看（先读里面的 AI_INDEX.md）。")
print("   现在可以安全关闭窗口：按 Ctrl+C，或在本窗口输入任意内容再回车。")
print("=" * 68)

# 等用户看完再退出（双击运行时窗口不会一闪而过）
try:
    while not _exit_now.is_set():
        time.sleep(0.2)
except KeyboardInterrupt:
    pass

print("👋 已退出。")
try:
    sys.stdout.flush()
finally:
    os._exit(0)   # 保证命令行立刻回到提示符，不留后台线程