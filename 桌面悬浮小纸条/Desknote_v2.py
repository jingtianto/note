import sys
import os
import json
import math
import signal
import subprocess
from pathlib import Path
import pyperclip
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QMenu, QDialog, QSpinBox, QSizePolicy, QFrame, QMessageBox,
    QCheckBox,
    QPlainTextEdit, QRadioButton, QLineEdit, QFileDialog, QGridLayout
)
from PySide6.QtCore import (
    Qt, QTimer, QPoint, QPointF, QRect, QRectF, QObject, Signal,
    QPropertyAnimation
)
from PySide6.QtGui import (
    QCursor, QColor, QFont, QFontMetrics, QPainter, QPen, QBrush,
    QPainterPath, QPixmap
)

# ===================== 在这里修改你的10条常用文本 =====================
SNIPPETS = [
    ("文本1：你的第一段常用内容", "#FFB6C1"),
    ("文本2：你的第二段常用内容", "#FFD6A5"),
    ("文本3：你的第三段常用内容", "#FDFFB6"),
    ("文本4：你的第四段常用内容", "#CAFFBF"),
    ("文本5：你的第五段常用内容", "#9BF6FF"),
    ("文本6：你的第六段常用内容", "#A0C4FF"),
    ("文本7：你的第七段常用内容", "#BDB2FF"),
    ("文本8：你的第八段常用内容", "#FFC6FF"),
    ("文本9：你的第九段常用内容", "#FFADAD"),
    ("文本10：你的第十段常用内容", "#C8E7FF"),
]
# =====================================================================

BAR_WIDTH = 4                     # 竖线默认粗细
BAR_HEIGHT = 160                  # 竖线的可视高度
BAR_HIT_WIDTH = 44                # 竖线窗口总宽（含两侧隐形可点击区）
BAR_PAD_TOP = 38                  # 竖线上方留白（放 ✕ 退出按钮）
BAR_PAD_BOTTOM = 16               # 竖线下方留白
LINE_Y = BAR_PAD_TOP              # 竖线在窗口里的纵坐标

LEAVE_DELAY = 0.15        # 鼠标移开后多久收起面板
PANEL_SLIDE = 3           # 滑出动画距离（最终停在紧贴竖线的位置）
DRAG_THRESHOLD = 4        # 按下后移动超过这么多像素才算“拖动”
GRAB_MARGIN = 8           # 面板四周的隐形可拖动边（比可视区域大一圈）

ICON_COL_WIDTH = 22       # 纸条左侧“颜色”列和“编辑”列的宽度
COLOR_DOT_SIZE = 12       # 颜色小圆点的直径
NOTE_MAX_TEXT_WIDTH = 800 # 文字超过这个宽度就折行
NOTE_MIN_TEXT_WIDTH = 40
NOTE_PAD_LEFT = 0         # 文字紧贴色块左边缘（“顶住左边界”）
NOTE_PAD_RIGHT = 8        # 文字与 copy/run 徽标之间的空隙
NOTE_V_PADDING = 0        # 纸条上下内边距（0 = 上下不留白，最紧凑）

# 纸条显示内容的三个上限：任意一个达标就在末尾加省略号。
# 只影响“显示”，完整原文仍在 text_content 里（复制、编辑不受影响）。
# 默认值在这里，实际以设置里的为准。
NOTE_MAX_LINES = 6        # 最多显示几行
NOTE_MAX_CHARS = 180      # 最多显示几个字符
NOTE_MAX_TEXT_HEIGHT = 120  # 文字区最高多少像素
# 纸条末尾的 copy / run 小徽标
BADGE_MARGIN = 6          # 徽标与色块右边缘之间的空隙
BADGE_COPY_COLOR = "#4A90D9"
BADGE_RUN_COLOR = "#3FAE6A"
BADGE_COPY_TEXT = "copy"
BADGE_RUN_TEXT = "run"

# 色盘：一圈固定颜色的扇块
# 前 12 个是彩色的（挑得柔和一点，配深色文字都看得清），
# 后 4 个是灰色系，方便把纸条/竖线调回朴素的白灰。
SECTOR_COLORS = [
    "#FF6B6B",  # 珊瑚红
    "#FF9F68",  # 蜜橙
    "#FFC46B",  # 琥珀
    "#FFE066",  # 暖黄
    "#C9E86A",  # 青柠
    "#7BD88F",  # 嫩绿
    "#4FD1C5",  # 薄荷
    "#4BC0E8",  # 天青
    "#5B8DEF",  # 宝蓝
    "#9B8CFF",  # 薰衣草
    "#C77DFF",  # 紫罗兰
    "#FF8FC7",  # 樱粉
    "#FFFFFF",  # 白
    "#D9D9D9",  # 浅灰
    "#9E9E9E",  # 中灰
    "#555555",  # 深灰（竖线默认色）
]
PICKER_SIZE = 126         # 色盘直径（比之前的 168 小一圈）
PICKER_INNER = 0.40       # 内圈半径占外圈的比例
PICKER_GAP = 1.5          # 扇块之间的缝隙（度）
PICKER_GROW = 5           # 鼠标悬停时扇块向外突出的像素

MAX_SNIPPETS = 20         # 最多能放多少条

# 顶部那一条（提示文字 / ⚙ / ✕）的固定配色：白底黑字，朴素高对比
HEADER_BG = "#FFFFFF"
HEADER_FG = "#1F1F1F"
HEADER_ICON = "#5A5A66"
HEADER_RADIUS = 8          # 顶部圆角
BAR_RADIUS = 6             # 所有长条色块的圆角
HEADER_HEIGHT = 26         # 标题栏高度（提示气泡用同一高度才能和它平齐）

# 标题栏上方那条“剪贴板纸条”：黑边框 + 淡黄底
CLIP_BG = "#FFF8D6"        # 淡黄纸色
CLIP_BORDER = "#1F1F1F"    # 黑色边框
CLIP_PAD = 8               # 左右内边距
CLIP_MIN_TEXT_WIDTH = 60   # 文字区最小宽度
CLIP_MAX_TEXT_WIDTH = 320  # 文字区最大宽度（再宽就折行）
CLIP_MAX_CHARS = 600       # 文字太长只显示前面一段
CLIP_MAX_TEXT_HEIGHT = 240   # 文字部分最高多少（超出就再截短）
CLIP_MAX_IMAGE_HEIGHT = 240  # 图片缩放后的最大高度
CLIP_MAX_SOURCE_SIDE = 1200  # 原图超过这个边长就先缩下来再留着
CLIP_EMPTY_TEXT = "剪贴板：空"

DEFAULT_SETTINGS = {
    "panel_title": "快捷文本",   # 面板顶部那句提示文字
    "show_clipboard": True,   # 是否显示标题栏上方的剪贴板纸条
    "font_size": 14,
    "letter_spacing": 0,     # 字间距（px，可为负）
    "line_spacing": 100,     # 行距（%）
    "note_max_lines": NOTE_MAX_LINES,          # 纸条最多显示几行
    "note_max_chars": NOTE_MAX_CHARS,          # 纸条最多显示几个字
    "note_max_text_height": NOTE_MAX_TEXT_HEIGHT,  # 纸条文字区最高多少
    "bar_width": BAR_WIDTH,
    "bar_color": "#555555",
    "bar_opacity": 100,      # 竖线透明度（%）
    "note_opacity": 100,     # 纸条透明度（%）
    "panel_offset_y": 0,     # 纸条顶端比竖线顶端低多少像素
    "bar_x": -1,             # 竖线上次待的位置（-1 = 用默认位置）
    "bar_y": -1,
}


def app_dir():
    """配置文件放哪：源码运行放脚本旁边，打包成 exe 后放 exe 旁边。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def app_stem():
    """配置文件名跟着脚本（或 exe）的名字跑。

    这样把脚本复制成 Desknote2.py 再跑，就会用 Desknote2.snippets.json /
    Desknote2.settings.json，两个实例互不干扰。
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).stem
    return Path(__file__).stem


# 你编辑过的内容会保存到这个文件，下次启动自动读取
_DATA_NAME = f"{app_stem()}.snippets.json"
_SETTINGS_NAME = f"{app_stem()}.settings.json"
DATA_FILE = app_dir() / _DATA_NAME
SETTINGS_FILE = app_dir() / _SETTINGS_NAME

# 旧版用的固定文件名，首次运行自动搬过来
_LEGACY_FILES = (
    (app_dir() / "snippets.json", DATA_FILE),
    (app_dir() / "settings.json", SETTINGS_FILE),
)


def migrate_legacy_config():
    """如果只有旧版固定名的配置文件，就把内容搬到新的文件名下。"""
    for legacy, target in _LEGACY_FILES:
        if legacy.exists() and not target.exists():
            try:
                target.write_text(
                    legacy.read_text(encoding="utf-8"), encoding="utf-8"
                )
            except Exception as exc:
                # 旧文件可能是别的编码：搬不动就算了，绝不能因此启动不了
                print(f"[警告] 迁移配置失败 {legacy}：{exc}")


def load_snippets():
    """优先读取 snippets.json，没有就用上面的默认值。

    每条是 {text, color, action, target}：
    - action = "copy" → 点击复制 text 到剪贴板
    - action = "run"  → 点击用系统默认方式打开 target（相当于双击它）
    旧版只有 text/color 的文件会自动补上 action="copy"。
    """
    if DATA_FILE.exists():
        try:
            data = json.loads(DATA_FILE.read_text(encoding="utf-8"))
            items = []
            for entry in data:
                text = str(entry.get("text", "")).strip()
                if not text:
                    continue
                action = str(entry.get("action", "copy"))
                if action not in ("copy", "run"):
                    action = "copy"
                items.append({
                    "text": text,
                    "color": str(entry.get("color", "#FFD6A5")),
                    "action": action,
                    "target": str(entry.get("target", "")),
                })
            if items:
                return items
        except Exception:
            pass
    return [
        {"text": t, "color": c, "action": "copy", "target": ""}
        for t, c in SNIPPETS
    ]


def _write_json(path, payload):
    """写配置文件；打包成 exe 后目录可能不可写，失败不慌张地报个错。"""
    try:
        path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return True
    except Exception as exc:
        # 存配置失败最多是这次改动没保存，绝不能把整个程序带崩
        print(f"[警告] 保存失败 {path}：{exc}")
        return False


def save_snippets(items):
    return _write_json(DATA_FILE, items)


def run_target(target):
    """用系统默认方式打开程序 / 文件 / 文件夹 / 网址，相当于双击它。"""
    target = (target or "").strip()
    if not target:
        return False
    try:
        if hasattr(os, "startfile"):
            os.startfile(target)          # Windows
        elif sys.platform == "darwin":
            subprocess.Popen(["open", target])
        else:
            subprocess.Popen(["xdg-open", target])
        return True
    except Exception as exc:
        print(f"[警告] 打开失败 {target}：{exc}")
        return False


def load_settings():
    settings = dict(DEFAULT_SETTINGS)
    if SETTINGS_FILE.exists():
        try:
            data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                for key, default in DEFAULT_SETTINGS.items():
                    if key in data:
                        settings[key] = type(default)(data[key])
        except Exception:
            pass

    # 范围保护，防止配置文件被改坏
    settings["font_size"] = min(48, max(8, settings["font_size"]))
    settings["bar_width"] = min(12, max(2, settings["bar_width"]))
    settings["bar_opacity"] = min(100, max(10, settings["bar_opacity"]))
    settings["note_opacity"] = min(100, max(10, settings["note_opacity"]))
    settings["panel_offset_y"] = min(200, max(0, settings["panel_offset_y"]))
    settings["letter_spacing"] = min(12, max(-2, settings["letter_spacing"]))
    settings["line_spacing"] = min(250, max(80, settings["line_spacing"]))
    settings["note_max_lines"] = min(50, max(1, settings["note_max_lines"]))
    settings["note_max_chars"] = min(2000, max(20, settings["note_max_chars"]))
    settings["note_max_text_height"] = min(
        600, max(20, settings["note_max_text_height"])
    )
    if not str(settings["panel_title"]).strip():
        settings["panel_title"] = DEFAULT_SETTINGS["panel_title"]
    settings["show_clipboard"] = bool(settings["show_clipboard"])
    return settings


def save_settings(settings):
    return _write_json(SETTINGS_FILE, settings)


class BarWidget(QWidget):
    """竖线。

    注意：这个窗口本身就是一块“实心”的隐形矩形（BAR_HIT_WIDTH × 总高），
    只有中间 4px 的线是看得见的。这样鼠标在竖线周围的生效区里移动、点击，
    都会落在本程序上，不会穿透点到下层的桌面图标或其他软件。
    """

    # 拖动后通知主程序，让面板跟着走
    moved = Signal()
    # 点了 ✕ 或右键菜单里的“退出程序”
    quit_requested = Signal()

    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            Qt.FramelessWindowHint
            | Qt.Tool
            | Qt.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(
            BAR_HIT_WIDTH, BAR_PAD_TOP + BAR_HEIGHT + BAR_PAD_BOTTOM
        )
        self.move(20, 240)

        self.drag_position = None
        self.thickness = DEFAULT_SETTINGS["bar_width"]
        self.color = DEFAULT_SETTINGS["bar_color"]

        # 看得见的那条竖线
        self.line = QLabel(self)

        # 悬停时才出现的 ✕（窗口内的子控件，不是独立窗口）
        self.close_button = QPushButton("✕", self)
        self.close_button.setGeometry(11, 6, 22, 22)
        self.close_button.setCursor(Qt.PointingHandCursor)
        self.close_button.setToolTip("退出程序")
        self.close_button.setStyleSheet("""
            QPushButton {
                background-color: #FF6B6B;
                color: white;
                border: none;
                border-radius: 11px;
                font-size: 12px;
                font-weight: 700;
            }
            QPushButton:hover {
                background-color: #D93B3B;
            }
        """)
        self.close_button.clicked.connect(
            lambda: self.quit_requested.emit()
        )
        self.close_button.hide()
        self.apply_style(self.thickness, self.color)

    def line_left(self):
        """竖线在本窗口内的左边缘 x。"""
        return (BAR_HIT_WIDTH - self.thickness) // 2

    def line_right(self):
        """竖线在本窗口内的右边缘 x（面板紧贴着它）。"""
        return self.line_left() + self.thickness

    def apply_style(self, thickness, color, opacity=100):
        """设置竖线的粗细、颜色和透明度（设置面板里改）。

        透明度用背景色的 alpha 实现，所以只影响竖线本身，
        不影响 ✕ 按钮等其他内容。
        """
        self.thickness = max(2, min(12, int(thickness)))
        self.color = color
        self.opacity = max(10, min(100, int(opacity)))

        rgb = QColor(color)
        alpha = round(self.opacity * 255 / 100)
        x = self.line_left()
        self.line.setGeometry(x, LINE_Y, self.thickness, BAR_HEIGHT)
        self.line.setStyleSheet(
            f"background-color: rgba({rgb.red()}, {rgb.green()},"
            f" {rgb.blue()}, {alpha});"
            f" border-radius: {self.thickness // 2}px;"
        )
        self.close_button.setGeometry(
            x + self.thickness // 2 - 11, 6, 22, 22
        )

    def paintEvent(self, event):
        """把整块窗口填成 alpha=1 的隐形底色。

        Windows 对分层窗口是按像素做命中测试的：alpha=0 的像素会把鼠标
        直接放行给下层窗口（也就是“点击穿透”）。填成 1/255 既看不出来，
        系统又认为它不透明，于是整块生效区都能接住鼠标。
        """
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor(0, 0, 0, 1))

    def show_close(self, visible):
        if visible == self.close_button.isVisible():
            return
        self.close_button.setVisible(visible)
        if visible:
            self.close_button.raise_()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_position = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if self.drag_position and event.buttons() == Qt.LeftButton:
            self.move(event.globalPosition().toPoint() - self.drag_position)
            self.moved.emit()
            event.accept()

    def mouseReleaseEvent(self, event):
        # 松手就把拖动状态清干净，不让它残留到下一次
        self.drag_position = None
        super().mouseReleaseEvent(event)

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        act_quit = menu.addAction("✕  退出程序")
        if menu.exec(event.globalPos()) == act_quit:
            self.quit_requested.emit()


class DragMixin:
    """按下不动 = 点击；按住拖动 = 移动整张纸条。

    子类需要自己声明 drag_started / drag_moved / drag_finished 三个信号，
    并可覆盖 on_drag_click() 来定义“点击”行为。
    """

    def _init_drag(self):
        self._press_pos = None
        self._dragging = False

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._press_pos = event.globalPosition().toPoint()
            self._dragging = False
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._press_pos is None or not (event.buttons() & Qt.LeftButton):
            super().mouseMoveEvent(event)
            return

        pos = event.globalPosition().toPoint()
        if not self._dragging:
            moved = (pos - self._press_pos).manhattanLength()
            if moved < DRAG_THRESHOLD:
                return
            self._dragging = True
            self.drag_started.emit(self._press_pos)
        self.drag_moved.emit(pos)
        event.accept()

    def mouseReleaseEvent(self, event):
        if self._press_pos is None or event.button() != Qt.LeftButton:
            super().mouseReleaseEvent(event)
            return

        dragged = self._dragging
        self._press_pos = None
        self._dragging = False
        if dragged:
            self.drag_finished.emit()
        else:
            self.on_drag_click()
        event.accept()

    def on_drag_click(self):
        """按下后没怎么移动就松开时调用，子类按需覆盖。"""


class HandleButton(QPushButton):
    """排序手柄（三条横线）：按住上下拖动就能调整纸条顺序。

    QPushButton 默认会吃掉鼠标事件，所以这里自己转发按下/移动/松开。
    """

    def __init__(self, on_press, on_move, on_release):
        super().__init__("≡")
        self._on_press = on_press
        self._on_move = on_move
        self._on_release = on_release
        self._pressed = False
        self.setFixedWidth(ICON_COL_WIDTH)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        self.setCursor(Qt.OpenHandCursor)
        self.setToolTip("按住上下拖动可调整顺序")

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._pressed = True
            self.setCursor(Qt.ClosedHandCursor)
            self._on_press(event.globalPosition().toPoint())
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._pressed and (event.buttons() & Qt.LeftButton):
            self._on_move(event.globalPosition().toPoint())
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._pressed and event.button() == Qt.LeftButton:
            self._pressed = False
            self.setCursor(Qt.OpenHandCursor)
            self._on_release(event.globalPosition().toPoint())
            event.accept()
            return
        super().mouseReleaseEvent(event)


class SnippetNote(DragMixin, QLabel):
    """一条彩色纸条。

    - 文字紧贴色块左边缘，末尾带一个 copy / run 小徽标
    - 宽度随文字长度自适应，超过上限就折行；高度随行数增长
    - 字号 / 字间距 / 行距 / 透明度都由设置控制
    - 单击 = 执行动作（复制或打开）；按住拖动 = 移动整张纸条
    """

    clicked = Signal()
    drag_started = Signal(QPoint)
    drag_moved = Signal(QPoint)
    drag_finished = Signal()

    def __init__(self, text, color, action, font_size,
                 alpha=255, letter_spacing=0, line_spacing=100):
        super().__init__()
        self._init_drag()
        self.text_content = text
        self.color = color
        self.action = action
        self.font_size = font_size
        self.alpha = alpha
        self.letter_spacing = letter_spacing
        self.line_spacing = line_spacing
        self._right_pad = 0
        self._wrapping = False
        # 三个显示上限（行数 / 字符数 / 高度），由设置决定
        self.caps = (NOTE_MAX_LINES, NOTE_MAX_CHARS, NOTE_MAX_TEXT_HEIGHT)
        self._display_ready = False

        self.setWordWrap(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setTextInteractionFlags(Qt.NoTextInteraction)
        self.setTextFormat(Qt.RichText)

        # 末尾的小徽标（对鼠标透明，点击照样落在纸条上）
        self.badge = QLabel(self)
        self.badge.setAlignment(Qt.AlignCenter)
        self.badge.setAttribute(Qt.WA_TransparentForMouseEvents, True)

        self.apply(text, color, action, font_size)

    # ---------- 外观 ----------
    def _font(self, size):
        font = QFont()
        font.setPixelSize(size)
        font.setLetterSpacing(QFont.AbsoluteSpacing, self.letter_spacing)
        return font

    def _html(self, text):
        """转成富文本：这样才能用 line-height 调行距。"""
        body = (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\n", "<br>")
        )
        return (
            f'<div style="line-height:{self.line_spacing}%;'
            f' white-space:pre-wrap">{body}</div>'
        )

    def _badge_metrics(self, font_size):
        bfont = QFont()
        bfont.setPixelSize(max(8, font_size - 5))
        bfont.setBold(True)
        fm = QFontMetrics(bfont)
        text = BADGE_RUN_TEXT if self.action == "run" else BADGE_COPY_TEXT
        return text, bfont, fm.horizontalAdvance(text) + 10, fm.height() + 2

    def _style(self, bg, right_pad):
        rgb = QColor(bg)
        return (
            f"background-color: rgba({rgb.red()}, {rgb.green()},"
            f" {rgb.blue()}, {self.alpha});"
            f" color: #1F1F1F;"          # 文字固定不透明，调透明度不影响它
            f" padding: {NOTE_V_PADDING}px {right_pad}px"
            f" {NOTE_V_PADDING}px {NOTE_PAD_LEFT}px;"
        )

    def set_alpha(self, alpha, refit=True):
        """改背景透明度。文字颜色始终不透明，不受影响。

        refit=False 时只记下数值，留给调用方自己重排一次。
        """
        self.alpha = max(0, min(255, int(alpha)))
        if refit and self._display_ready:
            self.apply(
                self.text_content, self.color, self.action, self.font_size
            )

    def on_drag_click(self):
        self.clicked.emit()

    # ---------- 尺寸 ----------
    def apply(self, text, color, action, font_size,
              letter_spacing=None, line_spacing=None):
        self.text_content = text          # 完整原文：复制、编辑都用它
        self.color = color
        self.action = action
        self.font_size = font_size
        if letter_spacing is not None:
            self.letter_spacing = letter_spacing
        if line_spacing is not None:
            self.line_spacing = line_spacing

        self.setFont(self._font(font_size))

        # 徽标先量出来，右边要给它留位置
        badge_text, badge_font, bw, bh = self._badge_metrics(font_size)
        self._right_pad = NOTE_PAD_RIGHT + bw + BADGE_MARGIN
        self.setStyleSheet(self._style(color, self._right_pad))

        self._fit_display()
        self._layout_badge(badge_text, badge_font, bw, bh, action)
        self._display_ready = True

    def set_caps(self, lines, chars, height, refit=True):
        """设置显示上限：行数 / 字符数 / 高度，任意一个达标就截断。

        设完就按新上限重排一次，不依赖调用方再 apply 一遍。
        refit=False 时只记下数值，留给调用方自己重排一次。
        """
        self.caps = (
            max(1, int(lines)), max(1, int(chars)), max(8, int(height))
        )
        if refit and self._display_ready:
            self._fit_display()

    # ---------- 显示内容的截断 ----------
    def _clip_by_count(self, text):
        """先按 行数 和 字符数 各裁一刀（只会变短，不会变长）。"""
        max_lines, max_chars, _ = self.caps
        lines = text.split("\n")
        if len(lines) > max_lines:
            text = "\n".join(lines[:max_lines])
        if len(text) > max_chars:
            text = text[:max_chars]
        return text

    def _measure(self, candidate):
        """量一段文字，返回（纸条宽, 纸条高）。"""
        # 先解开上一次的固定尺寸，否则量出来还是旧值
        self.setMinimumWidth(0)
        self.setMaximumWidth(16777215)
        self.setMinimumHeight(0)
        self.setMaximumHeight(16777215)

        self.setText(self._html(candidate))
        self.setWordWrap(False)
        natural = self.sizeHint()
        text_w = natural.width() - self._right_pad - NOTE_PAD_LEFT

        if text_w <= NOTE_MAX_TEXT_WIDTH:
            self._wrapping = False
            self.setWordWrap(False)
            width = max(
                natural.width(), NOTE_MIN_TEXT_WIDTH + self._right_pad
            )
            height = natural.height()
        else:
            self._wrapping = True
            width = NOTE_MAX_TEXT_WIDTH + self._right_pad + NOTE_PAD_LEFT
            self.setWordWrap(True)
            height = self.heightForWidth(width)
        return width, height

    def _fit_display(self):
        """定下最终显示的文字和尺寸。

        行数 / 字符数 / 高度 三个上限，**任意一个达标就截断**，
        末尾补一个省略号。注意 self.text_content 始终是完整原文。
        """
        full = self.text_content
        max_height = self.caps[2]
        base = self._clip_by_count(full)
        width, height = self._measure(base)

        if height <= max_height:
            # 高度没超：只有行数/字符数裁过才需要省略号
            text = base + " …" if base != full else base
            width, height = self._measure(text)
        else:
            # 高度超了：二分找出能放下的最长前缀
            low, high = 0, len(base)
            while low < high:
                mid = (low + high + 1) // 2
                _, h = self._measure(base[:mid] + " …")
                if h <= max_height:
                    low = mid
                else:
                    high = mid - 1
            text = base[:low].rstrip() + " …"
            width, height = self._measure(text)

        self.setText(self._html(text))
        self.setWordWrap(self._wrapping)
        self.setFixedWidth(width)
        self.setFixedHeight(height)

    def _layout_badge(self, badge_text, badge_font, bw, bh, action):
        """摆好末尾的 copy / run 徽标。"""
        badge_rgb = QColor(
            BADGE_RUN_COLOR if action == "run" else BADGE_COPY_COLOR
        )
        self.badge.setFont(badge_font)
        self.badge.setText(badge_text)
        self.badge.setStyleSheet(
            f"background-color: rgba({badge_rgb.red()},"
            f" {badge_rgb.green()}, {badge_rgb.blue()}, {self.alpha});"
            f" color: #FFFFFF; border: none;"
            f" padding: 0px;"      # 必须清掉，否则会继承纸条的 padding
            f" border-radius: {bh // 2}px;"
        )
        self.badge.setFixedSize(bw, bh)
        self.badge.move(
            self.width() - bw - BADGE_MARGIN,
            max(0, (self.height() - bh) // 2),
        )
        self.badge.show()

    def enterEvent(self, event):
        self.setStyleSheet(
            self._style(
                QColor(self.color).darker(110).name(), self._right_pad
            )
        )
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.setStyleSheet(self._style(self.color, self._right_pad))
        super().leaveEvent(event)


class ColorDotButton(QPushButton):
    """纸条左边的颜色按钮：画一个当前颜色的小圆点。"""

    def __init__(self, color):
        super().__init__()
        self.color = QColor(color)
        self._hover = False
        self.setFixedWidth(ICON_COL_WIDTH)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        self.setCursor(Qt.PointingHandCursor)
        self.setToolTip("改这条纸条的颜色")
        self.setStyleSheet(
            "QPushButton { background: transparent; border: none; }"
        )

    def apply_color(self, color):
        self.color = QColor(color)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        d = max(
            6, min(COLOR_DOT_SIZE, self.width() - 6, self.height() - 6)
        )
        rect = QRectF(
            (self.width() - d) / 2,
            (self.height() - d) / 2,
            d,
            d,
        )
        painter.setBrush(QBrush(self.color))
        pen = QPen(QColor("#4A90D9") if self._hover else QColor(0, 0, 0, 45))
        pen.setWidth(1)
        painter.setPen(pen)
        painter.drawEllipse(rect)

    def enterEvent(self, event):
        self._hover = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._hover = False
        self.update()
        super().leaveEvent(event)


class SectorColorPicker(QWidget):
    """一圈固定颜色的扇形色块；鼠标悬停的那一块会向外突出。"""

    picked = Signal(QColor)

    def __init__(self, current=None):
        super().__init__()
        self.setFixedSize(PICKER_SIZE, PICKER_SIZE)
        self.setCursor(Qt.PointingHandCursor)
        self.setMouseTracking(True)

        self.hover_index = None
        color = QColor(current) if current else QColor()
        self.current = color if color.isValid() else QColor(SECTOR_COLORS[0])

    # ---------- 几何 ----------
    def _outer_radius(self):
        return self.width() / 2 - PICKER_GROW - 2

    def _inner_radius(self):
        return self._outer_radius() * PICKER_INNER

    def _center(self):
        return QPointF(self.width() / 2, self.height() / 2)

    def _sector_path(self, index, grow=0.0):
        """第 index 个扇块的路径（环形扇块）。"""
        count = len(SECTOR_COLORS)
        span = 360.0 / count
        mid = 90.0 - index * span          # 第 0 块在正上方，顺时针排
        start = mid + span / 2 - PICKER_GAP / 2
        sweep = span - PICKER_GAP

        center = self._center()
        ro = self._outer_radius() + grow
        ri = self._inner_radius()
        outer = QRectF(center.x() - ro, center.y() - ro, ro * 2, ro * 2)
        inner = QRectF(center.x() - ri, center.y() - ri, ri * 2, ri * 2)

        path = QPainterPath()
        path.arcMoveTo(outer, start)
        path.arcTo(outer, start, -sweep)
        path.arcTo(inner, start - sweep, sweep)
        path.closeSubpath()
        return path

    def index_at(self, pos):
        """点在哪个扇块上？不在色盘上返回 None。"""
        center = self._center()
        dx = pos.x() - center.x()
        dy = pos.y() - center.y()
        dist = math.hypot(dx, dy)
        if dist < self._inner_radius():
            return None
        if dist > self._outer_radius() + PICKER_GROW:
            return None

        angle = math.degrees(math.atan2(-dy, dx)) % 360.0
        span = 360.0 / len(SECTOR_COLORS)
        return int(((90.0 - angle) % 360.0) // span) % len(SECTOR_COLORS)

    # ---------- 画 ----------
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        for index, color in enumerate(SECTOR_COLORS):
            hovered = index == self.hover_index
            painter.setBrush(QBrush(QColor(color)))
            if hovered:
                painter.setPen(QPen(QColor("#FFFFFF"), 2))
            else:
                painter.setPen(QPen(QColor(0, 0, 0, 30), 1))
            painter.drawPath(
                self._sector_path(index, PICKER_GROW if hovered else 0.0)
            )

        # 中间的圆点显示当前颜色
        radius = self._inner_radius() - 7
        if radius > 4:
            painter.setPen(QPen(QColor(0, 0, 0, 35), 1))
            painter.setBrush(QBrush(self.current))
            painter.drawEllipse(self._center(), radius, radius)

    # ---------- 交互 ----------
    def mouseMoveEvent(self, event):
        index = self.index_at(event.position())
        if index != self.hover_index:
            self.hover_index = index
            self.setToolTip(
                SECTOR_COLORS[index] if index is not None else ""
            )
            self.update()

    def leaveEvent(self, event):
        if self.hover_index is not None:
            self.hover_index = None
            self.update()
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        index = self.index_at(event.position())
        if index is not None:
            self.picked.emit(QColor(SECTOR_COLORS[index]))
            event.accept()


class ColorPopup(QDialog):
    """点颜色按钮后弹出来的选色小窗。"""

    def __init__(self, current, parent=None):
        super().__init__(parent, Qt.Popup | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.chosen = None

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        card = QFrame(self)
        card.setStyleSheet(
            "QFrame { background-color: #FFFFFF; border-radius: 10px;"
            " border: 1px solid #DDDDDD; }"
        )
        outer.addWidget(card)

        box = QVBoxLayout(card)
        box.setContentsMargins(10, 10, 10, 6)
        box.setSpacing(4)

        self.picker = SectorColorPicker(current)
        self.picker.picked.connect(self._choose)
        box.addWidget(self.picker, 0, Qt.AlignCenter)

        hint = QLabel("点扇块选颜色 · Esc 取消")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet(
            "color: #999999; font-size: 10px; border: none;"
        )
        box.addWidget(hint)

    def _choose(self, color):
        self.chosen = color
        self.accept()

    def pick(self, global_pos):
        self.adjustSize()
        screen = (
            QApplication.screenAt(global_pos)
            or QApplication.primaryScreen()
        )
        area = screen.availableGeometry()
        x = min(global_pos.x(), area.right() - self.width() - 4)
        y = min(global_pos.y(), area.bottom() - self.height() - 4)
        self.move(max(x, area.left() + 4), max(y, area.top() + 4))
        self.exec()
        return self.chosen


class SnippetRow(DragMixin, QWidget):
    """一行：[● 颜色][≡ 拖动排序][彩色纸条][✎ 编辑][✕ 删除]。

    行内和行间都不留空隙，左侧圆角归手柄、右侧圆角归 ✕。
    """

    clicked = Signal()
    drag_started = Signal(QPoint)
    drag_moved = Signal(QPoint)
    drag_finished = Signal()

    def __init__(self, index, item, settings, on_edit, on_color, on_delete):
        super().__init__()
        self._init_drag()
        self.index = index
        self.item = dict(item)
        self.settings = settings

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 颜色圆点（最左边）
        self.color_button = ColorDotButton(self.item["color"])
        self.color_button.clicked.connect(lambda: on_color(self))
        layout.addWidget(self.color_button)

        # 排序手柄（色点后面、纸条内容前面）
        self.reorder_controller = None
        self.handle_button = HandleButton(
            lambda gp: self._reorder("start", gp),
            lambda gp: self._reorder("move", gp),
            lambda gp: self._reorder("end", gp),
        )
        layout.addWidget(self.handle_button)

        # 彩色纸条本体
        self.note = SnippetNote(
            self.item["text"], self.item["color"], self.item["action"],
            settings["font_size"],
            alpha=round(settings["note_opacity"] * 255 / 100),
            letter_spacing=settings["letter_spacing"],
            line_spacing=settings["line_spacing"],
        )
        self.note.clicked.connect(self.clicked.emit)
        self.note.drag_started.connect(self.drag_started.emit)
        self.note.drag_moved.connect(self.drag_moved.emit)
        self.note.drag_finished.connect(self.drag_finished.emit)
        layout.addWidget(self.note)

        # 纸条末尾的 ✎ 编辑 和 ✕ 删除（背景跟着纸条颜色，像纸条的延伸）
        self.edit_button = QPushButton("✎")
        self.delete_button = QPushButton("✕")
        for button, tip, callback in (
            (self.edit_button, "编辑这条内容", lambda: on_edit(self)),
            (self.delete_button, "删除这条纸条", lambda: on_delete(self)),
        ):
            button.setFixedWidth(ICON_COL_WIDTH)
            button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
            button.setCursor(Qt.PointingHandCursor)
            button.setToolTip(tip)
            button.clicked.connect(callback)
        layout.addWidget(self.edit_button)
        layout.addWidget(self.delete_button)
        layout.addStretch(1)

        self.refresh()

    def _reorder(self, phase, global_pos):
        """把排序拖动的三个阶段转给面板处理。"""
        controller = self.reorder_controller
        if controller is None:
            return
        getattr(controller, f"reorder_{phase}")(self, global_pos)

    def _paint_buttons(self):
        """把 ✎ / ✕ / ≡ 的背景刷成纸条的颜色（带同样透明度）。

        这几个控件和纸条连成一条完整的长条：
        ≡ 负责左侧圆角，✕ 负责右侧圆角，中间的 ✎ 直角。
        """
        rgb = QColor(self.item["color"])
        alpha = round(self.settings["note_opacity"] * 255 / 100)
        common = (
            f"background-color: rgba({rgb.red()}, {rgb.green()},"
            f" {rgb.blue()}, {alpha}); color: #2B2B2B; border: none;"
            f" padding: 0px; font-size: 14px; font-weight: 600;"
        )
        self.edit_button.setStyleSheet(
            f"QPushButton {{ {common} border-radius: 0px; }}"
            "QPushButton:hover { background-color: #4A90D9;"
            " color: #FFFFFF; }"
        )
        # 排序手柄：跟纸条同色，左圆角由它负责
        self.handle_button.setStyleSheet(
            f"QPushButton {{ {common}"
            f" border-top-left-radius: {BAR_RADIUS}px;"
            f" border-bottom-left-radius: {BAR_RADIUS}px;"
            f" border-top-right-radius: 0px;"
            f" border-bottom-right-radius: 0px; }}"
            "QPushButton:hover { background-color: #4A90D9;"
            " color: #FFFFFF; }"
        )
        self.delete_button.setStyleSheet(
            f"QPushButton {{ {common}"
            f" border-top-left-radius: 0px; border-bottom-left-radius: 0px;"
            f" border-top-right-radius: {BAR_RADIUS}px;"
            f" border-bottom-right-radius: {BAR_RADIUS}px; }}"
            "QPushButton:hover { background-color: #E15555;"
            " color: #FFFFFF; }"
        )

    # ---------- 便捷读取当前这条的属性 ----------
    @property
    def text_content(self):
        return self.item["text"]

    @property
    def color(self):
        return self.item["color"]

    @property
    def action(self):
        return self.item["action"]

    @property
    def target(self):
        return self.item["target"]

    # ---------- 更新 ----------
    def set_item(self, item):
        self.item = dict(item)
        self.refresh()

    def apply_settings(self, settings):
        self.settings = settings
        self.refresh()

    def refresh(self):
        """按当前的 item / settings 重画这一行。"""
        self.color_button.apply_color(self.item["color"])
        self._paint_buttons()
        # 上限和透明度先记进去（refit=False 不重排），
        # 最后统一 apply 一次：一行只量一次尺寸，
        # 设置里拖滑块实时预览时 20 条纸条也不会卡。
        self.note.set_caps(
            self.settings["note_max_lines"],
            self.settings["note_max_chars"],
            self.settings["note_max_text_height"],
            refit=False,
        )
        self.note.set_alpha(
            round(self.settings["note_opacity"] * 255 / 100), refit=False
        )
        self.note.apply(
            self.item["text"], self.item["color"], self.item["action"],
            self.settings["font_size"],
            letter_spacing=self.settings["letter_spacing"],
            line_spacing=self.settings["line_spacing"],
        )


DIALOG_QSS = """
    QDialog { background: transparent; }
    QFrame#card {
        background-color: #FFFFFF;
        border: 1px solid #E4E4EA;
        border-radius: 12px;
    }
    QLabel { color: #55555F; font-size: 12px; border: none; }
    QLabel#title { color: #22222A; font-size: 14px; font-weight: 600; }
    QLabel#hint { color: #A0A0AA; font-size: 11px; }
    QSpinBox, QLineEdit, QPlainTextEdit {
        background-color: #F7F7FA;
        border: 1px solid #E1E1E8;
        border-radius: 6px;
        padding: 1px 6px;
        color: #33333A;
        font-size: 12px;
        min-height: 17px;
        selection-background-color: #4A90D9;
    }
    QSpinBox:focus, QLineEdit:focus, QPlainTextEdit:focus {
        border: 1px solid #4A90D9;
        background-color: #FFFFFF;
    }
    QSpinBox:disabled, QLineEdit:disabled {
        color: #B8B8C0; background-color: #F2F2F5;
    }
    QRadioButton { color: #44444C; font-size: 12px; spacing: 6px; }
    QCheckBox { color: #44444C; font-size: 12px; spacing: 6px; }
    QPushButton {
        background-color: #F1F1F6;
        border: 1px solid #E1E1E8;
        border-radius: 7px;
        padding: 5px 14px;
        color: #44444C;
        font-size: 12px;
    }
    QPushButton:hover { background-color: #E7E7EF; }
    QPushButton:disabled { color: #B8B8C0; background-color: #F6F6F9; }
    QPushButton#primary {
        background-color: #4A90D9;
        border: 1px solid #4A90D9;
        color: #FFFFFF;
        font-weight: 600;
    }
    QPushButton#primary:hover { background-color: #3F81C6; }
"""


def make_dialog_card(dialog, box_margins=(18, 16, 18, 14), spacing=10):
    """给对话框套一张白卡片，做出圆角 + 细边框的观感。"""
    dialog.setAttribute(Qt.WA_TranslucentBackground)
    outer = QVBoxLayout(dialog)
    outer.setContentsMargins(0, 0, 0, 0)

    card = QFrame(dialog)
    card.setObjectName("card")
    outer.addWidget(card)

    box = QVBoxLayout(card)
    box.setContentsMargins(*box_margins)
    box.setSpacing(spacing)
    dialog.setStyleSheet(DIALOG_QSS)
    return box


class SettingsDialog(QDialog):
    """设置面板：字号、字间距、行距、竖线、透明度、纸条下移。

    任何一项改动都会立刻发出 preview_changed，让主界面实时预览。
    """

    preview_changed = Signal(dict)

    def __init__(self, parent, settings):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setWindowFlags(
            Qt.FramelessWindowHint | Qt.Tool | Qt.WindowStaysOnTopHint
        )
        self.values = dict(settings)

        box = make_dialog_card(self, (16, 12, 16, 12), 7)

        title = QLabel("设置")
        title.setObjectName("title")
        box.addWidget(title)

        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(4)

        grid.addWidget(QLabel("顶部提示文字"), 0, 0)
        self.title_edit = QLineEdit(self.values["panel_title"])
        self.title_edit.setFixedWidth(140)
        self.title_edit.setPlaceholderText("例如：快捷文本")
        grid.addWidget(self.title_edit, 0, 1, Qt.AlignLeft)

        self.spin = self._add_row(
            grid, 1, "文字大小", "font_size", 8, 48, " px"
        )
        self.letter_spin = self._add_row(
            grid, 2, "字间距", "letter_spacing", -2, 12, " px"
        )
        self.line_spin = self._add_row(
            grid, 3, "行距", "line_spacing", 80, 250, " %"
        )
        # 纸条显示内容的三个上限，任一达标就出现省略号
        self.max_lines_spin = self._add_row(
            grid, 4, "纸条最多行数", "note_max_lines", 1, 50, " 行"
        )
        self.max_chars_spin = self._add_row(
            grid, 5, "纸条最多字数", "note_max_chars", 20, 2000, " 字"
        )
        self.max_height_spin = self._add_row(
            grid, 6, "纸条文字最高", "note_max_text_height", 20, 600, " px"
        )
        self.bar_spin = self._add_row(
            grid, 7, "竖线粗细", "bar_width", 2, 12, " px"
        )
        self.bar_alpha = self._add_row(
            grid, 8, "竖线透明度", "bar_opacity", 10, 100, " %"
        )
        self.note_alpha = self._add_row(
            grid, 9, "纸条透明度", "note_opacity", 10, 100, " %"
        )
        self.offset_spin = self._add_row(
            grid, 10, "纸条顶端下移", "panel_offset_y", 0, 200, " px"
        )

        grid.addWidget(QLabel("竖线颜色"), 11, 0)
        self.color_button = QPushButton()
        self.color_button.setFixedSize(58, 22)
        self.color_button.setCursor(Qt.PointingHandCursor)
        self.color_button.setToolTip("点这里选竖线颜色")
        self.color_button.clicked.connect(self.pick_bar_color)
        grid.addWidget(self.color_button, 11, 1, Qt.AlignLeft)
        self._refresh_color_button()

        self.clip_check = QCheckBox("在标题栏上方显示剪贴板纸条")
        self.clip_check.setChecked(bool(self.values["show_clipboard"]))
        grid.addWidget(self.clip_check, 12, 0, 1, 2)

        box.addLayout(grid)

        hint = QLabel("透明度只影响背景，文字始终清晰。")
        hint.setObjectName("hint")
        hint.setWordWrap(True)
        box.addWidget(hint)

        buttons = QHBoxLayout()
        buttons.addStretch()
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        ok = QPushButton("确定")
        ok.setObjectName("primary")
        ok.setDefault(True)
        ok.clicked.connect(self.accept)
        buttons.addWidget(cancel)
        buttons.addWidget(ok)
        box.addLayout(buttons)

        # 所有控件都建好之后再接信号，避免初始化时就触发预览
        self.title_edit.textChanged.connect(self._emit_preview)
        self.clip_check.toggled.connect(self._emit_preview)
        for spin in (self.spin, self.letter_spin, self.line_spin,
                     self.max_lines_spin, self.max_chars_spin,
                     self.max_height_spin, self.bar_spin, self.bar_alpha,
                     self.note_alpha, self.offset_spin):
            spin.valueChanged.connect(self._emit_preview)

    def _add_row(self, grid, row, label, key, low, high, suffix):
        grid.addWidget(QLabel(label), row, 0)
        spin = QSpinBox()
        spin.setRange(low, high)
        spin.setValue(int(self.values[key]))
        spin.setSuffix(suffix)
        spin.setFixedWidth(96)
        grid.addWidget(spin, row, 1, Qt.AlignLeft)
        return spin

    def _emit_preview(self):
        self.preview_changed.emit(self.result_settings())

    def _refresh_color_button(self):
        color = self.values["bar_color"]
        self.color_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                border: 1px solid rgba(0, 0, 0, 0.20);
                border-radius: 6px;
                padding: 0px;
            }}
            QPushButton:hover {{
                border: 2px solid #4A90D9;
            }}
        """)

    def pick_bar_color(self):
        popup = ColorPopup(self.values["bar_color"], self)
        chosen = popup.pick(self.color_button.mapToGlobal(QPoint(0, 0)))
        popup.deleteLater()        # 用完就销毁，不留给设置面板当子控件
        if chosen is not None and chosen.isValid():
            self.values["bar_color"] = chosen.name()
            self._refresh_color_button()
            self._emit_preview()

    def result_settings(self):
        """从原设置出发只覆盖改动项，这样 bar_x / bar_y 之类不会丢。"""
        values = dict(self.values)
        values.update({
            "panel_title": self.title_edit.text().strip()
            or DEFAULT_SETTINGS["panel_title"],
            "show_clipboard": self.clip_check.isChecked(),
            "font_size": self.spin.value(),
            "letter_spacing": self.letter_spin.value(),
            "line_spacing": self.line_spin.value(),
            "note_max_lines": self.max_lines_spin.value(),
            "note_max_chars": self.max_chars_spin.value(),
            "note_max_text_height": self.max_height_spin.value(),
            "bar_width": self.bar_spin.value(),
            "bar_color": self.values["bar_color"],
            "bar_opacity": self.bar_alpha.value(),
            "note_opacity": self.note_alpha.value(),
            "panel_offset_y": self.offset_spin.value(),
        })
        return values


class SnippetEditDialog(QDialog):
    """编辑一条纸条：显示文字 + 点击动作（复制 / 打开程序、文件、文件夹、网址）。"""

    def __init__(self, parent, item):
        super().__init__(parent)
        self.setWindowTitle("编辑纸条")
        self.setWindowFlags(
            Qt.FramelessWindowHint | Qt.Tool | Qt.WindowStaysOnTopHint
        )
        self.setMinimumWidth(420)

        box = make_dialog_card(self, (18, 14, 18, 12), 7)

        title = QLabel("编辑纸条")
        title.setObjectName("title")
        box.addWidget(title)

        box.addWidget(QLabel("① 纸条上显示的文字"))
        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlainText(item["text"])
        self.text_edit.setFixedHeight(70)
        box.addWidget(self.text_edit)

        box.addWidget(QLabel("② 点击这条纸条时"))
        self.copy_radio = QRadioButton("复制上面的文字到剪贴板")
        self.run_radio = QRadioButton(
            "打开下面的程序 / 文件 / 文件夹 / 网址"
        )
        if item.get("action") == "run":
            self.run_radio.setChecked(True)
        else:
            self.copy_radio.setChecked(True)
        box.addWidget(self.copy_radio)
        box.addWidget(self.run_radio)

        self.target_edit = QLineEdit(item.get("target", ""))
        self.target_edit.setPlaceholderText(
            "例如 C:\\Tools\\tool.exe   或   D:\\工作  或   https://..."
        )
        box.addWidget(self.target_edit)

        row = QHBoxLayout()
        row.setSpacing(6)
        self.browse_file_btn = QPushButton("浏览文件…")
        self.browse_file_btn.clicked.connect(self._browse_file)
        self.browse_dir_btn = QPushButton("选文件夹…")
        self.browse_dir_btn.clicked.connect(self._browse_dir)
        row.addWidget(self.browse_file_btn)
        row.addWidget(self.browse_dir_btn)
        row.addStretch()
        box.addLayout(row)

        hint = QLabel(
            "支持 .exe / .ahk / .lnk 快捷方式 / 文件夹 / 文档 / 网址，"
            "点一下就等于双击它。"
        )
        hint.setObjectName("hint")
        hint.setWordWrap(True)
        box.addWidget(hint)

        buttons = QHBoxLayout()
        buttons.addStretch()
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        ok = QPushButton("确定")
        ok.setObjectName("primary")
        ok.setDefault(True)
        ok.clicked.connect(self.accept)
        buttons.addWidget(cancel)
        buttons.addWidget(ok)
        box.addLayout(buttons)

        self.copy_radio.toggled.connect(self._sync_enabled)
        self.run_radio.toggled.connect(self._sync_enabled)
        self._sync_enabled()

    def _sync_enabled(self, *_args):
        running = self.run_radio.isChecked()
        self.target_edit.setEnabled(running)
        self.browse_file_btn.setEnabled(running)
        self.browse_dir_btn.setEnabled(running)

    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择要打开的文件", "", "所有文件 (*.*)"
        )
        if path:
            self.target_edit.setText(path)
            self._fill_text_if_empty(path)

    def _browse_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择要打开的文件夹")
        if path:
            self.target_edit.setText(path)
            self._fill_text_if_empty(path)

    def _fill_text_if_empty(self, path):
        if not self.text_edit.toPlainText().strip():
            self.text_edit.setPlainText(Path(path).stem)

    def result_item(self):
        text = self.text_edit.toPlainText().strip()
        action = "run" if self.run_radio.isChecked() else "copy"
        target = self.target_edit.text().strip()
        if action == "run" and not text and target:
            text = Path(target).stem
        return {"text": text, "action": action, "target": target}


class ClipboardNote(QWidget):
    """标题栏上方那条固定的“剪贴板纸条”。

    - 黑边框 + 淡黄纸底
    - 实时显示剪贴板内容（文字或图片），高度按内容自适应
    - 图片只保留缩放后的小图，原图用完即丢，内存开销很小
    - 鼠标事件全部放行给下面的面板，所以不影响拖动
    """

    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.font_size = settings["font_size"]
        self.source_image = None
        self.pending_text = CLIP_EMPTY_TEXT
        # 当前该显示什么。不能用 isVisible() 判断：面板平时是隐藏的，
        # 那时子控件的 isVisible() 是 False，文字就永远不会被填上。
        self.mode = "text"

        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        # QWidget 子类必须开这个，样式表的背景和边框才会被画出来
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet(
            f"background-color: {CLIP_BG};"
            f" border: 1px solid {CLIP_BORDER};"
            f" border-radius: {BAR_RADIUS}px;"
        )

        self.box = QVBoxLayout(self)
        self.box.setContentsMargins(CLIP_PAD, 5, CLIP_PAD, 6)
        self.box.setSpacing(4)

        self.text_label = QLabel()
        self.text_label.setWordWrap(True)
        self.text_label.setTextInteractionFlags(Qt.NoTextInteraction)
        self.text_label.setStyleSheet(
            f"background: transparent; border: none; color: {HEADER_FG};"
        )
        self.box.addWidget(self.text_label)

        self.image_label = QLabel()
        self.image_label.setStyleSheet(
            "background: transparent; border: none;"
        )
        self.box.addWidget(self.image_label)

        self.apply_settings(settings)
        self.refresh()

    # ---------------- 内容 ----------------
    def refresh(self):
        """重新读一次剪贴板（由 dataChanged 信号触发，不轮询）。"""
        clipboard = QApplication.clipboard()
        mime = clipboard.mimeData() if clipboard else None

        self.text_label.hide()
        self.image_label.hide()
        self.image_label.clear()
        self.source_image = None

        if mime is None:
            self._show_text(CLIP_EMPTY_TEXT)
            text = ""
        else:
            text = (clipboard.text() or "") if clipboard else ""

        if mime is not None and mime.hasImage():
            self._show_image(clipboard.image())
        elif text.strip():
            self._show_text(text)
        elif mime is not None and mime.hasUrls():
            self._show_text(
                "\n".join(
                    url.toLocalFile() or url.toString()
                    for url in mime.urls()
                )
            )
        else:
            self._show_text(CLIP_EMPTY_TEXT)

        self.relayout()

    def _show_text(self, text):
        text = (text or "").strip() or CLIP_EMPTY_TEXT
        if len(text) > CLIP_MAX_CHARS:
            text = text[:CLIP_MAX_CHARS] + " …"
        self.pending_text = text
        self.mode = "text"
        self.text_label.show()

    def _show_image(self, image):
        if image is None or image.isNull():
            self._show_text(CLIP_EMPTY_TEXT)
            return
        # 超大截图先缩到上限，避免几 MB 的原图一直占着内存
        if max(image.width(), image.height()) > CLIP_MAX_SOURCE_SIDE:
            image = image.scaled(
                CLIP_MAX_SOURCE_SIDE, CLIP_MAX_SOURCE_SIDE,
                Qt.KeepAspectRatio, Qt.SmoothTransformation,
            )
        self.source_image = QPixmap.fromImage(image)
        self.mode = "image"
        self.image_label.show()

    def _scale_image(self):
        if self.source_image is None or self.source_image.isNull():
            return
        self.image_label.setPixmap(
            self.source_image.scaled(
                CLIP_MAX_TEXT_WIDTH, CLIP_MAX_IMAGE_HEIGHT,
                Qt.KeepAspectRatio, Qt.SmoothTransformation,
            )
        )

    # ---------------- 尺寸：长和宽都按内容自适应 ----------------
    def relayout(self):
        if self.mode == "text":
            self._fit_text()
        else:
            self._scale_image()

        hint = self.box.sizeHint()
        # +2 是给 1px 边框留的余地
        self.setFixedWidth(hint.width() + 2)
        self.setFixedHeight(hint.height() + 2)

    def _measure(self, candidate):
        """量一段文字：返回（文字区宽度, 文字区高度）。"""
        label = self.text_label
        # 必须先解开上一次的固定尺寸，否则量出来还是旧值
        label.setMinimumWidth(0)
        label.setMaximumWidth(16777215)
        label.setMinimumHeight(0)
        label.setMaximumHeight(16777215)
        label.setText(candidate)
        label.setWordWrap(False)

        natural = label.sizeHint().width()
        width = max(
            CLIP_MIN_TEXT_WIDTH, min(natural, CLIP_MAX_TEXT_WIDTH)
        )
        wrapping = natural > CLIP_MAX_TEXT_WIDTH
        label.setWordWrap(wrapping)
        label.setFixedWidth(width)

        if wrapping:
            height = label.heightForWidth(width)
        else:
            height = label.sizeHint().height()
        return width, height

    def _fit_text(self):
        """按内容定宽定高；太高就往末尾截短，直到放进高度上限。"""
        text = self.pending_text
        width, height = self._measure(text)

        if height > CLIP_MAX_TEXT_HEIGHT:
            low, high = 0, len(text)
            while low < high:
                mid = (low + high + 1) // 2
                _, h = self._measure(text[:mid] + " …")
                if h <= CLIP_MAX_TEXT_HEIGHT:
                    low = mid
                else:
                    high = mid - 1
            text = text[:low] + " …"
            width, height = self._measure(text)

        label = self.text_label
        label.setText(text)
        label.setFixedWidth(width)
        label.setFixedHeight(height)

    def apply_settings(self, settings):
        self.font_size = settings["font_size"]
        font = QFont()
        font.setPixelSize(self.font_size)
        self.text_label.setFont(font)
        self.relayout()


class PanelWidget(DragMixin, QWidget):
    quit_requested = Signal()
    edit_started = Signal()
    edit_finished = Signal()
    drag_started = Signal(QPoint)
    drag_moved = Signal(QPoint)
    drag_finished = Signal()
    settings_changed = Signal(dict)
    settings_preview = Signal(dict)
    resized = Signal()

    def __init__(self, snippets, settings):
        super().__init__()
        self._init_drag()
        self.snippets = snippets
        self.settings = dict(settings)
        self.font_size = self.settings["font_size"]
        self.rows = []
        self.reordering = None

        self.setWindowFlags(
            Qt.FramelessWindowHint
            | Qt.Tool
            | Qt.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WA_TranslucentBackground)

        layout = QVBoxLayout(self)
        # 四周留一圈“隐形可拖动边”：窗口比看得见的纸条大一圈，
        # 这样竖线与 ✎ 之间不会出现能点到桌面图标的空洞
        layout.setContentsMargins(
            GRAB_MARGIN, GRAB_MARGIN, GRAB_MARGIN, GRAB_MARGIN
        )
        layout.setSpacing(0)          # 纸条之间不留间隙，靠颜色区分

        # 标题栏上方的剪贴板纸条
        self.clip_note = ClipboardNote(self.settings, self)
        self.clip_note.setVisible(self.settings["show_clipboard"])
        layout.addWidget(self.clip_note)
        # 事件驱动：只在剪贴板变化时回调一次，空闲时不占 CPU
        QApplication.clipboard().dataChanged.connect(
            self.on_clipboard_changed
        )

        # 顶部窄条：标题 + ⚙ 设置 + ✕ 退出
        # 用一块单独的控件做背景，颜色固定不跟纸条走
        header_box = QWidget(self)
        self.header_box = header_box
        header_box.setStyleSheet(
            f"background-color: {HEADER_BG};"
            f" border-radius: {BAR_RADIUS}px;"
        )
        # 高度固定下来：提示气泡要用同一个高度才能和它平齐
        header_box.setFixedHeight(HEADER_HEIGHT)
        header = QHBoxLayout(header_box)
        header.setContentsMargins(8, 4, 5, 4)
        header.setSpacing(4)

        title = QLabel(self.settings["panel_title"])
        title.setStyleSheet(
            f"background: transparent; color: {HEADER_FG};"
            f" font-size: 11px; font-weight: 600;"
        )
        self.title_label = title
        header.addWidget(title)
        header.addStretch()

        self.settings_button = QPushButton("⚙")
        self.settings_button.setFixedSize(18, 18)
        self.settings_button.setCursor(Qt.PointingHandCursor)
        self.settings_button.setToolTip("设置（文字 / 竖线 / 透明度）")
        self.settings_button.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {HEADER_ICON};
                border: none;
                border-radius: 9px;
                padding: 0px;
                font-size: 13px;
            }}
            QPushButton:hover {{
                background-color: #ECECF2;
                color: #1F1F1F;
            }}
        """)
        self.settings_button.clicked.connect(self.open_settings)
        header.addWidget(self.settings_button)

        self.close_button = QPushButton("✕")
        self.close_button.setFixedSize(18, 18)
        self.close_button.setCursor(Qt.PointingHandCursor)
        self.close_button.setToolTip("退出程序")
        self.close_button.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {HEADER_ICON};
                border: none;
                border-radius: 9px;
                padding: 0px;
                font-size: 12px;
                font-weight: 700;
            }}
            QPushButton:hover {{
                background-color: #E15555;
                color: #FFFFFF;
            }}
        """)
        self.close_button.clicked.connect(lambda: self.quit_requested.emit())
        header.addWidget(self.close_button)
        # 左对齐：标题栏背景只到 ✕ 就结束，长度随提示文字长度自适应
        layout.addWidget(header_box, 0, Qt.AlignLeft)
        # 记住标题栏高度：纸条要和竖线对齐时得减掉它
        self.header_height = HEADER_HEIGHT

        # 最后一条下面：＋ 新增
        # 白底黑框的小条，跟顶部标题栏一个观感；
        # 宽度只包住文字本身（不拉满整行），左边空出色点那一列，
        # 让它和上面每一条纸条的左边缘严格对齐
        self.add_button = QPushButton("＋ 新增一条")
        self.add_button.setCursor(Qt.PointingHandCursor)
        self.add_button.setToolTip(f"再加一条（最多 {MAX_SNIPPETS} 条）")
        self.add_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {HEADER_BG};
                color: {HEADER_FG};
                border: 1px solid {HEADER_FG};
                border-radius: {BAR_RADIUS}px;
                padding: 3px 10px;
                font-size: 11px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: #ECECF2;
            }}
            QPushButton:disabled {{
                background-color: #F2F2F5;
                color: #B8B8C0;
                border: 1px solid #C8C8D0;
            }}
        """)
        self.add_button.clicked.connect(self.add_snippet)
        add_row = QHBoxLayout()
        # 左边空出色点那一列、上面留一点缝，让白框和纸条左边缘对齐又不贴着
        add_row.setContentsMargins(ICON_COL_WIDTH, 4, 0, 0)
        add_row.setSpacing(0)
        add_row.addWidget(self.add_button)
        add_row.addStretch(1)
        layout.addLayout(add_row)

        self.build_rows()
        self.adjustSize()

    def build_rows(self):
        """按 self.snippets 生成所有行（新增 / 删除后重建）。"""
        for row in self.rows:
            self.layout().removeWidget(row)
            row.setParent(None)
            row.deleteLater()
        self.rows = []

        for index, item in enumerate(self.snippets):
            row = SnippetRow(
                index, item, self.settings,
                self.edit_row, self.edit_color, self.delete_row,
            )
            row.reorder_controller = self
            row.clicked.connect(self.activate_row)
            row.drag_started.connect(self.drag_started.emit)
            row.drag_moved.connect(self.drag_moved.emit)
            row.drag_finished.connect(self.drag_finished.emit)
            # 行必须插在标题栏后面、+ 按钮前面（前面还有剪贴板纸条）
            self.layout().insertWidget(
                self.layout().indexOf(self.header_box) + 1 + index, row
            )
            self.rows.append(row)

        full = len(self.snippets) >= MAX_SNIPPETS
        self.add_button.setEnabled(not full)
        self.add_button.setText(
            f"已达上限 {MAX_SNIPPETS} 条" if full else "＋ 新增一条"
        )
        QTimer.singleShot(0, self.refit)

    def add_snippet(self):
        """在最后一条下面加一条（最多 MAX_SNIPPETS 条）。"""
        if len(self.snippets) >= MAX_SNIPPETS:
            return

        used = {item["color"] for item in self.snippets}
        color = next(
            (c for c in SECTOR_COLORS if c not in used), SECTOR_COLORS[0]
        )
        self.snippets.append({
            "text": "新纸条", "color": color,
            "action": "copy", "target": "",
        })
        self.build_rows()
        save_snippets(self.snippets)

        # 直接弹出编辑框，省得再多点一下
        self.edit_row(self.rows[-1])

    def delete_row(self, row):
        """点 ✕ 删除这一条（先问一下，避免误删）。"""
        label = row.text_content.replace("\n", " ")
        if len(label) > 16:
            label = label[:16] + "…"

        self.edit_started.emit()
        try:
            answer = QMessageBox.question(
                self, "删除纸条",
                f"确定删除「{label}」这一条吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
        finally:
            self.edit_finished.emit()

        if answer != QMessageBox.Yes:
            return

        del self.snippets[row.index]
        self.build_rows()
        save_snippets(self.snippets)

    def open_settings(self):
        """点 ⚙ 打开设置。里面的修改会实时预览，确定后才写入文件。"""
        original = dict(self.settings)

        dialog = SettingsDialog(self, self.settings)
        dialog.preview_changed.connect(self.preview_settings)
        self.place_dialog(dialog, self.settings_button)

        self.edit_started.emit()
        try:
            accepted = dialog.exec() == QDialog.Accepted
            new_settings = dialog.result_settings()
        finally:
            self.edit_finished.emit()
            dialog.deleteLater()   # 用完就销毁，开一整天也不会积一堆隐藏窗口

        if accepted:
            self.preview_settings(new_settings)
            self.settings_changed.emit(self.settings)
        else:
            # 取消：把实时预览过的效果还原回去
            self.preview_settings(original)
            self.settings_preview.emit(original)

    def place_dialog(self, dialog, anchor):
        """把弹窗贴在某个按钮旁边：默认在按钮右侧，右边放不下就翻到左侧。

        竖线靠在屏幕右边时面板会翻到竖线左侧，这里的自动翻边也就跟着反过来。
        """
        corner = anchor.mapToGlobal(QPoint(0, 0))
        screen = QApplication.screenAt(corner) or QApplication.primaryScreen()
        area = screen.availableGeometry()
        hint = dialog.sizeHint()
        width = max(hint.width(), dialog.minimumWidth())
        height = hint.height()

        x = corner.x() + anchor.width() + 6
        if x + width > area.right() - 4:
            x = corner.x() - width - 6          # 翻到按钮左边
        x = max(x, area.left() + 4)
        y = min(corner.y(), area.bottom() - height - 4)
        y = max(y, area.top() + 4)
        dialog.move(x, y)

    def preview_settings(self, values):
        """改一下就马上看到效果（不写文件）。"""
        self.settings = dict(values)
        self.font_size = self.settings["font_size"]
        self.title_label.setText(self.settings["panel_title"])
        self.clip_note.apply_settings(self.settings)
        self.clip_note.setVisible(self.settings["show_clipboard"])
        for row in self.rows:
            row.apply_settings(self.settings)
        # 尺寸提示要等 LayoutRequest 事件处理完才刷新，所以延后一拍再算
        QTimer.singleShot(0, self.refit)
        self.settings_preview.emit(self.settings)

    def header_anchor(self):
        """提示气泡的落点：整条标题栏的右上角。

        用标题栏而不是那个 18px 的 ✕：气泡和标题栏同高，
        上边缘对齐了看起来就是一条平齐的兄弟条。
        """
        box = self.header_box
        corner = box.mapToGlobal(QPoint(box.width(), 0))
        return QPoint(corner.x() + 6, corner.y())

    def sync_clipboard(self):
        """手动同步一次剪贴板（打开面板时兜底用）。"""
        self.clip_note.refresh()
        QTimer.singleShot(0, self.refit)

    def on_clipboard_changed(self):
        """剪贴板变了：刷新那条纸条（事件驱动，不轮询）。"""
        self.clip_note.refresh()
        QTimer.singleShot(0, self.refit)

    def content_width_hint(self):
        """所有纸条里最宽的那条。"""
        return max(
            (row.sizeHint().width() for row in self.rows), default=180
        )

    def chrome_height(self):
        """标题栏 + 上方剪贴板纸条的总高度（纸条要和竖线对齐时用）。

        用设置里的 show_clipboard 而不是 clip_note.isVisible()：
        面板平时就是隐藏的，那时子控件的 isVisible() 是 False，
        会白白少算一块高度，让第一条纸条和竖线顶端对不齐。
        """
        height = self.header_height
        if self.settings["show_clipboard"]:
            height += self.clip_note.height()
        return height

    def refit(self):
        """按当前内容重新收紧面板尺寸。"""
        self.layout().activate()
        self.adjustSize()
        self.resized.emit()

    # ---------------- 拖动排序 ----------------
    def reorder_start(self, row, global_pos):
        """按住排序手柄：进入拖动排序状态。"""
        self.reordering = row
        self.edit_started.emit()      # 拖动期间别把面板收起来

    def reorder_move(self, row, global_pos):
        """拖动中：根据鼠标位置把这一行挪到对应位置。"""
        if self.reordering is not row:
            return

        local_y = self.mapFromGlobal(global_pos).y()
        target = 0
        for other in self.rows:
            if other is row:
                continue
            if local_y > other.y() + other.height() / 2:
                target += 1

        base = self.layout().indexOf(self.header_box) + 1
        self.layout().removeWidget(row)
        self.layout().insertWidget(base + target, row)
        self.layout().activate()

    def reorder_end(self, row, global_pos):
        """松手：把界面上看到的顺序写回数据并保存。"""
        if self.reordering is None:
            return
        self.reordering = None
        self.edit_finished.emit()

        visual = sorted(self.rows, key=lambda r: r.y())
        self.rows = visual
        # 原地改内容而不是换一个新 list：
        # SnippetApp 手里拿着的是同一个 list，换了对象两边就脱钩了
        self.snippets[:] = [r.item for r in visual]
        for i, item_row in enumerate(visual):
            item_row.index = i
        save_snippets(self.snippets)
        QTimer.singleShot(0, self.refit)

    def paintEvent(self, event):
        """同竖线：填 alpha=1 的隐形底色，防止四周那圈隐形边被点穿。"""
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor(0, 0, 0, 1))

    def edit_row(self, row):
        """点 ✎ 弹出编辑框：改显示文字，以及点击时是复制还是打开。"""
        dialog = SnippetEditDialog(self, row.item)
        # 贴在 ✎ 按钮右侧；右边放不下就自动翻到左侧
        self.place_dialog(dialog, row.edit_button)

        self.edit_started.emit()
        try:
            accepted = dialog.exec() == QDialog.Accepted
            result = dialog.result_item()
        finally:
            self.edit_finished.emit()
            dialog.deleteLater()   # 用完就销毁，开一整天也不会积一堆隐藏窗口

        if not accepted or not result["text"]:
            return

        item = dict(row.item)
        item.update(result)
        row.set_item(item)
        self.snippets[row.index] = item
        QTimer.singleShot(0, self.refit)
        save_snippets(self.snippets)

    def edit_color(self, row):
        """点颜色圆点弹出扇形色盘，选完立即生效并写入 snippets.json。"""
        popup = ColorPopup(row.color, self)

        self.edit_started.emit()
        try:
            chosen = popup.pick(
                row.color_button.mapToGlobal(QPoint(0, 0))
            )
        finally:
            self.edit_finished.emit()
            popup.deleteLater()    # 同上：弹窗用完就销毁

        if chosen is None or not chosen.isValid():
            return

        item = dict(row.item)
        item["color"] = chosen.name()
        row.set_item(item)
        self.snippets[row.index] = item
        save_snippets(self.snippets)

    def activate_row(self):
        """单击纸条：按 action 决定是复制文字，还是打开程序/文件。"""
        row = self.sender()
        if not isinstance(row, SnippetRow):
            return

        if row.action == "run":
            ok = run_target(row.target)
            tip = "▶ 已打开" if ok else "⚠ 打开失败，检查路径"
        else:
            pyperclip.copy(row.text_content)
            tip = "✅ 已复制"
        show = getattr(QApplication.instance(), "show_tooltip", None)
        if callable(show):
            show(tip)

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        act_quit = menu.addAction("✕  退出程序")
        if menu.exec(event.globalPos()) == act_quit:
            self.quit_requested.emit()


class TooltipWidget(QWidget):
    """提示气泡（✅ 已复制 / 📂 已打开）：高度对齐标题栏。"""

    def __init__(self, height=HEADER_HEIGHT):
        super().__init__()
        self.setWindowFlags(
            Qt.FramelessWindowHint
            | Qt.Tool
            | Qt.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WA_TranslucentBackground)
        # 提示条只是提示一下，不能挡住鼠标点击
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.setStyleSheet("background-color: rgba(0,0,0,0);")

        self._height = height
        self.label = QLabel("✅ 已复制", self)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet(f"""
            background-color: #222222;
            color: white;
            padding: 0px 12px;
            border-radius: {BAR_RADIUS}px;
            font-size: 12px;
        """)
        self.set_height(height)

    def set_height(self, height):
        """把气泡高度对齐到标题栏高度。"""
        self._height = max(18, int(height))
        self.label.setFixedHeight(self._height)
        self._sync()

    def set_text(self, text):
        self.label.setText(text)
        self._sync()

    def _sync(self):
        """宽度跟着文字走，高度始终是标题栏高度。"""
        width = self.label.sizeHint().width()
        self.resize(width, self._height)
        self.label.setGeometry(0, 0, width, self._height)


class SnippetApp(QObject):
    def __init__(self):
        super().__init__()
        self.snippets = load_snippets()
        self.settings = load_settings()

        self.bar = BarWidget()
        # 回到上次待的位置（每个实例各记各的）
        if self.settings["bar_x"] >= 0 or self.settings["bar_y"] >= 0:
            self.bar.move(self.settings["bar_x"], self.settings["bar_y"])
        self.apply_bar_settings()
        self.panel = PanelWidget(self.snippets, self.settings)
        self.tooltip = TooltipWidget(self.panel.header_height)

        self.panel_shown = False
        self.need_close = False
        self.hold = False          # 编辑弹窗打开时，挂起自动隐藏
        self.dragging = False      # 正在拖动整张纸条
        self.drag_offset = QPoint(0, 0)

        self.bar.show()

        self.close_timer = QTimer(self)
        self.close_timer.setSingleShot(True)
        self.close_timer.timeout.connect(self.auto_hide_panel)

        # 提示气泡的自动消失：用同一个定时器 restart，
        # 连点两条时不会被上一条留下的定时器提前收走
        self.tooltip_timer = QTimer(self)
        self.tooltip_timer.setSingleShot(True)
        self.tooltip_timer.timeout.connect(self.tooltip.hide)

        # 挪完位置过一会儿再把坐标写进配置
        self.pos_timer = QTimer(self)
        self.pos_timer.setSingleShot(True)
        self.pos_timer.timeout.connect(self.save_bar_pos)

        # 面板向右滑出动画
        self.slide_anim = QPropertyAnimation(self.panel, b"pos")
        self.slide_anim.setDuration(140)

        # 退出入口：竖线上的 ✕ / 竖线右键 / 面板右上角 ✕ / 面板右键
        self.bar.quit_requested.connect(self.quit_app)
        self.panel.quit_requested.connect(self.quit_app)

        # 编辑时不让面板自动收起
        self.panel.edit_started.connect(lambda: self.set_hold(True))
        self.panel.edit_finished.connect(lambda: self.set_hold(False))

        # 设置面板：改一下实时预览，点确定才写入文件
        self.panel.settings_preview.connect(self.preview_settings)
        self.panel.settings_changed.connect(self.commit_settings)
        self.panel.resized.connect(self.on_bar_moved)

        # 拖动竖线时，让 ✕ 和面板一起跟着走
        self.bar.moved.connect(self.on_bar_moved)

        # 拖动面板本体（纸条、空白处、隐形边）同样移动整张纸条
        self.panel.drag_started.connect(self.begin_drag)
        self.panel.drag_moved.connect(self.drag_to)
        self.panel.drag_finished.connect(self.end_drag)

        QApplication.instance().show_tooltip = self.show_tooltip

        # 关键：用 Qt 自己的全局鼠标坐标（逻辑像素），
        # 和窗口 geometry() 同一坐标系，高分屏缩放也不会错位
        self.monitor_timer = QTimer(self)
        self.monitor_timer.timeout.connect(self.check_mouse)
        self.monitor_timer.start(50)

        # 万一配置里的位置已经跑到屏幕外，先把竖线拉回可见区域
        self.clamp_bar_into_screen()

    def check_mouse(self):
        if self.hold or self.dragging:
            return

        pos = QCursor.pos()
        x, y = pos.x(), pos.y()

        hover = self.bar.geometry().contains(x, y)
        panel_in = self.panel_shown and self.panel.geometry().contains(x, y)

        # 鼠标进到竖线那块矩形里，才把 ✕ 显示出来
        self.bar.show_close(hover)

        if hover and not self.panel_shown:
            self.show_panel()
            return

        if self.panel_shown and not hover and not panel_in:
            self.schedule_hide_panel()

    def set_hold(self, hold):
        self.hold = hold
        if hold:
            self.need_close = False
            self.close_timer.stop()

    # ---------------- 拖动整张纸条 ----------------
    def begin_drag(self, global_pos):
        self.dragging = True
        self.drag_offset = global_pos - self.bar.pos()
        self.slide_anim.stop()

    def drag_to(self, global_pos):
        if not self.dragging:
            return
        self.bar.move(global_pos - self.drag_offset)
        self.on_bar_moved()

    def end_drag(self):
        self.dragging = False

    def preview_settings(self, values):
        """实时预览：只生效不写文件（取消时会被还原）。"""
        self.settings = dict(values)
        self.apply_bar_settings()
        self.on_bar_moved()

    def commit_settings(self, values):
        """点确定：应用并写入 settings.json。"""
        self.preview_settings(values)
        save_settings(self.settings)

    def apply_bar_settings(self):
        self.bar.apply_style(
            self.settings["bar_width"],
            self.settings["bar_color"],
            self.settings["bar_opacity"],
        )

    def panel_target(self, sliding=False):
        """面板该待的位置。右边放不下就自动翻到竖线左边。

        sliding=True 时返回从竖线底下钻出来的起点（用于滑出动画）。
        """
        bar = self.bar.geometry()
        width = self.panel.width()
        screen = (
            QApplication.screenAt(bar.center())
            or QApplication.primaryScreen()
        )
        area = screen.availableGeometry()

        line_left = bar.left() + self.bar.line_left()
        line_right = bar.left() + self.bar.line_right()
        # 第一条纸条的顶端与竖线顶端齐平（panel_offset_y 可再往下推）
        line_top = (
            bar.top() + LINE_Y - GRAB_MARGIN
            - self.panel.chrome_height()
            + self.settings["panel_offset_y"]
        )
        room_right = area.right() - line_right
        room_left = line_left - area.left()

        if room_right >= width or room_right >= room_left:
            x = line_right
            return QPoint(x - PANEL_SLIDE if sliding else x, line_top)

        x = line_left - width
        return QPoint(x + PANEL_SLIDE if sliding else x, line_top)

    def on_bar_moved(self):
        self.position_panel_now()
        self.pos_timer.start(500)

    def save_bar_pos(self):
        """把竖线当前位置记到配置里，下次启动还在原地。"""
        self.clamp_bar_into_screen()
        self.settings["bar_x"] = self.bar.x()
        self.settings["bar_y"] = self.bar.y()
        save_settings(self.settings)

    def clamp_bar_into_screen(self):
        """别让竖线被拖出屏幕（可见的那段必须完整留在屏幕里）。

        窗口比可见的线宽，所以允许窗口超出屏幕边缘，只要线还在就行。
        """
        screen = (
            QApplication.screenAt(self.bar.geometry().center())
            or QApplication.primaryScreen()
        )
        area = screen.availableGeometry()

        line_left = self.bar.x() + self.bar.line_left()
        line_right = line_left + self.bar.thickness
        line_top = self.bar.y() + LINE_Y
        line_bottom = line_top + BAR_HEIGHT

        dx = dy = 0
        if line_left < area.left():
            dx = area.left() - line_left
        elif line_right > area.right() + 1:
            dx = area.right() + 1 - line_right
        if line_top < area.top():
            dy = area.top() - line_top
        elif line_bottom > area.bottom() + 1:
            dy = area.bottom() + 1 - line_bottom

        if dx or dy:
            self.bar.move(self.bar.x() + dx, self.bar.y() + dy)
            self.position_panel_now()

    def position_panel_now(self):
        """立即把面板重新贴到竖线上（不动画）。"""
        if not getattr(self, "panel_shown", False):
            return
        anim = getattr(self, "slide_anim", None)
        if anim is not None:
            anim.stop()
        self.panel.move(self.panel_target())

    def hover_zone(self):
        """竖线窗口本身就是一块实心矩形，整块都算生效区。"""
        return self.bar.geometry()

    def is_inside_ui(self, x, y):
        if self.hover_zone().contains(x, y):
            return True
        if self.panel_shown and self.panel.geometry().contains(x, y):
            return True
        return False

    # ---------------- 退出 ----------------
    def quit_app(self, *args):
        self.monitor_timer.stop()
        self.close_timer.stop()
        self.tooltip_timer.stop()
        self.slide_anim.stop()
        # 刚拖完竖线就点退出时，补一次还没落盘的位置
        if self.pos_timer.isActive():
            self.pos_timer.stop()
            self.save_bar_pos()
        for widget in (self.panel, self.bar, self.tooltip):
            widget.hide()
        QApplication.instance().quit()

    def show_panel(self):
        if self.panel_shown:
            return
        self.panel_shown = True
        self.need_close = False
        self.close_timer.stop()

        # 打开前同步一次剪贴板，防止某次变化没收到信号
        self.panel.sync_clipboard()

        start = self.panel_target(sliding=True)
        end = self.panel_target()

        self.panel.move(start)
        self.panel.show()
        self.panel.raise_()

        # 从竖线底下平滑滑出，最后紧贴竖线（中间不留缝，点不到桌面图标）
        self.slide_anim.stop()
        self.slide_anim.setStartValue(start)
        self.slide_anim.setEndValue(end)
        self.slide_anim.start()

    def hide_panel(self):
        if not self.panel_shown:
            return
        self.panel_shown = False
        self.need_close = False
        self.close_timer.stop()
        self.panel.hide()

    def schedule_hide_panel(self):
        if self.panel_shown and not self.need_close:
            self.need_close = True
            self.close_timer.start(int(LEAVE_DELAY * 1000))

    def auto_hide_panel(self):
        if not self.panel_shown or self.hold:
            return

        pos = QCursor.pos()
        if not self.is_inside_ui(pos.x(), pos.y()):
            self.hide_panel()
        else:
            self.need_close = False

    def show_tooltip(self, text="✅ 已复制"):
        """提示气泡贴着标题栏右端弹出，和标题栏同高、上边缘平齐。"""
        self.tooltip.set_height(self.panel.header_height)
        self.tooltip.set_text(text)
        self.tooltip.show()
        self.tooltip.raise_()

        pos = self.panel.header_anchor()
        screen = QApplication.screenAt(pos) or QApplication.primaryScreen()
        area = screen.availableGeometry()
        tip_w, tip_h = self.tooltip.width(), self.tooltip.height()

        # 右边放不下就翻到标题栏左侧，免得飘出屏幕
        x = pos.x()
        if x + tip_w > area.right() - 4:
            x = pos.x() - self.panel.header_box.width() - 12 - tip_w

        self.tooltip.move(
            max(area.left() + 4, min(x, area.right() - tip_w - 4)),
            max(area.top() + 4, min(pos.y(), area.bottom() - tip_h - 4)),
        )
        self.tooltip_timer.start(1200)


if __name__ == "__main__":
    # 让 Ctrl+C 能直接结束脚本（不然会被 Qt 事件循环吞掉）
    signal.signal(signal.SIGINT, signal.SIG_DFL)

    # 把旧版固定名的配置搬到 "<脚本名>.snippets.json" 等文件里
    migrate_legacy_config()

    app = QApplication(sys.argv)

    snippet_app = SnippetApp()

    sys.exit(app.exec())
