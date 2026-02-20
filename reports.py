import tkinter as tk
from tkinter import messagebox, filedialog
import json
import os
from datetime import datetime

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                     Paragraph, Spacer, HRFlowable, KeepInFrame)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False


_AR_FORMS = {
    'ء': ('\uFE80', None,    None,    None   ),
    'آ': ('\uFE81', '\uFE82',None,    None   ),
    'أ': ('\uFE83', '\uFE84',None,    None   ),
    'ؤ': ('\uFE85', '\uFE86',None,    None   ),
    'إ': ('\uFE87', '\uFE88',None,    None   ),
    'ئ': ('\uFE89', '\uFE8A','\uFE8B','\uFE8C'),
    'ا': ('\uFE8D', '\uFE8E',None,    None   ),
    'ب': ('\uFE8F', '\uFE90','\uFE91','\uFE92'),
    'ة': ('\uFE93', '\uFE94',None,    None   ),
    'ت': ('\uFE95', '\uFE96','\uFE97','\uFE98'),
    'ث': ('\uFE99', '\uFE9A','\uFE9B','\uFE9C'),
    'ج': ('\uFE9D', '\uFE9E','\uFE9F','\uFEA0'),
    'ح': ('\uFEA1', '\uFEA2','\uFEA3','\uFEA4'),
    'خ': ('\uFEA5', '\uFEA6','\uFEA7','\uFEA8'),
    'د': ('\uFEA9', '\uFEAA',None,    None   ),
    'ذ': ('\uFEAB', '\uFEAC',None,    None   ),
    'ر': ('\uFEAD', '\uFEAE',None,    None   ),
    'ز': ('\uFEAF', '\uFEB0',None,    None   ),
    'س': ('\uFEB1', '\uFEB2','\uFEB3','\uFEB4'),
    'ش': ('\uFEB5', '\uFEB6','\uFEB7','\uFEB8'),
    'ص': ('\uFEB9', '\uFEBA','\uFEBB','\uFEBC'),
    'ض': ('\uFEBD', '\uFEBE','\uFEBF','\uFEC0'),
    'ط': ('\uFEC1', '\uFEC2','\uFEC3','\uFEC4'),
    'ظ': ('\uFEC5', '\uFEC6','\uFEC7','\uFEC8'),
    'ع': ('\uFEC9', '\uFECA','\uFECB','\uFECC'),
    'غ': ('\uFECD', '\uFECE','\uFECF','\uFED0'),
    'ف': ('\uFED1', '\uFED2','\uFED3','\uFED4'),
    'ق': ('\uFED5', '\uFED6','\uFED7','\uFED8'),
    'ك': ('\uFED9', '\uFEDA','\uFEDB','\uFEDC'),
    'ل': ('\uFEDD', '\uFEDE','\uFEDF','\uFEE0'),
    'م': ('\uFEE1', '\uFEE2','\uFEE3','\uFEE4'),
    'ن': ('\uFEE5', '\uFEE6','\uFEE7','\uFEE8'),
    'ه': ('\uFEE9', '\uFEEA','\uFEEB','\uFEEC'),
    'و': ('\uFEED', '\uFEEE',None,    None   ),
    'ى': ('\uFEEF', '\uFEF0',None,    None   ),
    'ي': ('\uFEF1', '\uFEF2','\uFEF3','\uFEF4'),
}

_LAM_ALEF = {
    'آ': ('\uFEF5', '\uFEF6'),
    'أ': ('\uFEF7', '\uFEF8'),
    'إ': ('\uFEF9', '\uFEFA'),
    'ا': ('\uFEFB', '\uFEFC'),
}


def _is_arabic(ch):
    return '\u0600' <= ch <= '\u06FF'


def _connects_left(ch):
    """Returns True if char has an initial/medial form (can connect to next char)."""
    return ch in _AR_FORMS and _AR_FORMS[ch][2] is not None


def _shape_run(chars):
    """Shape a list of Arabic characters into their correct presentation forms."""
    shaped, i = [], 0
    while i < len(chars):
        ch = chars[i]

        if ch not in _AR_FORMS:
            shaped.append(ch)
            i += 1
            continue

        if ch == 'ل' and i + 1 < len(chars) and chars[i + 1] in _LAM_ALEF:
            prev_conn = i > 0 and _connects_left(chars[i - 1])
            shaped.append(_LAM_ALEF[chars[i + 1]][1 if prev_conn else 0])
            i += 2
            continue

        f         = _AR_FORMS[ch]
        prev_conn = i > 0 and _connects_left(chars[i - 1])
        nxt       = chars[i + 1] if i + 1 < len(chars) else ''
        next_ar   = (_is_arabic(nxt) or nxt in _AR_FORMS) and nxt not in ' ،؟!–-\n'
        can_left  = f[2] is not None

        if   prev_conn and next_ar and can_left and f[3]: form = f[3]  # medial
        elif prev_conn and f[1]:                          form = f[1]  # final
        elif next_ar   and can_left and f[2]:             form = f[2]  # initial
        else:                                             form = f[0]  # isolated

        shaped.append(form)
        i += 1
    return shaped


def ar(text: str) -> str:
    """
    Shape + reverse Arabic text for correct PDF rendering.

    Usage:  Paragraph(ar("نص عربي"), my_style)
            canvas.drawRightString(x, y, ar("نص عربي"))

    Handles mixed Arabic/Latin lines: Arabic segments are shaped
    and reversed; Latin/numeric segments keep their original order;
    the whole line's segments are reversed for RTL reading direction.
    """
    if not text:
        return text

    result_lines = []
    for line in text.split('\n'):
        if not line.strip():
            result_lines.append('')
            continue

        segs, cur, cur_ar = [], [], None
        for ch in line:
            is_ar = _is_arabic(ch) or ch in '،؟!'
            if is_ar != cur_ar and cur:
                segs.append((cur_ar, ''.join(cur)))
                cur, cur_ar = [ch], is_ar
            else:
                cur.append(ch)
                if cur_ar is None:
                    cur_ar = is_ar
        if cur:
            segs.append((cur_ar, ''.join(cur)))

        parts = []
        for is_ar, seg in segs:
            if is_ar:
                shaped = _shape_run(list(seg.strip()))
                parts.append(('ar', ''.join(shaped)[::-1]))
            else:
                parts.append(('ltr', seg))

        parts.reverse()
        result_lines.append(''.join(p for _, p in parts))

    return '\n'.join(result_lines)


def _strip_emoji(text: str) -> str:
    """Strip emoji / pictographic characters that FreeSerif can't render."""
    import re
    return re.sub(
        r'[\U0001F300-\U0001F9FF\U00002600-\U000027BF\U0000FE00-\U0000FE0F\u200d]+',
        '', text, flags=re.UNICODE
    ).strip()


def _register_arabic_font():
    """
    Register a system font that supports Arabic glyphs.
    Tries FreeSerif (Linux), Times New Roman (Windows/Mac).
    Returns True if registration succeeded.
    """
    candidates_regular = [
        '/usr/share/fonts/truetype/freefont/FreeSerif.ttf',          # Linux
        'C:/Windows/Fonts/times.ttf',                                  # Windows
        'C:/Windows/Fonts/arial.ttf',
        '/Library/Fonts/Times New Roman.ttf',                          # macOS
        '/System/Library/Fonts/Supplemental/Times New Roman.ttf',
        '/System/Library/Fonts/Supplemental/Arial.ttf',
    ]
    candidates_bold = [
        '/usr/share/fonts/truetype/freefont/FreeSerifBold.ttf',
        'C:/Windows/Fonts/timesbd.ttf',
        'C:/Windows/Fonts/arialbd.ttf',
        '/Library/Fonts/Times New Roman Bold.ttf',
    ]

    def try_register(name, paths):
        for p in paths:
            if os.path.exists(p):
                try:
                    pdfmetrics.registerFont(TTFont(name, p))
                    return True
                except Exception:
                    pass
        return False

    ok_reg  = try_register('ArabicFont',     candidates_regular)
    ok_bold = try_register('ArabicFontBold', candidates_bold)

    if ok_reg and not ok_bold:
        for p in candidates_regular:
            if os.path.exists(p):
                try:
                    pdfmetrics.registerFont(TTFont('ArabicFontBold', p))
                    ok_bold = True
                    break
                except Exception:
                    pass

    return ok_reg and ok_bold

 
BG      = "#0f0f1a"
CARD    = "#1a1a2e"
CARD2   = "#16213e"
ACCENT  = "#e94560"
ACCENT2 = "#0f3460"
TEXT    = "#eaeaea"
SUBTEXT = "#a0a0b0"
NOTE_BG = "#1e1e32"
NOTE_FG = "#f5f0e8"
BORDER  = "#2a2a4a"
SUCCESS = "#4ecca3"



CANVASES = {
    "BMC": {
        "title": "🎯 Business Model Canvas",
        "subtitle": "Alexander Osterwalder",
        "color": "#e94560",
        "layout": [
            (0, 0, 2, 1, "key_partners",      "🤝 Key Partnerships\nالشراكات الرئيسية",          "شركاء، موردين، داعمين..."),
            (0, 1, 1, 1, "key_activities",    "⚙️ Key Activities\nالأنشطة الأساسية",             "أهم أنشطة بتعملها عشان تقدم القيمة..."),
            (0, 2, 2, 1, "value_proposition", "💎 Value Proposition\nالقيمة المقدمة",             "إيه المشكلة اللي بتحلها؟ وإيه اللي بيميزك؟"),
            (0, 3, 1, 1, "customer_relations","❤️ Customer Relationships\nعلاقة العملاء",         "دعم مباشر – خدمة ذاتية – اشتراك..."),
            (0, 4, 2, 1, "customer_segments", "👥 Customer Segments\nشرائح العملاء",              "طلبة – شركات – مكتبات – أفراد..."),
            (1, 1, 1, 1, "key_resources",     "🏗️ Key Resources\nالموارد الأساسية",              "فريق – تكنولوجيا – براند..."),
            (1, 3, 1, 1, "channels",          "📡 Channels\nقنوات التوزيع",                      "سوشيال ميديا – موقع – موزعين..."),
            (2, 0, 1, 2, "cost_structure",    "💸 Cost Structure\nهيكل التكاليف",                "أكبر مصاريف عندك إيه؟"),
            (2, 2, 1, 3, "revenue_streams",   "💰 Revenue Streams\nمصادر الإيرادات",             "بيع مباشر – اشتراكات – عمولات..."),
        ],
        "grid": (3, 5),
    },
    "Lean": {
        "title": "🧠 Lean Canvas",
        "subtitle": "Ash Maurya",
        "color": "#4ecca3",
        "layout": [
            (0, 0, 2, 1, "problem",         "⚡ Problem\nالمشكلة",                         "أهم 3 مشاكل عند العميل..."),
            (0, 1, 1, 1, "solution",        "💡 Solution\nالحل",                           "إزاي هتحل المشكلة؟"),
            (0, 2, 2, 1, "uvp",             "🌟 Unique Value Proposition\nالقيمة الفريدة", "وعد واضح ومختصر يميزك..."),
            (0, 3, 1, 1, "unfair_adv",      "🔒 Unfair Advantage\nميزة تنافسية",           "حاجة محدش يعرف يقلدها بسهولة..."),
            (0, 4, 2, 1, "customer_seg",    "👥 Customer Segments\nالعملاء",               "مين عنده المشكلة دي؟"),
            (1, 1, 1, 1, "key_metrics",     "📊 Key Metrics\nمؤشرات الأداء",               "إزاي تقيس نجاحك؟"),
            (1, 3, 1, 1, "channels",        "📡 Channels\nالقنوات",                        "هتوصل للعميل إزاي؟"),
            (2, 0, 1, 2, "cost_structure",  "💸 Cost Structure\nالتكاليف",                 "المصاريف الأساسية..."),
            (2, 2, 1, 3, "revenue_streams", "💰 Revenue Streams\nالإيرادات",               "طرق الربح..."),
        ],
        "grid": (3, 5),
    },
    "Marketing": {
        "title": "📊 Marketing Canvas",
        "subtitle": "خطة التسويق الشاملة",
        "color": "#ffd700",
        "layout": [
            (0, 0, 1, 1, "target_market", "🎯 Target Market\nالسوق المستهدف",        "مين العميل المثالي؟"),
            (0, 1, 1, 1, "positioning",   "📍 Positioning\nالتموضع",                "عايز الناس تشوفك إزاي؟"),
            (0, 2, 1, 1, "objectives",    "🏆 Marketing Objectives\nأهداف التسويق", "مبيعات؟ انتشار؟ Leads؟"),
            (1, 0, 1, 1, "product",       "📦 Product\nالمنتج",                      "وصف المنتج/الخدمة..."),
            (1, 1, 1, 1, "price",         "💲 Price\nالسعر",                         "استراتيجية التسعير..."),
            (1, 2, 1, 1, "place",         "🗺️ Place\nالمكان",                       "قنوات البيع والتوزيع..."),
            (2, 0, 1, 1, "promotion",     "📣 Promotion\nالترويج",                   "كيف ستروج للمنتج؟"),
            (2, 1, 1, 1, "budget",        "💰 Budget\nالميزانية",                    "كام هتصرف؟"),
            (2, 2, 1, 1, "kpis",          "📈 KPIs\nمؤشرات القياس",                  "معدل التحويل – تكلفة العميل..."),
        ],
        "grid": (3, 3),
    },
    "VPC": {
        "title": "🎯 Value Proposition Canvas",
        "subtitle": "Alexander Osterwalder",
        "color": "#a855f7",
        "layout": [
            (0, 0, 1, 1, "products_services","🛍️ Products & Services\nالمنتجات والخدمات","ماذا تقدم للعميل؟"),
            (0, 1, 1, 1, "customer_jobs",    "🔨 Customer Jobs\nمهام العميل",              "هو عايز يعمل إيه؟"),
            (1, 0, 1, 1, "pain_relievers",   "💊 Pain Relievers\nمخففات الألم",            "إزاي بتقلل الألم؟"),
            (1, 1, 1, 1, "pains",            "😣 Pains\nالمشاكل والمعاناة",               "المشاكل اللي بيعاني منها..."),
            (2, 0, 1, 1, "gain_creators",    "🚀 Gain Creators\nمحققو المكاسب",           "إزاي بتزود المكاسب؟"),
            (2, 1, 1, 1, "gains",            "🌟 Gains\nالمكاسب المتوقعة",               "المكاسب اللي مستنيها..."),
        ],
        "grid": (3, 2),
    },
    "SWOT": {
        "title": "📈 SWOT Analysis",
        "subtitle": "تحليل بيئي شامل",
        "color": "#38bdf8",
        "layout": [
            (0, 0, 1, 1, "strengths",     "💪 Strengths\nنقاط القوة",    "ما الذي تتميز به؟"),
            (0, 1, 1, 1, "weaknesses",    "⚠️ Weaknesses\nنقاط الضعف",  "ما الذي يحتاج تحسين؟"),
            (1, 0, 1, 1, "opportunities", "🌱 Opportunities\nالفرص",     "الفرص المتاحة في السوق..."),
            (1, 1, 1, 1, "threats",       "🔥 Threats\nالتهديدات",       "التهديدات الخارجية..."),
        ],
        "grid": (2, 2),
    },
    "AARRR": {
        "title": "🚀 Growth Hacking Funnel (AARRR)",
        "subtitle": "Dave McClure",
        "color": "#fb923c",
        "layout": [
            (0, 0, 1, 3, "acquisition", "🎣 Acquisition\nاكتساب العملاء",   "من أين يأتي عملاؤك؟ قنوات الاكتساب..."),
            (1, 0, 1, 3, "activation",  "⚡ Activation\nأول تجربة ناجحة",   "كيف تُشعِر العميل بالقيمة لأول مرة؟"),
            (2, 0, 1, 3, "retention",   "🔄 Retention\nالاحتفاظ بالعملاء",  "كيف تُبقي العملاء يعودون؟"),
            (3, 0, 1, 3, "revenue",     "💵 Revenue\nتحقيق الدخل",          "كيف تحقق إيراداً من عملائك؟"),
            (4, 0, 1, 3, "referral",    "📢 Referral\nالإحالة",             "كيف يُحيل عملاؤك أصدقاءهم إليك؟"),
        ],
        "grid": (5, 3),
    },
}


DATA_FILE = os.path.join(os.path.expanduser("~"), "business_canvas_data.json")


def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


class CanvasApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🧩 Business Canvas Studio")
        self.geometry("1280x800")
        self.minsize(900, 600)
        self.configure(bg=BG)
        self.data           = load_data()
        self._active        = None
        self._text_widgets  = {}
        self._build_ui()

    def _build_ui(self):
        self._setup_fonts()
        self._build_sidebar()
        self._build_main()
        self._show_welcome()

    def _setup_fonts(self):
        self.f_heading = ("Georgia",  14, "bold")
        self.f_sub     = ("Courier",  11)
        self.f_note    = ("Courier",  12)
        self.f_small   = ("Courier",   9)
        self.f_welcome = ("Georgia",  38, "bold")

    def _build_sidebar(self):
        sb = tk.Frame(self, bg=CARD2, width=240)
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)

        tk.Label(sb, text="🧩",            font=("Arial", 36),    bg=CARD2, fg=ACCENT).pack(pady=(24, 0))
        tk.Label(sb, text="Canvas Studio", font=self.f_heading,   bg=CARD2, fg=TEXT  ).pack()
        tk.Label(sb, text="مخططات الأعمال",font=self.f_sub,       bg=CARD2, fg=SUBTEXT).pack(pady=(0, 20))
        tk.Frame(sb, bg=BORDER, height=1).pack(fill="x", padx=16)

        btn_defs = [
            ("🏠", "الرئيسية",     "_show_welcome"),
            ("🎯", "BMC",          "BMC"),
            ("🧠", "Lean Canvas",  "Lean"),
            ("📊", "Marketing",    "Marketing"),
            ("🎯", "Value Prop.",  "VPC"),
            ("📈", "SWOT",         "SWOT"),
            ("🚀", "AARRR Funnel", "AARRR"),
        ]
        self.nav_buttons = {}
        for icon, label, key in btn_defs:
            self._sidebar_btn(sb, icon, label, key)

        tk.Frame(sb, bg=BORDER, height=1).pack(fill="x", padx=16, pady=8)

        for txt, cmd in [
            ("💾  حفظ",       self._save_all),
            ("📂  تحميل",     self._load_all),
            ("🗑️  مسح الكل", self._clear_all),
        ]:
            tk.Button(sb, text=txt, font=self.f_sub, bg=ACCENT2, fg=TEXT,
                      bd=0, cursor="hand2", activebackground=ACCENT,
                      activeforeground="white", pady=6, command=cmd
                      ).pack(fill="x", padx=16, pady=2)

        self.clock_lbl = tk.Label(sb, text="", font=self.f_small, bg=CARD2, fg=SUBTEXT)
        self.clock_lbl.pack(side="bottom", pady=12)
        self._tick()

    def _sidebar_btn(self, parent, icon, label, key):
        frame = tk.Frame(parent, bg=CARD2, cursor="hand2")
        frame.pack(fill="x", padx=8, pady=2)

        def on_enter(e): frame.configure(bg=ACCENT2)
        def on_leave(e): frame.configure(bg=CARD2 if self._active != key else ACCENT2)
        frame.bind("<Enter>", on_enter)
        frame.bind("<Leave>", on_leave)

        lbl = tk.Label(frame, text=f"  {icon}  {label}", font=self.f_sub,
                       bg=CARD2, fg=TEXT, anchor="w", padx=8, pady=8)
        lbl.pack(fill="x")
        lbl.bind("<Enter>", on_enter)
        lbl.bind("<Leave>", on_leave)

        def click(e=None, k=key):
            self._show_welcome() if k == "_show_welcome" else self._show_canvas(k)

        frame.bind("<Button-1>", click)
        lbl.bind("<Button-1>", click)
        self.nav_buttons[key] = frame

    def _build_main(self):
        self.main = tk.Frame(self, bg=BG)
        self.main.pack(side="right", fill="both", expand=True)

    def _clear_main(self):
        for w in self.main.winfo_children():
            w.destroy()

    def _show_welcome(self):
        self._active = "_show_welcome"
        self._clear_main()

        header = tk.Frame(self.main, bg=CARD, pady=40)
        header.pack(fill="x")
        tk.Label(header, text="مرحباً بك في",          font=("Georgia", 16), bg=CARD, fg=SUBTEXT).pack()
        tk.Label(header, text="Business Canvas Studio", font=self.f_welcome, bg=CARD, fg=TEXT).pack()
        tk.Label(header, text="ابدأ رسم مستقبل مشروعك الآن 🚀",
                 font=("Georgia", 14), bg=CARD, fg=ACCENT).pack(pady=(8, 0))
        tk.Frame(self.main, bg=BORDER, height=1).pack(fill="x")

        cards_frame = tk.Frame(self.main, bg=BG)
        cards_frame.pack(fill="both", expand=True, padx=30, pady=30)

        card_info = [
            ("🎯", ACCENT,    "Business Model Canvas", "9 خانات لنموذج عملك الكامل",  "BMC"),
            ("🧠", "#4ecca3", "Lean Canvas",           "للستارتاب والأفكار الجديدة",   "Lean"),
            ("📊", "#ffd700", "Marketing Canvas",      "خطة التسويق الشاملة",          "Marketing"),
            ("🎯", "#a855f7", "Value Proposition",     "تحليل قيمتك للعميل",           "VPC"),
            ("📈", "#38bdf8", "SWOT Analysis",         "تحليل بيئي داخلي وخارجي",      "SWOT"),
            ("🚀", "#fb923c", "AARRR Funnel",          "قمع النمو للستارتاب",           "AARRR"),
        ]

        for i, (icon, clr, title, desc, key) in enumerate(card_info):
            col, row = i % 3, i // 3
            card = tk.Frame(cards_frame, bg=CARD, bd=0,
                            highlightthickness=2, highlightbackground=clr, cursor="hand2")
            card.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")

            tk.Label(card, text=icon,  font=("Arial", 28),  bg=CARD).pack(pady=(20, 4))
            tk.Label(card, text=title, font=self.f_heading, bg=CARD, fg=clr).pack()
            tk.Label(card, text=desc,  font=self.f_sub,     bg=CARD, fg=SUBTEXT,
                     wraplength=180).pack(pady=(4, 20))

            def open_canvas(e=None, k=key): self._show_canvas(k)
            card.bind("<Button-1>", open_canvas)
            for child in card.winfo_children():
                child.bind("<Button-1>", open_canvas)

            def enter(e, f=card, c=clr):
                f.configure(bg=c)
                for ch in f.winfo_children(): ch.configure(bg=c)
            def leave(e, f=card):
                f.configure(bg=CARD)
                for ch in f.winfo_children(): ch.configure(bg=CARD)
            card.bind("<Enter>", enter)
            card.bind("<Leave>", leave)

        for c in range(3): cards_frame.columnconfigure(c, weight=1)
        for r in range(2): cards_frame.rowconfigure(r,    weight=1)

    def _show_canvas(self, canvas_key):
        self._active = canvas_key
        self._clear_main()
        cfg   = CANVASES[canvas_key]
        color = cfg["color"]

        title_bar = tk.Frame(self.main, bg=CARD, pady=14)
        title_bar.pack(fill="x")
        tk.Label(title_bar, text=cfg["title"],
                 font=self.f_heading, bg=CARD, fg=color).pack(side="left", padx=20)
        tk.Label(title_bar, text=f"by {cfg['subtitle']}",
                 font=self.f_small, bg=CARD, fg=SUBTEXT).pack(side="left")

        for txt, cmd in [
            ("📄 PDF", lambda k=canvas_key: self._export_pdf(k)),
            ("💾 حفظ", lambda k=canvas_key: self._save_canvas(k)),
            ("🗑️ مسح", lambda k=canvas_key: self._clear_canvas(k)),
            ("🏠 رجوع", self._show_welcome),
        ]:
            clr = SUCCESS if txt.startswith("📄") else ACCENT2
            tk.Button(title_bar, text=txt, font=self.f_small,
                      bg=clr, fg=TEXT if clr != SUCCESS else "#0f0f1a",
                      bd=0, padx=12, pady=6, cursor="hand2", command=cmd,
                      activebackground=ACCENT, activeforeground="white"
                      ).pack(side="right", padx=4)

        tk.Frame(self.main, bg=color, height=2).pack(fill="x")

        container = tk.Frame(self.main, bg=BG)
        container.pack(fill="both", expand=True)

        cv  = tk.Canvas(container, bg=BG, bd=0, highlightthickness=0)
        vsc = tk.Scrollbar(container, orient="vertical",   command=cv.yview)
        hsc = tk.Scrollbar(container, orient="horizontal", command=cv.xview)
        cv.configure(yscrollcommand=vsc.set, xscrollcommand=hsc.set)
        hsc.pack(side="bottom", fill="x")
        vsc.pack(side="right",  fill="y")
        cv.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(cv, bg=BG)
        cv.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(int(-1*(e.delta/120)), "units"))

        self._build_canvas_grid(inner, canvas_key, cfg, color)

    def _build_canvas_grid(self, parent, canvas_key, cfg, color):
        rows, cols = cfg["grid"]
        saved      = self.data.get(canvas_key, {})
        text_widgets = {}

        grid_frame = tk.Frame(parent, bg=BG)
        grid_frame.pack(padx=16, pady=16, fill="both", expand=True)

        for c in range(cols): grid_frame.columnconfigure(c, weight=1, minsize=220)
        for r in range(rows): grid_frame.rowconfigure(r,    weight=1, minsize=180)

        for row, col, rowspan, colspan, key, label, hint in cfg["layout"]:
            cell = tk.Frame(grid_frame, bg=NOTE_BG,
                            highlightthickness=1, highlightbackground=BORDER)
            cell.grid(row=row, column=col, rowspan=rowspan, columnspan=colspan,
                      padx=4, pady=4, sticky="nsew")

            hdr = tk.Frame(cell, bg=ACCENT2)
            hdr.pack(fill="x")
            tk.Label(hdr, text=label, font=("Courier", 9, "bold"),
                     bg=ACCENT2, fg=color, anchor="w", padx=8, pady=6,
                     justify="left", wraplength=200).pack(side="left", fill="x", expand=True)

            tk.Label(cell, text=f"💡 {hint}", font=("Courier", 8),
                     bg=NOTE_BG, fg=SUBTEXT, anchor="w", padx=6,
                     wraplength=200, justify="left").pack(fill="x", pady=(4, 0))

            txt_frame = tk.Frame(cell, bg=NOTE_BG)
            txt_frame.pack(fill="both", expand=True, padx=4, pady=4)

            txt = tk.Text(txt_frame, font=self.f_note, bg=NOTE_BG, fg=NOTE_FG,
                          insertbackground=color, relief="flat", bd=0,
                          padx=8, pady=8, wrap="word", undo=True,
                          selectbackground=ACCENT2, selectforeground=TEXT)
            sc = tk.Scrollbar(txt_frame, command=txt.yview, bg=BG)
            txt.configure(yscrollcommand=sc.set)
            sc.pack(side="right", fill="y")
            txt.pack(side="left", fill="both", expand=True)

            if key in saved:
                txt.insert("1.0", saved[key])

            def on_fi(e, c=cell, clr=color): c.configure(highlightbackground=clr,  highlightthickness=2)
            def on_fo(e, c=cell):            c.configure(highlightbackground=BORDER, highlightthickness=1)
            txt.bind("<FocusIn>",  on_fi)
            txt.bind("<FocusOut>", on_fo)

            text_widgets[key] = txt

        self._text_widgets[canvas_key] = text_widgets

    def _save_canvas(self, canvas_key):
        if canvas_key not in self._text_widgets:
            return
        self.data.setdefault(canvas_key, {})
        for key, txt in self._text_widgets[canvas_key].items():
            self.data[canvas_key][key] = txt.get("1.0", "end-1c")
        save_data(self.data)
        self._flash("✅ تم الحفظ!")

    def _save_all(self):
        for k in list(self._text_widgets):
            self._save_canvas(k)
        save_data(self.data)
        self._flash("✅ تم الحفظ!")

    def _load_all(self):
        self.data = load_data()
        if self._active and self._active != "_show_welcome":
            self._show_canvas(self._active)

    def _clear_canvas(self, canvas_key):
        if not messagebox.askyesno("تأكيد", f"هل تريد مسح بيانات {canvas_key}؟"):
            return
        if canvas_key in self._text_widgets:
            for txt in self._text_widgets[canvas_key].values():
                txt.delete("1.0", "end")
        self.data.pop(canvas_key, None)
        save_data(self.data)

    def _clear_all(self):
        if not messagebox.askyesno("تأكيد", "هل تريد مسح كل البيانات؟ لا يمكن التراجع!"):
            return
        self.data = {}
        self._text_widgets = {}
        save_data(self.data)
        self._show_welcome()

    def _flash(self, msg, bg_color=SUCCESS, width=240):
        w = tk.Toplevel(self)
        w.overrideredirect(True)
        w.configure(bg=bg_color)
        x = self.winfo_x() + self.winfo_width()  // 2 - width // 2
        y = self.winfo_y() + 40
        w.geometry(f"{width}x50+{x}+{y}")
        tk.Label(w, text=msg, font=self.f_heading, bg=bg_color, fg="white").pack(expand=True)
        self.after(1800, w.destroy)

    def _export_pdf(self, canvas_key):
        if not PDF_AVAILABLE:
            messagebox.showerror(
                "مكتبة ناقصة",
                "يرجى تثبيت reportlab:\npip install reportlab"
            )
            return

        self._save_canvas(canvas_key)
        cfg         = CANVASES[canvas_key]
        canvas_data = self.data.get(canvas_key, {})

        default_name = f"{canvas_key}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile=default_name,
            title="حفظ PDF",
        )
        if not path:
            return

        try:
            self._generate_pdf(path, canvas_key, cfg, canvas_data)
            self._flash("✅ تم تصدير PDF!", width=280)
        except Exception as e:
            messagebox.showerror("خطأ في التصدير", f"فشل تصدير PDF:\n{e}")

    def _generate_pdf(self, path, canvas_key, cfg, canvas_data):

        font_ok      = _register_arabic_font()
        ar_font      = "ArabicFont"     if font_ok else "Helvetica"
        ar_font_bold = "ArabicFontBold" if font_ok else "Helvetica-Bold"

        def hx(h):
            h = h.lstrip("#")
            return colors.Color(int(h[0:2],16)/255, int(h[2:4],16)/255, int(h[4:6],16)/255)

        accent     = hx(cfg["color"])
        dark_bg    = hx("#1a1a2e")
        cell_bg    = hx("#1e1e32")
        cell_bg2   = hx("#222240")
        txt_clr    = hx("#eaeaea")
        sub_clr    = hx("#a0a0b0")
        light_gray = hx("#2a2a4a")

        W, H = landscape(A4)
        doc  = SimpleDocTemplate(
            path, pagesize=landscape(A4),
            rightMargin=1.5*cm, leftMargin=1.5*cm,
            topMargin=1.5*cm,   bottomMargin=1.5*cm,
        )

        base = getSampleStyleSheet()

        def ps(name, font, size, clr, align=TA_RIGHT, leading=None):
            return ParagraphStyle(name, parent=base["Normal"],
                                  fontName=font, fontSize=size,
                                  textColor=clr, alignment=align,
                                  leading=leading or size * 1.45)

        title_s    = ps("TT",  ar_font_bold, 18, accent,   TA_LEFT)
        sub_s      = ps("SS",  ar_font,       9, sub_clr,  TA_RIGHT)
        cell_hdr_s = ps("CH",  ar_font_bold,  9, accent,   TA_RIGHT, 13)
        cell_en_s  = ps("CE",  ar_font_bold,  8, accent,   TA_LEFT,  11)
        cell_con_s = ps("CC",  ar_font,        9, txt_clr,  TA_RIGHT, 14)
        empty_s    = ps("EE",  ar_font,        8, sub_clr,  TA_RIGHT)
        footer_s   = ps("FF",  ar_font,        8, sub_clr,  TA_CENTER)

        story   = []
        now_str = datetime.now().strftime("%Y/%m/%d  %H:%M")
        title_en = _strip_emoji(cfg["title"])

        hdr_data = [[
            Paragraph(title_en, title_s),
            Paragraph(ar(f"by {cfg['subtitle']}   |   {now_str}"), sub_s),
        ]]
        hdr_tbl = Table(hdr_data, colWidths=[W * 0.55 - 3*cm, W * 0.45 - 3*cm])
        hdr_tbl.setStyle(TableStyle([
            ("BACKGROUND",   (0,0),(-1,-1), dark_bg),
            ("VALIGN",       (0,0),(-1,-1), "MIDDLE"),
            ("TOPPADDING",   (0,0),(-1,-1), 10),
            ("BOTTOMPADDING",(0,0),(-1,-1), 10),
            ("LEFTPADDING",  (0,0),(0,-1),  12),
            ("RIGHTPADDING", (-1,0),(-1,-1),12),
        ]))
        story.append(hdr_tbl)
        story.append(HRFlowable(width="100%", thickness=3, color=accent,
                                spaceAfter=8, spaceBefore=4))

        layout     = cfg["layout"]
        rows, cols = cfg["grid"]

        table_data    = [[None] * cols for _ in range(rows)]
        span_commands = []

        for row, col, rowspan, colspan, key, label, hint in layout:
            label_clean = _strip_emoji(label)
            parts_label = label_clean.split('\n')
            label_en_txt = parts_label[0].strip()
            label_ar_txt = parts_label[1].strip() if len(parts_label) > 1 else ''

            content = canvas_data.get(key, "").strip()

            cell_parts = []

            if label_en_txt:
                cell_parts.append(Paragraph(label_en_txt, cell_en_s))

            if label_ar_txt:
                cell_parts.append(Paragraph(ar(label_ar_txt), cell_hdr_s))

            cell_parts.append(
                HRFlowable(width="100%", thickness=0.5,
                           color=accent, spaceBefore=3, spaceAfter=5)
            )

            if content:
                for line in content.split("\n"):
                    line = line.strip()
                    if line:
                        cell_parts.append(Paragraph(ar("• " + line), cell_con_s))
            else:
                cell_parts.append(Paragraph(ar("لم تتم الإضافة بعد"), empty_s))

            table_data[row][col] = KeepInFrame(0, 0, cell_parts, mode="shrink")

            if rowspan > 1 or colspan > 1:
                span_commands.append(
                    ("SPAN", (col, row), (col + colspan - 1, row + rowspan - 1))
                )

        for r in range(rows):
            for c in range(cols):
                if table_data[r][c] is None:
                    table_data[r][c] = ""

        col_w = (W - 3*cm) / cols
        row_h = (H - 6*cm) / rows

        tbl = Table(table_data,
                    colWidths=[col_w] * cols,
                    rowHeights=[row_h] * rows)

        tbl_style = [
            ("GRID",         (0,0),(-1,-1), 1.5, accent),
            ("VALIGN",       (0,0),(-1,-1), "TOP"),
            ("LEFTPADDING",  (0,0),(-1,-1), 7),
            ("RIGHTPADDING", (0,0),(-1,-1), 7),
            ("TOPPADDING",   (0,0),(-1,-1), 7),
            ("BOTTOMPADDING",(0,0),(-1,-1), 7),
        ]
        for r in range(rows):
            tbl_style.append(
                ("BACKGROUND", (0,r),(-1,r), cell_bg if r % 2 == 0 else cell_bg2)
            )
        tbl_style.extend(span_commands)
        tbl.setStyle(TableStyle(tbl_style))
        story.append(tbl)

        story.append(Spacer(1, 6))
        story.append(HRFlowable(width="100%", thickness=1,
                                 color=light_gray, spaceAfter=4, spaceBefore=4))
        story.append(Paragraph(
            f"Business Canvas Studio  •  {title_en}  •  {now_str}",
            footer_s
        ))

        doc.build(story)

    def _tick(self):
        self.clock_lbl.configure(text=datetime.now().strftime("%Y/%m/%d\n%H:%M:%S"))
        self.after(1000, self._tick)



if __name__ == "__main__":
    app = CanvasApp()
    app.mainloop()