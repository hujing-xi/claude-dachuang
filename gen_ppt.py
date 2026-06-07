from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import copy

# ── Colors ──────────────────────────────────────────────────
DEEP_BLUE  = RGBColor(0x07, 0x1F, 0x65)
TEAL       = RGBColor(0x34, 0x8B, 0x9F)
RED        = RGBColor(0xC0, 0x00, 0x00)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
GRAY_BG    = RGBColor(0xE8, 0xE8, 0xE8)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
MID_GRAY   = RGBColor(0xAA, 0xAA, 0xAA)
DARK_TEXT  = RGBColor(0x20, 0x20, 0x20)
ROW_ALT    = RGBColor(0xEF, 0xF4, 0xFB)   # light blue-gray for alt rows

W  = Inches(13.33)
H  = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

blank_layout = prs.slide_layouts[6]   # completely blank

# ── Helper functions ─────────────────────────────────────────

def add_rect(slide, x, y, w, h, fill_color=None, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(1, x, y, w, h)  # MSO_SHAPE_TYPE.RECTANGLE=1
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape

def add_textbox(slide, x, y, w, h, text, font_name="微软雅黑", font_size=18,
                bold=False, color=DARK_TEXT, align=PP_ALIGN.LEFT,
                wrap=True, v_anchor=None):
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    if v_anchor:
        tf.vertical_anchor = v_anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox

def set_para(para, text, font_name="微软雅黑", font_size=18,
             bold=False, color=DARK_TEXT, align=PP_ALIGN.LEFT):
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color

def add_rect_text(slide, x, y, w, h, title, body=None,
                  bg=DEEP_BLUE, title_color=WHITE, body_color=WHITE,
                  title_size=20, body_size=16, title_bold=True):
    rect = add_rect(slide, x, y, w, h, fill_color=bg)
    txBox = slide.shapes.add_textbox(x + Inches(0.15), y + Inches(0.12),
                                      w - Inches(0.3), h - Inches(0.18))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    set_para(p, title, font_size=title_size, bold=title_bold, color=title_color)
    if body:
        from pptx.oxml.ns import qn
        p2 = tf.add_paragraph()
        set_para(p2, body, font_size=body_size, bold=False, color=body_color)
    return rect

def content_page_header(slide, chapter_tag, page_num):
    """Left: teal vertical bar + chapter tag. Right-bottom: page number."""
    # teal vertical bar
    add_rect(slide, Inches(0.25), Inches(0.22), Inches(0.12), Inches(0.72), fill_color=TEAL)
    # deep blue title background
    add_rect(slide, Inches(0.37), Inches(0.22), Inches(9.8), Inches(0.72), fill_color=DEEP_BLUE)
    add_textbox(slide, Inches(0.45), Inches(0.25), Inches(9.6), Inches(0.65),
                chapter_tag, font_size=26, bold=True, color=WHITE)
    # page number bottom-right
    add_textbox(slide, Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.4),
                str(page_num), font_size=16, bold=False, color=MID_GRAY, align=PP_ALIGN.RIGHT)

def section_page(slide, part_num, title, color=DEEP_BLUE):
    """Section transition page: gray left + deep blue right."""
    # full gray background
    add_rect(slide, 0, 0, W, H, fill_color=GRAY_BG)
    # right deep blue block (trapezoid-like via rectangle)
    add_rect(slide, Inches(5.5), 0, Inches(7.83), H, fill_color=DEEP_BLUE)
    # Part N text
    add_textbox(slide, Inches(5.8), Inches(2.4), Inches(2.5), Inches(1.0),
                f"Part {part_num}", font_size=22, bold=True, color=WHITE)
    add_textbox(slide, Inches(5.8), Inches(1.4), Inches(2.5), Inches(1.0),
                "0" + str(part_num), font_size=72, bold=True, color=WHITE)
    # Chapter name
    add_textbox(slide, Inches(8.5), Inches(3.0), Inches(4.5), Inches(1.2),
                title, font_size=40, bold=True, color=WHITE)
    # Decorative line on gray side
    add_rect(slide, Inches(1.0), Inches(3.5), Inches(3.8), Inches(0.06), fill_color=TEAL)
    # Part label on gray side
    add_textbox(slide, Inches(1.0), Inches(2.5), Inches(4.0), Inches(0.8),
                f"PART  {part_num:02d}", font_size=28, bold=True, color=DEEP_BLUE)

# ═══════════════════════════════════════════════════════════════
# SLIDE 1 — COVER
# ═══════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank_layout)

# Full background: light gray
add_rect(s1, 0, 0, W, H, fill_color=GRAY_BG)

# Dark gray left strip
add_rect(s1, 0, 0, Inches(1.0), H, fill_color=RGBColor(0x55,0x55,0x55))

# Deep blue title block
add_rect(s1, Inches(1.2), Inches(2.2), Inches(11.0), Inches(1.4), fill_color=DEEP_BLUE)
add_textbox(s1, Inches(1.4), Inches(2.28), Inches(10.7), Inches(1.2),
            "面向垂直场景的智能体RAG知识服务系统研究与实现",
            font_size=30, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Red accent line
add_rect(s1, Inches(1.2), Inches(3.62), Inches(11.0), Inches(0.08), fill_color=RED)

# Subtitle block
add_textbox(s1, Inches(1.2), Inches(3.8), Inches(11.0), Inches(0.6),
            "大学生创新创业训练计划立项汇报",
            font_size=22, bold=False, color=DEEP_BLUE, align=PP_ALIGN.CENTER)

# Info line
add_textbox(s1, Inches(1.2), Inches(4.7), Inches(11.0), Inches(0.5),
            "负责人：胡景熙   |   成员：孙泽恺、冷诗雨   |   指导教师：周航",
            font_size=18, bold=False, color=DARK_TEXT, align=PP_ALIGN.CENTER)
add_textbox(s1, Inches(1.2), Inches(5.2), Inches(11.0), Inches(0.5),
            "北京交通大学威海国际学院   |   2026年6月",
            font_size=18, bold=False, color=DARK_TEXT, align=PP_ALIGN.CENTER)

# Decorative circle (school emblem placeholder)
from pptx.util import Pt as PT
from pptx.enum.text import PP_ALIGN
ov = s1.shapes.add_shape(9, Inches(1.5), Inches(4.45), Inches(0.55), Inches(0.55))  # oval
ov.fill.solid(); ov.fill.fore_color.rgb = TEAL
ov.line.fill.background()

# ═══════════════════════════════════════════════════════════════
# SLIDE 2 — TABLE OF CONTENTS
# ═══════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank_layout)
add_rect(s2, 0, 0, W, H, fill_color=WHITE)

# Title
add_rect(s2, Inches(0.5), Inches(0.3), Inches(4.0), Inches(0.75), fill_color=DEEP_BLUE)
add_textbox(s2, Inches(0.6), Inches(0.32), Inches(3.8), Inches(0.7),
            "目  录  CONTENTS", font_size=24, bold=True, color=WHITE)

# Vertical divider
add_rect(s2, Inches(5.0), Inches(1.4), Inches(0.04), Inches(5.5), fill_color=TEAL)

toc_items = [
    ("01", "选题背景与意义"),
    ("02", "研究现状分析"),
    ("03", "主要研究内容"),
    ("04", "工作计划与预期成果"),
    ("05", "参考文献"),
]
colors_toc = [DEEP_BLUE, TEAL, RED, DEEP_BLUE, TEAL]

for i, (num, label) in enumerate(toc_items):
    y = Inches(1.5 + i * 1.0)
    # number circle
    ov = s2.shapes.add_shape(9, Inches(3.2), y + Inches(0.08), Inches(0.55), Inches(0.55))
    ov.fill.solid(); ov.fill.fore_color.rgb = colors_toc[i]
    ov.line.fill.background()
    add_textbox(s2, Inches(3.18), y + Inches(0.04), Inches(0.6), Inches(0.55),
                num, font_size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(s2, Inches(4.0), y, Inches(8.5), Inches(0.65),
                label, font_size=26, bold=False, color=DARK_TEXT)
    # thin line under each item
    add_rect(s2, Inches(3.0), y + Inches(0.72), Inches(9.8), Inches(0.03),
             fill_color=RGBColor(0xCC,0xCC,0xCC))

# Right decorative element
add_rect(s2, Inches(12.0), Inches(0.5), Inches(1.0), Inches(6.5), fill_color=DEEP_BLUE)
add_rect(s2, Inches(12.4), Inches(0.5), Inches(0.6), Inches(6.5), fill_color=TEAL)

# ═══════════════════════════════════════════════════════════════
# SLIDE 3 — Part 1 Transition
# ═══════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank_layout)
section_page(s3, 1, "选题背景与意义")

# ═══════════════════════════════════════════════════════════════
# SLIDE 4 — 1-1 选题背景
# ═══════════════════════════════════════════════════════════════
s4 = prs.slides.add_slide(blank_layout)
add_rect(s4, 0, 0, W, H, fill_color=WHITE)
content_page_header(s4, "1-1  选题背景", 4)

cards = [
    ("LLM 快速发展",
     "大语言模型能力迅速提升，成为\n知识密集型应用的核心基础设施"),
    ("RAG 成为主流方案",
     "检索增强生成（RAG）已是知识问答\n系统的主流技术路线 [1][10]"),
    ("用户请求具有异质性",
     "真实场景中同时存在事实查询、\n空间推理、路线规划三类不同任务"),
    ("垂直场景研究缺口",
     "景区/校园等小规模垂直场景的\n多策略RAG研究尚属空白 [6]"),
]

card_colors = [DEEP_BLUE, TEAL, RED, DEEP_BLUE]
card_x = [Inches(0.3), Inches(3.6), Inches(6.9), Inches(10.2)]

for i, (title, body) in enumerate(cards):
    cx = card_x[i]
    cy = Inches(1.3)
    cw = Inches(3.0)
    ch = Inches(5.8)
    # card background (light)
    add_rect(s4, cx, cy, cw, ch, fill_color=LIGHT_GRAY)
    # colored top bar
    add_rect(s4, cx, cy, cw, Inches(0.55), fill_color=card_colors[i])
    # number circle
    ov = s4.shapes.add_shape(9, cx + Inches(1.15), cy + Inches(0.65), Inches(0.7), Inches(0.7))
    ov.fill.solid(); ov.fill.fore_color.rgb = card_colors[i]
    ov.line.fill.background()
    add_textbox(s4, cx + Inches(1.1), cy + Inches(0.63), Inches(0.7), Inches(0.65),
                str(i+1), font_size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # title
    add_textbox(s4, cx + Inches(0.1), cy + Inches(1.5), cw - Inches(0.2), Inches(0.7),
                title, font_size=18, bold=True, color=card_colors[i], align=PP_ALIGN.CENTER)
    # body
    add_textbox(s4, cx + Inches(0.12), cy + Inches(2.3), cw - Inches(0.24), Inches(3.0),
                body, font_size=15, bold=False, color=DARK_TEXT, align=PP_ALIGN.LEFT)

# Ref line at bottom
add_textbox(s4, Inches(0.3), Inches(7.05), Inches(12.0), Inches(0.35),
            "引用：[1] Lewis et al., NeurIPS 2020  [6] Dong et al., ACL 2025  [10] Gao et al., 2024",
            font_size=11, bold=False, color=MID_GRAY)

# ═══════════════════════════════════════════════════════════════
# SLIDE 5 — 1-2 选题意义
# ═══════════════════════════════════════════════════════════════
s5 = prs.slides.add_slide(blank_layout)
add_rect(s5, 0, 0, W, H, fill_color=WHITE)
content_page_header(s5, "1-2  选题意义", 5)

col_data = [
    (DEEP_BLUE, "理论意义",
     "首次针对小规模垂直服务场景的用户请求异质性问题进行系统性建模，将请求划分为三类并设计对应检索路径，填补多策略RAG研究空白。"),
    (TEAL, "应用价值",
     "为洱海生态廊道等景区提供可复用的智能知识服务框架，支持事实问答、空间关系推理两类核心任务，具有直接落地价值。"),
    (RED, "研究延续性",
     "在团队去年大创成果（事实问答模块）的基础上扩展，引入图谱检索与多策略路由，研究脉络清晰，具备完整连续性。"),
]

for i, (col, title, body) in enumerate(col_data):
    cx = Inches(0.4 + i * 4.3)
    cy = Inches(1.25)
    cw = Inches(4.0)
    ch = Inches(5.9)
    add_rect(s5, cx, cy, cw, ch, fill_color=LIGHT_GRAY)
    add_rect(s5, cx, cy, cw, Inches(0.08), fill_color=col)
    # icon circle
    ov = s5.shapes.add_shape(9, cx + Inches(1.55), cy + Inches(0.2), Inches(0.9), Inches(0.9))
    ov.fill.solid(); ov.fill.fore_color.rgb = col
    ov.line.fill.background()
    add_textbox(s5, cx + Inches(1.5), cy + Inches(0.18), Inches(0.9), Inches(0.88),
                str(i+1), font_size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(s5, cx + Inches(0.1), cy + Inches(1.3), cw - Inches(0.2), Inches(0.65),
                title, font_size=22, bold=True, color=col, align=PP_ALIGN.CENTER)
    add_rect(s5, cx + Inches(0.5), cy + Inches(2.0), cw - Inches(1.0), Inches(0.04), fill_color=col)
    add_textbox(s5, cx + Inches(0.15), cy + Inches(2.15), cw - Inches(0.3), Inches(3.5),
                body, font_size=16, bold=False, color=DARK_TEXT)

# ═══════════════════════════════════════════════════════════════
# SLIDE 6 — Part 2 Transition
# ═══════════════════════════════════════════════════════════════
s6 = prs.slides.add_slide(blank_layout)
section_page(s6, 2, "研究现状分析")

# ═══════════════════════════════════════════════════════════════
# SLIDE 7 — 2-1 国内外研究现状
# ═══════════════════════════════════════════════════════════════
s7 = prs.slides.add_slide(blank_layout)
add_rect(s7, 0, 0, W, H, fill_color=WHITE)
content_page_header(s7, "2-1  国内外研究现状", 7)

# Left column: International
add_rect(s7, Inches(0.3), Inches(1.25), Inches(6.0), Inches(0.55), fill_color=DEEP_BLUE)
add_textbox(s7, Inches(0.4), Inches(1.27), Inches(5.8), Inches(0.5),
            "🌐  国际前沿", font_size=20, bold=True, color=WHITE)

int_items = [
    ("[4] Microsoft GraphRAG", "图结构检索，验证了知识图谱在关系推理任务上的显著优势"),
    ("[1] Meta REALM / RAG", "稠密检索与生成结合，奠定现代RAG基础范式"),
    ("[9] Self-RAG (华盛顿大学)", "自适应按需检索，引入反思令牌提升可靠性"),
    ("[8] ReAct (谷歌/普林斯顿)", "推理与行动交错，构建Agentic RAG重要基础"),
]
for i, (bold_t, body_t) in enumerate(int_items):
    iy = Inches(1.95 + i * 1.2)
    add_rect(s7, Inches(0.3), iy, Inches(0.08), Inches(0.5), fill_color=TEAL)
    add_textbox(s7, Inches(0.5), iy - Inches(0.02), Inches(5.6), Inches(0.45),
                bold_t, font_size=16, bold=True, color=DEEP_BLUE)
    add_textbox(s7, Inches(0.5), iy + Inches(0.38), Inches(5.6), Inches(0.55),
                body_t, font_size=14, bold=False, color=DARK_TEXT)

# Right column: Domestic
add_rect(s7, Inches(6.8), Inches(1.25), Inches(6.2), Inches(0.55), fill_color=TEAL)
add_textbox(s7, Inches(6.9), Inches(1.27), Inches(6.0), Inches(0.5),
            "🇨🇳  国内进展", font_size=20, bold=True, color=WHITE)

dom_items = [
    ("阿里 / 百度 垂直RAG", "探索面向垂直领域的多路径检索框架，覆盖电商、医疗、金融等场景"),
    ("[10] RAG综述 (同济/复旦)", "系统梳理Naive→Advanced→Modular RAG演化路径"),
    ("[11] CRAG (中科大/UCLA)", "纠错检索增强，引入置信度评估触发不同检索策略"),
    ("共同局限", "上述工作均未针对景区/校园等小规模垂直场景的任务异质性进行系统研究"),
]
dom_colors = [TEAL, TEAL, TEAL, RED]
for i, (bold_t, body_t) in enumerate(dom_items):
    iy = Inches(1.95 + i * 1.2)
    add_rect(s7, Inches(6.8), iy, Inches(0.08), Inches(0.5), fill_color=dom_colors[i])
    add_textbox(s7, Inches(7.0), iy - Inches(0.02), Inches(5.8), Inches(0.45),
                bold_t, font_size=16, bold=True, color=TEAL if i < 3 else RED)
    add_textbox(s7, Inches(7.0), iy + Inches(0.38), Inches(5.8), Inches(0.55),
                body_t, font_size=14, bold=False, color=DARK_TEXT)

# Divider
add_rect(s7, Inches(6.55), Inches(1.25), Inches(0.04), Inches(5.6), fill_color=MID_GRAY)

add_textbox(s7, Inches(0.3), Inches(7.05), Inches(12.5), Inches(0.35),
            "引用：[1][4][8][9][10][11]",
            font_size=11, bold=False, color=MID_GRAY)

# ═══════════════════════════════════════════════════════════════
# SLIDE 8 — 2-2 现有方法三大不足
# ═══════════════════════════════════════════════════════════════
s8 = prs.slides.add_slide(blank_layout)
add_rect(s8, 0, 0, W, H, fill_color=WHITE)
content_page_header(s8, "2-2  现有方法的三大不足", 8)

gaps = [
    ("❶  任务类型无法区分",
     "现有RAG系统将事实查询、空间关系推理、路线规划视作同类问题，\n采用统一检索策略，导致非事实类任务准确率大幅下降。"),
    ("❷  检索策略缺乏差异化",
     "单一向量检索对关系密集型（图结构）问题效果有限；\n纯关键词检索对语义相似查询召回率低。两种极端均不适用。"),
    ("❸  生成结果缺乏可靠验证",
     "大多数RAG系统在生成阶段缺少事后验证机制，\n面对时效性冲突或约束违反时无法主动纠错，幻觉风险高。"),
]

for i, (title, body) in enumerate(gaps):
    gy = Inches(1.3 + i * 1.75)
    # numbered bar
    add_rect(s8, Inches(0.3), gy, Inches(12.6), Inches(1.5), fill_color=LIGHT_GRAY)
    add_rect(s8, Inches(0.3), gy, Inches(0.1), Inches(1.5),
             fill_color=[RED, DEEP_BLUE, TEAL][i])
    add_textbox(s8, Inches(0.55), gy + Inches(0.1), Inches(12.0), Inches(0.5),
                title, font_size=20, bold=True, color=[RED, DEEP_BLUE, TEAL][i])
    add_textbox(s8, Inches(0.55), gy + Inches(0.6), Inches(12.0), Inches(0.75),
                body, font_size=16, bold=False, color=DARK_TEXT)

# Conclusion box
add_rect(s8, Inches(0.3), Inches(6.4), Inches(12.6), Inches(0.72), fill_color=DEEP_BLUE)
add_textbox(s8, Inches(0.5), Inches(6.45), Inches(12.2), Inches(0.62),
            "→  本项目提出「多策略 Agentic RAG 框架」，针对性解决以上三大问题",
            font_size=18, bold=True, color=WHITE)

add_textbox(s8, Inches(0.3), Inches(7.08), Inches(12.5), Inches(0.3),
            "引用：[2][3][5][6][7]",
            font_size=11, bold=False, color=MID_GRAY)

# ═══════════════════════════════════════════════════════════════
# SLIDE 9 — Part 3 Transition
# ═══════════════════════════════════════════════════════════════
s9 = prs.slides.add_slide(blank_layout)
section_page(s9, 3, "主要研究内容")

# ═══════════════════════════════════════════════════════════════
# SLIDE 10 — 3-0 系统架构总览
# ═══════════════════════════════════════════════════════════════
s10 = prs.slides.add_slide(blank_layout)
add_rect(s10, 0, 0, W, H, fill_color=WHITE)
content_page_header(s10, "3-0  系统架构总览", 10)

# 4-layer architecture: horizontal flow
layers = [
    (DEEP_BLUE, "Layer 1\n知识底座层",
     "文档库\n节点表\n图谱关系"),
    (TEAL, "Layer 2\n任务路由层",
     "意图识别\n任务分类\n动态分配"),
    (RED, "Layer 3\n多策略检索层",
     "Hybrid RAG\nGraphRAG\nPlanner Agent"),
    (DEEP_BLUE, "Layer 4\n验证输出层",
     "证据充分性\n时效性检查\n约束满足验证"),
]

lx = [Inches(0.35), Inches(3.55), Inches(6.75), Inches(9.95)]
for i, (col, title, body) in enumerate(layers):
    cx = lx[i]
    add_rect(s10, cx, Inches(1.3), Inches(3.0), Inches(5.5), fill_color=LIGHT_GRAY)
    add_rect(s10, cx, Inches(1.3), Inches(3.0), Inches(0.75), fill_color=col)
    add_textbox(s10, cx + Inches(0.1), Inches(1.32), Inches(2.8), Inches(0.72),
                title, font_size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(s10, cx + Inches(0.15), Inches(2.2), Inches(2.7), Inches(4.3),
                body, font_size=16, bold=False, color=DARK_TEXT, align=PP_ALIGN.CENTER)
    # Arrow (except last)
    if i < 3:
        add_rect(s10, cx + Inches(3.02), Inches(3.2), Inches(0.5), Inches(0.06), fill_color=MID_GRAY)
        # arrowhead triangle
        arrow = s10.shapes.add_shape(5, cx + Inches(3.4), Inches(3.08), Inches(0.15), Inches(0.3))
        arrow.fill.solid(); arrow.fill.fore_color.rgb = MID_GRAY
        arrow.line.fill.background()

# Scene label
add_rect(s10, Inches(0.35), Inches(6.95), Inches(12.6), Inches(0.4), fill_color=RGBColor(0xEB,0xF3,0xFB))
add_textbox(s10, Inches(0.5), Inches(6.97), Inches(12.2), Inches(0.35),
            "实验场景：洱海生态廊道    |    三类任务：事实规则查询 · 空间关系推理 · 约束路线规划",
            font_size=14, bold=False, color=DEEP_BLUE, align=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════
# SLIDE 11 — 3-1 场景化知识底座构建
# ═══════════════════════════════════════════════════════════════
s11 = prs.slides.add_slide(blank_layout)
add_rect(s11, 0, 0, W, H, fill_color=WHITE)
content_page_header(s11, "3-1  场景化知识底座构建", 11)

# Left: problem
add_rect(s11, Inches(0.3), Inches(1.3), Inches(5.8), Inches(0.55), fill_color=RED)
add_textbox(s11, Inches(0.4), Inches(1.32), Inches(5.6), Inches(0.5),
            "● 针对问题", font_size=20, bold=True, color=WHITE)
add_rect(s11, Inches(0.3), Inches(1.85), Inches(5.8), Inches(4.9), fill_color=LIGHT_GRAY)
add_textbox(s11, Inches(0.45), Inches(2.0), Inches(5.5), Inches(4.5),
            "景区公开资料散乱异构，缺乏统一知识表示形式\n\n"
            "• 文本资料：介绍手册、规则公告（非结构化）\n\n"
            "• 空间数据：景区节点坐标、设施信息（半结构化）\n\n"
            "• 关系数据：景点关联、路径限制（结构化）\n\n"
            "→ 三类数据格式差异导致统一检索困难",
            font_size=16, bold=False, color=DARK_TEXT)

# Right: solution
add_rect(s11, Inches(6.5), Inches(1.3), Inches(6.5), Inches(0.55), fill_color=TEAL)
add_textbox(s11, Inches(6.6), Inches(1.32), Inches(6.3), Inches(0.5),
            "● 解决方案", font_size=20, bold=True, color=WHITE)

kb_items = [
    (DEEP_BLUE, "文档库（非结构化）",
     "整理洱海生态廊道公开资料、规则性文档，建立向量索引，支持 Hybrid RAG 检索"),
    (TEAL, "结构化节点表",
     "提取空间节点属性（坐标、设施类型、容量限制），构建关系型数据库"),
    (RED, "图谱关系表",
     "抽取景点间路径关系、相邻关系、约束关系，构建知识图谱，支持 GraphRAG"),
]
for i, (col, ktitle, kbody) in enumerate(kb_items):
    ky = Inches(1.95 + i * 1.75)
    add_rect(s11, Inches(6.5), ky, Inches(6.5), Inches(1.6), fill_color=LIGHT_GRAY)
    add_rect(s11, Inches(6.5), ky, Inches(0.12), Inches(1.6), fill_color=col)
    add_textbox(s11, Inches(6.72), ky + Inches(0.1), Inches(6.1), Inches(0.45),
                ktitle, font_size=17, bold=True, color=col)
    add_textbox(s11, Inches(6.72), ky + Inches(0.58), Inches(6.1), Inches(0.9),
                kbody, font_size=14, bold=False, color=DARK_TEXT)

# Arrow between left and right
add_textbox(s11, Inches(5.9), Inches(3.5), Inches(0.7), Inches(0.6),
            "→", font_size=32, bold=True, color=TEAL)

# ═══════════════════════════════════════════════════════════════
# SLIDE 12 — 3-2 任务路由层设计
# ═══════════════════════════════════════════════════════════════
s12 = prs.slides.add_slide(blank_layout)
add_rect(s12, 0, 0, W, H, fill_color=WHITE)
content_page_header(s12, "3-2  任务路由层设计", 12)

# Input node
add_rect(s12, Inches(0.4), Inches(2.8), Inches(2.5), Inches(1.1), fill_color=DEEP_BLUE)
add_textbox(s12, Inches(0.5), Inches(2.85), Inches(2.3), Inches(1.0),
            "用户自然语言\n输入请求", font_size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Arrow to router
add_rect(s12, Inches(2.9), Inches(3.25), Inches(1.0), Inches(0.06), fill_color=DARK_TEXT)

# Router
add_rect(s12, Inches(3.9), Inches(2.65), Inches(2.8), Inches(1.4), fill_color=TEAL)
add_textbox(s12, Inches(4.0), Inches(2.75), Inches(2.6), Inches(1.2),
            "任务路由层\nTask Router", font_size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Three output branches
branches = [
    (Inches(1.5), Inches(1.3), DEEP_BLUE,
     "事实规则查询", "Hybrid RAG\n关键词+向量混合检索"),
    (Inches(1.5), Inches(3.5), TEAL,
     "空间关系推理", "GraphRAG\n知识图谱路径检索"),
    (Inches(1.5), Inches(5.7), RED,
     "约束路线规划", "Planner Agent\n约束满足规划"),
]

for bx, by, bcol, btype, bmethod in branches:
    # line from router
    add_rect(s12, Inches(6.7), by + Inches(0.55), bx - Inches(0.3), Inches(0.05), fill_color=bcol)
    # task type box
    add_rect(s12, bx + Inches(6.4), by, Inches(2.8), Inches(1.1), fill_color=bcol)
    add_textbox(s12, bx + Inches(6.5), by + Inches(0.1), Inches(2.6), Inches(0.9),
                btype, font_size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # method box
    add_rect(s12, bx + Inches(9.5), by, Inches(3.0), Inches(1.1), fill_color=LIGHT_GRAY)
    add_rect(s12, bx + Inches(9.5), by, Inches(0.1), Inches(1.1), fill_color=bcol)
    add_textbox(s12, bx + Inches(9.7), by + Inches(0.05), Inches(2.7), Inches(1.0),
                bmethod, font_size=15, bold=False, color=DARK_TEXT)

# Note
add_textbox(s12, Inches(0.4), Inches(7.05), Inches(12.5), Inches(0.35),
            "引用：[8] ReAct  [9] Self-RAG  [6] RAG-Critic  ——  路由准确率是核心评估指标之一",
            font_size=11, bold=False, color=MID_GRAY)

# ═══════════════════════════════════════════════════════════════
# SLIDE 13 — 3-3 多策略检索与验证层
# ═══════════════════════════════════════════════════════════════
s13 = prs.slides.add_slide(blank_layout)
add_rect(s13, 0, 0, W, H, fill_color=WHITE)
content_page_header(s13, "3-3  多策略检索模块与验证层", 13)

cols_info = [
    (DEEP_BLUE, "Hybrid RAG", "事实规则查询",
     ["BM25 关键词检索", "稠密向量检索 [1][2]", "分数融合排序", "上下文窗口截断 [5]"]),
    (TEAL, "GraphRAG", "空间关系推理",
     ["知识图谱三元组检索", "路径多跳推理", "子图抽取", "关系置信度排序"]),
    (RED, "Planner Agent", "约束路线规划",
     ["约束条件解析", "图上最优路径搜索", "ReAct 推理循环 [8]", "规划可行性验证"]),
]

for i, (col, method, task, steps) in enumerate(cols_info):
    cx = Inches(0.35 + i * 4.32)
    add_rect(s13, cx, Inches(1.3), Inches(4.0), Inches(5.2), fill_color=LIGHT_GRAY)
    add_rect(s13, cx, Inches(1.3), Inches(4.0), Inches(0.6), fill_color=col)
    add_textbox(s13, cx + Inches(0.1), Inches(1.32), Inches(3.8), Inches(0.55),
                method, font_size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(s13, cx + Inches(0.3), Inches(1.95), Inches(3.4), Inches(0.45),
             fill_color=RGBColor(0xDD,0xDD,0xDD))
    add_textbox(s13, cx + Inches(0.35), Inches(1.97), Inches(3.3), Inches(0.42),
                f"适用：{task}", font_size=14, bold=False, color=DARK_TEXT, align=PP_ALIGN.CENTER)
    for j, step in enumerate(steps):
        add_rect(s13, cx + Inches(0.2), Inches(2.55 + j * 0.85), Inches(0.06), Inches(0.4), fill_color=col)
        add_textbox(s13, cx + Inches(0.35), Inches(2.52 + j * 0.85), Inches(3.5), Inches(0.5),
                    step, font_size=15, bold=False, color=DARK_TEXT)

# Validation layer at bottom
add_rect(s13, Inches(0.35), Inches(6.6), Inches(12.6), Inches(0.72), fill_color=DEEP_BLUE)
add_textbox(s13, Inches(0.5), Inches(6.65), Inches(12.2), Inches(0.6),
            "验证层  ▶  证据充分性检查  |  信息时效性验证  |  约束满足性检查  →  降低幻觉风险",
            font_size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_textbox(s13, Inches(0.3), Inches(7.1), Inches(12.5), Inches(0.3),
            "引用：[1][2][5][8][11]",
            font_size=11, bold=False, color=MID_GRAY)

# ═══════════════════════════════════════════════════════════════
# SLIDE 14 — Part 4 Transition
# ═══════════════════════════════════════════════════════════════
s14 = prs.slides.add_slide(blank_layout)
section_page(s14, 4, "工作计划与预期成果")

# ═══════════════════════════════════════════════════════════════
# SLIDE 15 — 4 工作计划 + 预期成果
# ═══════════════════════════════════════════════════════════════
s15 = prs.slides.add_slide(blank_layout)
add_rect(s15, 0, 0, W, H, fill_color=WHITE)
content_page_header(s15, "4  工作计划与预期成果", 15)

# Timeline title
add_textbox(s15, Inches(0.35), Inches(1.3), Inches(5.0), Inches(0.45),
            "工作计划", font_size=20, bold=True, color=DEEP_BLUE)
add_rect(s15, Inches(0.35), Inches(1.75), Inches(0.04), Inches(4.5), fill_color=TEAL)

timeline = [
    ("2026.7–8",  "数据收集与知识底座构建", DEEP_BLUE),
    ("2026.9–10", "任务路由层与 Hybrid RAG 实现", TEAL),
    ("2026.11–12","GraphRAG 图谱检索模块开发", RED),
    ("2027.1–2",  "验证层设计与系统集成联调", DEEP_BLUE),
    ("2027.3–6",  "对比实验、分析与论文撰写", TEAL),
]
for i, (period, task, col) in enumerate(timeline):
    ty = Inches(1.82 + i * 0.88)
    # dot
    ov = s15.shapes.add_shape(9, Inches(0.2), ty + Inches(0.08), Inches(0.3), Inches(0.3))
    ov.fill.solid(); ov.fill.fore_color.rgb = col; ov.line.fill.background()
    add_textbox(s15, Inches(0.6), ty, Inches(1.5), Inches(0.45),
                period, font_size=13, bold=True, color=col)
    add_textbox(s15, Inches(2.1), ty, Inches(4.5), Inches(0.45),
                task, font_size=15, bold=False, color=DARK_TEXT)

# Right: Expected results
add_textbox(s15, Inches(7.0), Inches(1.3), Inches(5.5), Inches(0.45),
            "预期成果与经费", font_size=20, bold=True, color=DEEP_BLUE)

results = [
    (TEAL, "📄 学术论文",
     "发表人工智能或自然语言处理领域\n学术论文 1 篇（会议/期刊）"),
    (DEEP_BLUE, "💻 软件系统",
     "完成面向景区服务的多策略Agentic RAG\n系统，含可视化交互界面"),
    (RED, "💰 经费预算",
     "云计算 ¥3000  |  印刷 ¥1000\n论文版面费 ¥2000  |  合计 ¥6000"),
]
for i, (col, rtitle, rbody) in enumerate(results):
    ry = Inches(1.9 + i * 1.7)
    add_rect(s15, Inches(7.0), ry, Inches(5.8), Inches(1.55), fill_color=LIGHT_GRAY)
    add_rect(s15, Inches(7.0), ry, Inches(0.1), Inches(1.55), fill_color=col)
    add_textbox(s15, Inches(7.2), ry + Inches(0.1), Inches(5.5), Inches(0.45),
                rtitle, font_size=18, bold=True, color=col)
    add_textbox(s15, Inches(7.2), ry + Inches(0.6), Inches(5.5), Inches(0.85),
                rbody, font_size=14, bold=False, color=DARK_TEXT)

# ═══════════════════════════════════════════════════════════════
# SLIDE 16 — 参考文献
# ═══════════════════════════════════════════════════════════════
s16 = prs.slides.add_slide(blank_layout)
add_rect(s16, 0, 0, W, H, fill_color=WHITE)
content_page_header(s16, "参考文献", 16)

refs = [
    "[1]  Lewis P, et al. Retrieval-Augmented Generation for Knowledge-Intensive NLP Tasks. NeurIPS 2020.",
    "[2]  Gao L, et al. Precise Zero-Shot Dense Retrieval without Relevance Labels (HyDE). ACL 2023.",
    "[3]  Es S, et al. RAGAS: Automated Evaluation of Retrieval Augmented Generation. EACL 2024.",
    "[4]  Saad-Falcon J, et al. ARES: An Automated Evaluation Framework for RAG Systems. NAACL 2024.",
    "[5]  Liu N F, et al. Lost in the Middle: How Language Models Use Long Contexts. TACL 2024.",
    "[6]  Dong G, et al. RAG-Critic: Critic-Guided Agentic Workflow for RAG. ACL 2025.",
    "[7]  Lee D, et al. Shifting from Ranking to Set Selection for RAG (SETR). ACL 2025.",
    "[8]  Yao S, et al. ReAct: Synergizing Reasoning and Acting in Language Models. ICLR 2023.",
    "[9]  Asai A, et al. Self-RAG: Learning to Retrieve, Generate, and Critique. ICLR 2024.",
    "[10] Gao Y, et al. Retrieval-Augmented Generation for Large Language Models: A Survey. arXiv 2024.",
    "[11] Yan S-Q, et al. Corrective Retrieval Augmented Generation (CRAG). arXiv 2024.",
    "[12] Friel R, et al. RAGBench: Explainable Benchmark for RAG Systems. arXiv 2024.",
]

row_h = Inches(0.42)
# Header
add_rect(s16, Inches(0.3), Inches(1.25), Inches(12.6), row_h, fill_color=DEEP_BLUE)
add_textbox(s16, Inches(0.45), Inches(1.27), Inches(12.3), row_h - Inches(0.06),
            "文献列表  （共 12 篇）", font_size=16, bold=True, color=WHITE)

for i, ref in enumerate(refs):
    ry = Inches(1.68 + i * row_h)
    bg = LIGHT_GRAY if i % 2 == 0 else ROW_ALT
    add_rect(s16, Inches(0.3), ry, Inches(12.6), row_h - Inches(0.02), fill_color=bg)
    add_textbox(s16, Inches(0.42), ry + Inches(0.03), Inches(12.3), row_h - Inches(0.08),
                ref, font_size=13, bold=False, color=DARK_TEXT)

# ─── Save ───────────────────────────────────────────────────────
out_path = "/home/user/claude-dachuang/rag_ppt.pptx"
prs.save(out_path)
print(f"Saved: {out_path}")
print(f"Total slides: {len(prs.slides)}")
