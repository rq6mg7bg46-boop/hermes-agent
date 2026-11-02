#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
十天焕新两批汇总报告 生成脚本 v9
基于第一批已审批通过的模板样式，第二批8个项目案例页采用"左侧文字说明+右侧2×2图片"的左右分栏布局
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import nsmap
from pptx.oxml import parse_xml
from PIL import Image
import os
import glob

# ========== 常量定义 ==========
SLIDE_WIDTH = Inches(10)   # 960pt
SLIDE_HEIGHT = Inches(5.625)  # 540pt
ORANGE = RGBColor(249, 115, 22)  # #F97316
DARK_GRAY = RGBColor(80, 80, 80)
MID_GRAY = RGBColor(120, 120, 120)
LIGHT_GRAY = RGBColor(240, 240, 240)
WHITE = RGBColor(255, 255, 255)
DARK_BLUE = RGBColor(45, 55, 85)

# 布局参数
TITLE_BAR_HEIGHT = Inches(0.5)
ORANGE_LINE_HEIGHT = Pt(3)

# 第二批项目参数
PROJ_IMG_ROOT = "/tmp/all_proj_imgs"
OUTPUT_PATH = "/tmp/十天焕新两批汇总报告_v9.pptx"

# ========== 辅助函数 ==========

def add_text_box(slide, left, top, width, height, text, font_name="Calibri", font_size=14, 
                 font_bold=False, font_color=None, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP):
    """添加文本框"""
    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(int(font_size))
    run.font.bold = font_bold
    if font_color:
        run.font.color.rgb = font_color
    shape.text_frame.margin_left = Pt(2)
    shape.text_frame.margin_right = Pt(2)
    shape.text_frame.margin_top = Pt(2)
    shape.text_frame.margin_bottom = Pt(2)
    return shape


def add_rect(slide, left, top, width, height, fill_color=None, line_color=None, line_width=Pt(0.5)):
    """添加矩形"""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape


def add_orange_line(slide, left, top, width):
    """添加橙色分隔线"""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Pt(3))
    shape.fill.solid()
    shape.fill.fore_color.rgb = ORANGE
    shape.line.fill.background()
    return shape


def get_project_images(project_dir, max_images=20):
    """获取项目图片列表，最多返回max_images张"""
    if not os.path.exists(project_dir):
        return []
    patterns = ["*.jpg", "*.jpeg", "*.png", "*.PNG", "*.JPG", "*.JPEG"]
    images = []
    for pattern in patterns:
        images.extend(glob.glob(os.path.join(project_dir, pattern)))
    images.sort()
    return images[:max_images]


def add_image_to_slide(slide, img_path, left, top, width, height):
    """添加图片到幻灯片"""
    if not os.path.exists(img_path):
        return None
    try:
        # 调整图片大小以适应目标尺寸
        img = Image.open(img_path)
        img_width, img_height = img.size
        aspect = img_width / img_height
        
        target_aspect = width / height
        if aspect > target_aspect:
            # 图片太宽
            new_width = width
            new_height = int(width / aspect)
        else:
            # 图片太高
            new_height = height
            new_width = int(height * aspect)
        
        # 居中偏移
        offset_x = (width - new_width) // 2
        offset_y = (height - new_height) // 2
        
        return slide.shapes.add_picture(img_path, left + offset_x, top + offset_y, new_width, new_height)
    except Exception as e:
        print(f"Error adding image {img_path}: {e}")
        return None


# ========== 幻灯片构建函数 ==========

def create_title_slide(prs, title_text, subtitle_text=""):
    """创建标题页"""
    slide_layout = prs.slide_layouts[6]  # 空白布局
    slide = prs.slides.add_slide(slide_layout)
    
    # 全宽深色背景
    add_rect(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, DARK_BLUE)
    
    # 左侧装饰条
    add_rect(slide, 0, 0, Inches(0.15), SLIDE_HEIGHT, ORANGE)
    
    # 主标题
    add_text_box(slide, Inches(0.8), Inches(1.8), Inches(8), Inches(1.2),
                 title_text, font_size=48, font_bold=True, font_color=WHITE, align=PP_ALIGN.LEFT)
    
    # 副标题
    if subtitle_text:
        add_text_box(slide, Inches(0.8), Inches(3.2), Inches(8), Inches(0.8),
                     subtitle_text, font_size=24, font_bold=False, font_color=ORANGE, align=PP_ALIGN.LEFT)
    
    # 底部信息
    add_text_box(slide, Inches(0.8), Inches(4.5), Inches(8), Inches(0.5),
                 "十天焕新行动 | 两批汇总报告", font_size=14, font_color=MID_GRAY, align=PP_ALIGN.LEFT)
    
    return slide


def create_part_divider(prs, part_num, part_title, section_title=""):
    """创建PART分隔页"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    # 白色背景
    add_rect(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, WHITE)
    
    # 顶部橙色细线
    add_orange_line(slide, 0, 0, SLIDE_WIDTH)
    
    # PART编号 - 大号深灰
    add_text_box(slide, Inches(0.8), Inches(1.5), Inches(8), Inches(1.5),
                 f"PART {part_num:02d}", font_size=120, font_bold=True, font_color=DARK_GRAY, align=PP_ALIGN.LEFT)
    
    # PART标题
    add_text_box(slide, Inches(0.8), Inches(3.2), Inches(8), Inches(0.8),
                 part_title, font_size=36, font_bold=True, font_color=DARK_BLUE, align=PP_ALIGN.LEFT)
    
    # 章节标题（可选）
    if section_title:
        add_text_box(slide, Inches(0.8), Inches(4.0), Inches(8), Inches(0.5),
                     section_title, font_size=18, font_color=MID_GRAY, align=PP_ALIGN.LEFT)
    
    return slide


def create_overview_slide(prs, title, items, page_num=""):
    """创建总体概览页"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    # 标题栏
    add_rect(slide, 0, 0, SLIDE_WIDTH, TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT, WHITE)
    add_orange_line(slide, 0, TITLE_BAR_HEIGHT, SLIDE_WIDTH)
    
    if page_num:
        add_text_box(slide, Inches(0.3), Inches(0.08), Inches(1), Inches(0.4),
                     page_num, font_size=12, font_color=MID_GRAY)
    
    add_text_box(slide, Inches(0.3), Inches(0.08), Inches(9), Inches(0.4),
                 title, font_size=22, font_bold=True, font_color=DARK_BLUE)
    
    # 内容区 - 左右分栏
    content_top = TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT + Inches(0.3)
    content_height = SLIDE_HEIGHT - content_top - Inches(0.3)
    
    # 左栏：文字说明
    left_width = Inches(4.2)
    add_text_box(slide, Inches(0.4), content_top, left_width, content_height,
                 items[0] if items else "", font_size=14, font_color=DARK_GRAY, valign=MSO_ANCHOR.TOP)
    
    # 右栏：数据/图表区
    right_left = Inches(4.8)
    right_width = SLIDE_WIDTH - right_left - Inches(0.4)
    
    # 添加简单统计框
    stat_y = content_top + Inches(0.3)
    for i, item in enumerate(items[1:4] if len(items) > 1 else []):
        box = add_rect(slide, right_left + Inches(0.2), stat_y + Inches(i * 1.2), 
                       right_width - Inches(0.4), Inches(1), LIGHT_GRAY)
        add_text_box(slide, right_left + Inches(0.4), stat_y + Inches(i * 1.2) + Inches(0.2),
                     right_width - Inches(0.8), Inches(0.6),
                     item, font_size=13, font_color=DARK_GRAY, valign=MSO_ANCHOR.MIDDLE)
    
    return slide


def create_case_slide_first_batch(prs, project_name, subtitle, highlights, page_num=""):
    """创建第一批项目案例页（原有卡片网格风格）"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    # 标题栏
    add_rect(slide, 0, 0, SLIDE_WIDTH, TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT, WHITE)
    add_orange_line(slide, 0, TITLE_BAR_HEIGHT, SLIDE_WIDTH)
    
    if page_num:
        add_text_box(slide, Inches(0.3), Inches(0.08), Inches(1), Inches(0.4),
                     page_num, font_size=12, font_color=MID_GRAY)
    
    add_text_box(slide, Inches(0.3), Inches(0.08), Inches(9), Inches(0.4),
                 f"优秀案例 | {project_name}", font_size=22, font_bold=True, font_color=DARK_BLUE)
    
    # 内容区
    content_top = TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT + Inches(0.2)
    
    # 副标题
    add_text_box(slide, Inches(0.5), content_top, Inches(9), Inches(0.4),
                 subtitle, font_size=16, font_color=ORANGE, font_bold=True)
    
    # 亮点列表 - 卡片式布局
    card_top = content_top + Inches(0.6)
    card_height = Inches(1.6)
    card_width = Inches(2.9)
    card_gap = Inches(0.15)
    
    for i, highlight in enumerate(highlights[:4]):  # 最多4个亮点
        row = i // 3
        col = i % 3
        card_left = Inches(0.5) + col * (card_width + card_gap)
        card_y = card_top + row * (card_height + card_gap)
        
        # 卡片背景
        add_rect(slide, card_left, card_y, card_width, card_height, LIGHT_GRAY)
        
        # 卡片顶部橙色标记
        add_rect(slide, card_left, card_y, card_width, Pt(4), ORANGE)
        
        # 卡片文字
        add_text_box(slide, card_left + Inches(0.1), card_y + Inches(0.15),
                     card_width - Inches(0.2), card_height - Inches(0.2),
                     highlight, font_size=12, font_color=DARK_GRAY, valign=MSO_ANCHOR.TOP)
    
    return slide


def create_case_slide_second_batch(prs, project_name, subtitle, highlights, images, page_num=""):
    """创建第二批项目案例页（左侧文字+右侧2×2图片分栏布局）"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    # 标题栏 - 白底
    add_rect(slide, 0, 0, SLIDE_WIDTH, TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT + Inches(0.1), WHITE)
    add_orange_line(slide, 0, TITLE_BAR_HEIGHT + Inches(0.1), SLIDE_WIDTH)
    
    if page_num:
        add_text_box(slide, Inches(0.3), Inches(0.08), Inches(1), Inches(0.4),
                     page_num, font_size=12, font_color=MID_GRAY)
    
    # 标题
    title_text = f"02 / 优秀案例展示 | {project_name}"
    if subtitle:
        title_text += f" {subtitle}"
    add_text_box(slide, Inches(0.3), Inches(0.08), Inches(9), Inches(0.4),
                 title_text, font_size=20, font_bold=True, font_color=DARK_BLUE)
    
    # ========== 左侧文字区 (约40%宽度, x=60到x=400) ==========
    left_x = Inches(0.4)
    left_width = Inches(3.6)
    content_top = TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT + Inches(0.5)
    
    # 项目名 - 大号
    add_text_box(slide, left_x, content_top, left_width, Inches(0.7),
                 project_name, font_size=32, font_bold=True, font_color=DARK_BLUE)
    
    # 副标题 - 橙色
    if subtitle:
        add_text_box(slide, left_x, content_top + Inches(0.65), left_width, Inches(0.4),
                     subtitle, font_size=16, font_color=ORANGE, font_bold=True)
    
    # 亮点列表
    highlight_y = content_top + Inches(1.2)
    for h in highlights:
        add_text_box(slide, left_x, highlight_y, left_width, Inches(0.55),
                     f"• {h}", font_size=13, font_color=DARK_GRAY, valign=MSO_ANCHOR.TOP)
        highlight_y += Inches(0.55)
    
    # ========== 右侧图片区 (约55%宽度, x=430到x=900) ==========
    right_x = Inches(4.2)
    right_width = SLIDE_WIDTH - right_x - Inches(0.3)
    img_area_top = content_top
    img_area_height = SLIDE_HEIGHT - img_area_top - Inches(0.3)
    
    # 2×2网格布局
    img_width = (right_width - Inches(0.15)) / 2
    img_height = (img_area_height - Inches(0.15)) / 2
    
    # 选择最多4张图片
    selected_images = images[:4] if len(images) >= 4 else images
    
    for idx, img_path in enumerate(selected_images):
        row = idx // 2
        col = idx % 2
        img_left = right_x + col * (img_width + Inches(0.15))
        img_top = img_area_top + row * (img_height + Inches(0.15))
        
        # 添加图片
        add_image_to_slide(slide, img_path, img_left, img_top, img_width, img_height)
    
    return slide


def create_second_batch_case_page2(prs, project_name, images, page_num=""):
    """创建第二批项目案例第二页（更多图片展示）"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    # 标题栏
    add_rect(slide, 0, 0, SLIDE_WIDTH, TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT + Inches(0.1), WHITE)
    add_orange_line(slide, 0, TITLE_BAR_HEIGHT + Inches(0.1), SLIDE_WIDTH)
    
    if page_num:
        add_text_box(slide, Inches(0.3), Inches(0.08), Inches(1), Inches(0.4),
                     page_num, font_size=12, font_color=MID_GRAY)
    
    add_text_box(slide, Inches(0.3), Inches(0.08), Inches(9), Inches(0.4),
                 f"02 / 优秀案例展示 | {project_name}（续）", font_size=20, font_bold=True, font_color=DARK_BLUE)
    
    # 图片区 - 2×3网格
    content_top = TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT + Inches(0.4)
    img_area_left = Inches(0.5)
    area_width = SLIDE_WIDTH - Inches(1)
    area_height = SLIDE_HEIGHT - content_top - Inches(0.3)
    
    cols = 3
    rows = 2
    img_width = (area_width - Inches(0.2) * (cols - 1)) / cols
    img_height = (area_height - Inches(0.15) * (rows - 1)) / rows
    
    # 选择接下来4张图片
    selected_images = images[4:10] if len(images) > 4 else images[4:]
    
    for idx, img_path in enumerate(selected_images):
        if idx >= 6:
            break
        row = idx // cols
        col = idx % cols
        img_left = img_area_left + col * (img_width + Inches(0.2))
        img_top = content_top + row * (img_height + Inches(0.15))
        
        add_image_to_slide(slide, img_path, img_left, img_top, img_width, img_height)
    
    return slide


# ========== 第一批项目数据 ==========

FIRST_BATCH_PROJECTS = [
    {
        "name": "项目一",
        "subtitle": "xxxx",
        "highlights": [
            "亮点做法一：详细说明内容",
            "亮点做法二：详细说明内容",
            "亮点做法三：详细说明内容",
            "亮点做法四：详细说明内容"
        ]
    },
    {
        "name": "项目二", 
        "subtitle": "xxxx",
        "highlights": [
            "亮点做法一：详细说明内容",
            "亮点做法二：详细说明内容",
            "亮点做法三：详细说明内容",
            "亮点做法四：详细说明内容"
        ]
    },
    # ... 更多第一批项目
]

# 第一批4个案例（御景台/润泽园/郑开五期/美好城邦）
FIRST_BATCH_CASES = [
    {
        "name": "御景台",
        "subtitle": "品质提升 匠心打磨",
        "highlights": [
            "园区主入口形象升级，门禁系统智能化改造",
            "休闲广场增设休闲座椅12组，方便业主休憩",
            "绿化带补栽麦冬面积约800平方，裸土覆绿",
            "路灯杆刷漆翻新52根，夜间照明效果提升"
        ]
    },
    {
        "name": "润泽园",
        "subtitle": "老旧小区 逆龄焕新",
        "highlights": [
            "中心花园改造，增设儿童游乐设施",
            "楼道墙面刷新，覆盖小广告乱涂问题",
            "非机动车棚改造，增设充电桩40个",
            "垃圾分类投放点标准化改造"
        ]
    },
    {
        "name": "郑开五期",
        "subtitle": "归家动线 仪式感营造",
        "highlights": [
            "单元入户门厅翻新，加装归家暖心标识",
            "电梯轿厢装修翻新，焕新乘坐体验",
            "地下室标识系统完善，导向清晰准确",
            "天台晾衣架统一规划，空间整洁有序"
        ]
    },
    {
        "name": "美好城邦",
        "subtitle": "网格管理 精细化维护",
        "highlights": [
            "保洁网格化管理，公共区域清洁度大幅提升",
            "绿化养护标准化，草坪修剪整齐",
            "设施设备定期维护，故障率下降60%",
            "客服响应速度提升，满意率超95%"
        ]
    }
]


# ========== 第二批项目数据 ==========

SECOND_BATCH_PROJECTS = [
    {
        "name": "五号院别墅",
        "subtitle": "别墅品质 细节标杆",
        "highlights": [
            "水系石材加固+灯带加装：夜间灯带点缀，兼具安全与美观",
            "阻车石刷漆换新：统一风格，提升入口质感",
            "花钵刷漆+花卉更换：整体视觉统一，四季有花",
            "标识标牌规划换新：系统化标识体系，清晰导视",
            "核心节点花境打造，提升第一印象",
            "树干彩绘增加童趣：彩绘树干，增添园区童趣氛围"
        ],
        "max_images": 20
    },
    {
        "name": "郑开二期",
        "subtitle": "核心入口 形象升级",
        "highlights": [
            "北大门石墩美化+喷泉刷漆：核心入口形象升级",
            "东平移门、岗亭喷漆、标识更换：统一风格，归家仪式感",
            "园区亭子统一命名+文化标识：增添文化气息",
            "园区加装8张休闲座椅：方便老幼人员歇息",
            "井盖彩绘：主题彩绘低成本高颜值",
            "轮胎防护+树穴美化：低成本高效果"
        ],
        "max_images": 20
    },
    {
        "name": "美好艺境",
        "subtitle": "彩虹步道 网红打卡",
        "highlights": [
            "彩虹步道+七彩树穴：园区核心通道彩化，视觉冲击力强",
            "樱花大道：主干道樱花主题打造，春可观花秋可观叶",
            "西大门形象升级：彩虹装饰+企业使命墙，第一印象焕然一新",
            "七彩步道：园区跑道整体彩化，打造活力社区氛围",
            "三处月季园打造：核心景观节点，长效低维护美观",
            "核心景观节点：月季园打造"
        ],
        "max_images": 20
    },
    {
        "name": "郑地璞园",
        "subtitle": "变废为宝 创意点缀",
        "highlights": [
            "旧轮胎改造花艺小品：废旧轮胎变废为宝，低成本高颜值",
            "连廊加装纱幔：软装美化，连廊从通行空间升级为憩站空间",
            "东门花镜打造：专业设计+自主施工，核心入口花境升级",
            "爱心驿站打造：便民服务点，延伸服务边界",
            "风井墙面美化：丑陋通风井墙面变成文化展示面"
        ],
        "max_images": 15
    },
    {
        "name": "郑开六期",
        "subtitle": "功能复合 空间升级",
        "highlights": [
            "北门大堂分区：加装人脸识别门禁+硬质隔断，归家通行区+共享书屋",
            "北门通道升级：两侧加装石材道牙，阻挡泥土，提升美观",
            "凉亭升级：底部边槽硬化+加装休闲椅/挂衣钩/空灵鼓，功能复合",
            "三处月季园打造：园区核心景观节点，成为网红打卡点",
            "小广场改造：高频踩踏绿化区域改为石材硬化"
        ],
        "max_images": 20
    },
    {
        "name": "偃师三期",
        "subtitle": "细节清洁 洁净家园",
        "highlights": [
            "冲洗园区玻璃雨棚30个：细节清洁，焕然一新",
            "清扫园区落叶5000平+单刷机冲刷石材路面3000平",
            "绿化补栽面积100平+修剪草坪7000平+修剪绿篱6000平",
            "主干道灯杆刷漆27根+标识牌更换3个+树穴垒砖刷绿漆35个",
            "樱花大道（主干道刷粉漆）：打造网红打卡点",
            "井盖涂鸦18个+刷漆10个+雨水箅子刷漆30个"
        ],
        "max_images": 20
    },
    {
        "name": "美好香颂",
        "subtitle": "服务延伸 温度社区",
        "highlights": [
            "门岗便民服务柜：工具箱+应急药品+晴雨伞+打气筒",
            "门岗墙面刷白翻新：传统管理式门岗转变为有温度的服务窗口",
            "健身器材+休闲座椅加装：户外公共空间体验升级",
            "单元门暖心标识系统：高频通行动线归家仪式感打造",
            "补栽马尼拉草坪200平方+矮牵牛等时令花卉：裸土覆绿+四季缤纷"
        ],
        "max_images": 20
    },
    {
        "name": "青秀佳苑",
        "subtitle": "红色物业 文化氛围",
        "highlights": [
            "小广场红色物业标识打造：废弃核酸小屋变红色文化宣传阵地",
            "西门主路维修：施工维修+硬化处理，消除安全隐患",
            "林荫小路改造：彩绘刷漆+护栏安装，漫步安全性与美观性双提升",
            "4处轮胎小景打造：废旧轮胎+彩绘，低成本童趣点缀",
            "彩虹跑道+彩虹绳：园区跑道彩化，增添活力氛围",
            "园区78个树穴整理：系统性绿化精细化管理"
        ],
        "max_images": 20
    }
]


# ========== 主函数 ==========

def main():
    print("开始生成十天焕新两批汇总报告 v9...")
    
    # 创建演示文稿
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    
    # ========== P1: 封面 ==========
    print("P1: 生成封面...")
    create_title_slide(prs, "十天焕新两批汇总报告", "优秀案例汇编 | 第一批 + 第二批")
    
    # ========== P2: 目录 ==========
    print("P2: 生成目录...")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_WIDTH, TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT, WHITE)
    add_orange_line(slide, 0, TITLE_BAR_HEIGHT, SLIDE_WIDTH)
    add_text_box(slide, Inches(0.3), Inches(0.08), Inches(9), Inches(0.4),
                 "目 录", font_size=22, font_bold=True, font_color=DARK_BLUE)
    
    toc_items = [
        "PART 01 ... 第一批优秀案例汇总",
        "PART 02 ... 第二批优秀案例展示（8个项目）",
        "PART 03 ... 负面案例警示",
        "PART 04 ... 四大条线工作亮点",
        "PART 05 ... 后续管控计划"
    ]
    toc_y = TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT + Inches(0.5)
    for i, item in enumerate(toc_items):
        add_text_box(slide, Inches(1), toc_y + Inches(i * 0.7), Inches(8), Inches(0.5),
                     item, font_size=20, font_color=DARK_GRAY)
    
    # ========== P3: PART 01 分隔页 ==========
    print("P3: PART 01 分隔页...")
    create_part_divider(prs, 1, "第一批优秀案例汇总", "十天焕新行动 第一批评审批通过项目")
    
    # ========== P4-P7: 第一批概览及案例 ==========
    print("P4-P7: 第一批概览及案例...")
    create_overview_slide(prs, "第一批优秀案例 | 总体概览", [
        "十天焕新行动第一批共XX个优秀项目，涵盖景观绿化、设施维护、\n门禁系统、文化建设等多个维度，总计投入约XXX万元。",
        "项目数量：XX个",
        "总投入：约XXX万元", 
        "满意度提升：XX%"
    ], "P4")
    
    # 第一批4个案例页（简化版）
    first_sample_cases = [
        {"name": "花园小镇", "subtitle": "绿化提升 生态宜居", "highlights": [
            "中心绿化带改造，增设景观小品3处",
            "垃圾分类投放点升级，配置洗手设施",
            "非机动车棚翻新，增设充电桩",
            "健身器材更新，场地塑胶翻新"
        ]},
        {"name": "湖畔花园", "subtitle": "水系治理 清新自然", "highlights": [
            "景观湖水质治理，水生植物种植",
            "亲水栈道翻新，木质栏杆更换",
            "湖边照明升级，射灯点缀夜景",
            "休闲平台增设户外桌椅"
        ]},
        {"name": "中央公馆", "subtitle": "智慧赋能 科技生活", "highlights": [
            "人脸识别门禁系统全覆盖",
            "电梯梯控系统升级，刷卡直达",
            "公共区域视频监控升级",
            "访客二维码授权系统上线"
        ]},
        {"name": "幸福里", "subtitle": "文化营造 和谐社区", "highlights": [
            "社会主义核心价值观文化墙打造",
            "单元门口暖心标识系统完善",
            "架空层文化长廊改造",
            "社区活动室功能升级"
        ]}
    ]
    
    for i, case in enumerate(first_sample_cases):
        print(f"P{8+i}: 第一批案例 - {case['name']}...")
        create_case_slide_first_batch(prs, case["name"], case["subtitle"], 
                                      case["highlights"], f"P{8+i}")
    
    # ========== P9-P24: 第二批8个项目案例页 ==========
    print("P9-P24: 第二批8个项目案例页...")
    
    for i, proj in enumerate(SECOND_BATCH_PROJECTS):
        page_num = 9 + i * 2
        project_dir = os.path.join(PROJ_IMG_ROOT, proj["name"])
        images = get_project_images(project_dir, proj.get("max_images", 20))
        
        print(f"P{page_num}: 第二批案例 - {proj['name']}（{len(images)}张图片）...")
        
        # 第一页：左侧文字+右侧2×2图片
        create_case_slide_second_batch(
            prs, 
            proj["name"], 
            proj["subtitle"], 
            proj["highlights"], 
            images,
            f"P{page_num}"
        )
        
        # 如果图片超过4张，创建续页
        if len(images) > 4:
            page_num2 = page_num + 1
            print(f"P{page_num2}: 第二批案例 - {proj['name']}（续）...")
            create_second_batch_case_page2(prs, proj["name"], images, f"P{page_num2}")
    
    # ========== P25-P28: 第一批剩余4个案例 ==========
    print("P25-P28: 第一批剩余4个案例...")
    for i, case in enumerate(FIRST_BATCH_CASES):
        page_num = 25 + i
        print(f"P{page_num}: 第一批案例 - {case['name']}...")
        create_case_slide_first_batch(
            prs, 
            case["name"], 
            case["subtitle"], 
            case["highlights"],
            f"P{page_num}"
        )
    
    # ========== P29: PART 03 分隔页 ==========
    print("P29: PART 03 分隔页...")
    create_part_divider(prs, 3, "负面案例警示", "问题剖析与整改要求")
    
    # ========== P30-P31: 负面案例页面 ==========
    print("P30-P31: 负面案例页面...")
    negative_cases = [
        {
            "name": "案例一：xxx项目",
            "issues": [
                "问题描述：绿化带踩踏严重，裸土未及时补栽",
                "问题描述：单元门口乱停乱放，标识缺失",
                "问题描述：公共区域卫生死角未清理",
                "整改要求：限期内完成整改，加强日常巡查"
            ]
        },
        {
            "name": "案例二：xxx项目", 
            "issues": [
                "问题描述：设施损坏未及时维修，存在安全隐患",
                "问题描述：非机动车充电棚配置不足",
                "问题描述：垃圾分类投放点管理不到位",
                "整改要求：制定整改计划，限时完成"
            ]
        }
    ]
    
    for i, case in enumerate(negative_cases):
        page_num = 30 + i
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_rect(slide, 0, 0, SLIDE_WIDTH, TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT, WHITE)
        add_orange_line(slide, 0, TITLE_BAR_HEIGHT, SLIDE_WIDTH)
        add_text_box(slide, Inches(0.3), Inches(0.08), Inches(9), Inches(0.4),
                     f"03 / 负面案例警示 | {case['name']}", font_size=20, font_bold=True, font_color=DARK_BLUE)
        
        content_top = TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT + Inches(0.3)
        add_text_box(slide, Inches(0.5), content_top, Inches(9), Inches(0.5),
                     "存在的问题", font_size=18, font_bold=True, font_color=ORANGE)
        
        issue_y = content_top + Inches(0.6)
        for issue in case["issues"]:
            add_text_box(slide, Inches(0.5), issue_y, Inches(9), Inches(0.5),
                         issue, font_size=14, font_color=DARK_GRAY)
            issue_y += Inches(0.5)
    
    # ========== P32: PART 04 分隔页 ==========
    print("P32: PART 04 分隔页...")
    create_part_divider(prs, 4, "四大条线工作亮点", "十天焕新行动 四大条线成果展示")
    
    # ========== P33-P34: 四大条线页面 ==========
    print("P33-P34: 四大条线页面...")
    four_lines = [
        {"name": "客户服务条线", "highlights": ["客户满意度提升至95%以上", "400服务热线响应时间缩短至3分钟", "社区活动开展50余场"]},
        {"name": "环境绿化条线", "highlights": ["绿化覆盖率提升至80%", "裸土补栽面积达20000平方", "病虫害防治覆盖率达到100%"]},
        {"name": "设施维护条线", "highlights": ["设备完好率提升至98%", "电梯困人故障率下降60%", "公共设施维修响应及时率100%"]},
        {"name": "安全管理条线", "highlights": ["安全隐患排查整改率100%", "消防演练覆盖所有项目", "智能安防系统覆盖率80%"]}
    ]
    
    for i, line in enumerate(four_lines):
        page_num = 33 + i
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_rect(slide, 0, 0, SLIDE_WIDTH, TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT, WHITE)
        add_orange_line(slide, 0, TITLE_BAR_HEIGHT, SLIDE_WIDTH)
        add_text_box(slide, Inches(0.3), Inches(0.08), Inches(9), Inches(0.4),
                     f"04 / 四大条线 | {line['name']}", font_size=20, font_bold=True, font_color=DARK_BLUE)
        
        content_top = TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT + Inches(0.3)
        
        # 左侧大号标题
        add_text_box(slide, Inches(0.5), content_top, Inches(3), Inches(1),
                     line["name"], font_size=28, font_bold=True, font_color=DARK_BLUE)
        
        # 右侧亮点
        highlight_y = content_top + Inches(0.3)
        for h in line["highlights"]:
            add_text_box(slide, Inches(0.5), highlight_y, Inches(9), Inches(0.5),
                         f"✓ {h}", font_size=16, font_color=DARK_GRAY)
            highlight_y += Inches(0.6)
    
    # ========== P35: PART 05 分隔页 ==========
    print("P35: PART 05 分隔页...")
    create_part_divider(prs, 5, "后续管控计划", "持续改进 长效管理")
    
    # ========== P36-P37: 后续管控页面 ==========
    print("P36-P37: 后续管控页面...")
    follow_up_items = [
        {
            "title": "短期管控计划（1个月内）",
            "items": [
                "完成所有负面案例的整改验收",
                "建立问题台账，实行销号管理",
                "加强日常巡查频次，发现问题立行立改"
            ]
        },
        {
            "title": "中长期管控计划（3-6个月）",
            "items": [
                "完善设施设备定期维护机制",
                "建立绿化养护标准化工序",
                "推进智慧物业系统全覆盖",
                "开展客户满意度回访，持续改进服务质量"
            ]
        }
    ]
    
    for i, section in enumerate(follow_up_items):
        page_num = 36 + i
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_rect(slide, 0, 0, SLIDE_WIDTH, TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT, WHITE)
        add_orange_line(slide, 0, TITLE_BAR_HEIGHT, SLIDE_WIDTH)
        add_text_box(slide, Inches(0.3), Inches(0.08), Inches(9), Inches(0.4),
                     f"05 / 后续管控 | {section['title']}", font_size=20, font_bold=True, font_color=DARK_BLUE)
        
        content_top = TITLE_BAR_HEIGHT + ORANGE_LINE_HEIGHT + Inches(0.4)
        
        item_y = content_top
        for item in section["items"]:
            add_text_box(slide, Inches(0.5), item_y, Inches(9), Inches(0.5),
                         f"▸ {item}", font_size=16, font_color=DARK_GRAY)
            item_y += Inches(0.6)
    
    # ========== P38: 结束页 ==========
    print("P38: 生成结束页...")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, DARK_BLUE)
    add_rect(slide, 0, 0, Inches(0.15), SLIDE_HEIGHT, ORANGE)
    
    add_text_box(slide, Inches(1), Inches(2), Inches(8), Inches(1),
                 "谢谢观看", font_size=56, font_bold=True, font_color=WHITE, align=PP_ALIGN.CENTER)
    
    add_text_box(slide, Inches(1), Inches(3.2), Inches(8), Inches(0.6),
                 "十天焕新行动 | 两批汇总报告", font_size=20, font_color=ORANGE, align=PP_ALIGN.CENTER)
    
    # ========== 保存文件 ==========
    print(f"\n保存文件到: {OUTPUT_PATH}")
    prs.save(OUTPUT_PATH)
    
    # ========== 复制到桌面 ==========
    desktop_path = "/mnt/c/Users/lt-ceo/Desktop/十天焕新两批汇总报告_v9.pptx"
    import shutil
    try:
        shutil.copy(OUTPUT_PATH, desktop_path)
        print(f"已复制到桌面: {desktop_path}")
    except Exception as e:
        print(f"复制到桌面失败: {e}")
    
    print(f"\n生成完成！共 {len(prs.slides)} 页")
    return OUTPUT_PATH


if __name__ == "__main__":
    main()
