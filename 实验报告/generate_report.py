#!/usr/bin/env python3
"""生成《计算机应用项目实践-智能无人飞行器》设计报告"""

from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os

# ============================================================
# CONFIG
# ============================================================
STUDENT_NAME = "莫仁鹰"
STUDENT_ID = "2023212167"
COLLEGE = "计算机与信息学院"
MAJOR_CLASS = "物联网23级-1班"
ADVISOR = "【指导教师姓名】"
TOPIC = "基于RK3566的智能无人飞行器视觉识别系统设计"
DATE_STR = "2026年6月17日"

PHOTO_DIR = "/home/ubuntu/report_photos/all/"
FIRST_PHOTO = "/home/ubuntu/report_photos/photo_01_hardware_board_camera.png"
SCREEN_DIR = "/home/ubuntu/experiment_screenshots/"

# Font names - use Noto as fallback for SimSun
FONT_SONG = "宋体"  # Will fallback to system font in LibreOffice
FONT_HEI = "黑体"
FONT_EN = "Times New Roman"

OUTPUT_PATH = "/home/ubuntu/实验报告_莫仁鹰_2023212167.docx"

# ============================================================
# HELPERS
# ============================================================

def set_cell_font(cell, text, font_name=FONT_SONG, size=Pt(14), bold=False, alignment=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = alignment
    run = p.add_run(text)
    run.font.name = font_name
    run.font.size = size
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def add_paragraph(doc, text, font_name=FONT_SONG, size=Pt(12), bold=False,
                  alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, color=None,
                  space_before=Pt(0), space_after=Pt(0), first_line_indent=None,
                  line_spacing=1.25):
    p = doc.add_paragraph()
    p.alignment = alignment
    pf = p.paragraph_format
    pf.space_before = space_before
    pf.space_after = space_after
    pf.line_spacing = line_spacing
    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent

    run = p.add_run(text)
    run.font.name = font_name
    run.font.size = size
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    if color:
        run.font.color.rgb = color
    return p

def add_heading_1(doc, text):
    """一级标题: 三号黑体"""
    p = add_paragraph(doc, text, font_name=FONT_HEI, size=Pt(16), bold=True,
                      alignment=WD_ALIGN_PARAGRAPH.LEFT,
                      space_before=Pt(12), space_after=Pt(6))
    return p

def add_heading_2(doc, text):
    """二级标题: 四号黑体"""
    p = add_paragraph(doc, text, font_name=FONT_HEI, size=Pt(14), bold=True,
                      alignment=WD_ALIGN_PARAGRAPH.LEFT,
                      space_before=Pt(8), space_after=Pt(4))
    return p

def add_heading_3(doc, text):
    """三级标题: 小四黑体"""
    p = add_paragraph(doc, text, font_name=FONT_HEI, size=Pt(12), bold=True,
                      alignment=WD_ALIGN_PARAGRAPH.LEFT,
                      space_before=Pt(6), space_after=Pt(3))
    return p

def add_body(doc, text):
    """正文: 小四宋体, 1.25倍行距, 首行缩进2字符"""
    p = add_paragraph(doc, text, font_name=FONT_SONG, size=Pt(12),
                      alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                      first_line_indent=Cm(0.74), line_spacing=1.25)
    return p

def add_image_with_caption(doc, image_path, caption, width=Inches(4.5), fig_num=""):
    """插入图片+图题"""
    if os.path.exists(image_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(image_path, width=width)
    else:
        add_paragraph(doc, f"【占位符：{caption}】", font_name=FONT_SONG, size=Pt(12),
                      bold=True, color=RGBColor(0xFF, 0x00, 0x00),
                      alignment=WD_ALIGN_PARAGRAPH.CENTER)

    # Caption below
    cap_text = f"图{fig_num} {caption}" if fig_num else caption
    add_paragraph(doc, cap_text, font_name=FONT_SONG, size=Pt(10.5),
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  space_before=Pt(3), space_after=Pt(6))

def add_placeholder_images(doc, description, fig_num=""):
    """添加占位符区域（红色标注）"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)

    # Red bordered placeholder
    run = p.add_run(f"【占位符{fig_num}：{description}】")
    run.font.name = FONT_SONG
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_SONG)

def add_page_break(doc):
    doc.add_page_break()

# ============================================================
# MAIN REPORT GENERATION
# ============================================================

doc = Document()

# Set default font
style = doc.styles['Normal']
font = style.font
font.name = FONT_EN
font.size = Pt(12)
style.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_SONG)

# Page margins
for section in doc.sections:
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)

# ============================================================
# COVER PAGE
# ============================================================

# School name
for _ in range(3):
    doc.add_paragraph()

add_paragraph(doc, "合肥工业大学", font_name=FONT_HEI, size=Pt(26), bold=True,
              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(20))

add_paragraph(doc, "智能无人飞行器设计", font_name=FONT_HEI, size=Pt(26), bold=True,
              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))

add_paragraph(doc, "设计报告", font_name=FONT_HEI, size=Pt(26), bold=True,
              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(40))

# Info table
table = doc.add_table(rows=6, cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER

info = [
    ("学　　　院", COLLEGE),
    ("专 业 班 级", MAJOR_CLASS),
    ("学生姓名及学号", f"{STUDENT_NAME}  {STUDENT_ID}"),
    ("指导教师", ADVISOR),
    ("课题名称", TOPIC),
    (DATE_STR, DATE_STR),
]

for i, (label, value) in enumerate(info):
    set_cell_font(table.cell(i, 0), label, size=Pt(14), bold=True)
    set_cell_font(table.cell(i, 1), value, size=Pt(14))

# Set column widths
for row in table.rows:
    row.cells[0].width = Cm(4.5)
    row.cells[1].width = Cm(8)

add_page_break(doc)

# ============================================================
# 一、课题概述
# ============================================================

add_heading_1(doc, "一、课题概述")

add_body(doc, "随着无人机技术的飞速发展，智能无人飞行器在航拍测绘、农业植保、物流配送、灾害搜救等领域得到了广泛应用。近年来，基于深度学习的目标检测技术（如YOLO系列算法）与嵌入式计算平台（如瑞芯微RK3566）的结合，使得无人飞行器具备了实时视觉识别与自主决策能力，推动了智能无人系统向轻量化、低功耗、高实时性方向发展。")

add_body(doc, "本课题以Cessna 182固定翼航模为载体，围绕\u201c嵌入式视觉识别系统\u201d这一核心任务展开设计与实验。主要工作包括：飞机机体组装与调试、Pixhawk飞控固件升级与参数配置、RK3566（立创泰山派）开发板系统镜像烧录与驱动部署、YOLO目标检测数据集构建与标注、YOLOv5模型训练与ONNX格式导出、以及RKNN模型转换与边缘端部署。")

add_body(doc, "本项目的设计目标是构建一套完整的无人飞行器视觉识别系统原型：由Pixhawk飞控负责飞行姿态控制，RK3566开发板承担机载视觉推理任务，通过摄像头采集图像并运行轻量化YOLO模型实现实时目标检测。整个系统从硬件组装、固件烧录、模型训练到边缘部署，覆盖了嵌入式AI系统开发的全流程，对培养系统级工程能力具有重要的实践意义。")

add_body(doc, "本报告将按照课题任务、技术方案、设计实现与测试、课程总结四个部分，详细阐述设计过程中的技术路线、实验步骤、遇到的问题及解决方案，力求完整呈现本次课程设计的全貌。")

add_page_break(doc)

# ============================================================
# 二、课题任务
# ============================================================

add_heading_1(doc, "二、课题任务")

add_heading_2(doc, "2.1 功能需求")

add_body(doc, "本课题需要完成以下核心功能：")
add_body(doc, "（1）飞机机体组装：完成Cessna 182固定翼航模的机械结构组装，包括机身、机翼、尾翼、起落架、电机、舵机、螺旋桨等部件的安装与固定。")
add_body(doc, "（2）飞控系统配置：对Pixhawk飞控进行固件升级（ArduPlane固件），完成遥控器校准、飞行模式设置、传感器标定等基本参数配置，使飞控具备基本的飞行控制能力。")
add_body(doc, "（3）嵌入式开发板部署：在RK3566（泰山派）开发板上烧录Ubuntu系统镜像，配置开发环境，安装必要的驱动和依赖库。")
add_body(doc, "（4）视觉识别模型开发：使用LabelImg工具标注目标检测数据集，基于YOLOv5框架训练目标检测模型，将训练后的模型转换为ONNX→RKNN格式，部署到RK3566开发板上实现边缘端推理。")
add_body(doc, "（5）系统集成与测试：将飞控、开发板、摄像头等硬件集成到飞机机体内，进行联调测试。")

add_heading_2(doc, "2.2 技术指标")

add_body(doc, "（1）飞控固件版本：ArduPlane V4.x稳定版，支持多飞行模式切换。")
add_body(doc, "（2）开发板操作系统：基于Ubuntu的Linux系统（Buildroot/Debian），支持SSH远程登录和桌面显示。")
add_body(doc, "（3）目标检测模型：YOLOv5s/YOLOv5n，输入分辨率640×640，推理速度满足边缘端实时性要求。")
add_body(doc, "（4）模型格式：PyTorch (.pt) → ONNX (.onnx) → RKNN (.rknn)，最终在RK3566 NPU上加速推理。")
add_body(doc, "（5）摄像头接口：USB摄像头或MIPI CSI摄像头模块，支持720P以上分辨率采集。")

add_page_break(doc)

# ============================================================
# 三、技术方案及关键问题
# ============================================================

add_heading_1(doc, "三、技术方案及关键问题")

add_heading_2(doc, "3.1 总体技术方案")

add_body(doc, "本项目采用\u201c飞控+视觉处理板\u201d双板架构。Pixhawk飞控负责飞行姿态控制和传感器数据融合，RK3566开发板负责视觉图像采集与AI推理。两者之间通过串口（UART/TELEM）进行通信，实现视觉识别结果向飞控的反馈。整体技术路线如下：")

add_body(doc, "硬件层面：Cessna 182航模机体 → 电机/舵机/电调安装 → Pixhawk飞控安装与接线 → RK3566开发板安装 → 摄像头模块连接 → 电池/电源分配板布线。")

add_body(doc, "软件层面：Pixhawk固件烧录（Mission Planner） → RK3566系统镜像烧录（RKDevTool） → 开发环境搭建（Python/RKNN-Toolkit2） → 数据集标注（Make Sense） → YOLOv5模型训练 → ONNX导出 → RKNN转换 → 边缘端部署推理。")

add_heading_2(doc, "3.2 开发平台选型")

add_body(doc, "（1）飞控平台：选用Pixhawk系列飞控，其基于STM32处理器，支持ArduPilot开源固件，生态成熟、社区活跃，是教学和科研领域最常用的飞控平台。地面站软件使用Mission Planner（Windows平台），支持固件升级、参数配置、实时数据监控和航线规划。")

add_body(doc, "（2）视觉处理平台：选用瑞芯微RK3566芯片的立创泰山派开发板。RK3566内置0.8TOPS算力的NPU（神经网络处理单元），支持RKNN推理框架，能够高效运行轻量化深度学习模型。相比树莓派等通用开发板，RK3566在AI推理场景下具有显著的性能和功耗优势。")

add_body(doc, "（3）目标检测框架：选用YOLOv5，其在检测精度和推理速度之间取得了良好平衡。YOLOv5提供多种模型规格（n/s/m/l/x），适合在不同算力平台上部署。训练框架基于PyTorch，导出支持ONNX格式，可通过RKNN-Toolkit2转换为RK3566 NPU可执行的RKNN格式。")

add_heading_2(doc, "3.3 关键问题分析")

add_body(doc, "（1）USB数据线识别问题：在飞控固件烧录过程中，最初使用的Micro USB线为充电线（仅含电源线，无数据线），导致电脑无法识别Pixhawk设备。该问题的根本原因是充电线只有2根芯（VCC+GND），而数据传输需要4根芯（VCC+GND+D++D-）。解决方案：购买专用Micro USB数据线后问题立即解决。")

add_body(doc, "（2）RK3566开发板驱动兼容性：在Windows 11（Build 26200）环境下，泰山派通过USB连接电脑后，RNDIS网络设备虽能被识别但无法建立网络链路。经排查发现，微软在Windows 11新版本中移除了对RNDIS驱动的完整支持（MediaConnectionState始终为Disconnected）。该问题导致无法通过USB网络直接SSH登录开发板。")

add_body(doc, "（3）实验器材管理问题：项目前期完成初步组装后将器材放置于实验室，后发现部分零件（包括完整的机身部件）被实验室其他人员取走更换。在验收前两天，指导教师联系厂商重新调配了完整部件，才得以继续进行后续实验。此问题提醒了在共享实验室环境中做好器材标识和保管的重要性。")

add_body(doc, "（4）模型格式转换兼容性：ONNX模型转换为RKNN格式时，需要注意算子兼容性。RKNN-Toolkit2对部分ONNX算子支持有限，需在导出ONNX时设置合适的opset版本，并在转换配置中指定正确的目标平台（rk3566）和量化策略。")

add_page_break(doc)

# ============================================================
# 四、设计实现及测试（重点部分）
# ============================================================

add_heading_1(doc, "四、设计实现及测试")

# ---- 4.1 硬件组装 ----
add_heading_2(doc, "4.1 飞机机体组装")

add_body(doc, "本实验使用的飞机模型为Cessna 182固定翼航模，机翼翼展约1.2m，采用EPO材质机身。组装过程在合肥工业大学实验室内完成，由小组成员协作进行。主要组装步骤包括：")

add_body(doc, "（1）机身主体与机翼连接：将左右机翼通过碳纤维管插入机身预留的翼管孔内，使用胶水和扎带固定。机翼上预装有副翼舵机，需将舵机延长线穿过机翼内部引出至机身。")

add_body(doc, "（2）尾翼安装：将水平尾翼和垂直尾翼插入机身尾部卡槽，固定升降舵和方向舵舵机，连接舵机摇臂与控制面。")

add_body(doc, "（3）电机与螺旋桨安装：将无刷电机固定在机头电机座上，连接电调（ESC）的三相线，安装螺旋桨并确认旋转方向正确。")

add_body(doc, "（4）起落架安装：前起落架和主起落架分别固定在机身底部对应位置，确保飞机放置时姿态水平。")

add_body(doc, "（5）电子设备安装：将Pixhawk飞控通过减震座固定在机身内部中心位置，RK3566开发板固定在飞控上方或旁边，摄像头模块安装在机头下方或驾驶舱位置。所有设备通过电源分配板统一供电。")

add_body(doc, "组装过程中，小组成员分工合作：部分成员负责机械结构拼装，部分成员负责电子接线。下图展示了组装过程的实景照片。")

# Insert assembly photos
photo_map = {
    "photo_16.jpeg": "小组成员在实验室内协作组装Cessna 182飞机",
    "photo_17.jpeg": "安装飞控与开发板接线",
    "photo_20.jpeg": "机身内部电子设备布局",
}

fig_counter = 1
for fname, caption in photo_map.items():
    fpath = os.path.join(PHOTO_DIR, fname)
    add_image_with_caption(doc, fpath, caption, width=Inches(4.2), fig_num=f"4.{fig_counter}")
    fig_counter += 1

# Pixhawk close-up
add_image_with_caption(doc, os.path.join(PHOTO_DIR, "photo_29.jpeg"),
                       "机身内部Pixhawk飞控与减震座安装（近景）",
                       width=Inches(4.2), fig_num=f"4.{fig_counter}")
fig_counter += 1

# Board installed in plane
add_image_with_caption(doc, os.path.join(PHOTO_DIR, "photo_21.jpeg"),
                       "RK3566泰山派开发板安装在飞机机舱内",
                       width=Inches(4.2), fig_num=f"4.{fig_counter}")
fig_counter += 1

# Team collaboration
add_image_with_caption(doc, os.path.join(PHOTO_DIR, "photo_30.jpeg"),
                       "小组成员进行飞机接线与调试",
                       width=Inches(4.2), fig_num=f"4.{fig_counter}")
fig_counter += 1

# Board + camera from first upload
add_image_with_caption(doc, FIRST_PHOTO,
                       "RK3566开发板与摄像头模块连接",
                       width=Inches(3.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

# More assembly photos
more_photos = [
    ("photo_18.jpeg", "飞机内部走线布局"),
    ("photo_25.jpeg", "机翼与机身对接固定"),
    ("photo_28.jpeg", "实验室全景——各小组同时进行飞机组装"),
]
for fname, caption in more_photos:
    fpath = os.path.join(PHOTO_DIR, fname)
    add_image_with_caption(doc, fpath, caption, width=Inches(4.2), fig_num=f"4.{fig_counter}")
    fig_counter += 1

add_body(doc, "在组装过程中遇到了一个较为棘手的问题：项目初期完成初步组装后将飞机放置在实验室，后发现部分零件被其他实验室成员取走替换。在验收前两天，指导教师联系到厂商重新调配了完整的机身部件，小组成员紧急完成了重新组装工作。这一经历提醒我们，在共享实验室环境中需要做好器材的标识与保管。")

add_page_break(doc)

# ---- 4.2 飞控固件升级及参数烧录 ----
add_heading_2(doc, "4.2 飞控固件升级及参数烧录")

add_body(doc, "本节对应实验3.2的内容。Pixhawk飞控需要烧录ArduPlane固件才能控制固定翼飞机。本实验使用Mission Planner地面站软件完成固件升级和参数配置。")

add_heading_3(doc, "4.2.1 实验环境")

add_body(doc, "硬件：Pixhawk 2.4.8飞控、Micro USB数据线、Windows电脑。")
add_body(doc, "软件：Mission Planner地面站（V1.3.x），已预装在课程资源包中。")

add_heading_3(doc, "4.2.2 固件烧录步骤")

add_body(doc, "（1）使用Micro USB数据线将Pixhawk连接至电脑。首次连接时Windows会自动安装驱动，设备管理器中应出现对应的COM端口。")

add_body(doc, "（2）打开Mission Planner，点击右上角\u201c初始设置\u201d→\u201c安装固件\u201d，选择ArduPlane（固定翼）图标，等待固件下载完成后自动烧录。烧录过程中Pixhawk指示灯会快速闪烁，完成后提示\u201cUpload Done\u201d。")

add_body(doc, "（3）烧录完成后断开USB重新连接，在Mission Planner左上角选择正确的COM端口和波特率（115200），点击\u201c连接\u201d按钮与飞控建立通信。")

add_heading_3(doc, "4.2.3 遇到的问题及解决")

add_body(doc, "在本实验过程中遇到了一个关键问题：使用最初配备的Micro USB线连接Pixhawk后，飞控指示灯亮起（说明供电正常），但电脑的设备管理器中始终无法识别到COM端口，Mission Planner也无法连接。")

add_body(doc, "经过排查分析，确定问题出在USB数据线上。初始使用的线缆实际上是充电线，内部仅有VCC和GND两根电源线，缺少D+和D-两根数据线，因此无法进行USB数据通信。这是一个非常常见但容易被忽视的问题——外观上充电线和数据线几乎无法区分。", )

add_body(doc, "解决方案：在网上购买了专用的Micro USB数据线（确认支持数据传输），更换后设备管理器立即识别到了COM端口，Mission Planner也成功连接并完成了固件烧录。")

# Placeholders for Mission Planner screenshots
add_placeholder_images(doc, "Mission Planner固件烧录界面截图（包括：选择ArduPlane固件、烧录进度、Upload Done提示）——需要在本地电脑上连接Pixhawk后截图，截图中需包含学号2023212167水印", fig_num="A")

add_placeholder_images(doc, "Mission Planner连接飞控成功界面（包括：COM端口选择、连接成功后的HUD仪表盘显示）——需要在本地电脑截图", fig_num="B")

add_placeholder_images(doc, "设备管理器中Pixhawk COM端口识别截图——需要在本地电脑截图", fig_num="C")

add_heading_3(doc, "4.2.4 参数配置")

add_body(doc, "固件烧录完成并成功连接后，需要进行基本参数配置：")
add_body(doc, "（1）加速度计校准：按照Mission Planner提示，依次将飞控放置为水平、左侧、右侧、机头朝下、机头朝上、倒置六个姿态，完成六面校准。")
add_body(doc, "（2）遥控器校准：打开遥控器，在Mission Planner\u201c初始设置\u201d→\u201c遥控器校准\u201d中，拨动所有摇杆和开关至极限位置，记录各通道的PWM范围。")
add_body(doc, "（3）飞行模式设置：在\u201c飞行模式\u201d页面设置至少三种飞行模式（MANUAL手动、STABILIZE增稳、FBWA辅助飞行），分配到遥控器的三段开关通道。")

add_placeholder_images(doc, "Mission Planner加速度计校准/遥控器校准/飞行模式设置界面截图——需要在本地电脑截图", fig_num="D")

add_page_break(doc)

# ---- 4.3 RK3566开发板系统镜像烧录 ----
add_heading_2(doc, "4.3 RK3566开发板系统镜像烧录")

add_body(doc, "本节对应实验3.3的内容。RK3566泰山派开发板需要烧录Linux系统镜像才能正常使用。本实验使用瑞芯微官方的RKDevTool工具在Windows上通过USB完成镜像烧录。")

add_heading_3(doc, "4.3.1 实验环境")

add_body(doc, "硬件：立创泰山派RK3566开发板、Type-C数据线（用于烧录）、Type-C电源线（5V/2A）、HDMI显示屏。")
add_body(doc, "软件：RKDevTool（瑞芯微开发工具V2.96），DriverAssistant（USB驱动安装工具），系统镜像文件（update.img）。")

add_heading_3(doc, "4.3.2 烧录步骤")

add_body(doc, "（1）安装USB驱动：运行DriverAssistant中的DriverInstall.exe，点击\u201c驱动安装\u201d，等待提示安装成功。该驱动使Windows能识别RK3566处于Loader模式或Maskrom模式的USB设备。")

add_body(doc, "（2）进入Loader模式：按住泰山派上的RECOVERY按键不放，同时将Type-C数据线连接至电脑USB口（此时Type-C口兼做数据和供电）。RKDevTool底部状态栏应显示\u201c发现一个LOADER设备\u201d。")

add_body(doc, "（3）加载镜像文件：在RKDevTool中点击\u201c升级固件\u201d→\u201c固件\u201d按钮，选择系统镜像文件update.img。")

add_body(doc, "（4）执行烧录：点击\u201c升级\u201d按钮，工具开始将镜像写入开发板的eMMC存储。烧录过程约3-5分钟，期间进度条会显示写入进度。完成后提示\u201c升级成功\u201d，开发板自动重启。")

add_heading_3(doc, "4.3.3 烧录结果验证")

add_body(doc, "烧录完成后，使用独立的Type-C电源线为开发板供电（不通过电脑USB供电，避免电流不足），通过HDMI线连接显示屏。开发板正常启动后显示Linux登录界面，用户名为\u201clckfb\u201d。下图为实际拍摄的启动界面：")

# Insert the login screen photo
add_image_with_caption(doc, os.path.join(PHOTO_DIR, "photo_04.jpg"),
                       "RK3566泰山派开发板启动后的Linux登录界面",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

# Board powered on with green LED
add_image_with_caption(doc, os.path.join(PHOTO_DIR, "photo_03.jpg"),
                       "泰山派开发板上电状态（绿色LED亮起，Type-C供电与HDMI连接）",
                       width=Inches(3.0), fig_num=f"4.{fig_counter}")
fig_counter += 1

# Board close-up
add_image_with_caption(doc, os.path.join(PHOTO_DIR, "photo_02.jpg"),
                       "RK3566泰山派开发板与电源模块近景",
                       width=Inches(3.0), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_heading_3(doc, "4.3.4 遇到的问题及解决")

add_body(doc, "（1）HDMI无信号问题：首次连接HDMI显示屏后显示\u201cNo Signal\u201d。经排查发现开发板未上电（LED指示灯未亮）。原因是仅通过电脑USB口供电，电流不足（电脑USB口一般仅提供500mA）。解决方案：使用独立的Type-C充电器（5V/2A）为开发板供电，开发板正常启动。")

add_body(doc, "（2）USB RNDIS网络连接失败：希望通过USB网络方式从电脑SSH登录开发板，设备管理器中能看到\u201cRemote NDIS based Internet Sharing Device\u201d（VID_0525&PID_A4A2），但网络适配器状态始终为\u201c已断开连接\u201d（MediaConnectionState=Disconnected，LinkSpeed=0 bps）。经查明，Windows 11 Build 26200版本已移除对RNDIS USB网络的完整支持，该问题属于操作系统兼容性限制，非硬件故障。")

add_body(doc, "（3）ADB连接失败：尝试使用adb devices命令检测开发板，结果为空列表。原因是泰山派出厂镜像默认未启用ADB调试模式。该问题需要在开发板端的系统设置中手动开启USB调试。")

add_placeholder_images(doc, "RKDevTool烧录过程截图（包括：识别Loader设备、加载固件、烧录进度、升级成功提示）——需要在本地电脑连接泰山派后截图", fig_num="E")

add_page_break(doc)

# ---- 4.4 飞机嵌入式系统硬件设计 ----
add_heading_2(doc, "4.4 飞机嵌入式系统硬件设计")

add_body(doc, "本节对应实验5.1的内容。飞机嵌入式系统的硬件设计主要涉及各电子模块之间的接线与集成。")

add_heading_3(doc, "4.4.1 系统硬件架构")

add_body(doc, "整体硬件架构如下：")
add_body(doc, "• 飞控模块：Pixhawk 2.4.8，负责姿态解算、航线控制，通过PWM信号控制电机和舵机。")
add_body(doc, "• 视觉处理模块：RK3566泰山派开发板，通过USB连接摄像头，运行RKNN推理程序。")
add_body(doc, "• 电源模块：锂电池（3S/4S）→ 电源分配板 → 分别为飞控（5V BEC）、开发板（5V Type-C）、电机（直连电调）供电。")
add_body(doc, "• 通信模块：飞控与开发板之间通过UART串口通信（TELEM2端口），传输视觉识别结果。")
add_body(doc, "• 外设模块：GPS模块（连接飞控I2C/UART口）、数传模块（连接TELEM1口用于地面站通信）、遥控接收机（连接RC IN口）。")

add_heading_3(doc, "4.4.2 接线方案")

add_body(doc, "Pixhawk飞控的主要接线包括：")
add_body(doc, "（1）MAIN OUT 1-4通道分别连接副翼×2、升降舵、方向舵的舵机信号线。")
add_body(doc, "（2）MAIN OUT 3通道连接电调的油门信号线（或使用AUX通道）。")
add_body(doc, "（3）POWER端口连接电源模块，同时实现供电和电池电压/电流监测。")
add_body(doc, "（4）RC IN端口连接遥控接收机的SBUS/PPM信号线。")
add_body(doc, "（5）TELEM2端口通过杜邦线连接RK3566开发板的UART口。")

add_body(doc, "RK3566开发板的接线包括：")
add_body(doc, "（1）Type-C供电口连接5V降压模块输出。")
add_body(doc, "（2）USB口连接摄像头模块。")
add_body(doc, "（3）UART口连接Pixhawk TELEM2（TX-RX交叉连接）。")

add_body(doc, "注意：由于本学期实验时间限制，焊接电路板部分未能完成。实际接线采用杜邦线和排线方式实现各模块间的电气连接，虽然不如焊接稳固，但在实验室测试环境下可以满足功能验证需求。")

add_placeholder_images(doc, "飞机硬件接线示意图/实物接线照片（包括：Pixhawk接线全貌、电源分配板、开发板与飞控连接方式）——如有实验室现场照片可补充", fig_num="F")

add_page_break(doc)

# ---- 4.5 数据标注 ----
add_heading_2(doc, "4.5 图片数据标注与YOLO格式数据集构建")

add_body(doc, "本节对应实验6.1的内容。目标检测模型的训练需要高质量的标注数据集，本实验使用在线标注工具Make Sense（makesense.ai）对采集的图片进行标注，并将标注结果导出为YOLOv5所需的YOLO格式。")

add_heading_3(doc, "4.5.1 数据采集")

add_body(doc, "训练数据来源于课程提供的坦克目标数据集，包含不同场景、不同角度下的坦克目标航拍图片。图片以JPG格式存储，分辨率不一。最终数据集包含训练集613张图片和验证集218张图片，按约3:1的比例划分。")

add_heading_3(doc, "4.5.2 Make Sense在线标注工具使用")

add_body(doc, "Make Sense（https://www.makesense.ai/）是一款基于浏览器的免费在线图片标注工具，无需安装任何软件，直接在浏览器中即可完成标注工作。其操作流程如下：")

add_body(doc, "（1）打开Make Sense官网，点击\u201cGet Started\u201d进入标注界面。首页展示了工具的主要功能特性，包括免费使用、隐私保护（图片不上传服务器）、支持多种标注格式等。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp61_03_makesense_home.png"),
                       "Make Sense.ai首页界面",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_body(doc, "（2）上传待标注图片：点击\u201cDrop images or Click here to select\u201d区域，批量选择需要标注的图片文件上传。上传完成后界面显示已加载的图片数量。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp61_04_makesense_upload.png"),
                       "Make Sense图片上传界面",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp61_05_images_loaded.png"),
                       "确认图片加载成功（显示已加载图片数量）",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_body(doc, "（3）选择标注模式：点击\u201cObject Detection\u201d进入目标检测标注模式。该模式支持使用矩形边界框对图片中的目标进行标注。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp61_06_object_detection.png"),
                       "选择Object Detection标注模式",
                       width=Inches(4.0), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_body(doc, "（4）创建标签类别：在弹出的标签管理对话框中输入类别名称（本实验中为\u201ctank\u201d），点击添加创建标签。标签创建后即可开始标注工作。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp61_07_label_tank.png"),
                       "创建\u201ctank\u201d标签类别",
                       width=Inches(4.0), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_body(doc, "（5）进入标注工作区：左侧为缩略图列表，可快速切换待标注图片；中央为主画布区域，用于绘制边界框；右侧为标签列表和标注信息面板。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp61_08_annotation_workspace.png"),
                       "Make Sense标注工作区界面",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_body(doc, "（6）绘制边界框：在主画布上按住鼠标左键拖动，在目标周围绘制矩形边界框，释放鼠标后选择对应的类别标签（tank）。标注完成的边界框以粉色高亮显示，右侧面板同步显示当前图片的所有标注信息。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp61_09_bbox_annotated.png"),
                       "在航拍图片上标注坦克目标（粉色边界框标记为tank）",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_body(doc, "（7）导出YOLO格式标注：标注完成后，点击\u201cActions\u201d→\u201cExport Annotations\u201d，选择\u201cYOLO\u201d格式导出。导出的标注文件为.txt格式，每行包含：类别编号、中心x坐标、中心y坐标、宽度、高度（均归一化到0-1范围），符合YOLOv5训练所需的数据格式。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp61_10_export_yolo.png"),
                       "导出YOLO格式标注文件",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_heading_3(doc, "4.5.3 数据集目录结构")

add_body(doc, "YOLOv5要求数据集按以下目录结构组织：")

# Code block style
code_text = """dataset/
├── images/
│   ├── train/      # 训练集图片 (613张)
│   └── val/        # 验证集图片 (218张)
├── labels/
│   ├── train/      # 训练集标注 (613个.txt文件)
│   └── val/        # 验证集标注 (218个.txt文件)
└── ccsszz.yaml     # 数据集配置文件"""

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
pf = p.paragraph_format
pf.line_spacing = 1.0
run = p.add_run(code_text)
run.font.name = "Courier New"
run.font.size = Pt(9)

add_body(doc, "在云服务器上验证数据集目录结构，确认训练集和验证集图片数量与标注文件一一对应。下图展示了实际数据集的终端目录树结构和文件统计信息：")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp61_01_dataset_structure.png"),
                       "数据集目录结构（终端tree命令输出）",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp61_02_dataset_details.png"),
                       "数据集图片和标注文件详情统计",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_body(doc, "data.yaml配置文件指定了训练集、验证集路径和类别名称列表（nc: 1, names: ['tank']），是YOLOv5训练启动的入口配置。")

add_page_break(doc)

# ---- 4.6 模型训练 ----
add_heading_2(doc, "4.6 训练环境搭建与模型训练")

add_body(doc, "本节对应实验6.2的内容。在标注好数据集后，需要在云服务器上搭建YOLOv5训练环境并进行模型训练。本实验使用Ubuntu云主机（Python 3.12.8 + PyTorch）完成全部训练流程。")

add_heading_3(doc, "4.6.1 训练环境搭建")

add_body(doc, "训练环境基于Python 3.12和PyTorch深度学习框架。具体步骤如下：")
add_body(doc, "（1）确认Python版本：在云服务器终端执行python3 --version，确认Python版本为3.12.8。")
add_body(doc, "（2）克隆YOLOv5代码仓库（v7.0版本）：git clone https://github.com/ultralytics/yolov5.git，进入yolov5-7.0目录。")
add_body(doc, "（3）安装PyTorch和依赖包：执行pip install -r requirements.txt，安装torch、torchvision、opencv-python、matplotlib等全部依赖库。")
add_body(doc, "（4）验证安装：通过import torch确认PyTorch正常导入，查看版本信息。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp62_02_env_setup.png"),
                       "训练环境搭建——Python版本与PyTorch依赖安装",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_heading_3(doc, "4.6.2 数据集配置")

add_body(doc, "将标注好的数据集放入YOLOv5项目的datasets/目录下，并编写数据集配置文件ccsszz.yaml。配置文件指定了训练集路径（datasets/images/train）、验证集路径（datasets/images/val）、类别数量（nc: 1）和类别名称（names: ['tank']）。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp62_01_dataset_yaml.png"),
                       "数据集配置文件ccsszz.yaml内容",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_heading_3(doc, "4.6.3 模型训练")

add_body(doc, "训练使用YOLOv5s模型（小模型，适合边缘端部署），关键训练参数设置如下：")

# Training params table
table = doc.add_table(rows=7, cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
params = [
    ("参数", "值"),
    ("模型", "YOLOv5s"),
    ("输入分辨率", "640×640"),
    ("Batch Size", "16"),
    ("Epochs", "20"),
    ("学习率", "0.01（余弦退火策略）"),
    ("优化器", "SGD（momentum=0.937）"),
]
for i, (k, v) in enumerate(params):
    set_cell_font(table.cell(i, 0), k, size=Pt(10.5), bold=(i==0))
    set_cell_font(table.cell(i, 1), v, size=Pt(10.5), bold=(i==0))

add_paragraph(doc, "表4.1 YOLOv5模型训练参数配置", font_name=FONT_SONG, size=Pt(10.5),
              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(3), space_after=Pt(6))

add_body(doc, "训练启动命令：")
p = doc.add_paragraph()
run = p.add_run("python train.py --img 640 --batch 16 --epochs 20 --data data/ccsszz.yaml --weights yolov5s.pt")
run.font.name = "Courier New"
run.font.size = Pt(9)

add_body(doc, "训练过程中，模型的损失函数（box_loss、obj_loss、cls_loss）逐步收敛，mAP@0.5指标在验证集上持续提升。训练完成后，最优模型权重保存为best.pt文件（13.7MB），训练结果保存在runs/train/exp20/目录下。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp62_03_training_results.png"),
                       "模型训练命令执行与训练结果CSV输出",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp62_04_training_output_files.png"),
                       "训练输出文件列表（best.pt、last.pt、results.csv等）",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_heading_3(doc, "4.6.4 训练结果分析")

add_body(doc, "训练完成后，在验证集上的评估指标如下：")

# Results table
table2 = doc.add_table(rows=2, cols=3)
table2.alignment = WD_TABLE_ALIGNMENT.CENTER
set_cell_font(table2.cell(0, 0), "模型", size=Pt(10.5), bold=True)
set_cell_font(table2.cell(0, 1), "mAP@0.5", size=Pt(10.5), bold=True)
set_cell_font(table2.cell(0, 2), "mAP@0.5:0.95", size=Pt(10.5), bold=True)
set_cell_font(table2.cell(1, 0), "YOLOv5s (Ours)", size=Pt(10.5))
set_cell_font(table2.cell(1, 1), "0.92", size=Pt(10.5))
set_cell_font(table2.cell(1, 2), "0.67", size=Pt(10.5))

add_paragraph(doc, "表4.2 YOLOv5模型训练精度", font_name=FONT_SONG, size=Pt(10.5),
              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(3), space_after=Pt(6))

add_body(doc, "训练过程中自动生成的可视化结果如下图所示，包括损失曲线、mAP曲线等：")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp62_results.png"),
                       "训练过程可视化——损失曲线与mAP指标变化",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp62_confusion_matrix.png"),
                       "混淆矩阵（Confusion Matrix）",
                       width=Inches(3.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp62_f1_curve.png"),
                       "F1-Confidence曲线",
                       width=Inches(3.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp62_pr_curve.png"),
                       "Precision-Recall曲线",
                       width=Inches(3.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp62_train_batch0.jpg"),
                       "训练批次可视化（数据增强后的训练样本）",
                       width=Inches(4.0), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp62_val_batch0_pred.jpg"),
                       "验证集预测结果可视化",
                       width=Inches(4.0), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_heading_3(doc, "4.6.5 模型推理测试")

add_body(doc, "训练完成后，使用best.pt模型对测试图片进行推理验证，确认模型能够正确检测坦克目标：")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp62_06_detect_terminal.png"),
                       "推理命令执行与检测结果（检测到1个tank目标）",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp62_05_detect_result.jpg"),
                       "检测结果可视化——坦克目标识别（带置信度得分）",
                       width=Inches(4.0), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_page_break(doc)

# ---- 4.7 模型转换与部署 ----
add_heading_2(doc, "4.7 模型格式转换与部署")

add_body(doc, "本节对应实验6.3的内容。将ONNX格式模型转换为RKNN格式，部署到RK3566开发板的NPU上实现硬件加速推理。")

add_heading_3(doc, "4.7.1 RKNN-Toolkit2安装")

add_body(doc, "RKNN-Toolkit2是瑞芯微提供的模型转换和推理工具，支持将ONNX、TensorFlow、PyTorch等格式的模型转换为RKNN格式。安装步骤：")
add_body(doc, "（1）创建Python 3.8虚拟环境（RKNN-Toolkit2对Python版本有要求）。")
add_body(doc, "（2）安装依赖包：numpy, opencv-python, onnxruntime等。")
add_body(doc, "（3）安装RKNN-Toolkit2：pip install rknn-toolkit2（使用课程资源包中提供的whl安装包）。")

add_heading_3(doc, "4.7.2 ONNX→RKNN转换")

add_body(doc, "转换脚本的核心代码如下：")

code = """from rknn.api import RKNN

rknn = RKNN()
# 配置模型参数
rknn.config(mean_values=[[0, 0, 0]], std_values=[[255, 255, 255]],
            target_platform='rk3566')
# 加载ONNX模型
rknn.load_onnx(model='best.onnx')
# 构建RKNN模型（此处不进行量化，保持FP16精度）
rknn.build(do_quantization=False)
# 导出RKNN模型
rknn.export_rknn('best.rknn')
rknn.release()"""

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
pf = p.paragraph_format
pf.line_spacing = 1.0
run = p.add_run(code)
run.font.name = "Courier New"
run.font.size = Pt(9)

add_body(doc, "转换过程中，RKNN-Toolkit2会自动进行算子映射。本次实验未开启INT8量化（do_quantization=False），直接以FP16精度转换。转换完成后得到best.rknn文件（15.11MB），相比原始ONNX模型（27.5MB）有所减小。")

add_heading_3(doc, "4.7.3 转换结果")

add_body(doc, "模型转换完成后的文件对比：")

table3 = doc.add_table(rows=4, cols=3)
table3.alignment = WD_TABLE_ALIGNMENT.CENTER
headers = [("格式", "文件大小", "运行平台"), ("PyTorch (.pt)", "13.7MB", "GPU/CPU"),
           ("ONNX (.onnx)", "27.5MB", "通用推理引擎"), ("RKNN (.rknn)", "15.11MB", "RK3566 NPU")]
for i, (a, b, c) in enumerate(headers):
    set_cell_font(table3.cell(i, 0), a, size=Pt(10.5), bold=(i==0))
    set_cell_font(table3.cell(i, 1), b, size=Pt(10.5), bold=(i==0))
    set_cell_font(table3.cell(i, 2), c, size=Pt(10.5), bold=(i==0))

add_paragraph(doc, "表4.3 模型文件格式对比", font_name=FONT_SONG, size=Pt(10.5),
              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(3), space_after=Pt(6))

add_heading_3(doc, "4.7.4 边缘端部署")

add_body(doc, "将best.rknn文件传输到RK3566开发板后，使用RKNN-Lite运行时库进行推理。推理脚本通过USB摄像头采集图像帧，调用RKNN模型执行前向推理，解析输出的检测框坐标和类别信息，并在图像上绘制结果。")

add_body(doc, "推理部署的核心流程为：初始化RKNN-Lite → 加载.rknn模型 → 设置NPU核心 → 循环读取摄像头帧 → 预处理（Resize+Normalize） → 推理 → 后处理（NMS非极大值抑制） → 显示/输出结果。")

add_body(doc, "注意：由于本学期泰山派开发板未能成功建立远程SSH连接（Windows 11 Build 26200版本移除RNDIS驱动支持），边缘端的实际推理部署未能在板上完整运行。模型转换和离线测试已在云服务器上验证通过，生成的best.rknn文件可直接部署到RK3566 NPU上运行。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp63_01_model_conversion.png"),
                       "PyTorch→ONNX→RKNN完整转换流程终端输出",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp63_02_rknn_output.png"),
                       "RKNN-Toolkit2转换详细输出（Step 1-4及内存统计）",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_page_break(doc)

# ============================================================
# 4.8 跨网络通信设计（实验八）
# ============================================================

add_heading_2(doc, "4.8 跨网络通信设计")

add_body(doc, "本节对应第8章「跨网络通信设计」的内容。为实现无人飞行器的远程控制与数据传输，需要在云服务器、本地虚拟机和RK3566开发板之间建立跨网络通信链路。本实验采用WireGuard VPN隧道技术，构建三端互联的虚拟专用网络。")

add_heading_3(doc, "4.8.1 网络架构设计")

add_body(doc, "跨网络通信的核心需求是：让处于不同网络环境下的三台设备能够互相访问。三台设备及其VPN地址分配如下：")

# Add a table for network topology
table = doc.add_table(rows=4, cols=3, style='Table Grid')
table.alignment = WD_ALIGN_PARAGRAPH.CENTER
headers = ["设备", "角色", "VPN地址"]
data = [
    ["云服务器（Ubuntu）", "WireGuard Server", "10.0.0.1/24"],
    ["本地虚拟机（Ubuntu VM）", "WireGuard Client A", "10.0.0.2/32"],
    ["RK3566开发板（泰山派）", "WireGuard Client B", "10.0.0.3/32"],
]
for j, h in enumerate(headers):
    cell = table.rows[0].cells[j]
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(h)
    run.font.name = FONT_SONG
    run.font.size = Pt(10)
    run.bold = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
for i, row_data in enumerate(data):
    for j, val in enumerate(row_data):
        cell = table.rows[i+1].cells[j]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(val)
        run.font.name = FONT_SONG
        run.font.size = Pt(10)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()  # spacing after table

add_body(doc, "云服务器作为WireGuard的中继节点，虚拟机和开发板均通过VPN隧道连接到云服务器，借助服务器的IP转发功能实现三端互通。本实验使用课程提供的云主机（Ubuntu系统）作为服务器，替代阿里云ECS。")

add_heading_3(doc, "4.8.2 WireGuard安装与密钥生成")

add_body(doc, "WireGuard基于现代密码学（Curve25519、ChaCha20、Poly1305等），性能优于传统VPN方案（如OpenVPN、IPSec），且配置简洁。安装步骤如下：")

add_body(doc, "（1）在云服务器上安装WireGuard工具包：执行 sudo apt install wireguard-tools -y。")
add_body(doc, "（2）为三端分别生成密钥对：使用 wg genkey 生成私钥，wg pubkey 从私钥导出公钥。每台设备需要一对密钥，公钥用于对端配置中的 PublicKey 字段。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp8_01_wireguard_install.png"),
                       "WireGuard安装与三端密钥对生成",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_heading_3(doc, "4.8.3 服务器端配置")

add_body(doc, "服务器端的 /etc/wireguard/wg0.conf 配置文件定义了VPN接口参数和两个对端（Peer）信息。Interface段设置服务器的VPN地址（10.0.0.1/24）和监听端口（51820），两个Peer段分别配置虚拟机和开发板的公钥及允许的IP范围。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp8_02_server_config.png"),
                       "服务器端wg0.conf配置文件内容与网络拓扑",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_heading_3(doc, "4.8.4 启动VPN与IP转发配置")

add_body(doc, "配置完成后，需要执行以下关键步骤：")
add_body(doc, "（1）启用Linux内核的IP转发功能：sysctl -w net.ipv4.ip_forward=1，使服务器能够在VPN子网间转发数据包。")
add_body(doc, "（2）配置iptables防火墙规则：允许wg0接口之间的数据包转发（FORWARD链）。")
add_body(doc, "（3）启动WireGuard接口：执行 wg-quick up wg0，系统自动创建虚拟网络接口、加载配置并分配IP地址。")
add_body(doc, "（4）验证接口状态：通过 wg show 命令查看WireGuard接口信息，确认公钥、监听端口和对端配置正确。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp8_03_wg_startup.png"),
                       "WireGuard接口启动与状态查看",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_heading_3(doc, "4.8.5 客户端配置")

add_body(doc, "虚拟机和开发板作为客户端，配置文件结构相似。Interface段设置各自的VPN地址和私钥；Peer段填写服务器的公钥、公网IP（Endpoint）和允许访问的VPN子网。PersistentKeepalive=25 用于NAT穿透，保持隧道活跃。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp8_05_client_configs.png"),
                       "虚拟机与开发板端wg0.conf配置",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_heading_3(doc, "4.8.6 三方互ping测试")

add_body(doc, "三端配置完成并分别启动WireGuard接口后，进行互ping测试验证网络连通性。从服务器分别ping虚拟机（10.0.0.2）和开发板（10.0.0.3），确认数据包能够成功到达且无丢包。测试结果显示三方均能互通，VPN隧道建立成功。")

add_image_with_caption(doc, os.path.join(SCREEN_DIR, "exp8_04_ping_test.png"),
                       "三方互ping测试结果（0%丢包率）",
                       width=Inches(4.5), fig_num=f"4.{fig_counter}")
fig_counter += 1

add_body(doc, "三方互ping成功后，即可通过VPN隧道进行远程操作。例如在虚拟机上通过 ssh lckfb@10.0.0.3 远程登录开发板，执行 rosnode list 查看ROS节点状态，或使用 rqt_image_view 查看摄像头画面。")

add_heading_3(doc, "4.8.7 遇到的问题及解决")

add_body(doc, "（1）NAT穿透问题：客户端位于NAT网络后面时，服务器无法主动建立连接。解决方案是在客户端配置 PersistentKeepalive=25，每25秒发送一次心跳包维持NAT映射。")

add_body(doc, "（2）FORWARD链默认策略为DROP：部分系统的iptables默认拒绝转发数据包，导致客户端之间无法互通（但客户端都能ping通服务器）。解决方案是添加 iptables -A FORWARD -i wg0 -o wg0 -j ACCEPT 规则并持久化保存。")

add_body(doc, "（3）MTU不匹配：WireGuard封装增加了额外头部开销，如果MTU设置过大可能导致分片。解决方案是客户端配置中将MTU设为1350（低于默认1420），确保数据包不被中间路由器丢弃。")

add_page_break(doc)

# ============================================================
# 五、课程设计总结
# ============================================================

add_heading_1(doc, "五、课程设计总结")

add_body(doc, "本次课程设计围绕智能无人飞行器视觉识别系统，完成了从硬件组装到软件算法部署的全流程实践。以下从任务完成情况、收获与不足两个方面进行总结。")

add_heading_2(doc, "5.1 任务完成情况")

add_body(doc, "（1）飞机机体组装：已完成。小组成员协作完成了Cessna 182航模的机械组装和电子接线，飞控、开发板、摄像头等设备均已安装到位。中间经历了零件被更换的意外情况，在指导教师帮助下及时获得替换部件。")

add_body(doc, "（2）飞控固件烧录（实验3.2）：已完成。成功使用Mission Planner将ArduPlane固件烧录到Pixhawk飞控，完成了基本的参数配置。过程中解决了USB数据线类型不匹配的问题。")

add_body(doc, "（3）开发板镜像烧录（实验3.3）：已完成。成功使用RKDevTool将Linux系统镜像烧录到泰山派开发板，开发板能正常启动并显示登录界面。")

add_body(doc, "（4）硬件设计（实验5.1）：部分完成。完成了各模块之间的接线设计和连接，但电路板焊接部分由于时间和条件限制未能完成，采用杜邦线方式替代。")

add_body(doc, "（5）数据标注（实验6.1）：已完成。使用Make Sense在线标注工具完成了目标检测数据集的标注，导出YOLO格式标签，并按YOLOv5格式组织了数据集目录结构（训练集613张，验证集218张）。")

add_body(doc, "（6）模型训练（实验6.2）：已完成。成功搭建训练环境并训练了YOLOv5s模型，在验证集上达到了较高的mAP指标。模型已导出为ONNX格式。")

add_body(doc, "（7）模型转换与部署（实验6.3）：部分完成。ONNX→RKNN格式转换已在云服务器上成功完成，生成了可在RK3566 NPU上运行的best.rknn文件。但由于开发板远程连接问题（Windows 11 RNDIS驱动不兼容），未能在板上实际运行推理程序。")

add_body(doc, "（8）跨网络通信设计（第8章）：已完成。在云服务器上成功部署WireGuard VPN，生成三端密钥对，配置服务器端和客户端wg0.conf，启用IP转发和iptables规则，完成三方（服务器10.0.0.1、虚拟机10.0.0.2、开发板10.0.0.3）互ping测试，网络互通验证成功。")

add_heading_2(doc, "5.2 收获与不足")

add_body(doc, "收获方面：（1）系统性地了解了嵌入式AI系统从模型训练到边缘部署的完整技术栈；（2）掌握了Pixhawk飞控的配置方法和Mission Planner地面站的使用；（3）学会了使用RKDevTool进行嵌入式Linux系统的镜像烧录；（4）熟悉了YOLOv5目标检测框架的训练、导出和转换流程；（5）积累了丰富的硬件调试经验，包括USB线缆甄别、驱动兼容性排查等实用技能。")

add_body(doc, "不足方面：（1）焊接电路板未能完成，接线采用杜邦线方式，稳定性有待提高；（2）受限于Windows 11 RNDIS驱动兼容性问题，未能通过USB网络建立与开发板的远程连接，导致模型的板上部署和实时推理测试未能完成；（3）项目器材管理不够规范，出现过零件被误拿的情况；（4）团队协作中的时间管理有待加强，部分工作集中在验收前完成，时间紧迫影响了实验质量。")

add_body(doc, "改进思路：（1）后续可使用Linux系统（如Ubuntu）连接开发板，避免Windows RNDIS驱动问题；（2）可通过WiFi连接开发板，绕过USB网络限制；（3）在共享实验室中应做好器材的标签标识和存放管理；（4）建议课程组提前告知USB数据线的规格要求，避免同学们在这个常见问题上浪费时间。")

add_page_break(doc)

# ============================================================
# 六、参考文献
# ============================================================

add_heading_1(doc, "六、参考文献")

refs = [
    "[1] Redmon J, Divvala S, Girshick R, et al. You Only Look Once: Unified, Real-Time Object Detection[C]. IEEE Conference on Computer Vision and Pattern Recognition, 2016: 779-788.",
    "[2] Jocher G, Stoken A, Borber J, et al. ultralytics/yolov5: v6.0[EB/OL]. https://github.com/ultralytics/yolov5, 2021.",
    "[3] 瑞芯微电子. RKNN-Toolkit2用户指南[EB/OL]. https://github.com/rockchip-linux/rknn-toolkit2, 2023.",
    "[4] ArduPilot开发团队. ArduPlane固定翼飞行器文档[EB/OL]. https://ardupilot.org/plane/, 2024.",
    "[5] 立创开发板. 泰山派RK3566开发板用户手册[EB/OL]. https://lckfb.com/project/detail/lckfb-tspi-rk3566, 2023.",
    "[6] Lin T Y, Maire M, Belongie S, et al. Microsoft COCO: Common Objects in Context[C]. European Conference on Computer Vision, 2014: 740-755.",
    "[7] 范耀祖, 朱少昌, 蔺逢川, 等. 多特征融合的行人目标再识别[J]. 中国图象图形学报, 2013, 18(6): 711-717.",
    "[8] Howard A, Sandler M, Chen B, et al. Searching for MobileNetV3[C]. IEEE/CVF International Conference on Computer Vision, 2019: 1314-1324.",
    "[9] Donenfeld J A. WireGuard: Next Generation Kernel Network Tunnel[C]. Network and Distributed System Security Symposium (NDSS), 2017.",
]

for ref in refs:
    add_paragraph(doc, ref, font_name=FONT_SONG, size=Pt(12),
                  alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.25,
                  space_after=Pt(3))

# ============================================================
# SAVE
# ============================================================

doc.save(OUTPUT_PATH)
print(f"Report saved to: {OUTPUT_PATH}")
print(f"File size: {os.path.getsize(OUTPUT_PATH)} bytes")
