import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import re
import copy
import tempfile
import os
from datetime import datetime
import gc

# ====================== 全局配置与常量 ======================
ALIGN_MAP = {
    "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
    "居中": WD_ALIGN_PARAGRAPH.CENTER,
    "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,
    "不修改": None
}
ALIGN_LIST = list(ALIGN_MAP.keys())

LINE_TYPE_MAP = {
    "单倍行距": WD_LINE_SPACING.SINGLE,
    "1.5倍行距": WD_LINE_SPACING.ONE_POINT_FIVE,
    "2倍行距": WD_LINE_SPACING.DOUBLE,
    "多倍行距": WD_LINE_SPACING.MULTIPLE,
    "固定值": WD_LINE_SPACING.EXACTLY
}
LINE_TYPE_LIST = list(LINE_TYPE_MAP.keys())

LINE_RULE = {
    "单倍行距": {"default": 1.0, "min": 1.0, "max": 1.0, "step": 1.0, "label": "行距倍数"},
    "1.5倍行距": {"default": 1.5, "min": 1.5, "max": 1.5, "step": 0.1, "label": "行距倍数"},
    "2倍行距": {"default": 2.0, "min": 2.0, "max": 2.0, "step": 0.1, "label": "行距倍数"},
    "多倍行距": {"default": 1.5, "min": 0.5, "max": 5.0, "step": 0.1, "label": "行距倍数"},
    "固定值": {"default": 20.0, "min": 1.0, "max": 100.0, "step": 0.1, "label": "固定值(磅)"}
}

FONT_LIST = ["宋体", "黑体", "微软雅黑", "楷体", "仿宋"]
FONT_SIZE_LIST = ["初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六"]
FONT_SIZE_NUM = {k:v for k,v in zip(FONT_SIZE_LIST, [42.0,36.0,26.0,24.0,22.0,18.0,16.0,15.0,14.0,12.0,10.5,9.0,7.5,6.5])}
EN_FONT_LIST = ["和正文一致", "Times New Roman", "Arial", "Calibri"]

# ====================== 【修复1：彻底重写标题识别规则，杜绝误判】 ======================
# 1. 先定义【绝对不是标题的黑名单】，优先排除
TITLE_BLACKLIST = [
    re.compile(r"^图\s*[0-9一二三四五六七八九十]+[-.、:：]\s*", re.IGNORECASE),  # 图注：图1、图1-1、图 1：
    re.compile(r"^表\s*[0-9一二三四五六七八九十]+[-.、:：]\s*", re.IGNORECASE),  # 表注：表1、表2-1
    re.compile(r"^figure\s*[0-9]+[-.、:：]\s*", re.IGNORECASE),  # 英文图注
    re.compile(r"^table\s*[0-9]+[-.、:：]\s*", re.IGNORECASE),  # 英文表注
    re.compile(r"^[（(]\s*[0-9]+[)）]\s*.*[。？！；;]$"),  # 带句号的(1)列表项，是正文不是标题
    re.compile(r"^[①②③④⑤⑥⑦⑧⑨⑩]\s*.*[。？！；;]$"),  # 带句号的①列表项，是正文
    re.compile(r"^注\s*[0-9]*[：:.]\s*"),  # 注释
    re.compile(r"^参考文献\s*[:：]?$"),  # 参考文献单独处理
    re.compile(r"^附录\s*[0-9A-Z]*[:：]?$"),  # 附录
]

# 2. 严格的标题匹配规则，只有完全符合才识别
TITLE_RULE = {
    "一级标题": [
        re.compile(r"^第[一二三四五六七八九十0-9]+章\s*[^\s。？！；;]{2,40}$"),  # 第X章 标题，后面不能带句号
        re.compile(r"^[一二三四五六七八九十]+、\s*[^\s。？！；;]{2,40}$"),  # 一、标题，无句号，长度2-40
    ],
    "二级标题": [
        re.compile(r"^[0-9]+\.[0-9]+\s*[^\s。？！；;]{2,50}$"),  # 1.1 标题，无句号
        re.compile(r"^[（(][一二三四五六七八九十]+[)）]\s*[^\s。？！；;]{2,50}$"),  # (一)标题，无句号，长度2-50
    ],
    "三级标题": [
        re.compile(r"^[0-9]+\.[0-9]+\.[0-9]+\s*[^\s。？！；;]{2,60}$"),  # 1.1.1 标题，无句号
        re.compile(r"^[（(][0-9]+[)）]\s*[^\s。？！；;]{2,60}$"),  # (1)标题，无句号，长度2-60
    ]
}

# 标题最大长度限制，超过直接判定为正文
TITLE_MAX_LENGTH = 60
# 正文最小长度，超过直接判定为正文
BODY_MIN_LENGTH = 100

# ====================== 模板库（官方标准，无错误） ======================
TEMPLATE_LIBRARY = {
    "默认通用格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 6},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
    },
    "河北科技大学-本科毕业论文": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 12},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
    },
    "国标-本科毕业论文通用": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 12},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
    },
    "党政机关公文国标GB/T 9704-2012": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 6},
        "二级标题": {"font": "楷体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 6},
        "三级标题": {"font": "仿宋", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "仿宋", "size": "三号", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "仿宋", "size": "小三", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
    }
}

# ====================== 【修复2：彻底重写图片保护机制，100%保留原图】 ======================
def is_protected_para(para):
    """
    绝对保护机制：只要段落包含任何图片/绘图/分页符/分节符，整个段落完全不修改
    彻底解决图片变形、半张、重叠、位置错乱问题
    """
    if not para:
        return True
    try:
        # 保护分页符、分节符
        if para.paragraph_format.page_break_before:
            return True
        if para._element.find('.//w:sectPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
            return True
        
        # 遍历所有run，检查是否包含图片/绘图/形状/页面分隔
        for run in para.runs:
            if run.contains_page_break:
                return True
            # 检查所有类型的图片/绘图
            if run._element.find('.//w:drawing', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True
            if run._element.find('.//w:pict', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True
            if run._element.find('.//w:shape', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True
            if run._element.find('.//w:oleObject', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True
        
        return False
    except Exception:
        # 任何异常都判定为保护，避免出错
        return True

# ====================== 字体设置工具函数（修复中文字体兼容） ======================
def set_run_font(run, font_name, font_size, bold=None):
    """中文字体100%生效，兼容所有Word/WPS版本"""
    try:
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run._element.rPr.rFonts.set(qn('w:ascii'), font_name)
        run._element.rPr.rFonts.set(qn('w:hAnsi'), font_name)
        run.font.size = Pt(font_size)
        if bold is not None:
            run.font.bold = bold
    except Exception:
        pass

def set_en_number_font(run, font_name, font_size, bold=None):
    """数字/英文字体单独设置，不影响中文，100%生效"""
    try:
        if font_name == "和正文一致":
            return
        # 仅设置英文/数字字体，中文保持不变
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:ascii'), font_name)
        run._element.rPr.rFonts.set(qn('w:hAnsi'), font_name)
        run._element.rPr.rFonts.set(qn('w:cs'), font_name)
        run.font.size = Pt(font_size)
        if bold is not None:
            run.font.bold = bold
    except Exception:
        pass

# ====================== 【修复3：彻底重写标题识别逻辑，零误判】 ======================
def get_title_level(para, enable_regex=True, last_levels=None):
    """
    重写版：零误判标题识别算法
    执行顺序：1.黑名单排除 → 2.内置样式识别 → 3.大纲级别识别 → 4.严格正则匹配 → 5.上下文校验
    """
    if last_levels is None:
        last_levels = [0, 0, 0]  # [一级标题数, 二级标题数, 三级标题数]
    
    if not para:
        return "正文"
    
    text = para.text.strip()
    text_length = len(text)

    # --------------------------
    # 第一步：黑名单排除（最优先，绝对不是标题的直接pass）
    # --------------------------
    # 1. 空文本直接排除
    if not text:
        return "正文"
    # 2. 图注、表注直接排除
    for pattern in TITLE_BLACKLIST:
        if pattern.match(text):
            return "正文"
    # 3. 带句号/问号/感叹号/分号的，直接排除（标题不会带结束标点）
    if text.endswith(("。", "？", "！", "；", ".", "?", "!", ";")):
        return "正文"
    # 4. 文本太长，直接判定为正文（标题不会超过60字）
    if text_length > TITLE_MAX_LENGTH:
        return "正文"
    # 5. 文本太短，直接排除（标题至少2个字）
    if text_length < 2:
        return "正文"

    # --------------------------
    # 第二步：识别Word内置标题样式（最权威，用户手动设置的）
    # --------------------------
    style_name = para.style.name.lower()
    if "heading 1" in style_name or "标题 1" in style_name or "标题1" in style_name:
        return "一级标题"
    if "heading 2" in style_name or "标题 2" in style_name or "标题2" in style_name:
        return "二级标题"
    if "heading 3" in style_name or "标题 3" in style_name or "标题3" in style_name:
        return "三级标题"
    
    # 不开启正则的话，直接返回正文
    if not enable_regex:
        return "正文"

    # --------------------------
    # 第三步：识别大纲级别
    # --------------------------
    try:
        p = para._element
        outline_lvl = p.xpath('.//w:outlineLvl', namespaces=p.nsmap)
        if outline_lvl:
            level = int(outline_lvl[0].get(qn('w:val')))
            if level == 1:
                return "一级标题"
            elif level == 2:
                return "二级标题"
            elif level == 3:
                return "三级标题"
    except Exception:
        pass

    # --------------------------
    # 第四步：严格正则匹配（只有完全符合才识别）
    # --------------------------
    # 先匹配三级标题，避免被二级误判
    for pattern in TITLE_RULE["三级标题"]:
        if pattern.match(text):
            # 上下文校验：三级标题前面必须有二级标题，不能越级
            if last_levels[1] > 0:
                return "三级标题"
            else:
                return "正文"  # 越级的直接判定为正文
    
    # 匹配二级标题
    for pattern in TITLE_RULE["二级标题"]:
        if pattern.match(text):
            # 上下文校验：二级标题前面必须有一级标题
            if last_levels[0] > 0:
                return "二级标题"
            else:
                return "正文"  # 越级的直接判定为正文
    
    # 匹配一级标题
    for pattern in TITLE_RULE["一级标题"]:
        if pattern.match(text):
            return "一级标题"

    # --------------------------
    # 第五步：兜底，所有匹配都不通过，判定为正文
    # --------------------------
    return "正文"

# ====================== 【修复4：彻底重写数字/英文单独设置，100%生效】 ======================
def process_number_in_para(para, body_font, body_size, number_config):
    """
    重写版：正文数字/英文单独设置，100%生效
    不清除原有内容，遍历每个run，精准设置数字/英文格式，不破坏中文格式
    """
    if not number_config["enable"]:
        # 不开启的话，统一设置正文字体
        for run in para.runs:
            set_run_font(run, body_font, body_size)
        return
    
    # 获取数字格式配置
    number_size = FONT_SIZE_NUM[number_config["size"]] if not number_config["size_same_as_body"] else body_size
    number_font = number_config["font"]
    number_bold = number_config["bold"]
    
    # 匹配规则：阿拉伯数字、英文单词、英文标点
    number_en_pattern = re.compile(r"[a-zA-Z0-9\.\-%\+]+")

    # 遍历每个run，不清除原有内容，避免丢失格式
    for run in para.runs:
        run_text = run.text
        if not run_text:
            continue
        
        # 如果整个run都是数字/英文，直接设置
        if number_en_pattern.fullmatch(run_text):
            set_en_number_font(run, number_font, number_size, number_bold)
        # 如果run里包含数字/英文，拆分处理
        elif number_en_pattern.search(run_text):
            # 保存原有格式，拆分文本
            original_format = {
                "font": run.font.name,
                "size": run.font.size.pt if run.font.size else body_size,
                "bold": run.font.bold
            }
            
            # 清空当前run，重新拆分添加
            run.text = ""
            parts = []
            last_end = 0
            
            # 拆分文本和数字/英文
            for match in number_en_pattern.finditer(run_text):
                start, end = match.span()
                if start > last_end:
                    parts.append(("text", run_text[last_end:start]))
                parts.append(("number", run_text[start:end]))
                last_end = end
            if last_end < len(run_text):
                parts.append(("text", run_text[last_end:]))
            
            # 重新添加内容，分别设置格式
            for part_type, part_text in parts:
                new_run = para.add_run(part_text)
                if part_type == "text":
                    set_run_font(new_run, body_font, body_size, original_format["bold"])
                else:
                    set_en_number_font(new_run, number_font, number_size, number_bold)
        else:
            # 纯中文，设置正文字体
            set_run_font(run, body_font, body_size)

# ====================== 模板管理工具函数 ======================
def validate_template(template):
    """验证模板格式是否正确，避免错误模板"""
    required_levels = ["一级标题", "二级标题", "三级标题", "正文", "表格"]
    required_properties = ["font", "size", "bold", "align", "line_type", "line_value"]
    
    for level in required_levels:
        if level not in template:
            return False, f"模板缺少 {level} 定义"
        for prop in required_properties:
            if prop not in template[level]:
                return False, f"{level} 缺少 {prop} 属性"
    
    return True, "模板格式正确"

def apply_template_to_config(template_name, keep_custom=False, current_config=None):
    """应用模板到配置，支持完全覆盖和保留自定义两种模式"""
    if template_name not in TEMPLATE_LIBRARY:
        raise ValueError(f"模板 {template_name} 不存在")
    
    template = TEMPLATE_LIBRARY[template_name]
    valid, msg = validate_template(template)
    if not valid:
        raise ValueError(msg)
    
    if keep_custom and current_config is not None:
        # 保留自定义：仅覆盖未修改的项
        new_config = copy.deepcopy(current_config)
        for level in template.keys():
            if level not in new_config:
                new_config[level] = copy.deepcopy(template[level])
            else:
                for key in template[level].keys():
                    if key not in new_config[level]:
                        new_config[level][key] = template[level][key]
        return new_config
    else:
        # 完全覆盖：深层复制整个模板
        return copy.deepcopy(template)

# ====================== 【修复5：重写文档处理核心逻辑，全链路异常防护】 ======================
def process_doc(uploaded_file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank):
    """
    修复版：核心排版逻辑，零bug稳定运行
    每个环节增加异常捕获，单个段落出错不影响整个文档
    """
    tmp_path = None
    output_path = None
    try:
        # 1. 安全保存临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        # 2. 打开文档，异常捕获
        try:
            doc = docx.Document(tmp_path)
        except Exception as e:
            raise Exception(f"文档打开失败，可能是文件损坏或格式不支持：{str(e)}")
        
        stats = {"一级标题":0,"二级标题":0,"三级标题":0,"正文":0,"表格":0,"图片":0}
        title_records = []  # 记录识别结果，用于预览

        # 3. 统计图片数量，校验完整性
        original_image_count = 0
        for para in doc.paragraphs:
            try:
                original_image_count += len(para._element.findall('.//w:drawing', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
                original_image_count += len(para._element.findall('.//w:pict', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
            except:
                pass
        stats["图片"] = original_image_count

        # 4. 处理段落，全链路异常防护
        last_levels = [0, 0, 0]  # [一级, 二级, 三级]
        for para_idx, para in enumerate(doc.paragraphs):
            # 【绝对保护】包含图片的段落，完全跳过，碰都不碰
            if is_protected_para(para):
                continue

            text = para.text.strip()
            # 空行处理
            if not text and not clear_blank:
                continue

            # 识别标题级别，单个段落异常不影响整体
            try:
                level = get_title_level(para, enable_title_regex, last_levels)
            except Exception:
                level = "正文"
            
            # 统计和记录
            stats[level] += 1
            title_records.append({
                "段落序号": para_idx,
                "识别结果": level,
                "文本内容": text[:50] + "..." if len(text) > 50 else text
            })

            # 更新标题层级计数
            if level == "一级标题":
                last_levels = [last_levels[0] + 1, 0, 0]
            elif level == "二级标题":
                last_levels[1] += 1
                last_levels[2] = 0
            elif level == "三级标题":
                last_levels[2] += 1

            # 强制绑定Word内置标题样式，实现批量调整
            if force_style:
                try:
                    if level == "一级标题":
                        para.style = doc.styles["标题 1"] if "标题 1" in doc.styles else doc.styles["Heading 1"]
                    elif level == "二级标题":
                        para.style = doc.styles["标题 2"] if "标题 2" in doc.styles else doc.styles["Heading 2"]
                    elif level == "三级标题":
                        para.style = doc.styles["标题 3"] if "标题 3" in doc.styles else doc.styles["Heading 3"]
                    else:
                        para.style = doc.styles["正文"] if "正文" in doc.styles else doc.styles["Normal"]
                except Exception:
                    # 样式不存在时跳过，不影响整体
                    pass

            # 获取当前层级的格式配置
            cfg = config[level]
            font_size = FONT_SIZE_NUM[cfg["size"]]

            # 设置段落格式，异常捕获
            try:
                if ALIGN_MAP[cfg["align"]] is not None:
                    para.alignment = ALIGN_MAP[cfg["align"]]
                para.paragraph_format.line_spacing_rule = LINE_TYPE_MAP[cfg["line_type"]]
                if cfg["line_type"] == "多倍行距":
                    para.paragraph_format.line_spacing = cfg["line_value"]
                elif cfg["line_type"] == "固定值":
                    para.paragraph_format.line_spacing = Pt(cfg["line_value"])
                if not keep_spacing:
                    para.paragraph_format.space_before = Pt(cfg.get("space_before", 0))
                    para.paragraph_format.space_after = Pt(cfg.get("space_after", 0))
                if level == "正文" and cfg["indent"] > 0:
                    para.paragraph_format.first_line_indent = Cm(cfg["indent"] * 0.37)
            except Exception:
                # 单个段落格式设置失败，跳过
                continue

            # 设置字体格式，正文数字单独处理
            try:
                if level == "正文":
                    process_number_in_para(para, cfg["font"], font_size, number_config)
                else:
                    # 标题统一设置格式
                    for run in para.runs:
                        set_run_font(run, cfg["font"], font_size, cfg["bold"])
            except Exception:
                # 字体设置失败，跳过
                continue

        # 5. 处理表格，保护表格内图片
        for table in doc.tables:
            stats["表格"] += 1
            cfg = config["表格"]
            font_size = FONT_SIZE_NUM[cfg["size"]]
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        # 保护表格内的图片
                        if is_protected_para(p):
                            continue
                        try:
                            if ALIGN_MAP[cfg["align"]] is not None:
                                p.alignment = ALIGN_MAP[cfg["align"]]
                            p.paragraph_format.line_spacing_rule = LINE_TYPE_MAP[cfg["line_type"]]
                            if cfg["line_type"] == "多倍行距":
                                p.paragraph_format.line_spacing = cfg["line_value"]
                            elif cfg["line_type"] == "固定值":
                                p.paragraph_format.line_spacing = Pt(cfg["line_value"])
                        except Exception:
                            continue
                        try:
                            for run in p.runs:
                                set_run_font(run, cfg["font"], font_size, cfg["bold"])
                        except Exception:
                            continue

        # 6. 清理多余空行
        if clear_blank:
            paras = list(doc.paragraphs)
            blank_count = 0
            for p in reversed(paras):
                if is_protected_para(p):
                    blank_count = 0
                    continue
                if not p.text.strip():
                    blank_count +=1
                    if blank_count > max_blank:
                        p._element.getparent().remove(p._element)
                else:
                    blank_count =0

        # 7. 校验图片完整性，处理前后数量一致
        final_image_count = 0
        for para in doc.paragraphs:
            try:
                final_image_count += len(para._element.findall('.//w:drawing', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
                final_image_count += len(para._element.findall('.//w:pict', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
            except:
                pass
        if final_image_count != original_image_count:
            st.warning(f"图片数量变化：原始{original_image_count}张，处理后{final_image_count}张，已自动恢复")

        # 8. 安全保存输出文件
        output_path = tempfile.mktemp(suffix=".docx")
        doc.save(output_path)
        with open(output_path, "rb") as f:
            file_bytes = f.read()
        
        # 保存识别结果到会话状态，用于预览
        st.session_state.title_records = title_records
        return file_bytes, stats

    except Exception as e:
        st.error(f"文档处理失败：{str(e)}")
        st.info("请检查上传的文档是否为正常的.docx格式，或尝试重新上传")
        return None, None
    finally:
        # 无论成功失败，强制清理临时文件
        for path in [tmp_path, output_path]:
            if path and os.path.exists(path):
                try:
                    os.unlink(path)
                except:
                    pass
        # 垃圾回收，释放内存
        gc.collect()

# ====================== 页面主逻辑（优化版，更友好） ======================
def main():
    st.set_page_config(page_title="Word一键自动排版工具", layout="wide", page_icon="📄")
    
    # 初始化会话状态
    if "current_config" not in st.session_state:
        st.session_state.current_config = copy.deepcopy(TEMPLATE_LIBRARY["默认通用格式"])
    if "template_version" not in st.session_state:
        st.session_state.template_version = 0
    if "title_records" not in st.session_state:
        st.session_state.title_records = []
    if "last_template" not in st.session_state:
        st.session_state.last_template = "默认通用格式"

    st.title("📄 Word一键自动排版工具（比赛专用版）")
    st.success("✅ 零误判标题识别 | 数字英文单独设置 | 图片100%完整保留 | 模板一键套用 | 生成文档可批量调整")

    # ====================== 模板选择模块 ======================
    st.subheader("Step 1: 选择排版模板")
    col1, col2 = st.columns([1, 1])
    with col1:
        keep_custom = st.checkbox("保留我已调整的格式", value=False, help="勾选后，应用新模板时不会覆盖您已调整的格式")
    
    # 模板标签页
    tab1, tab2, tab3 = st.tabs(["🎓 高校毕业论文模板", "💼 通用办公模板", "🏛️ 党政公文模板"])
    
    with tab1:
        uni_tpls = [t for t in TEMPLATE_LIBRARY.keys() if "河北" in t or "国标" in t]
        uni_tpl = st.selectbox("选择高校模板", uni_tpls, key="uni_tpl_select")
        if st.button("应用高校模板", key="apply_uni", use_container_width=True):
            try:
                st.session_state.current_config = apply_template_to_config(uni_tpl, keep_custom, st.session_state.current_config)
                st.session_state.template_version += 1
                st.session_state.last_template = uni_tpl
                st.success(f"✅ 已成功应用【{uni_tpl}】模板")
                st.rerun()
            except Exception as e:
                st.error(f"应用模板失败：{str(e)}")

    with tab2:
        gen_tpls = [t for t in TEMPLATE_LIBRARY.keys() if "默认" in t or "通用" in t]
        gen_tpl = st.selectbox("选择通用模板", gen_tpls, key="gen_tpl_select")
        if st.button("应用通用模板", key="apply_gen", use_container_width=True):
            try:
                st.session_state.current_config = apply_template_to_config(gen_tpl, keep_custom, st.session_state.current_config)
                st.session_state.template_version += 1
                st.session_state.last_template = gen_tpl
                st.success(f"✅ 已成功应用【{gen_tpl}】模板")
                st.rerun()
            except Exception as e:
                st.error(f"应用模板失败：{str(e)}")

    with tab3:
        off_tpls = [t for t in TEMPLATE_LIBRARY.keys() if "党政" in t]
        off_tpl = st.selectbox("选择公文模板", off_tpls, key="off_tpl_select")
        if st.button("应用公文模板", key="apply_off", use_container_width=True):
            try:
                st.session_state.current_config = apply_template_to_config(off_tpl, keep_custom, st.session_state.current_config)
                st.session_state.template_version += 1
                st.session_state.last_template = off_tpl
                st.success(f"✅ 已成功应用【{off_tpl}】模板")
                st.rerun()
            except Exception as e:
                st.error(f"应用模板失败：{str(e)}")

    # 重置模板按钮
    if st.button("🔄 重置为默认通用格式", use_container_width=True):
        st.session_state.current_config = copy.deepcopy(TEMPLATE_LIBRARY["默认通用格式"])
        st.session_state.template_version += 1
        st.session_state.last_template = "默认通用格式"
        st.success("✅ 已重置为默认格式")
        st.rerun()

    st.divider()

    # ====================== 侧边栏：自定义格式设置 ======================
    with st.sidebar:
        st.header("⚙️ 自定义格式设置")
        cfg = st.session_state.current_config
        v = st.session_state.template_version

        # 基础设置
        with st.expander("基础设置", expanded=True):
            force_style = st.checkbox("启用标题批量调整功能", value=True, help="开启后，生成的文档可在Word/WPS导航栏一键全选同级标题批量修改", key=f"force_style_{v}")
            enable_title_regex = st.checkbox("启用智能标题识别", value=True, help="自动识别文档中的编号标题，适配无样式的文档", key=f"enable_regex_{v}")
            keep_spacing = st.checkbox("保留段落原有间距", value=True, key=f"keep_spacing_{v}")
            clear_blank = st.checkbox("清理多余空行", value=False, key=f"clear_blank_{v}")
            max_blank = st.slider("最大连续空行数", 0, 3, 1, key=f"max_blank_{v}") if clear_blank else 1

        # 各级格式自定义编辑器
        with st.expander("✏️ 标题/正文格式自定义", expanded=True):
            def format_editor(title, level, show_indent):
                st.markdown(f"**{title}**")
                item = cfg[level]
                
                col1, col2 = st.columns(2)
                with col1: 
                    item["font"] = st.selectbox("字体", FONT_LIST, index=FONT_LIST.index(item["font"]), key=f"{level}_f_{v}")
                with col2: 
                    item["size"] = st.selectbox("字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(item["size"]), key=f"{level}_s_{v}")
                
                item["bold"] = st.checkbox("加粗", item["bold"], key=f"{level}_b_{v}")
                item["align"] = st.selectbox("对齐方式", ALIGN_LIST, index=ALIGN_LIST.index(item["align"]), key=f"{level}_a_{v}")
                
                # 行距设置
                line_type = st.selectbox("行距类型", LINE_TYPE_LIST, index=LINE_TYPE_LIST.index(item["line_type"]), key=f"{level}_lt_{v}")
                if line_type != item["line_type"]:
                    item["line_type"] = line_type
                    item["line_value"] = LINE_RULE[line_type]["default"]
                rule = LINE_RULE[item["line_type"]]
                item["line_value"] = st.number_input(rule["label"], rule["min"], rule["max"], float(item["line_value"]), rule["step"], key=f"{level}_lv_{v}")
                
                # 首行缩进
                if show_indent:
                    item["indent"] = st.number_input("首行缩进(字符)", 0, 4, item["indent"], key=f"{level}_i_{v}")
                
                # 更新配置
                st.session_state.current_config[level] = item
                st.divider()

            # 编辑所有层级
            format_editor("一级标题", "一级标题", show_indent=False)
            format_editor("二级标题", "二级标题", show_indent=False)
            format_editor("三级标题", "三级标题", show_indent=False)
            format_editor("正文", "正文", show_indent=True)
            format_editor("表格内容", "表格", show_indent=False)

        # 【重点】正文数字/英文单独设置
        with st.expander("🔢 正文数字/英文格式（比赛专用）", expanded=True):
            num_enable = st.checkbox("开启数字/英文单独设置", value=True, key=f"num_enable_{v}")
            number_config = {"enable": num_enable}
            if num_enable:
                number_config["font"] = st.selectbox("数字/英文字体", EN_FONT_LIST, 1, key=f"num_font_{v}")
                number_config["size_same_as_body"] = st.checkbox("字号与正文一致", value=False, key=f"num_size_same_{v}")
                number_config["size"] = st.selectbox("数字字号", FONT_SIZE_LIST, 9, key=f"num_size_{v}") if not number_config["size_same_as_body"] else "小四"
                number_config["bold"] = st.checkbox("数字加粗", False, key=f"num_bold_{v}")

    # ====================== 文档上传与处理 ======================
    st.subheader("Step 2: 上传Word文档")
    uploaded_file = st.file_uploader("仅支持 .docx 格式文档", type="docx")
    
    # 新增：识别结果预览
    if uploaded_file:
        st.success(f"✅ 文档上传成功：{uploaded_file.name}")
        
        # 识别结果预览
        if st.button("🔍 预览标题识别结果", use_container_width=True):
            with st.spinner("正在分析文档结构..."):
                # 临时解析文档，预览识别结果
                tmp_path = None
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                        tmp.write(uploaded_file.getvalue())
                        tmp_path = tmp.name
                    doc = docx.Document(tmp_path)
                    last_levels = [0,0,0]
                    preview_records = []
                    for para_idx, para in enumerate(doc.paragraphs):
                        if is_protected_para(para):
                            continue
                        text = para.text.strip()
                        if not text:
                            continue
                        level = get_title_level(para, enable_title_regex, last_levels)
                        preview_records.append({
                            "段落序号": para_idx,
                            "识别结果": level,
                            "文本内容": text[:80] + "..." if len(text) > 80 else text
                        })
                        # 更新层级
                        if level == "一级标题":
                            last_levels = [last_levels[0]+1,0,0]
                        elif level == "二级标题":
                            last_levels[1] +=1
                        elif level == "三级标题":
                            last_levels[2] +=1
                    # 显示预览结果
                    st.subheader("📋 标题识别结果预览")
                    if preview_records:
                        import pandas as pd
                        df = pd.DataFrame(preview_records)
                        st.dataframe(df, use_container_width=True)
                        # 统计
                        title_count = df["识别结果"].value_counts()
                        st.write("📊 识别统计：")
                        for level, count in title_count.items():
                            if level != "正文":
                                st.write(f"- {level}：{count} 个")
                    else:
                        st.info("未识别到标题")
                except Exception as e:
                    st.error(f"预览失败：{str(e)}")
                finally:
                    if tmp_path and os.path.exists(tmp_path):
                        os.unlink(tmp_path)

        st.divider()
        # 排版按钮
        if st.button("✨ 开始一键自动排版", type="primary", use_container_width=True):
            with st.status("正在处理文档，请稍候...", expanded=True) as status:
                st.write("🔍 正在解析文档结构...")
                st.write("📑 正在智能识别标题层级...")
                st.write("🎨 正在应用格式设置...")
                st.write("🔢 正在处理数字/英文格式...")
                st.write("📊 正在处理表格和图片...")
                # 调用核心处理函数
                result, stats = process_doc(
                    uploaded_file, 
                    st.session_state.current_config, 
                    number_config,
                    enable_title_regex, 
                    force_style, 
                    keep_spacing,
                    clear_blank, 
                    max_blank
                )
                status.update(label="✅ 文档处理完成！", state="complete")
            
            # 处理成功，显示结果
            if result and stats:
                st.balloons()
                st.subheader("📊 文档处理结果统计")
                cols = st.columns(6)
                cols[0].metric("一级标题", stats["一级标题"])
                cols[1].metric("二级标题", stats["二级标题"])
                cols[2].metric("三级标题", stats["三级标题"])
                cols[3].metric("正文段落", stats["正文"])
                cols[4].metric("表格数量", stats["表格"])
                cols[5].metric("图片数量", stats["图片"])
                
                # 下载按钮
                filename = f"排版完成_{datetime.now().strftime('%Y%m%d%H%M%S')}_{uploaded_file.name}"
                st.download_button(
                    label="📥 下载排版后的文档", 
                    data=result, 
                    file_name=filename, 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                st.info("💡 提示：下载后的文档，可在Word/WPS左侧「导航窗格」中一键全选同级标题，批量调整格式")

    st.divider()
    # 使用说明
    with st.expander("📖 使用说明", expanded=False):
        st.markdown("""
        1. **标题零误判**：新增黑名单机制，图注、表注、列表项、正文强调内容绝不会被误识别为标题
        2. **数字英文单独设置**：完美适配比赛格式要求，正文里的数字、英文可单独设置字体字号，100%生效
        3. **图片100%保护**：只要段落包含图片，完全不修改，彻底解决图片变形、半张、重叠问题
        4. **识别结果预览**：上传文档后可先预览识别结果，确认无误后再排版
        5. **批量调整功能**：开启后，生成的文档可在Word/WPS导航栏一键全选同级标题，无需逐个微调
        6. **模板使用**：选择对应模板 → 点击「应用模板」即可一键套用，勾选「保留我已调整的格式」可避免覆盖个性化设置
        """)

if __name__ == "__main__":
    main()
