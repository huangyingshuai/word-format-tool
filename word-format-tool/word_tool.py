import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.oxml.ns import qn
import requests
import os
import tempfile
import re

# ====================== 页面全局配置（必须放在最前面）======================
st.set_page_config(page_title="文式通 - Word格式智能处理系统", layout="wide")

# ====================== 全局常量定义（修复渲染bug核心）======================
# 对齐方式映射表
ALIGN_MAP = {
    "不修改": None,
    "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
    "居中": WD_ALIGN_PARAGRAPH.CENTER,
    "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "右对齐": WD_ALIGN_PARAGRAPH.RIGHT
}
ALIGN_REVERSE_MAP = {v: k for k, v in ALIGN_MAP.items()}

# 行距类型配置
LINE_TYPE_DEFAULT = {
    "单倍行距": 1,
    "1.5倍行距": 1.5,
    "2倍行距": 2,
    "多倍行距": 1.5,
    "固定值": 20
}
LINE_TYPE_RANGE = {
    "多倍行距": {"min": 0.5, "max": 5.0, "step": 0.1},
    "固定值": {"min": 6, "max": 100, "step": 1}
}
LINE_SPACING_TYPE_MAP = {
    "单倍行距": WD_LINE_SPACING.SINGLE,
    "1.5倍行距": WD_LINE_SPACING.ONE_POINT_FIVE,
    "2倍行距": WD_LINE_SPACING.DOUBLE,
    "多倍行距": WD_LINE_SPACING.MULTIPLE,
    "固定值": WD_LINE_SPACING.EXACTLY
}

# 字体映射表
FONT_MAP = {
    "宋体": "SimSun",
    "微软雅黑": "Microsoft YaHei",
    "黑体": "SimHei",
    "楷体": "KaiTi",
    "仿宋": "FangSong"
}
EN_FONT_MAP = {
    "Times New Roman": "Times New Roman",
    "Arial": "Arial",
    "Calibri": "Calibri",
    "和正文一致": "same"
}
FONT_SIZE_MAP = {
    "初号": 42,
    "小初": 36,
    "一号": 26,
    "小一": 24,
    "二号": 22,
    "小二": 18,
    "三号": 16,
    "小三": 15,
    "四号": 14,
    "小四": 12,
    "五号": 10.5,
    "小五": 9,
    "六号": 7.5,
    "小六": 6.5
}

# 标题识别正则
TITLE_PATTERNS = {
    "一级标题": re.compile(r"^[一二三四五六七八九十]+、\s*[^，。？！；]{0,40}$|^第[一二三四五六七八九十]+章\s*[^，。？！；]{0,40}$|^第\d+章\s*[^，。？！；]{0,40}$|^\d+、\s*[^，。？！；]{0,40}$"),
    "二级标题": re.compile(r"^[（(][一二三四五六七八九十]+[）)]\s*[^，。？！；]{0,50}$|^\d+\.\s+[^，。？！；]{0,50}$|^\d+、\s*[^，。？！；]{0,50}$"),
    "三级标题": re.compile(r"^[（(]\d+[）)]\s*[^，。？！；]{0,60}$|^\d+\.\d+\s+[^，。？！；]{0,60}$|^\d+\.\d+\.\d+\s*[^，。？！；]{0,60}$|^\d+\）\s*[^，。？！；]{0,60}$")
}
STYLE_NAME_MAP = {
    "一级标题": ["标题 1", "Heading 1"],
    "二级标题": ["标题 2", "Heading 2"],
    "三级标题": ["标题 3", "Heading 3"],
    "正文": ["正文", "Normal"]
}

# ====================== 模板库 ======================
GENERAL_TEMPLATES = {
    "默认格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1, "indent": 0}
    }
}

UNIVERSITY_TEMPLATES = {
    "河北科技大学-本科毕业论文模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1, "indent": 0}
    },
    "河北工业大学-本科毕业论文模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "楷体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1, "indent": 0}
    },
    "燕山大学-本科毕业论文模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "固定值", "line_value": 20, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1, "indent": 0}
    },
    "河北大学-本科毕业论文模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "仿宋", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1, "indent": 0}
    },
    "国标-本科毕业论文通用模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1, "indent": 0}
    }
}

OFFICIAL_TEMPLATES = {
    "党政机关公文国标GB/T 9704-2012模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "楷体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "仿宋", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "仿宋", "size": "三号", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "仿宋", "size": "小三", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1, "indent": 0}
    }
}

WORD_NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# ====================== 全局API配置 ======================
try:
    DOUBAO_API_KEY = st.secrets["DOUBAO_API_KEY"]
    DOUBAO_MODEL = st.secrets.get("DOUBAO_MODEL", "ep-20250628104918-7rqxd")
except:
    DOUBAO_API_KEY = ""
    DOUBAO_MODEL = "ep-20250628104918-7rqxd"
DOUBAO_API_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"

# ====================== 核心工具函数 ======================
def is_protected_para(para):
    """检查受保护段落（带图片、分页符、域代码），不修改内容只改字体"""
    if para.paragraph_format.page_break_before:
        return True
    for run in para.runs:
        if run.contains_page_break:
            return True
        if run._element.find('.//w:sectPr', WORD_NS) is not None:
            return True
        if run._element.find('.//w:drawing', WORD_NS) is not None:
            return True
        if run._element.find('.//w:pict', WORD_NS) is not None:
            return True
        if run._element.find('.//w:fldChar', WORD_NS) is not None:
            return True
    return False

def set_run_font(run, font_name, font_size_pt, is_bold=None):
    try:
        font_en = FONT_MAP.get(font_name, font_name)
        run.font.name = font_en
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_en)
        run.font.size = Pt(font_size_pt)
        if is_bold is not None:
            run.bold = is_bold
    except:
        pass

def set_en_run_font(run, font_name, font_size_pt, is_bold=None):
    try:
        if font_name == "same":
            return
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:ascii'), font_name)
        run._element.rPr.rFonts.set(qn('w:hAnsi'), font_name)
        run.font.size = Pt(font_size_pt)
        if is_bold is not None:
            run.bold = is_bold
    except:
        pass

def get_title_level(para, enable_regex=True, last_title_level=None):
    """带上下文推理的标题识别"""
    style_name = para.style.name
    if "Heading 1" in style_name or style_name == "标题 1" or "TOC 1" in style_name:
        return "一级标题"
    elif "Heading 2" in style_name or style_name == "标题 2" or "TOC 2" in style_name:
        return "二级标题"
    elif "Heading 3" in style_name or style_name == "标题 3" or "TOC 3" in style_name:
        return "三级标题"
    
    if not enable_regex:
        return "正文"
    
    text = para.text.strip()
    if not text:
        return "正文"
    
    if len(text) < 60 and last_title_level == "一级标题":
        for level, pattern in TITLE_PATTERNS.items():
            if pattern.match(text):
                return level
    
    if len(text) < 60 and last_title_level == "二级标题":
        if TITLE_PATTERNS["三级标题"].match(text):
            return "三级标题"
        for level, pattern in TITLE_PATTERNS.items():
            if pattern.match(text):
                return level
    
    for level in ["一级标题", "二级标题", "三级标题"]:
        pattern = TITLE_PATTERNS[level]
        if pattern.match(text):
            return level
    
    return "正文"

def process_number_text(para, body_font, body_size_pt, number_config):
    """安全处理数字，不改变文本位置"""
    number_size_pt = body_size_pt if number_config["size_same_as_body"] else FONT_SIZE_MAP[number_config["size"]]
    number_pattern = re.compile(r"-?\d+\.?\d*%?")
    
    for run in para.runs:
        text = run.text
        if not text:
            set_run_font(run, body_font, body_size_pt, is_bold=False)
            continue
        
        if not number_pattern.search(text):
            set_run_font(run, body_font, body_size_pt, is_bold=False)
            continue
        
        set_run_font(run, body_font, body_size_pt, is_bold=False)
        
        for match in number_pattern.finditer(text):
            start_idx = match.start()
            end_idx = match.end()
            if start_idx > 0:
                run = run.split(start_idx)
            number_run = run.split(end_idx - start_idx)
            set_en_run_font(number_run, number_config["font"], number_size_pt, is_bold=number_config["bold"])

# ====================== AI功能函数 ======================
def ai_punctuation_correct(text):
    if not DOUBAO_API_KEY or not text.strip():
        return text
    prompt = f"请对以下文本进行标点符号规范修正，要求：1. 中文内容使用中文全角标点，英文/数字内容使用英文半角标点；2. 修正错误的标点符号，统一标点规范；3. 完全保留原文的内容、语序、换行符；4. 仅输出修正后的文本，不要额外解释。\n原文：{text}"
    try:
        response = requests.post(DOUBAO_API_URL, json={
            "model": DOUBAO_MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.3,
            "max_tokens": 4096
        }, headers={"Authorization": f"Bearer {DOUBAO_API_KEY}"}, timeout=30)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except:
        return text

def ai_text_optimize(text):
    if not DOUBAO_API_KEY or not text.strip():
        return text
    prompt = f"请对以下文本进行优化，要求：1. 修正所有错别字、语病、不通顺的语句；2. 优化语句流畅度，让表达更通顺自然；3. 完全保留原文的核心意思、专业术语、数字、段落结构、换行符；4. 仅输出优化后的文本，不要额外解释。\n原文：{text}"
    try:
        response = requests.post(DOUBAO_API_URL, json={
            "model": DOUBAO_MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.5,
            "max_tokens": 4096
        }, headers={"Authorization": f"Bearer {DOUBAO_API_KEY}"}, timeout=30)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except:
        return text

def ai_process_text(text, mode="润色"):
    if not DOUBAO_API_KEY or not text.strip():
        return text
    prompt = f"请对以下文本进行{mode}处理，要求：1. 完全保留原文核心意思、专业术语、数字、标点符号；2. 不改变原文段落结构、换行符；3. 输出仅返回处理后的文本，不要额外解释。\n原文：{text}"
    try:
        response = requests.post(DOUBAO_API_URL, json={
            "model": DOUBAO_MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7,
            "max_tokens": 4096
        }, headers={"Authorization": f"Bearer {DOUBAO_API_KEY}"}, timeout=30)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except:
        return text

# ====================== 文档处理主函数 ======================
def process_document(
    uploaded_file, 
    format_config, 
    number_config, 
    ai_mode, 
    enable_title_regex, 
    force_style, 
    keep_original_spacing, 
    remove_blank_lines, 
    max_blank_lines,
    enable_punctuation_correct,
    enable_text_optimize
):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    
    try:
        doc = docx.Document(tmp_path)
        stats = {"一级标题": 0, "二级标题": 0, "三级标题": 0, "正文": 0, "表格": 0, "图片": 0}

        for para in doc.paragraphs:
            try:
                stats["图片"] += len(para._element.findall('.//w:drawing'))
            except:
                pass

        last_title_level = None
        for para in doc.paragraphs:
            if is_protected_para(para):
                for run in para.runs:
                    set_run_font(run, format_config["正文"]["font"], FONT_SIZE_MAP[format_config["正文"]["size"]])
                continue
            
            is_blank = not para.text.strip()
            if is_blank:
                continue
            
            level = get_title_level(para, enable_title_regex, last_title_level)
            stats[level] += 1
            if level in ["一级标题", "二级标题", "三级标题"]:
                last_title_level = level

            processed_text = para.text
            if ai_mode != "不使用AI" and "TOC" not in para.style.name:
                processed_text = ai_process_text(processed_text, ai_mode)
            if enable_punctuation_correct and "TOC" not in para.style.name:
                processed_text = ai_punctuation_correct(processed_text)
            if enable_text_optimize and "TOC" not in para.style.name:
                processed_text = ai_text_optimize(processed_text)
            
            if processed_text != para.text:
                para.text = processed_text

            config = format_config[level]
            font_size_pt = FONT_SIZE_MAP[config["size"]]

            if force_style:
                style_names = STYLE_NAME_MAP.get(level, ["正文", "Normal"])
                for style_name in style_names:
                    try:
                        if style_name in doc.styles:
                            para.style = style_name
                            break
                    except:
                        continue

            try:
                if config["align"] != "不修改":
                    para.alignment = ALIGN_MAP[config["align"]]
                para.paragraph_format.line_spacing_rule = LINE_SPACING_TYPE_MAP[config["line_type"]]
                if config["line_type"] == "多倍行距":
                    para.paragraph_format.line_spacing = config["line_value"]
                elif config["line_type"] == "固定值":
                    para.paragraph_format.line_spacing = Pt(config["line_value"])
                
                if not keep_original_spacing:
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                
                if level == "正文" and config["indent"] > 0:
                    para.paragraph_format.first_line_indent = Cm(config["indent"] * 0.35)
                
                para.paragraph_format.page_break_before = False
                para.paragraph_format.keep_with_next = False
            except:
                pass

            if level == "正文" and number_config["enable"]:
                process_number_text(para, config["font"], font_size_pt, number_config)
            else:
                for run in para.runs:
                    set_run_font(run, config["font"], font_size_pt, config["bold"])

        for table in doc.tables:
            stats["表格"] += 1
            config = format_config["表格"]
            font_size_pt = FONT_SIZE_MAP[config["size"]]
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if is_protected_para(para):
                            for run in para.runs:
                                set_run_font(run, config["font"], font_size_pt, config["bold"])
                            continue
                        
                        if not para.text.strip():
                            continue
                        
                        if force_style:
                            try:
                                para.style = "正文"
                            except:
                                pass
                        
                        try:
                            if config["align"] != "不修改":
                                para.alignment = ALIGN_MAP[config["align"]]
                            para.paragraph_format.line_spacing_rule = LINE_SPACING_TYPE_MAP[config["line_type"]]
                            if config["line_type"] == "多倍行距":
                                para.paragraph_format.line_spacing = config["line_value"]
                            elif config["line_type"] == "固定值":
                                para.paragraph_format.line_spacing = Pt(config["line_value"])
                        except:
                            pass
                        
                        for run in para.runs:
                            set_run_font(run, config["font"], font_size_pt, config["bold"])

        if remove_blank_lines:
            paragraphs = list(doc.paragraphs)
            blank_count = 0
            for i in range(len(paragraphs)-1, -1, -1):
                para = paragraphs[i]
                is_blank = not para.text.strip()
                if is_protected_para(para):
                    blank_count = 0
                    continue
                if is_blank:
                    blank_count += 1
                    if blank_count > max_blank_lines:
                        p = para._element
                        p.getparent().remove(p)
                else:
                    blank_count = 0

        output_path = tempfile.mktemp(suffix=".docx")
        doc.save(output_path)
        with open(output_path, "rb") as f:
            file_bytes = f.read()
        
        return file_bytes, stats
    
    except Exception as e:
        st.error(f"文档处理失败：{str(e)}")
        return None, None
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        if 'output_path' in locals() and os.path.exists(output_path):
            os.unlink(output_path)

# ====================== 【修复渲染bug核心】页面主逻辑 ======================
def main():
    # 1. 先初始化session_state（页面最前面执行）
    if "format_config" not in st.session_state:
        st.session_state.format_config = GENERAL_TEMPLATES["默认格式"]
    
    # 2. 页面标题与声明
    st.title("📄 文式通 - Word格式智能处理系统")
    st.warning("⚠️ 重要声明：此工具仅能减少复杂的格式调整工作量，处理完成后仍需您手动与原文进行对比核对，确保内容与格式无误。")
    st.markdown("✅ 100%保留图片/目录/原排版 | ✅ 高校论文格式一键适配 | ✅ 标点规范/错别字修正 | ✅ 全国大学生计算机设计大赛参赛作品")

    # 3. 模板选择模块（先执行模板选择，再渲染侧边栏）
    st.subheader("📋 一键套用标准格式模板")
    template_type = st.radio("模板类型", ["通用办公模板", "高校毕业论文模板", "党政公文模板"], horizontal=True)

    if template_type == "通用办公模板":
        template_dict = GENERAL_TEMPLATES
    elif template_type == "高校毕业论文模板":
        template_dict = UNIVERSITY_TEMPLATES
    else:
        template_dict = OFFICIAL_TEMPLATES

    template_name = st.selectbox("选择目标格式", list(template_dict.keys()), index=0)
    
    # 【修复核心】应用模板按钮逻辑，点击后立刻更新session_state并刷新页面
    if st.button("📌 应用选中的模板", use_container_width=True, type="primary"):
        st.session_state.format_config = template_dict[template_name]
        st.success(f"✅ 已成功应用【{template_name}】，左侧格式参数已同步更新！")
        st.rerun()

    # 4. 【修复核心】侧边栏格式设置（在模板按钮之后渲染，确保参数同步）
    with st.sidebar:
        st.header("🎨 详细格式设置")
        format_config = st.session_state.format_config

        st.subheader("🏷️ 核心设置")
        force_style = st.checkbox("强制统一套用Word原生样式", value=True, help="开启后，标题自动套用「标题1/2/3」样式，正文套用「正文」样式，Word导航窗格可直接识别")
        enable_title_regex = st.checkbox("启用上下文智能标题识别", value=True, help="开启后，结合上下文推理标题层级，解决三级标题识别不清的问题")
        keep_original_spacing = st.checkbox("保留原段落段前/段后距", value=True, help="默认开启，彻底解决空行空页、排版错乱问题")

        st.divider()
        st.subheader("📄 空行清理设置")
        remove_blank_lines = st.checkbox("清除多余空行", value=False)
        max_blank_lines = st.slider("最多保留连续空行数", min_value=0, max_value=2, value=1) if remove_blank_lines else 1

        st.divider()
        st.subheader("🔤 文本优化设置")
        if not DOUBAO_API_KEY:
            st.info("ℹ️ 填写豆包API密钥即可启用以下功能")
        enable_punctuation_correct = st.checkbox("启用中英文标点规范修正", value=False, disabled=not DOUBAO_API_KEY)
        enable_text_optimize = st.checkbox("启用错别字修正+语句流畅度优化", value=False, disabled=not DOUBAO_API_KEY)

        # 格式设置生成函数
        def create_format_config(title, key_prefix, level):
            st.divider()
            st.subheader(title)
            config = format_config[level]
            
            # 字体、字号、加粗
            config["font"] = st.selectbox("字体", list(FONT_MAP.keys()), key=f"{key_prefix}_font", index=list(FONT_MAP.keys()).index(config["font"]))
            config["size"] = st.selectbox("字号", list(FONT_SIZE_MAP.keys()), key=f"{key_prefix}_size", index=list(FONT_SIZE_MAP.keys()).index(config["size"]))
            config["bold"] = st.checkbox("加粗", value=config["bold"], key=f"{key_prefix}_bold")
            
            # 对齐方式
            current_align_text = ALIGN_REVERSE_MAP.get(ALIGN_MAP.get(config["align"], config["align"]), "不修改")
            config["align"] = st.selectbox(
                "对齐方式", 
                list(ALIGN_MAP.keys()), 
                key=f"{key_prefix}_align", 
                index=list(ALIGN_MAP.keys()).index(current_align_text)
            )
            
            # 行距设置
            line_type_list = list(LINE_SPACING_TYPE_MAP.keys())
            config["line_type"] = st.selectbox("行距类型", line_type_list, key=f"{key_prefix}_line_type", index=line_type_list.index(config["line_type"]))
            
            # 动态行距数值
            if config["line_type"] in ["多倍行距", "固定值"]:
                range_config = LINE_TYPE_RANGE[config["line_type"]]
                config["line_value"] = st.number_input(
                    "行距倍数" if config["line_type"] == "多倍行距" else "固定值(磅)",
                    min_value=range_config["min"],
                    max_value=range_config["max"],
                    value=float(config["line_value"]),
                    step=range_config["step"],
                    key=f"{key_prefix}_line_value"
                )
            else:
                config["line_value"] = LINE_TYPE_DEFAULT[config["line_type"]]
                st.number_input("行距倍数", value=config["line_value"], disabled=True, key=f"{key_prefix}_line_value_disabled")
            
            # 首行缩进
            if "indent" in config:
                config["indent"] = st.number_input("首行缩进(字符)", min_value=0, max_value=4, value=config["indent"], key=f"{key_prefix}_indent")
            
            # 更新session_state
            st.session_state.format_config[level] = config
            return config

        # 各级格式设置
        format_config["一级标题"] = create_format_config("📌 一级标题", "h1", "一级标题")
        format_config["二级标题"] = create_format_config("📌 二级标题", "h2", "二级标题")
        format_config["三级标题"] = create_format_config("📌 三级标题", "h3", "三级标题")
        format_config["正文"] = create_format_config("📝 正文内容", "body", "正文")
        format_config["表格"] = create_format_config("📊 表格内容", "table", "表格")

        # 正文数字格式设置
        st.divider()
        st.subheader("🔢 正文数字格式设置")
        number_config = {}
        number_config["enable"] = st.checkbox("启用数字单独格式设置", value=True)
        if number_config["enable"]:
            number_config["font"] = st.selectbox("数字字体", list(EN_FONT_MAP.keys()), index=0)
            number_config["size_same_as_body"] = st.checkbox("字号和正文一致", value=True)
            if not number_config["size_same_as_body"]:
                number_config["size"] = st.selectbox("数字字号", list(FONT_SIZE_MAP.keys()), index=list(FONT_SIZE_MAP.keys()).index("小四"))
            else:
                number_config["size"] = "小四"
            number_config["bold"] = st.checkbox("数字加粗", value=False)

    # 5. 主界面上传&处理
    uploaded_file = st.file_uploader("📤 上传Word文档（仅支持.docx格式）", type="docx")
    
    if uploaded_file:
        st.success("✅ 文档上传成功！")
        
        # AI功能设置
        if not DOUBAO_API_KEY:
            ai_mode = "不使用AI"
        else:
            ai_mode = st.radio("🤖 AI文本处理", ["不使用AI", "润色", "降重"], horizontal=True)
        
        # 处理按钮
        if st.button("🚀 开始处理文档", type="primary", use_container_width=True):
            progress_bar = st.progress(0, text="文档处理准备中...")
            try:
                progress_bar.progress(10, text="正在解析文档...")
                file_bytes, stats = process_document(
                    uploaded_file, 
                    format_config, 
                    number_config, 
                    ai_mode, 
                    enable_title_regex, 
                    force_style, 
                    keep_original_spacing, 
                    remove_blank_lines, 
                    max_blank_lines,
                    enable_punctuation_correct,
                    enable_text_optimize
                )
                progress_bar.progress(80, text="文档处理完成，正在生成下载文件...")
                
                if file_bytes and stats:
                    progress_bar.progress(100, text="处理完成！")
                    st.subheader("📋 文档内容分类识别结果")
                    res_col1, res_col2, res_col3, res_col4, res_col5, res_col6 = st.columns(6)
                    res_col1.metric("一级标题", stats["一级标题"])
                    res_col2.metric("二级标题", stats["二级标题"])
                    res_col3.metric("三级标题", stats["三级标题"])
                    res_col4.metric("正文段落", stats["正文"])
                    res_col5.metric("表格数量", stats["表格"])
                    res_col6.metric("图片数量", stats["图片"])
                    
                    st.info("ℹ️ 再次提醒：此工具仅能减少格式调整工作量，下载后请务必与原文档进行对比核对，确认内容与格式无误。")
                    
                    st.download_button(
                        label="📥 下载处理完成的文档",
                        data=file_bytes,
                        file_name=f"格式调整完成_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        type="secondary"
                    )
                    st.success("🎉 文档处理完成！点击上方按钮下载")
            except Exception as e:
                st.error(f"文档处理失败：{str(e)}")
            finally:
                progress_bar.empty()

    # 底部版权声明
    st.divider()
    with st.expander("📖 关于本作品 | 开源版权声明"):
        st.markdown("""
        本作品为全国大学生计算机设计大赛参赛作品，是一款面向大学生与办公人群的Word文档格式智能处理工具，旨在解决文档格式调整耗时费力、格式不规范的行业痛点。
        本作品使用的开源组件及对应协议：
        - Streamlit（Apache 2.0协议）
        - python-docx（MIT协议）
        - requests（Apache 2.0协议）
        所有开源组件均已遵守对应开源协议要求，无侵权行为。
        """)
        st.markdown("⚠️ 工具声明：本工具仅能减少复杂的格式调整工作量，处理完成后仍需与原文进行对比核对，确保内容与格式无误。")

if __name__ == "__main__":
    main()
