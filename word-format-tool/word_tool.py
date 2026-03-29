import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.oxml.ns import qn
import requests
import os
import tempfile
import re

# ====================== 全局配置（部署适配版）======================
# 豆包API配置（部署时在Streamlit Cloud的Secrets里填写，本地开发不填也能正常用格式功能）
try:
    DOUBAO_API_KEY = st.secrets["DOUBAO_API_KEY"]
    DOUBAO_MODEL = st.secrets.get("DOUBAO_MODEL", "ep-20250628104918-7rqxd")
except:
    DOUBAO_API_KEY = ""
    DOUBAO_MODEL = "ep-20250628104918-7rqxd"
DOUBAO_API_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"

# ====================== 标准映射表（和WPS/Word完全一致）======================
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

TITLE_PATTERNS = {
    "一级标题": re.compile(r"^[一二三四五六七八九十]+、\s*[^，。？！；]{0,30}$|^第[一二三四五六七八九十]+章\s*[^，。？！；]{0,30}$|^第\d+章\s*[^，。？！；]{0,30}$"),
    "二级标题": re.compile(r"^[（(][一二三四五六七八九十]+[）)]\s*[^，。？！；]{0,40}$|^\d+\、\s*[^，。？！；]{0,40}$|^\d+\.\s+[^，。？！；]{0,40}$"),
    "三级标题": re.compile(r"^[（(]\d+[）)]\s*[^，。？！；]{0,50}$|^\d+\.\d+\s+[^，。？！；]{0,50}$|^\d+\.\d+\.\d+\s+[^，。？！；]{0,50}$")
}

LINE_SPACING_TYPE_MAP = {
    "单倍行距": WD_LINE_SPACING.SINGLE,
    "1.5倍行距": WD_LINE_SPACING.ONE_POINT_FIVE,
    "2倍行距": WD_LINE_SPACING.DOUBLE,
    "多倍行距": WD_LINE_SPACING.MULTIPLE,
    "固定值": WD_LINE_SPACING.EXACTLY
}

STYLE_NAME_MAP = {
    "一级标题": ["标题 1", "Heading 1"],
    "二级标题": ["标题 2", "Heading 2"],
    "三级标题": ["标题 3", "Heading 3"],
    "正文": ["正文", "Normal"]
}

# 标准格式模板库
FORMAT_TEMPLATES = {
    "默认格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1}
    },
    "党政公文标准格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5},
        "二级标题": {"font": "楷体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5},
        "三级标题": {"font": "仿宋", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5},
        "正文": {"font": "仿宋", "size": "三号", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "仿宋", "size": "小三", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1}
    },
    "本科毕业论文格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1}
    }
}

# 高校毕业论文模板库（参赛核心亮点）
UNIVERSITY_TEMPLATES = {
    "河北科技大学-本科毕业论文模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1}
    },
    "国标-本科毕业论文通用模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1}
    },
    "党政机关公文国标GB/T 9704-2012模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5},
        "二级标题": {"font": "楷体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5},
        "三级标题": {"font": "仿宋", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5},
        "正文": {"font": "仿宋", "size": "三号", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "仿宋", "size": "小三", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1}
    }
}

WORD_NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# ====================== 核心工具函数 ======================
def iter_block_items(parent):
    from docx.document import Document
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl

    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, docx.table._Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("不支持的父元素类型")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def set_paragraph_style(para, level, doc):
    style_names = STYLE_NAME_MAP.get(level, ["正文", "Normal"])
    for style_name in style_names:
        try:
            if style_name in doc.styles:
                para.style = style_name
                return
        except:
            continue
    pass

def is_special_page_para(para):
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

def get_title_level(para, enable_regex=True):
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
    for level, pattern in TITLE_PATTERNS.items():
        if pattern.match(text):
            return level
    return "正文"

def process_number_text(para, body_font, body_size_pt, number_config):
    number_size_pt = body_size_pt if number_config["size_same_as_body"] else FONT_SIZE_MAP[number_config["size"]]
    number_pattern = re.compile(r"-?\d+\.?\d*%?")
    
    for run in para.runs:
        text = run.text
        if not text:
            continue
        if not number_pattern.search(text):
            set_run_font(run, body_font, body_size_pt, is_bold=False)
            continue
        parts = number_pattern.split(text)
        numbers = number_pattern.findall(text)
        run.text = ""
        for i in range(len(parts)):
            if parts[i]:
                new_run = para.add_run(parts[i])
                set_run_font(new_run, body_font, body_size_pt, is_bold=False)
            if i < len(numbers):
                new_run = para.add_run(numbers[i])
                set_en_run_font(new_run, number_config["font"], number_size_pt, is_bold=number_config["bold"])

def remove_extra_blank_lines(doc, max_blank_lines=1):
    paragraphs = list(doc.paragraphs)
    blank_count = 0
    for i in range(len(paragraphs)-1, -1, -1):
        para = paragraphs[i]
        is_blank = not para.text.strip()
        if is_special_page_para(para):
            blank_count = 0
            continue
        if is_blank:
            blank_count += 1
            if blank_count > max_blank_lines:
                p = para._element
                p.getparent().remove(p)
        else:
            blank_count = 0

# ====================== AI降重/润色核心函数 ======================
def ai_process_text(text, mode="润色"):
    if not DOUBAO_API_KEY or not text.strip():
        return text
    
    prompt = f"请对以下文本进行{mode}处理，要求：1. 完全保留原文核心意思、专业术语、数字、标点符号；2. 不改变原文段落结构、换行符；3. 输出仅返回处理后的文本，不要额外解释。\n原文：{text}"
    
    headers = {
        "Authorization": f"Bearer {DOUBAO_API_KEY}",
        "Content-Type": "application/json"
    }
    data = {
        "model": DOUBAO_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7,
        "max_tokens": 4096
    }
    
    try:
        response = requests.post(DOUBAO_API_URL, json=data, headers=headers, timeout=30)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        st.warning(f"AI处理失败：{str(e)}，将保留原文")
        return text

# ====================== 文档处理主函数 ======================
def process_document(uploaded_file, format_config, number_config, ai_mode, enable_title_regex, force_style, keep_original_spacing, remove_blank_lines, max_blank_lines):
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

        for para in doc.paragraphs:
            if is_special_page_para(para):
                continue
            
            is_blank = not para.text.strip()
            if is_blank:
                continue
            
            level = get_title_level(para, enable_title_regex)
            stats[level] += 1

            processed_text = para.text
            if ai_mode != "不使用AI" and "TOC" not in para.style.name:
                processed_text = ai_process_text(para.text, ai_mode)
                if processed_text != para.text:
                    para.text = processed_text

            config = format_config[level]
            font_size_pt = FONT_SIZE_MAP[config["size"]]

            if force_style:
                set_paragraph_style(para, level, doc)

            try:
                if config["align"] is not None:
                    para.alignment = config["align"]
                para.paragraph_format.line_spacing_rule = config["line_spacing_rule"]
                if config["line_spacing_rule"] == WD_LINE_SPACING.MULTIPLE:
                    para.paragraph_format.line_spacing = config["line_spacing_value"]
                elif config["line_spacing_rule"] == WD_LINE_SPACING.EXACTLY:
                    para.paragraph_format.line_spacing = Pt(config["line_spacing_value"])
                
                if not keep_original_spacing:
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                
                if level == "正文" and config["first_line_indent"] > 0:
                    para.paragraph_format.first_line_indent = Cm(config["first_line_indent"] * 0.35)
                
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
                        if not para.text.strip():
                            continue
                        if force_style:
                            set_paragraph_style(para, "正文", doc)
                        try:
                            if config["align"] is not None:
                                para.alignment = config["align"]
                            para.paragraph_format.line_spacing_rule = config["line_spacing_rule"]
                            if config["line_spacing_rule"] == WD_LINE_SPACING.MULTIPLE:
                                para.paragraph_format.line_spacing = config["line_spacing_value"]
                            elif config["line_spacing_rule"] == WD_LINE_SPACING.EXACTLY:
                                para.paragraph_format.line_spacing = Pt(config["line_spacing_value"])
                        except:
                            pass
                        for run in para.runs:
                            set_run_font(run, config["font"], font_size_pt, config["bold"])

        if remove_blank_lines:
            remove_extra_blank_lines(doc, max_blank_lines)

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

# ====================== Streamlit网页界面 ======================
def main():
    st.set_page_config(page_title="Word格式智能调整工具", layout="wide")
    st.title("📄 Word格式智能调整工具")
    
    st.warning("⚠️ 重要声明：此工具仅能减少复杂的格式调整工作量，处理完成后仍需您手动与原文进行对比核对，确保内容与格式无误。")
    
    st.markdown("✅ 100%保留图片/目录/原格式 | ✅ 高校论文格式一键适配 | ✅ 办公自动化格式处理 | ✅ 全国大学生计算机设计大赛参赛作品")

    st.subheader("📋 一键套用标准格式模板")
    template_type = st.radio("模板类型", ["通用办公模板", "高校毕业论文模板", "党政公文模板"], horizontal=True)

    if template_type == "通用办公模板":
        template_list = list(FORMAT_TEMPLATES.keys())
        template_dict = FORMAT_TEMPLATES
    elif template_type == "高校毕业论文模板":
        template_list = list(UNIVERSITY_TEMPLATES.keys())
        template_dict = UNIVERSITY_TEMPLATES
    else:
        template_list = ["党政机关公文国标GB/T 9704-2012模板"]
        template_dict = UNIVERSITY_TEMPLATES

    template_name = st.selectbox("选择目标格式", template_list, index=0)
    selected_template = template_dict[template_name]

    with st.sidebar:
        st.header("🎨 详细格式设置")
        format_config = {}

        st.subheader("🏷️ 样式统一设置")
        force_style = st.checkbox("强制统一套用Word原生样式", value=True, help="开启后，识别的标题自动套用「标题1/2/3」样式，正文套用「正文」样式，Word导航窗格可直接识别，目录可正常更新")
        enable_title_regex = st.checkbox("启用正则智能识别标题", value=True, help="关闭后仅识别Word自带的标题样式，彻底解决正文误判为标题的问题")

        st.divider()
        st.subheader("📄 空行空页修复设置")
        keep_original_spacing = st.checkbox("保留原段落段前/段后距", value=True, help="默认开启，彻底解决段落间距变大导致的空行空页问题")
        remove_blank_lines = st.checkbox("清除多余空行", value=False, help="开启后自动清理文档中连续的多余空行")
        if remove_blank_lines:
            max_blank_lines = st.slider("最多保留连续空行数", min_value=0, max_value=2, value=1, help="0=清除所有空行，1=只保留单个空行")
        else:
            max_blank_lines = 1

        def create_format_config(title, key_prefix, default_config):
            st.divider()
            st.subheader(title)
            config = {}
            config["font"] = st.selectbox("字体", list(FONT_MAP.keys()), key=f"{key_prefix}_font", index=list(FONT_MAP.keys()).index(default_config["font"]))
            config["size"] = st.selectbox("字号", list(FONT_SIZE_MAP.keys()), key=f"{key_prefix}_size", index=list(FONT_SIZE_MAP.keys()).index(default_config["size"]))
            config["bold"] = st.checkbox("加粗", value=default_config["bold"], key=f"{key_prefix}_bold")
            align_options = ["不修改", "左对齐", "居中", "两端对齐", "右对齐"]
            default_align = default_config["align"]
            config["align"] = st.selectbox("对齐方式", align_options, key=f"{key_prefix}_align", index=align_options.index(default_align))
            align_map = {
                "不修改": None,
                "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
                "居中": WD_ALIGN_PARAGRAPH.CENTER,
                "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY,
                "右对齐": WD_ALIGN_PARAGRAPH.RIGHT
            }
            config["align"] = align_map[config["align"]]
            
            default_line_type = default_config["line_type"]
            config["line_spacing_type"] = st.selectbox("行距类型", list(LINE_SPACING_TYPE_MAP.keys()), key=f"{key_prefix}_line_type", index=list(LINE_SPACING_TYPE_MAP.keys()).index(default_line_type))
            config["line_spacing_rule"] = LINE_SPACING_TYPE_MAP[config["line_spacing_type"]]
            
            default_line_value = default_config["line_value"]
            if config["line_spacing_type"] == "多倍行距":
                config["line_spacing_value"] = st.number_input("行距倍数", min_value=0.5, max_value=5.0, value=default_line_value, step=0.1, key=f"{key_prefix}_line_value")
            elif config["line_spacing_type"] == "固定值":
                config["line_spacing_value"] = st.number_input("固定值(磅)", min_value=6, max_value=100, value=default_line_value, step=1, key=f"{key_prefix}_line_value")
            else:
                config["line_spacing_value"] = 1
            
            if "indent" in default_config:
                config["first_line_indent"] = st.number_input("首行缩进(字符)", min_value=0, max_value=4, value=default_config["indent"], key=f"{key_prefix}_indent")
            else:
                config["first_line_indent"] = 0
            
            return config

        format_config["一级标题"] = create_format_config("📌 一级标题", "h1", selected_template["一级标题"])
        format_config["二级标题"] = create_format_config("📌 二级标题", "h2", selected_template["二级标题"])
        format_config["三级标题"] = create_format_config("📌 三级标题", "h3", selected_template["三级标题"])
        format_config["正文"] = create_format_config("📝 正文内容", "body", selected_template["正文"])
        format_config["表格"] = create_format_config("📊 表格内容", "table", selected_template["表格"])

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

    uploaded_file = st.file_uploader("📤 上传Word文档（仅支持.docx格式）", type="docx")
    
    if uploaded_file:
        st.success("✅ 文档上传成功！")
        
        if not DOUBAO_API_KEY:
            st.info("ℹ️ 填写豆包API密钥即可启用AI降重/润色功能，不填写可正常使用所有格式调整功能")
            ai_mode = "不使用AI"
        else:
            ai_mode = st.radio("🤖 AI文本处理", ["不使用AI", "润色", "降重"], horizontal=True)
        
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
                    max_blank_lines
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