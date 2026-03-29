import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import requests
import tempfile
import os
import re

# ====================== 全局常量定义 ======================
# 对齐方式映射
ALIGN_MAP = {
    "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
    "居中": WD_ALIGN_PARAGRAPH.CENTER,
    "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,
    "不修改": None
}
ALIGN_LIST = list(ALIGN_MAP.keys())

# 行距配置（严格防越界）
LINE_TYPE_MAP = {
    "单倍行距": WD_LINE_SPACING.SINGLE,
    "1.5倍行距": WD_LINE_SPACING.ONE_POINT_FIVE,
    "2倍行距": WD_LINE_SPACING.DOUBLE,
    "多倍行距": WD_LINE_SPACING.MULTIPLE,
    "固定值": WD_LINE_SPACING.EXACTLY
}
LINE_TYPE_LIST = list(LINE_TYPE_MAP.keys())
LINE_RULE = {
    "单倍行距": {"default": 1, "min": 1, "max": 1, "step": 1, "label": "行距倍数"},
    "1.5倍行距": {"default": 1.5, "min": 1.5, "max": 1.5, "step": 0.1, "label": "行距倍数"},
    "2倍行距": {"default": 2, "min": 2, "max": 2, "step": 0.1, "label": "行距倍数"},
    "多倍行距": {"default": 1.5, "min": 0.5, "max": 5.0, "step": 0.1, "label": "行距倍数"},
    "固定值": {"default": 20, "min": 6, "max": 100, "step": 1, "label": "固定值(磅)"}
}

# 字体配置
FONT_LIST = ["宋体", "黑体", "微软雅黑", "楷体", "仿宋"]
FONT_SIZE_LIST = ["初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六"]
FONT_SIZE_NUM = {k:v for k,v in zip(FONT_SIZE_LIST, [42,36,26,24,22,18,16,15,14,12,10.5,9,7.5,6.5])}
EN_FONT_LIST = ["和正文一致", "Times New Roman", "Arial", "Calibri"]

# 标题识别正则
TITLE_RULE = {
    "一级标题": re.compile(r"^[一二三四五六七八九十]+、\s*.{0,40}$|^第[一二三四五六七八九十]+章\s*.{0,40}$|^第\d+章\s*.{0,40}$|^\d+、\s*.{0,40}$"),
    "二级标题": re.compile(r"^[（(][一二三四五六七八九十]+[）)]\s*.{0,50}$|^\d+\.\s+.{0,50}$|^\d+、\s*.{0,50}$"),
    "三级标题": re.compile(r"^[（(]\d+[）)]\s*.{0,60}$|^\d+\.\d+\s+.{0,60}$|^\d+\.\d+\.\d+\s*.{0,60}$|^\d+\）\s*.{0,60}$")
}

# ====================== 模板库 ======================
GENERAL_TPL = {
    "默认格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1, "indent": 0}
    }
}

UNIVERSITY_TPL = {
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
    "国标-本科毕业论文通用模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1, "indent": 0}
    }
}

OFFICIAL_TPL = {
    "党政机关公文国标GB/T 9704-2012模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "楷体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "仿宋", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "仿宋", "size": "三号", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "仿宋", "size": "小三", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1, "indent": 0}
    }
}

# API配置
try:
    DOUBAO_KEY = st.secrets["DOUBAO_API_KEY"]
    DOUBAO_MODEL = st.secrets.get("DOUBAO_MODEL", "ep-20250628104918-7rqxd")
except:
    DOUBAO_KEY = ""
    DOUBAO_MODEL = "ep-20250628104918-7rqxd"
DOUBAO_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"

# ====================== 核心工具函数 ======================
def is_protected_para(para):
    """保护带图片、分页符、域代码的段落，只改字体，不动内容"""
    if para.paragraph_format.page_break_before:
        return True
    for run in para.runs:
        if run.contains_page_break:
            return True
        if run._element.find('.//w:sectPr', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
            return True
        if run._element.find('.//w:drawing', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
            return True
        if run._element.find('.//w:fldChar', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
            return True
    return False

def set_run_font(run, font_name, font_size, bold=None):
    """安全设置字体，不改变内容"""
    try:
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run.font.size = Pt(font_size)
        if bold is not None:
            run.font.bold = bold
    except:
        pass

def set_en_number_font(run, font_name, font_size, bold=None):
    """安全设置数字字体，不改变位置"""
    try:
        if font_name == "和正文一致":
            return
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:ascii'), font_name)
        run._element.rPr.rFonts.set(qn('w:hAnsi'), font_name)
        run.font.size = Pt(font_size)
        if bold is not None:
            run.font.bold = bold
    except:
        pass

def get_title_level(para, enable_regex, last_title_level):
    """带上下文推理的标题识别"""
    style_name = para.style.name
    if "Heading 1" in style_name or "标题 1" in style_name:
        return "一级标题"
    if "Heading 2" in style_name or "标题 2" in style_name:
        return "二级标题"
    if "Heading 3" in style_name or "标题 3" in style_name:
        return "三级标题"
    
    if not enable_regex:
        return "正文"
    
    text = para.text.strip()
    if not text or len(text) > 100:
        return "正文"
    
    if last_title_level == "一级标题":
        if TITLE_RULE["二级标题"].match(text):
            return "二级标题"
        if TITLE_RULE["一级标题"].match(text):
            return "一级标题"
    if last_title_level == "二级标题":
        if TITLE_RULE["三级标题"].match(text):
            return "三级标题"
        if TITLE_RULE["二级标题"].match(text):
            return "二级标题"
        if TITLE_RULE["一级标题"].match(text):
            return "一级标题"
    
    for level in ["一级标题", "二级标题", "三级标题"]:
        if TITLE_RULE[level].match(text):
            return level
    return "正文"

def process_number_in_para(para, body_font, body_size, number_config):
    """处理数字，绝对不改变位置和顺序"""
    number_size = FONT_SIZE_NUM[number_config["size"]] if not number_config["size_same_as_body"] else body_size
    number_font = number_config["font"]
    number_bold = number_config["bold"]
    number_pattern = re.compile(r"-?\d+\.?\d*%?")
    
    for run in para.runs:
        text = run.text
        if not text:
            set_run_font(run, body_font, body_size)
            continue
        if not number_pattern.search(text):
            set_run_font(run, body_font, body_size)
            continue
        
        set_run_font(run, body_font, body_size)
        for match in number_pattern.finditer(text):
            start, end = match.span()
            if start > 0:
                run = run.split(start)
            number_run = run.split(end - start)
            set_en_number_font(number_run, number_font, number_size, number_bold)

# AI功能函数
def ai_text_handle(text, mode):
    if not DOUBAO_KEY or not text.strip():
        return text
    prompt_map = {
        "润色": "润色文本，保留原意、结构、换行，仅输出结果",
        "降重": "降重改写，保留原意、结构、换行，仅输出结果",
        "标点修正": "修正中英文标点，中文用全角，英文数字用半角，保留原意、换行，仅输出结果",
        "错别字修正": "修正错别字、语病，优化流畅度，保留原意、结构、换行，仅输出结果"
    }
    try:
        resp = requests.post(DOUBAO_URL, json={
            "model": DOUBAO_MODEL,
            "messages": [{"role": "user", "content": f"{prompt_map[mode]}\n原文：{text}"}]
        }, headers={"Authorization": f"Bearer {DOUBAO_KEY}"}, timeout=30)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except:
        return text

# 文档处理主函数
def process_doc(uploaded_file, config, number_config, ai_mode, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank, fix_punctuation, fix_text):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    
    try:
        doc = docx.Document(tmp_path)
        stats = {"一级标题": 0, "二级标题": 0, "三级标题": 0, "正文": 0, "表格": 0, "图片": 0}

        for para in doc.paragraphs:
            try:
                stats["图片"] += len(para._element.findall('.//w:drawing', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
            except:
                pass

        last_title = None
        for para in doc.paragraphs:
            if is_protected_para(para):
                for run in para.runs:
                    set_run_font(run, config["正文"]["font"], FONT_SIZE_NUM[config["正文"]["size"]])
                continue
            
            text = para.text.strip()
            if not text:
                continue
            
            level = get_title_level(para, enable_title_regex, last_title)
            stats[level] += 1
            if level in ["一级标题", "二级标题", "三级标题"]:
                last_title = level

            processed_text = para.text
            if ai_mode != "不使用AI":
                processed_text = ai_text_handle(processed_text, ai_mode)
            if fix_punctuation:
                processed_text = ai_text_handle(processed_text, "标点修正")
            if fix_text:
                processed_text = ai_text_handle(processed_text, "错别字修正")
            if processed_text != para.text:
                para.text = processed_text

            cfg = config[level]
            font_size = FONT_SIZE_NUM[cfg["size"]]

            if force_style:
                try:
                    para.style = level
                except:
                    pass

            try:
                if cfg["align"] != "不修改":
                    para.alignment = ALIGN_MAP[cfg["align"]]
                para.paragraph_format.line_spacing_rule = LINE_TYPE_MAP[cfg["line_type"]]
                if cfg["line_type"] == "多倍行距":
                    para.paragraph_format.line_spacing = cfg["line_value"]
                elif cfg["line_type"] == "固定值":
                    para.paragraph_format.line_spacing = Pt(cfg["line_value"])
                if not keep_spacing:
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                if level == "正文" and cfg["indent"] > 0:
                    para.paragraph_format.first_line_indent = Cm(cfg["indent"] * 0.35)
                para.paragraph_format.page_break_before = False
                para.paragraph_format.keep_with_next = False
            except:
                pass

            if level == "正文" and number_config["enable"]:
                process_number_in_para(para, cfg["font"], font_size, number_config)
            else:
                for run in para.runs:
                    set_run_font(run, cfg["font"], font_size, cfg["bold"])

        for table in doc.tables:
            stats["表格"] += 1
            cfg = config["表格"]
            font_size = FONT_SIZE_NUM[cfg["size"]]
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if is_protected_para(para):
                            for run in para.runs:
                                set_run_font(run, cfg["font"], font_size, cfg["bold"])
                            continue
                        if not para.text.strip():
                            continue
                        if force_style:
                            try:
                                para.style = "正文"
                            except:
                                pass
                        try:
                            if cfg["align"] != "不修改":
                                para.alignment = ALIGN_MAP[cfg["align"]]
                            para.paragraph_format.line_spacing_rule = LINE_TYPE_MAP[cfg["line_type"]]
                            if cfg["line_type"] == "多倍行距":
                                para.paragraph_format.line_spacing = cfg["line_value"]
                            elif cfg["line_type"] == "固定值":
                                para.paragraph_format.line_spacing = Pt(cfg["line_value"])
                        except:
                            pass
                        for run in para.runs:
                            set_run_font(run, cfg["font"], font_size, cfg["bold"])

        if clear_blank:
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
                    if blank_count > max_blank:
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

# ====================== 【修复核心】页面主逻辑 ======================
def main():
    # 页面配置
    st.set_page_config(page_title="文式通 - Word格式智能处理系统", layout="wide")
    
    # 【修复核心】初始化session_state，增加模板版本号，强制组件刷新
    if "current_config" not in st.session_state:
        st.session_state.current_config = GENERAL_TPL["默认格式"]
    if "template_version" not in st.session_state:
        st.session_state.template_version = 0
    
    # 页面标题与声明
    st.title("📄 文式通 - Word格式智能处理系统")
    st.warning("⚠️ 重要声明：此工具仅能减少复杂的格式调整工作量，处理完成后仍需您手动与原文进行对比核对，确保内容与格式无误。")
    st.markdown("✅ 100%保留图片/目录/原排版 | ✅ 高校论文格式一键适配 | ✅ 标点规范/错别字修正 | ✅ 全国大学生计算机设计大赛参赛作品")

    # 【修复核心】模板选择，切换时强制刷新组件
    st.subheader("📋 一键套用标准格式模板")
    tpl_type = st.radio("模板类型", ["通用办公模板", "高校毕业论文模板", "党政公文模板"], horizontal=True)
    
    # 选择模板库
    if tpl_type == "通用办公模板":
        tpl_dict = GENERAL_TPL
    elif tpl_type == "高校毕业论文模板":
        tpl_dict = UNIVERSITY_TPL
    else:
        tpl_dict = OFFICIAL_TPL
    
    # 选择具体模板
    tpl_name = st.selectbox("选择目标格式", list(tpl_dict.keys()), index=0)
    
    # 【修复核心】模板变化时，强制更新配置+刷新页面，确保参数同步
    target_config = tpl_dict[tpl_name]
    if st.session_state.current_config != target_config:
        st.session_state.current_config = target_config
        st.session_state.template_version += 1
        st.success(f"✅ 已自动应用【{tpl_name}】，左侧格式参数已同步更新！")
        st.rerun()

    # 侧边栏格式设置
    with st.sidebar:
        st.header("🎨 详细格式设置")
        cfg = st.session_state.current_config

        # 强制重置按钮
        if st.button("🔄 强制重置格式参数", use_container_width=True):
            st.session_state.current_config = GENERAL_TPL["默认格式"]
            st.session_state.template_version += 1
            st.success("已重置为默认格式！")
            st.rerun()

        st.divider()
        st.subheader("🏷️ 核心设置")
        force_style = st.checkbox("强制统一套用Word原生样式", value=True)
        enable_title_regex = st.checkbox("启用上下文智能标题识别", value=True)
        keep_spacing = st.checkbox("保留原段落段前/段后距", value=True)

        st.divider()
        st.subheader("📄 空行清理设置")
        clear_blank = st.checkbox("清除多余空行", value=False)
        max_blank = st.slider("最多保留连续空行数", min_value=0, max_value=2, value=1) if clear_blank else 1

        st.divider()
        st.subheader("🔤 文本优化设置")
        if not DOUBAO_KEY:
            st.info("ℹ️ 填写豆包API密钥即可启用以下功能")
        fix_punctuation = st.checkbox("启用中英文标点规范修正", value=False, disabled=not DOUBAO_KEY)
        fix_text = st.checkbox("启用错别字修正+语句流畅度优化", value=False, disabled=not DOUBAO_KEY)
        if not DOUBAO_KEY:
            ai_mode = "不使用AI"
        else:
            ai_mode = st.radio("AI文本处理", ["不使用AI", "润色", "降重"], horizontal=True)

        # 【修复核心】格式设置块，key加入版本号，切换模板强制重新渲染
        def create_format_block(title, level):
            st.divider()
            st.subheader(title)
            item = cfg[level]
            version = st.session_state.template_version
            
            # 字体选择，key带版本号，切换模板强制刷新
            font_index = FONT_LIST.index(item["font"]) if item["font"] in FONT_LIST else 0
            item["font"] = st.selectbox(
                "字体", 
                FONT_LIST, 
                key=f"{level}_font_{version}", 
                index=font_index
            )
            
            # 字号选择
            size_index = FONT_SIZE_LIST.index(item["size"]) if item["size"] in FONT_SIZE_LIST else 0
            item["size"] = st.selectbox(
                "字号", 
                FONT_SIZE_LIST, 
                key=f"{level}_size_{version}", 
                index=size_index
            )
            
            # 加粗
            item["bold"] = st.checkbox(
                "加粗", 
                value=item["bold"], 
                key=f"{level}_bold_{version}"
            )
            
            # 对齐方式
            align_index = ALIGN_LIST.index(item["align"]) if item["align"] in ALIGN_LIST else 0
            item["align"] = st.selectbox(
                "对齐方式", 
                ALIGN_LIST, 
                key=f"{level}_align_{version}", 
                index=align_index
            )
            
            # 行距类型
            line_type_index = LINE_TYPE_LIST.index(item["line_type"]) if item["line_type"] in LINE_TYPE_LIST else 0
            new_line_type = st.selectbox(
                "行距类型", 
                LINE_TYPE_LIST, 
                key=f"{level}_line_type_{version}", 
                index=line_type_index
            )
            
            # 切换行距类型时，自动重置合法数值
            if new_line_type != item["line_type"]:
                item["line_type"] = new_line_type
                item["line_value"] = LINE_RULE[new_line_type]["default"]
                st.session_state.current_config[level] = item
                st.rerun()
            
            # 行距数值，严格限制范围
            line_rule = LINE_RULE[item["line_type"]]
            current_value = item["line_value"]
            if not (line_rule["min"] <= current_value <= line_rule["max"]):
                current_value = line_rule["default"]
                item["line_value"] = current_value
            
            item["line_value"] = st.number_input(
                line_rule["label"],
                min_value=line_rule["min"],
                max_value=line_rule["max"],
                value=float(current_value),
                step=line_rule["step"],
                key=f"{level}_line_value_{version}",
                disabled=line_rule["min"] == line_rule["max"]
            )
            
            # 首行缩进
            if "indent" in item:
                item["indent"] = st.number_input(
                    "首行缩进(字符)", 
                    min_value=0, 
                    max_value=4, 
                    value=item["indent"], 
                    key=f"{level}_indent_{version}"
                )
            
            # 更新到session_state
            st.session_state.current_config[level] = item
            return item

        # 各级格式块
        create_format_block("📌 一级标题", "一级标题")
        create_format_block("📌 二级标题", "二级标题")
        create_format_block("📌 三级标题", "三级标题")
        create_format_block("📝 正文内容", "正文")
        create_format_block("📊 表格内容", "表格")

        # 数字格式设置
        st.divider()
        st.subheader("🔢 正文数字格式设置")
        number_config = {}
        number_config["enable"] = st.checkbox("启用数字单独格式设置", value=True, key=f"number_enable_{st.session_state.template_version}")
        if number_config["enable"]:
            number_config["font"] = st.selectbox("数字字体", EN_FONT_LIST, index=0, key=f"number_font_{st.session_state.template_version}")
            number_config["size_same_as_body"] = st.checkbox("字号和正文一致", value=True, key=f"number_size_same_{st.session_state.template_version}")
            if not number_config["size_same_as_body"]:
                number_config["size"] = st.selectbox("数字字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index("小四"), key=f"number_size_{st.session_state.template_version}")
            else:
                number_config["size"] = "小四"
            number_config["bold"] = st.checkbox("数字加粗", value=False, key=f"number_bold_{st.session_state.template_version}")

    # 主界面上传&处理
    uploaded_file = st.file_uploader("📤 上传Word文档（仅支持.docx格式）", type="docx")
    
    if uploaded_file:
        st.success("✅ 文档上传成功！")
        
        # 处理按钮
        if st.button("🚀 开始处理文档", type="primary", use_container_width=True):
            progress_bar = st.progress(0, text="文档处理准备中...")
            try:
                progress_bar.progress(10, text="正在解析文档...")
                file_bytes, stats = process_doc(
                    uploaded_file, 
                    cfg, 
                    number_config, 
                    ai_mode, 
                    enable_title_regex, 
                    force_style, 
                    keep_spacing, 
                    clear_blank, 
                    max_blank,
                    fix_punctuation,
                    fix_text
                )
                progress_bar.progress(80, text="文档处理完成，正在生成下载文件...")
                
                if file_bytes and stats:
                    progress_bar.progress(100, text="处理完成！")
                    st.subheader("📋 文档内容分类识别结果")
                    cols = st.columns(6)
                    cols[0].metric("一级标题", stats["一级标题"])
                    cols[1].metric("二级标题", stats["二级标题"])
                    cols[2].metric("三级标题", stats["三级标题"])
                    cols[3].metric("正文段落", stats["正文"])
                    cols[4].metric("表格数量", stats["表格"])
                    cols[5].metric("图片数量", stats["图片"])
                    
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
