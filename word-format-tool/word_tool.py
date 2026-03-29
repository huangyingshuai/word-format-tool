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
    "固定值": {"default": 20.0, "min": 6.0, "max": 100.0, "step": 1.0, "label": "固定值(磅)"}
}

FONT_LIST = ["宋体", "黑体", "微软雅黑", "楷体", "仿宋"]
FONT_SIZE_LIST = ["初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六"]
FONT_SIZE_NUM = {k:v for k,v in zip(FONT_SIZE_LIST, [42.0,36.0,26.0,24.0,22.0,18.0,16.0,15.0,14.0,12.0,10.5,9.0,7.5,6.5])}
EN_FONT_LIST = ["和正文一致", "Times New Roman", "Arial", "Calibri"]

TITLE_RULE = {
    "一级标题": re.compile(r"^[一二三四五六七八九十]+、\s*.{0,40}$|^第[一二三四五六七八九十]+章\s*.{0,40}$|^第\d+章\s*.{0,40}$|^\d+、\s*.{0,40}$"),
    "二级标题": re.compile(r"^[（(][一二三四五六七八九十]+[）)]\s*.{0,50}$|^\d+\.\s+.{0,50}$|^\d+、\s*.{0,50}$"),
    "三级标题": re.compile(r"^[（(]\d+[）)]\s*.{0,60}$|^\d+\.\d+\s+.{0,60}$|^\d+\.\d+\.\d+\s*.{0,60}$|^\d+\）\s*.{0,60}$")
}

DOUBAO_MODEL = "ep-20250628104918-7rqxd"
DOUBAO_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"

# ====================== 高校/公文格式模板库 ======================
GENERAL_TPL = {
    "默认格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0}
    }
}

UNIVERSITY_TPL = {
    "河北科技大学-本科毕业论文模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0}
    },
    "河北工业大学-本科毕业论文模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "楷体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0}
    },
    "燕山大学-本科毕业论文模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "固定值", "line_value": 20.0, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0}
    },
    "国标-本科毕业论文通用模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0}
    }
}

OFFICIAL_TPL = {
    "党政机关公文国标GB/T 9704-2012模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "楷体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "仿宋", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "仿宋", "size": "三号", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "仿宋", "size": "小三", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0}
    }
}

# ====================== 核心工具函数 ======================
def is_protected_para(para):
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
    try:
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run.font.size = Pt(font_size)
        if bold is not None:
            run.font.bold = bold
    except:
        pass

def set_en_number_font(run, font_name, font_size, bold=None):
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
    number_size = FONT_SIZE_NUM[number_config["size"]] if not number_config["size_same_as_body"] else body_size
    number_font = number_config["font"]
    number_bold = number_config["bold"]
    number_pattern = re.compile(r"-?\d+\.?\d*%?")
    new_runs = []
    for run in para.runs:
        text = run.text
        if not text:
            new_runs.append(run)
            continue
        if not number_pattern.search(text):
            set_run_font(run, body_font, body_size)
            new_runs.append(run)
            continue
        parts = []
        last_end = 0
        for match in number_pattern.finditer(text):
            start, end = match.span()
            if start > last_end:
                parts.append(("text", text[last_end:start]))
            parts.append(("number", text[start:end]))
            last_end = end
        if last_end < len(text):
            parts.append(("text", text[last_end:]))
        run.text = ""
        for part_type, part_text in parts:
            new_run = para.add_run(part_text)
            if part_type == "text":
                set_run_font(new_run, body_font, body_size)
            else:
                set_en_number_font(new_run, number_font, number_size, number_bold)
            new_runs.append(new_run)
    for run in para.runs:
        run._element.getparent().remove(run._element)
    for new_run in new_runs:
        para._element.append(new_run._element)

# ====================== 【专业AI降重】严格遵循你的降重方法论 ======================
def ai_text_handle(text, mode, doubao_key):
    if not doubao_key or not text.strip():
        return text
    
    # 降重prompt：完全按照你给的全场景通用降重方法论编写
    if mode == "专业降重":
        prompt = f"""你是专业学术降重工程师，严格遵循以下降重规则处理文本，仅输出降重后的结果，不解释、不添加额外内容：
1. 核心原则：破连续字符匹配+破AI生成特征，不改动原意、核心数据、专有名词、法律法规
2. 语义重构四步法：拆解核心要素→更换叙事逻辑→补充场景化描述→删除套话
3. 句式要求：1长句搭配1-2短句，制造句式波动，提升困惑度
4. 替换AI套话：首先/其次→从落地场景来看；综上所述→结合全维度分析；一方面→站在需求端
5. 注入人类特征：加入合理过渡词、限定范围，避免泛泛而谈
6. 严禁：同义词替换、中英互译、修改核心术语、大段删减
原文：{text}"""
    elif mode == "润色":
        prompt = f"润色文本，保留原意、结构、换行，语句更流畅，仅输出结果：{text}"
    elif mode == "标点修正":
        prompt = f"修正中英文标点，中文全角、英文半角，保留原意换行，仅输出结果：{text}"
    elif mode == "错别字修正":
        prompt = f"修正错别字、语病，优化流畅度，保留原意结构，仅输出结果：{text}"
    
    try:
        resp = requests.post(DOUBAO_URL, json={
            "model": DOUBAO_MODEL,
            "messages": [{"role": "user", "content": prompt}]
        }, headers={"Authorization": f"Bearer {doubao_key}"}, timeout=30)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except:
        return text

def process_doc(uploaded_file, config, number_config, ai_mode, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank, fix_punctuation, fix_text, doubao_key):
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
            
            # AI文本处理（含专业降重）
            processed_text = para.text
            if ai_mode == "专业降重":
                processed_text = ai_text_handle(processed_text, "专业降重", doubao_key)
            elif ai_mode == "润色":
                processed_text = ai_text_handle(processed_text, "润色", doubao_key)
            if fix_punctuation:
                processed_text = ai_text_handle(processed_text, "标点修正", doubao_key)
            if fix_text:
                processed_text = ai_text_handle(processed_text, "错别字修正", doubao_key)
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

# ====================== 页面主逻辑（新增专业降重选项） ======================
def main():
    st.set_page_config(page_title="文式通 - Word格式智能处理系统", layout="wide")
    if "current_config" not in st.session_state:
        st.session_state.current_config = GENERAL_TPL["默认格式"]
    if "template_version" not in st.session_state:
        st.session_state.template_version = 0
    if "doubao_api_key" not in st.session_state:
        st.session_state.doubao_api_key = st.secrets.get("DOUBAO_API_KEY", "")

    st.title("📄 文式通 - Word格式智能处理系统")
    st.warning("⚠️ 重要声明：此工具仅能减少复杂的格式调整工作量，处理完成后仍需您手动与原文进行对比核对，确保内容与格式无误。")
    st.markdown("✅ 100%保留图片/目录/原排版 | ✅ 高校论文格式一键适配 | ✅ 专业AI降重 | ✅ 标点规范/错别字修正")

    # 模板选择
    st.subheader("📋 一键套用标准格式模板")
    tpl_type = st.radio("模板类型", ["通用办公模板", "高校毕业论文模板", "党政公文模板"], horizontal=True)
    tpl_dict = GENERAL_TPL if tpl_type == "通用办公模板" else UNIVERSITY_TPL if tpl_type == "高校毕业论文模板" else OFFICIAL_TPL
    tpl_name = st.selectbox("选择目标格式", list(tpl_dict.keys()), index=0)
    
    target_config = tpl_dict[tpl_name]
    if st.session_state.current_config != target_config:
        st.session_state.current_config = target_config
        st.session_state.template_version += 1
        st.success(f"✅ 已自动应用【{tpl_name}】，左侧格式参数已同步更新！")
        st.rerun()

    # 侧边栏设置
    with st.sidebar:
        st.header("🎨 详细格式设置")
        cfg = st.session_state.current_config
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
        max_blank = st.slider("最多保留连续空行数", 0, 2, 1) if clear_blank else 1

        st.divider()
        st.subheader("🔤 AI文本优化（含专业降重）")
        st.text_input("豆包API密钥", type="password", key="doubao_api_key", placeholder="输入密钥启用AI功能")
        DOUBAO_KEY = st.session_state.get("doubao_api_key", "") or st.secrets.get("DOUBAO_API_KEY", "")
        
        if not DOUBAO_KEY:
            st.info("ℹ️ 填写API密钥即可启用降重/润色功能")
        fix_punctuation = st.checkbox("修正中英文标点", False, disabled=not DOUBAO_KEY)
        fix_text = st.checkbox("修正错别字/语病", False, disabled=not DOUBAO_KEY)
        
        # 新增：专业降重选项
        ai_mode = "不使用AI"
        if DOUBAO_KEY:
            ai_mode = st.radio("AI处理模式", ["不使用AI", "润色", "专业降重"], horizontal=True)

        # 格式块生成
        def create_format_block(title, level):
            st.divider()
            st.subheader(title)
            item = cfg[level]
            version = st.session_state.template_version
            
            font_idx = FONT_LIST.index(item["font"]) if item["font"] in FONT_LIST else 0
            item["font"] = st.selectbox("字体", FONT_LIST, key=f"{level}_font_{version}", index=font_idx)
            size_idx = FONT_SIZE_LIST.index(item["size"]) if item["size"] in FONT_SIZE_LIST else 0
            item["size"] = st.selectbox("字号", FONT_SIZE_LIST, key=f"{level}_size_{version}", index=size_idx)
            item["bold"] = st.checkbox("加粗", item["bold"], key=f"{level}_bold_{version}")
            align_idx = ALIGN_LIST.index(item["align"]) if item["align"] in ALIGN_LIST else 0
            item["align"] = st.selectbox("对齐方式", ALIGN_LIST, key=f"{level}_align_{version}", index=align_idx)
            
            line_type_idx = LINE_TYPE_LIST.index(item["line_type"]) if item["line_type"] in LINE_TYPE_LIST else 0
            new_line_type = st.selectbox("行距类型", LINE_TYPE_LIST, key=f"{level}_line_type_{version}", index=line_type_idx)
            if new_line_type != item["line_type"]:
                item["line_type"] = new_line_type
                item["line_value"] = LINE_RULE[new_line_type]["default"]
                st.session_state.current_config[level] = item
                st.rerun()
            
            line_rule = LINE_RULE[item["line_type"]]
            curr_val = float(item["line_value"])
            if not (line_rule["min"] <= curr_val <= line_rule["max"]):
                curr_val = line_rule["default"]
                item["line_value"] = curr_val
            item["line_value"] = st.number_input(line_rule["label"], line_rule["min"], line_rule["max"], curr_val, line_rule["step"], key=f"{level}_line_value_{version}", disabled=line_rule["min"]==line_rule["max"])
            
            if "indent" in item:
                item["indent"] = st.number_input("首行缩进(字符)", 0,4,item["indent"],key=f"{level}_indent_{version}")
            st.session_state.current_config[level] = item
            return item

        create_format_block("📌 一级标题", "一级标题")
        create_format_block("📌 二级标题", "二级标题")
        create_format_block("📌 三级标题", "三级标题")
        create_format_block("📝 正文内容", "正文")
        create_format_block("📊 表格内容", "表格")

        # 数字格式
        st.divider()
        st.subheader("🔢 正文数字格式设置")
        number_config = {"enable": st.checkbox("启用数字单独格式", True, key=f"num_en_{st.session_state.template_version}")}
        if number_config["enable"]:
            number_config["font"] = st.selectbox("数字字体", EN_FONT_LIST, 0, key=f"num_font_{st.session_state.template_version}")
            number_config["size_same_as_body"] = st.checkbox("字号同正文", True, key=f"num_size_{st.session_state.template_version}")
            number_config["size"] = st.selectbox("数字字号", FONT_SIZE_LIST, FONT_SIZE_LIST.index("小四"), key=f"num_size_val_{st.session_state.template_version}") if not number_config["size_same_as_body"] else "小四"
            number_config["bold"] = st.checkbox("数字加粗", False, key=f"num_bold_{st.session_state.template_version}")

    # 主界面上传处理
    uploaded_file = st.file_uploader("📤 上传Word文档（.docx）", type="docx")
    if uploaded_file:
        st.success("✅ 文档上传成功！")
        if st.button("🚀 开始处理（格式调整+AI降重/润色）", type="primary", use_container_width=True):
            bar = st.progress(0, "准备处理...")
            try:
                bar.progress(10, "解析文档结构")
                data, stats = process_doc(uploaded_file, cfg, number_config, ai_mode, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank, fix_punctuation, fix_text, DOUBAO_KEY)
                bar.progress(80, "生成处理后文档")
                if data and stats:
                    bar.progress(100, "处理完成！")
                    st.subheader("📋 文档识别统计")
                    c1,c2,c3,c4,c5,c6 = st.columns(6)
                    c1.metric("一级标题", stats["一级标题"])
                    c2.metric("二级标题", stats["二级标题"])
                    c3.metric("三级标题", stats["三级标题"])
                    c4.metric("正文段落", stats["正文"])
                    c5.metric("表格", stats["表格"])
                    c6.metric("图片", stats["图片"])
                    
                    st.download_button("📥 下载处理完成文档", data, f"降重排版完成_{uploaded_file.name}", use_container_width=True)
                    st.success("🎉 文档处理完成！")
            except Exception as e:
                st.error(f"失败：{str(e)}")
            finally:
                bar.empty()

    st.divider()
    with st.expander("📖 版权与声明"):
        st.markdown("本作品为大学生计算机设计大赛参赛作品，基于Streamlit、python-docx开发，遵守开源协议。")
        st.markdown("⚠️ 降重后内容建议人工核对，确保学术规范与原意一致。")

if __name__ == "__main__":
    main()
