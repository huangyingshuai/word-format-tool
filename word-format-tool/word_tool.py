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
    "三级标题": re.compile(r"^[（(]\d+[）)]\s*.{0,60}$|^\d+\.\d+\s+.{0,60}$|^\d+\.\d+\.\d+\s+.{0,60}$|^\d+\）\s*.{0,60}$")
}

# ====================== 你的扣子降重智能体（已替换） ======================
COZE_AI_URL = "https://www.coze.cn/s/Dtw5_DzeCIo/"
BOT_NAME = "专业降重智能体"

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

# ====================== AI 降重函数（已对接你的扣子智能体） ======================
def ai_rewrite(text, mode):
    if not text.strip():
        return text

    if mode == "专业降重":
        prompt = f"""请对下面文本进行学术降重，不改变原意、专业名词、数据，语句通顺自然，只返回结果：
{text}"""
    elif mode == "润色":
        prompt = f"请润色下面文本，语句更通顺，保留原意，只返回结果：{text}"
    elif mode == "标点修正":
        prompt = f"修正下面文本的中英文标点，只返回结果：{text}"
    elif mode == "错别字修正":
        prompt = f"修正下面文本的错别字和语病，保留原意，只返回结果：{text}"
    else:
        return text

    try:
        # 调用你的扣子降重智能体
        resp = requests.post(
            COZE_AI_URL,
            json={"query": prompt},
            timeout=60
        )
        if resp.status_code == 200:
            data = resp.json()
            if "content" in data:
                return data["content"].strip()
            elif "choices" in data and len(data["choices"]) > 0:
                return data["choices"][0]["message"]["content"].strip()
        return text
    except Exception as e:
        st.warning(f"AI降重暂时不可用，已保留原文")
        return text

# ====================== 文档处理 ======================
def process_doc(uploaded_file, config, number_config, ai_mode, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank, fix_punctuation, fix_text):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    try:
        doc = docx.Document(tmp_path)
        stats = {"一级标题": 0, "二级标题": 0, "三级标题": 0, "正文": 0, "表格": 0, "图片": 0}

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
                processed_text = ai_rewrite(processed_text, ai_mode)
            if fix_punctuation:
                processed_text = ai_rewrite(processed_text, "标点修正")
            if fix_text:
                processed_text = ai_rewrite(processed_text, "错别字修正")

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
        st.error(f"处理失败：{str(e)}")
        return None, None
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        if 'output_path' in locals() and os.path.exists(output_path):
            os.unlink(output_path)

# ====================== 页面主逻辑 ======================
def main():
    st.set_page_config(page_title="文式通 - Word格式智能处理系统", layout="wide")
    if "current_config" not in st.session_state:
        st.session_state.current_config = GENERAL_TPL["默认格式"]
    if "template_version" not in st.session_state:
        st.session_state.template_version = 0

    st.title("📄 文式通 - Word格式智能处理系统")
    st.warning("⚠️ 处理后请人工核对内容与格式，确保无误。")
    st.markdown("✅ 100%保留图片/目录/公式 | ✅ 高校论文/公文模板 | ✅ AI智能降重 | ✅ 标点/错别字修正")
    # 展示你的扣子降重智能体链接
    st.info(f"🔗 当前使用AI降重智能体：[{BOT_NAME}]({COZE_AI_URL})")

    # 模板选择
    st.subheader("📋 选择格式模板")
    tpl_type = st.radio("模板类型", ["通用办公模板", "高校毕业论文模板", "党政公文模板"], horizontal=True)
    tpl_dict = GENERAL_TPL if tpl_type == "通用办公模板" else UNIVERSITY_TPL if tpl_type == "高校毕业论文模板" else OFFICIAL_TPL
    tpl_name = st.selectbox("选择目标格式", list(tpl_dict.keys()), index=0)
    target_config = tpl_dict[tpl_name]

    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("✅ 应用选中模板", type="primary", use_container_width=True):
            st.session_state.current_config = target_config
            st.session_state.template_version += 1
            st.rerun()
    with col2:
        st.caption("选择模板后点击应用，左侧参数自动同步")

    # 侧边栏
    with st.sidebar:
        st.header("🎨 格式详细设置")
        cfg = st.session_state.current_config
        if st.button("🔄 重置为默认格式", use_container_width=True):
            st.session_state.current_config = GENERAL_TPL["默认格式"]
            st.session_state.template_version += 1
            st.rerun()

        st.divider()
        st.subheader("核心设置")
        force_style = st.checkbox("强制统一Word样式", value=True)
        enable_title_regex = st.checkbox("智能标题识别", value=True)
        keep_spacing = st.checkbox("保留段前/段后距", value=True)

        st.divider()
        st.subheader("空行清理")
        clear_blank = st.checkbox("清除多余空行", value=False)
        max_blank = st.slider("最多保留连续空行数", 0, 3, 1) if clear_blank else 1

        st.divider()
        st.subheader("AI 降重 / 润色")
        fix_punctuation = st.checkbox("修正标点")
        fix_text = st.checkbox("修正错别字")
        ai_mode = st.radio("AI处理模式", ["不使用AI", "润色", "专业降重"], horizontal=True)

        # 格式块
        def create_format_block(title, level):
            st.divider()
            st.subheader(title)
            item = cfg[level]
            v = st.session_state.template_version

            item["font"] = st.selectbox("字体", FONT_LIST, index=FONT_LIST.index(item["font"]), key=f"{level}_font_{v}")
            item["size"] = st.selectbox("字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(item["size"]), key=f"{level}_size_{v}")
            item["bold"] = st.checkbox("加粗", item["bold"], key=f"{level}_bold_{v}")
            item["align"] = st.selectbox("对齐", ALIGN_LIST, index=ALIGN_LIST.index(item["align"]), key=f"{level}_align_{v}")

            new_line = st.selectbox("行距类型", LINE_TYPE_LIST, index=LINE_TYPE_LIST.index(item["line_type"]), key=f"{level}_lt_{v}")
            if new_line != item["line_type"]:
                item["line_type"] = new_line
                item["line_value"] = LINE_RULE[new_line]["default"]
                st.session_state.current_config[level] = item
                st.rerun()

            rule = LINE_RULE[item["line_type"]]
            item["line_value"] = st.number_input(rule["label"], rule["min"], rule["max"], float(item["line_value"]), rule["step"], key=f"{level}_lv_{v}", disabled=rule["min"]==rule["max"])

            if "indent" in item:
                item["indent"] = st.number_input("首行缩进(字符)", 0,4,item["indent"],key=f"{level}_indent_{v}")
            st.session_state.current_config[level] = item
            return item

        create_format_block("一级标题", "一级标题")
        create_format_block("二级标题", "二级标题")
        create_format_block("三级标题", "三级标题")
        create_format_block("正文", "正文")
        create_format_block("表格", "表格")

        st.divider()
        st.subheader("数字/英文格式")
        number_config = {"enable": st.checkbox("启用数字单独格式", True, key=f"num_en_{st.session_state.template_version}")}
        if number_config["enable"]:
            number_config["font"] = st.selectbox("数字字体", EN_FONT_LIST, 0, key=f"num_font_{st.session_state.template_version}")
            number_config["size_same_as_body"] = st.checkbox("字号同正文", True, key=f"num_same_{st.session_state.template_version}")
            number_config["size"] = st.selectbox("数字字号", FONT_SIZE_LIST, FONT_SIZE_LIST.index("小四"), key=f"num_size_{st.session_state.template_version}") if not number_config["size_same_as_body"] else "小四"
            number_config["bold"] = st.checkbox("数字加粗", False, key=f"num_bold_{st.session_state.template_version}")

    # 上传与处理
    st.divider()
    uploaded_file = st.file_uploader("上传 .docx 文档", type="docx")
    if uploaded_file:
        st.success("✅ 文档上传成功")
        if st.button("🚀 开始处理（格式+AI降重）", type="primary", use_container_width=True):
            with st.spinner("正在处理..."):
                data, stats = process_doc(uploaded_file, cfg, number_config, ai_mode, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank, fix_punctuation, fix_text)
                if data and stats:
                    st.subheader("📋 处理完成")
                    c1,c2,c3,c4,c5,c6 = st.columns(6)
                    c1.metric("一级标题", stats["一级标题"])
                    c2.metric("二级标题", stats["二级标题"])
                    c3.metric("三级标题", stats["三级标题"])
                    c4.metric("正文", stats["正文"])
                    c5.metric("表格", stats["表格"])
                    c6.metric("图片", stats["图片"])

                    st.download_button("📥 下载处理后文档", data, f"已降重排版_{uploaded_file.name}", use_container_width=True)
                    st.success("🎉 处理完成！")

    st.divider()
    with st.expander("📖 版权与声明"):
        st.markdown("本工具用于论文/公文格式自动化处理，仅辅助使用。")
        st.markdown(f"AI降重由扣子智能体提供：{COZE_AI_URL}")

if __name__ == "__main__":
    main()
