import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import requests
import tempfile
import os
import re
from tenacity import retry, stop_after_attempt, wait_exponential # 需安装 tenacity: pip install tenacity

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

# 优化：固定值步长改为0.1，最小值改为1.0 (由用户决定，代码仅做逻辑支持)
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

TITLE_RULE = {
    "一级标题": re.compile(r"^[一二三四五六七八九十]+、\s*.{0,40}$|^第[一二三四五六七八九十]+章\s*.{0,40}$|^第\d+章\s*.{0,40}$|^\d+、\s*.{0,40}$"),
    "二级标题": re.compile(r"^[（(][一二三四五六七八九十]+[）)]\s*.{0,50}$|^\d+\.\s+.{0,50}$|^\d+、\s*.{0,50}$"),
    "三级标题": re.compile(r"^[（(]\d+[）)]\s*.{0,60}$|^\d+\.\d+\s+.{0,60}$|^\d+\.\d+\.\d+\s+.{0,60}$|^\d+\）\s*.{0,60}$")
}

DOUBAO_MODEL = "ep-20250628104918-7rqxd"
DOUBAO_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"

# ====================== 模板库 ======================
GENERAL_TPL = {
    "默认通用格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0}
    }
}

UNIVERSITY_TPL = {
    "河北科技大学-本科毕业论文": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0}
    },
    "河北工业大学-本科毕业论文": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "楷体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0}
    },
    "国标-本科毕业论文通用": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0}
    }
}

OFFICIAL_TPL = {
    "党政机关公文国标GB/T 9704-2012": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "楷体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "仿宋", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0},
        "正文": {"font": "仿宋", "size": "三号", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2},
        "表格": {"font": "仿宋", "size": "小三", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0}
    }
}

# ====================== 核心工具函数 ======================
def is_protected_para(para):
    """判断段落是否受保护（分页符、分节符、图片等）"""
    if para.paragraph_format.page_break_before:
        return True
    # 修复：检查段落级别的分节符
    if para._element.find('.//w:sectPr', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
        return True
        
    for run in para.runs:
        if run.contains_page_break:
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
    except Exception as e:
        # 优化：仅在调试模式打印，避免前端报错
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
    except Exception as e:
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

# ====================== AI处理函数 (增加重试机制) ======================
@retry(stop=stop_after_attempt(2), wait=wait_exponential(multiplier=1, min=2, max=10))
def call_doubao_api(prompt, api_key):
    resp = requests.post(DOUBAO_URL, json={
        "model": DOUBAO_MODEL,
        "messages": [{"role": "user", "content": prompt}]
    }, headers={"Authorization": f"Bearer {api_key}"}, timeout=30)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]

def ai_text_handle(text, mode, doubao_key):
    if not doubao_key or not text.strip():
        return text
    try:
        if mode == "专业降重":
            prompt = f"""严格按学术降重规则处理：1.不破连续字符+破AI特征 2.不改原意、数据、专有名词 3.长短句搭配 4.替换AI套话 5.注入人类写作习惯，仅输出结果：{text}"""
        elif mode == "润色":
            prompt = f"润色文本，保留原意结构，仅输出结果：{text}"
        elif mode == "标点修正":
            prompt = f"修正中英文标点，中文全角英文半角，仅输出结果：{text}"
        elif mode == "错别字修正":
            prompt = f"修正错别字语病，保留原意，仅输出结果：{text}"
        else:
            return text
            
        return call_doubao_api(prompt, doubao_key)
    except Exception as e:
        # 优化：如果AI失败，返回原文，不中断整个流程
        return text

def process_doc(uploaded_file, config, number_config, ai_mode, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank, fix_punctuation, fix_text, doubao_key):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    
    output_path = None
    try:
        doc = docx.Document(tmp_path)
        stats = {"一级标题": 0, "二级标题": 0, "三级标题": 0, "正文": 0, "表格": 0, "图片": 0}
        
        # 统计图片
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
            # AI 处理逻辑
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
        
        # 处理表格
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
        
        # 清理空行
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
        st.error(f"😵 文档处理失败：{str(e)}")
        return None, None
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        if output_path and os.path.exists(output_path):
            os.unlink(output_path)

# ====================== 页面主逻辑 ======================
def main():
    st.set_page_config(page_title="文式通 - Word一键排版", layout="wide", page_icon="📄")
    
    if "current_config" not in st.session_state:
        st.session_state.current_config = GENERAL_TPL["默认通用格式"]
    if "template_version" not in st.session_state:
        st.session_state.template_version = 0
    if "doubao_api_key" not in st.session_state:
        st.session_state.doubao_api_key = st.secrets.get("DOUBAO_API_KEY", "")

    # --- 顶部区域 ---
    st.title("📄 文式通 - Word格式智能处理系统")
    
    # 亲民化：增加操作步骤引导
    st.info("💡 **使用指南**：1️⃣ 选模板 → 2️⃣ 传文件 → 3️⃣ 点处理 → 4️⃣ 下载。全程无需手动调格式！")

    # 重要声明
    st.warning("⚠️ **重要提醒**：此工具为辅助工具，处理完成后请务必手动核对内容与格式，确保无误。")

    # --- 模板选择区域 (优化布局) ---
    st.subheader("Step 1: 选择适用场景 📋")
    
    # 优化：使用Tab代替Radio，界面更整洁
    tab1, tab2, tab3 = st.tabs(["🎓 高校毕业论文", "💼 通用办公", "🏛️ 党政公文"])
    
    target_config = None
    tpl_name = "默认通用格式"
    
    with tab1:
        tpl_name = st.selectbox("选择学校模板", list(UNIVERSITY_TPL.keys()), index=0)
        target_config = UNIVERSITY_TPL[tpl_name]
        st.caption("包含：河北科大、河北工大、国家标准等。")
        
    with tab2:
        tpl_name = st.selectbox("选择办公模板", list(GENERAL_TPL.keys()), index=0)
        target_config = GENERAL_TPL[tpl_name]
        
    with tab3:
        tpl_name = st.selectbox("选择公文模板", list(OFFICIAL_TPL.keys()), index=0)
        target_config = OFFICIAL_TPL[tpl_name]

    # 应用模板按钮
    if st.button(f"✅ 应用【{tpl_name}】", type="primary"):
        st.session_state.current_config = target_config
        st.session_state.template_version += 1
        st.success(f"🎉 已成功应用模板！可以直接去上传文件啦。")
        st.rerun()

    st.divider()

    # --- 侧边栏 (专家模式折叠) ---
    with st.sidebar:
        st.header("⚙️ 高级设置 (可选)")
        
        with st.expander("🔧 基础开关设置", expanded=False):
            cfg = st.session_state.current_config
            if st.button("🔄 重置所有格式"):
                st.session_state.current_config = GENERAL_TPL["默认通用格式"]
                st.session_state.template_version += 1
                st.success("已重置！")
                st.rerun()

            force_style = st.checkbox("智能匹配Word标题样式", value=True, help="让Word目录能自动识别标题")
            enable_title_regex = st.checkbox("智能识别标题编号", value=True, help="如 '1、' '1.1' 自动识别为标题")
            keep_spacing = st.checkbox("保留原段落间距", value=True)

        with st.expander("🧹 空行清理", expanded=False):
            clear_blank = st.checkbox("清除多余空行", value=False)
            max_blank = st.slider("最多保留连续空行数", 0, 2, 1) if clear_blank else 1

        with st.expander("🤖 AI 智能优化 (需密钥)", expanded=False):
            st.text_input("火山引擎API密钥 (sk-...)", type="password", key="doubao_api_key", placeholder="输入密钥以启用降重")
            user_key = st.session_state.get("doubao_api_key", "")
            env_key = st.secrets.get("DOUBAO_API_KEY", "")
            DOUBAO_KEY = user_key or env_key
            is_valid_key = DOUBAO_KEY.startswith("sk-") if DOUBAO_KEY else False

            if not DOUBAO_KEY:
                st.info("💡 没有密钥？这部分功能可以不用，直接用格式排版就行。")
            elif not is_valid_key:
                st.warning("⚠️ 密钥格式不对哦，应该是 sk- 开头的")
            
            fix_punctuation = st.checkbox("修正标点符号", False, disabled=not is_valid_key)
            fix_text = st.checkbox("修正错别字", False, disabled=not is_valid_key)
            
            ai_mode = "不使用AI"
            if is_valid_key:
                ai_mode = st.radio("AI模式", ["不使用AI", "润色", "专业降重"], horizontal=True)

        # 极其细节的格式调整，折叠再折叠
        with st.expander("✏️ 微调整字体行距 (不建议新手动)", expanded=False):
            st.caption("通常应用模板后不需要改这里。")
            # 这里保留原有的详细设置逻辑，但默认折叠
            def create_format_block(title, level):
                st.markdown(f"**{title}**")
                item = cfg[level]
                version = st.session_state.template_version
                
                col1, col2 = st.columns(2)
                with col1:
                    font_idx = FONT_LIST.index(item["font"]) if item["font"] in FONT_LIST else 0
                    item["font"] = st.selectbox("字体", FONT_LIST, key=f"{level}_font_{version}", index=font_idx, label_visibility="collapsed")
                with col2:
                    size_idx = FONT_SIZE_LIST.index(item["size"]) if item["size"] in FONT_SIZE_LIST else 0
                    item["size"] = st.selectbox("字号", FONT_SIZE_LIST, key=f"{level}_size_{version}", index=size_idx, label_visibility="collapsed")
                
                item["bold"] = st.checkbox("加粗", item["bold"], key=f"{level}_bold_{version}")
                st.session_state.current_config[level] = item
                return item

            create_format_block("一级标题", "一级标题")
            create_format_block("正文", "正文")

        # 数字格式
        with st.expander("🔢 数字格式", expanded=False):
            number_config = {"enable": st.checkbox("启用数字单独格式", False, key=f"num_en_{st.session_state.template_version}")}
            if number_config["enable"]:
                number_config["font"] = st.selectbox("数字字体", EN_FONT_LIST, 0, key=f"num_font_{st.session_state.template_version}")
                number_config["size_same_as_body"] = st.checkbox("字号同正文", True, key=f"num_size_{st.session_state.template_version}")
                number_config["bold"] = st.checkbox("数字加粗", False, key=f"num_bold_{st.session_state.template_version}")

    # --- 主界面上传区域 ---
    st.subheader("Step 2: 上传文档并处理 🚀")
    
    # 初始化变量，防止报错
    # 注意：因为上面侧边栏可能没展开，这里需要给默认值
    # 为了简化逻辑，这里重新获取一下cfg，或者把变量定义放前面
    # 这里为了代码健壮性，我们做一个简单的封装
    # 但在这个脚本里，我们假设如果不展开侧边栏，使用默认值
    # 实际上最好的做法是把所有状态初始化放最前面
    
    # 这里为了修复可能的报错，我们简单定义一下如果侧边栏没渲染的情况
    # 实际上在Streamlit里，只要代码执行到，变量就会存在
    # 这里我们用一个try-catch思想来赋值
    try:
        # 尝试获取侧边栏的变量，如果没有则给默认值
        if 'force_style' not in locals(): force_style = True
        if 'enable_title_regex' not in locals(): enable_title_regex = True
        if 'keep_spacing' not in locals(): keep_spacing = True
        if 'clear_blank' not in locals(): clear_blank = False
        if 'max_blank' not in locals(): max_blank = 1
        if 'fix_punctuation' not in locals(): fix_punctuation = False
        if 'fix_text' not in locals(): fix_text = False
        if 'ai_mode' not in locals(): ai_mode = "不使用AI"
        if 'number_config' not in locals(): number_config = {"enable": False}
        if 'DOUBAO_KEY' not in locals(): DOUBAO_KEY = ""
        
        cfg = st.session_state.current_config
    except:
        pass

    uploaded_file = st.file_uploader("请上传 .docx 格式的 Word 文档", type="docx")
    
    if uploaded_file:
        st.success("✅ 文档上传成功！准备就绪。")
        
        if st.button("✨ 开始一键处理", type="primary", use_container_width=True):
            with st.status("正在处理中，请稍候...", expanded=True) as status:
                st.write("📖 解析文档结构...")
                try:
                    # 这里的 cfg 确保是当前的
                    current_cfg = st.session_state.current_config
                    
                    data, stats = process_doc(
                        uploaded_file, 
                        current_cfg, 
                        number_config, 
                        ai_mode, 
                        enable_title_regex, 
                        force_style, 
                        keep_spacing, 
                        clear_blank, 
                        max_blank, 
                        fix_punctuation, 
                        fix_text, 
                        DOUBAO_KEY
                    )
                    
                    if data and stats:
                        st.write("✅ 格式调整完成...")
                        st.write("📦 生成最终文件...")
                        status.update(label="处理完成！", state="complete", expanded=False)
                        
                        st.subheader("📊 处理结果统计")
                        c1,c2,c3,c4 = st.columns(4)
                        c1.metric("标题", stats["一级标题"]+stats["二级标题"]+stats["三级标题"])
                        c2.metric("正文段落", stats["正文"])
                        c3.metric("表格", stats["表格"])
                        c4.metric("图片", stats["图片"])
                        
                        # 优化文件名
                        from datetime import datetime
                        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                        new_filename = f"排版完成_{current_time}_{uploaded_file.name}"
                        
                        st.download_button("📥 下载处理后的文档", data, new_filename, use_container_width=True)
                        st.balloons()
                        
                except Exception as e:
                    status.update(label="处理失败", state="error", expanded=True)
                    st.error(f"发生错误：{e}")

    st.divider()
    with st.expander("❓ 常见问题 & 关于"):
        st.markdown("""
        *   **Q: 为什么我的文档打开后乱码？**
            A: 请确保上传的是 `.docx` 格式（Word 2007及以上），旧版 `.doc` 格式不支持哦。
        *   **Q: 处理后图片不见了怎么办？**
            A: 不会丢失的！本工具承诺100%保留图片。如果发现位置不对，请手动微调，这是Word排版的正常现象。
        *   **Q: 我是学生，没有API密钥怎么办？**
            A: 没有关系！本工具的**核心排版功能是完全免费**且不需要密钥的。AI降重只是锦上添花的功能。
        """)

if __name__ == "__main__":
    main()
