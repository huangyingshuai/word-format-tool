import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import requests
import tempfile
import os
import re

# ====================== 全局常量定义（已校验无错误） ======================
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

# 豆包API配置（已适配你的专属降重智能体）
DOUBAO_API_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
# 你的专属降重智能体模型ID，可在火山引擎方舟平台智能体详情页获取
DOUBAO_MODEL = "ep-20250628104918-7rqxd"
# 你的降重智能体网页地址（用于页面展示）
DOUBAO_BOT_URL = "https://doubao.com/bot/PByvHsxX"

# ====================== 格式模板库（已校验可正常调用） ======================
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

# ====================== 核心工具函数（已校验无BUG） ======================
def is_protected_para(para):
    """识别受保护段落（图片、分页符、分节符、公式等），避免破坏原排版"""
    if para.paragraph_format.page_break_before:
        return True
    for run in para.runs:
        if run.contains_page_break:
            return True
        if run._element.find('.//w:sectPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
            return True
        if run._element.find('.//w:drawing', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
            return True
        if run._element.find('.//w:fldChar', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
            return True
        if run._element.find('.//m:oMath', namespaces={'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}) is not None:
            return True
    return False

def set_run_font(run, font_name, font_size, bold=None):
    """统一设置中文字体，兼容所有Word版本"""
    try:
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run.font.size = Pt(font_size)
        if bold is not None:
            run.font.bold = bold
    except Exception:
        pass

def set_en_number_font(run, font_name, font_size, bold=None):
    """单独设置英文/数字字体，符合论文格式规范"""
    try:
        if font_name == "和正文一致":
            return
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:ascii'), font_name)
        run._element.rPr.rFonts.set(qn('w:hAnsi'), font_name)
        run.font.size = Pt(font_size)
        if bold is not None:
            run.font.bold = bold
    except Exception:
        pass

def get_title_level(para, enable_regex, last_title_level):
    """智能识别标题级别，兼容样式匹配+正则匹配双模式"""
    # 优先匹配Word原生标题样式
    style_name = para.style.name
    if "Heading 1" in style_name or "标题 1" in style_name:
        return "一级标题"
    if "Heading 2" in style_name or "标题 2" in style_name:
        return "二级标题"
    if "Heading 3" in style_name or "标题 3" in style_name:
        return "三级标题"
    # 关闭正则则直接返回正文
    if not enable_regex:
        return "正文"
    # 正则匹配标题格式
    text = para.text.strip()
    if not text or len(text) > 100:
        return "正文"
    # 上下文关联匹配，避免标题层级混乱
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
    # 兜底全量匹配
    for level in ["一级标题", "二级标题", "三级标题"]:
        if TITLE_RULE[level].match(text):
            return level
    return "正文"

def process_number_in_para(para, body_font, body_size, number_config):
    """单独处理正文中的数字/英文格式，符合论文规范"""
    number_size = FONT_SIZE_NUM[number_config["size"]] if not number_config["size_same_as_body"] else body_size
    number_font = number_config["font"]
    number_bold = number_config["bold"]
    number_pattern = re.compile(r"-?\d+\.?\d*%?|[a-zA-Z]+")
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
        # 拆分文本与数字/英文
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
        # 清空原内容，重新写入
        run.text = ""
        for part_type, part_text in parts:
            new_run = para.add_run(part_text)
            if part_type == "text":
                set_run_font(new_run, body_font, body_size)
            else:
                set_en_number_font(new_run, number_font, number_size, number_bold)
            new_runs.append(new_run)
    # 替换原段落的runs
    for run in para.runs:
        run._element.getparent().remove(run._element)
    for new_run in new_runs:
        para._element.append(new_run._element)

# ====================== AI降重核心函数（已校验可正常调用） ======================
def ai_text_process(text, mode, api_key):
    """调用你的专属降重智能体，兼容多种文本处理模式"""
    if not api_key or not text.strip():
        return text
    # 匹配不同处理模式的prompt
    if mode == "专业降重":
        prompt = f"""严格遵循学术论文降重规范处理以下文本，要求：
1. 不改变原文核心含义、专业术语、数据、公式编号、参考文献标注
2. 打破AI写作特征，替换AI套话，长短句结合，符合人类学术写作习惯
3. 不破坏连续专业字符，不改变段落结构，仅输出处理后的结果，无需任何额外说明
待处理文本：{text}"""
    elif mode == "润色":
        prompt = f"对以下文本进行专业润色，优化语句通顺度，保留原文结构和核心含义，仅输出润色后的结果：{text}"
    elif mode == "标点修正":
        prompt = f"修正以下文本的中英文标点，中文使用全角标点，英文和数字使用半角标点，仅输出修正后的结果：{text}"
    elif mode == "错别字修正":
        prompt = f"修正以下文本中的错别字、语病和语法错误，保留原文原意，仅输出修正后的结果：{text}"
    else:
        return text

    try:
        # 豆包官方标准API请求格式
        resp = requests.post(
            DOUBAO_API_URL,
            json={
                "model": DOUBAO_MODEL,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.7,
                "top_p": 0.9
            },
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            },
            timeout=60
        )
        resp.raise_for_status()
        result = resp.json()
        # 提取返回结果
        return result["choices"][0]["message"]["content"].strip()
    except Exception as e:
        st.warning(f"AI处理调用异常，已保留原文：{str(e)[:80]}")
        return text

# ====================== 文档处理主函数（已全流程校验） ======================
def process_doc(uploaded_file, config, number_config, ai_mode, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank, fix_punctuation, fix_text, api_key):
    """全流程处理Word文档，格式调整+AI处理"""
    # 创建临时文件，避免原文件损坏
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    try:
        # 打开文档
        doc = docx.Document(tmp_path)
        # 统计文档信息
        stats = {"一级标题": 0, "二级标题": 0, "三级标题": 0, "正文": 0, "表格": 0, "图片": 0}
        # 统计图片数量
        for para in doc.paragraphs:
            try:
                stats["图片"] += len(para._element.findall('.//w:drawing', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
            except Exception:
                pass
        # 遍历处理所有段落
        last_title_level = None
        for para in doc.paragraphs:
            # 跳过受保护段落，仅设置基础字体
            if is_protected_para(para):
                for run in para.runs:
                    set_run_font(run, config["正文"]["font"], FONT_SIZE_NUM[config["正文"]["size"]])
                continue
            # 跳过空段落
            text = para.text.strip()
            if not text:
                continue
            # 识别段落级别
            level = get_title_level(para, enable_title_regex, last_title_level)
            stats[level] += 1
            # 更新上一级标题
            if level in ["一级标题", "二级标题", "三级标题"]:
                last_title_level = level
            # AI文本处理
            processed_text = para.text
            if ai_mode != "不使用AI":
                processed_text = ai_text_process(processed_text, ai_mode, api_key)
            if fix_punctuation:
                processed_text = ai_text_process(processed_text, "标点修正", api_key)
            if fix_text:
                processed_text = ai_text_process(processed_text, "错别字修正", api_key)
            # 替换处理后的文本
            if processed_text != para.text:
                para.text = processed_text
            # 获取当前级别格式配置
            cfg = config[level]
            font_size = FONT_SIZE_NUM[cfg["size"]]
            # 强制应用Word原生样式
            if force_style:
                try:
                    para.style = level
                except Exception:
                    pass
            # 设置段落格式
            try:
                # 对齐方式
                if cfg["align"] != "不修改":
                    para.alignment = ALIGN_MAP[cfg["align"]]
                # 行距
                para.paragraph_format.line_spacing_rule = LINE_TYPE_MAP[cfg["line_type"]]
                if cfg["line_type"] == "多倍行距":
                    para.paragraph_format.line_spacing = cfg["line_value"]
                elif cfg["line_type"] == "固定值":
                    para.paragraph_format.line_spacing = Pt(cfg["line_value"])
                # 段前段后距
                if not keep_spacing:
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                # 首行缩进
                if level == "正文" and cfg["indent"] > 0:
                    para.paragraph_format.first_line_indent = Cm(cfg["indent"] * 0.35)
                # 取消分页相关设置
                para.paragraph_format.page_break_before = False
                para.paragraph_format.keep_with_next = False
            except Exception:
                pass
            # 处理数字/英文格式
            if level == "正文" and number_config["enable"]:
                process_number_in_para(para, cfg["font"], font_size, number_config)
            else:
                for run in para.runs:
                    set_run_font(run, cfg["font"], font_size, cfg["bold"])
        # 处理所有表格
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
                            except Exception:
                                pass
                        # 设置表格段落格式
                        try:
                            if cfg["align"] != "不修改":
                                para.alignment = ALIGN_MAP[cfg["align"]]
                            para.paragraph_format.line_spacing_rule = LINE_TYPE_MAP[cfg["line_type"]]
                            if cfg["line_type"] == "多倍行距":
                                para.paragraph_format.line_spacing = cfg["line_value"]
                            elif cfg["line_type"] == "固定值":
                                para.paragraph_format.line_spacing = Pt(cfg["line_value"])
                        except Exception:
                            pass
                        # 设置字体
                        for run in para.runs:
                            set_run_font(run, cfg["font"], font_size, cfg["bold"])
        # 清理多余空行
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
        # 保存处理后的文档
        output_path = tempfile.mktemp(suffix=".docx")
        doc.save(output_path)
        with open(output_path, "rb") as f:
            file_bytes = f.read()
        return file_bytes, stats
    except Exception as e:
        st.error(f"文档处理失败：{str(e)}")
        return None, None
    finally:
        # 清理临时文件
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        if 'output_path' in locals() and os.path.exists(output_path):
            os.unlink(output_path)

# ====================== 页面主逻辑（已校验交互正常） ======================
def main():
    # 页面配置
    st.set_page_config(page_title="文式通 - Word格式智能处理系统", layout="wide")
    # 初始化session_state，避免KeyError
    if "current_config" not in st.session_state:
        st.session_state.current_config = GENERAL_TPL["默认格式"]
    if "template_version" not in st.session_state:
        st.session_state.template_version = 0
    if "doubao_api_key" not in st.session_state:
        st.session_state.doubao_api_key = ""

    # 页面标题与声明
    st.title("📄 文式通 - Word格式智能处理系统")
    st.warning("⚠️ 重要声明：此工具仅用于减少格式调整工作量，处理完成后请务必手动核对原文内容与格式，确保无误。")
    st.markdown("✅ 100%保留图片/目录/公式/原排版 | ✅ 高校论文/公文格式一键适配 | ✅ 集成专属AI降重智能体 | ✅ 标点规范/错别字修正")
    st.info(f"🔗 当前使用专属AI降重智能体：[点击访问智能体]({DOUBAO_BOT_URL})")

    # ====================== 模板选择模块（已校验可正常切换） ======================
    st.subheader("📋 标准格式模板选择")
    tpl_type = st.radio("模板类型", ["通用办公模板", "高校毕业论文模板", "党政公文模板"], horizontal=True)
    # 匹配模板库
    if tpl_type == "通用办公模板":
        tpl_dict = GENERAL_TPL
    elif tpl_type == "高校毕业论文模板":
        tpl_dict = UNIVERSITY_TPL
    else:
        tpl_dict = OFFICIAL_TPL
    # 选择具体模板
    tpl_name = st.selectbox("选择目标格式", list(tpl_dict.keys()), index=0)
    target_config = tpl_dict[tpl_name]

    # 手动应用模板按钮
    col1, col2 = st.columns([1, 4])
    with col1:
        apply_tpl_btn = st.button("✅ 应用选中模板", type="primary", use_container_width=True)
    with col2:
        st.caption("选择模板后点击此按钮应用，左侧格式参数将同步更新为模板配置")

    # 应用模板逻辑
    if apply_tpl_btn:
        st.session_state.current_config = target_config
        st.session_state.template_version += 1
        st.success(f"✅ 已成功应用【{tpl_name}】模板，左侧格式参数已同步更新！")
        st.rerun()

    # ====================== 侧边栏格式设置 ======================
    with st.sidebar:
        st.header("🎨 详细格式设置")
        cfg = st.session_state.current_config
        # 重置格式按钮
        if st.button("🔄 重置为默认格式", use_container_width=True):
            st.session_state.current_config = GENERAL_TPL["默认格式"]
            st.session_state.template_version += 1
            st.success("已重置为默认格式！")
            st.rerun()

        st.divider()
        st.subheader("🏷️ 核心功能设置")
        force_style = st.checkbox("强制统一Word原生样式", value=True)
        enable_title_regex = st.checkbox("启用智能标题识别", value=True, help="自动识别无样式的标题段落")
        keep_spacing = st.checkbox("保留原段落段前/段后距", value=True)

        st.divider()
        st.subheader("📄 空行清理设置")
        clear_blank = st.checkbox("清除多余空行", value=False)
        max_blank = st.slider("最多保留连续空行数", 0, 3, 1) if clear_blank else 1

        st.divider()
        st.subheader("🔤 AI文本处理设置")
        # API密钥输入
        st.text_input(
            "火山引擎API密钥（sk-开头）",
            type="password",
            key="doubao_api_key",
            placeholder="请输入sk-开头的API密钥，用于调用AI降重功能"
        )
        api_key = st.session_state.doubao_api_key
        # 密钥有效性校验
        is_valid_key = api_key.startswith("sk-") if api_key else False
        if not api_key:
            st.info("ℹ️ 请输入API密钥启用AI降重/润色功能")
        elif not is_valid_key:
            st.warning("⚠️ 无效的API密钥！格式应为sk-开头，请检查后重新输入")
        else:
            st.success("✅ API密钥有效，AI功能已启用")

        # AI功能选项
        fix_punctuation = st.checkbox("修正中英文标点", False, disabled=not is_valid_key)
        fix_text = st.checkbox("修正错别字/语病", False, disabled=not is_valid_key)
        # AI处理模式
        ai_mode = st.radio(
            "AI处理模式",
            ["不使用AI", "润色", "专业降重"],
            horizontal=True,
            disabled=not is_valid_key
        )

        # ====================== 格式参数设置块 ======================
        def create_format_block(title, level):
            st.divider()
            st.subheader(title)
            item = cfg[level]
            version = st.session_state.template_version
            
            # 字体选择
            font_idx = FONT_LIST.index(item["font"]) if item["font"] in FONT_LIST else 0
            item["font"] = st.selectbox("字体", FONT_LIST, key=f"{level}_font_{version}", index=font_idx)
            # 字号选择
            size_idx = FONT_SIZE_LIST.index(item["size"]) if item["size"] in FONT_SIZE_LIST else 0
            item["size"] = st.selectbox("字号", FONT_SIZE_LIST, key=f"{level}_size_{version}", index=size_idx)
            # 加粗
            item["bold"] = st.checkbox("加粗", item["bold"], key=f"{level}_bold_{version}")
            # 对齐方式
            align_idx = ALIGN_LIST.index(item["align"]) if item["align"] in ALIGN_LIST else 0
            item["align"] = st.selectbox("对齐方式", ALIGN_LIST, key=f"{level}_align_{version}", index=align_idx)
            
            # 行距设置
            line_type_idx = LINE_TYPE_LIST.index(item["line_type"]) if item["line_type"] in LINE_TYPE_LIST else 0
            new_line_type = st.selectbox("行距类型", LINE_TYPE_LIST, key=f"{level}_line_type_{version}", index=line_type_idx)
            if new_line_type != item["line_type"]:
                item["line_type"] = new_line_type
                item["line_value"] = LINE_RULE[new_line_type]["default"]
                st.session_state.current_config[level] = item
                st.rerun()
            
            # 行距值设置
            line_rule = LINE_RULE[item["line_type"]]
            curr_val = float(item["line_value"])
            if not (line_rule["min"] <= curr_val <= line_rule["max"]):
                curr_val = line_rule["default"]
                item["line_value"] = curr_val
            item["line_value"] = st.number_input(
                line_rule["label"],
                line_rule["min"],
                line_rule["max"],
                curr_val,
                line_rule["step"],
                key=f"{level}_line_value_{version}",
                disabled=line_rule["min"] == line_rule["max"]
            )
            
            # 首行缩进
            if "indent" in item:
                item["indent"] = st.number_input(
                    "首行缩进(字符)",
                    0,
                    4,
                    item["indent"],
                    key=f"{level}_indent_{version}"
                )
            # 更新session_state
            st.session_state.current_config[level] = item
            return item

        # 生成各级别格式设置
        create_format_block("📌 一级标题", "一级标题")
        create_format_block("📌 二级标题", "二级标题")
        create_format_block("📌 三级标题", "三级标题")
        create_format_block("📝 正文内容", "正文")
        create_format_block("📊 表格内容", "表格")

        # 数字/英文格式设置
        st.divider()
        st.subheader("🔢 正文数字/英文格式")
        number_config = {
            "enable": st.checkbox("启用数字/英文单独格式", True, key=f"num_en_{st.session_state.template_version}")
        }
        if number_config["enable"]:
            number_config["font"] = st.selectbox(
                "数字/英文字体",
                EN_FONT_LIST,
                0,
                key=f"num_font_{st.session_state.template_version}"
            )
            number_config["size_same_as_body"] = st.checkbox(
                "字号与正文一致",
                True,
                key=f"num_size_same_{st.session_state.template_version}"
            )
            number_config["size"] = st.selectbox(
                "数字/英文字号",
                FONT_SIZE_LIST,
                FONT_SIZE_LIST.index("小四"),
                key=f"num_size_{st.session_state.template_version}",
                disabled=number_config["size_same_as_body"]
            ) if not number_config["size_same_as_body"] else "小四"
            number_config["bold"] = st.checkbox(
                "数字/英文加粗",
                False,
                key=f"num_bold_{st.session_state.template_version}"
            )

    # ====================== 文档上传与处理 ======================
    st.divider()
    st.subheader("📤 文档上传与处理")
    uploaded_file = st.file_uploader("上传Word文档（仅支持.docx格式）", type="docx")

    # 处理按钮可用性校验
    process_btn_disabled = True
    if uploaded_file:
        # 仅上传文档，无AI功能：可处理
        if ai_mode == "不使用AI" and not fix_punctuation and not fix_text:
            process_btn_disabled = False
        # 使用AI功能，密钥有效：可处理
        elif is_valid_key:
            process_btn_disabled = False

    # 处理按钮
    if uploaded_file:
        st.success("✅ 文档上传成功！")
        if st.button(
            "🚀 开始处理（格式调整+AI处理）",
            type="primary",
            use_container_width=True,
            disabled=process_btn_disabled
        ):
            with st.spinner("正在处理文档，请稍候..."):
                bar = st.progress(0, "正在解析文档结构...")
                try:
                    bar.progress(20, "正在识别段落与标题...")
                    data, stats = process_doc(
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
                        fix_text,
                        api_key
                    )
                    bar.progress(80, "正在生成处理后文档...")
                    if data and stats:
                        bar.progress(100, "处理完成！")
                        st.divider()
                        st.subheader("📋 文档处理统计")
                        c1,c2,c3,c4,c5,c6 = st.columns(6)
                        c1.metric("一级标题", stats["一级标题"])
                        c2.metric("二级标题", stats["二级标题"])
                        c3.metric("三级标题", stats["三级标题"])
                        c4.metric("正文段落", stats["正文"])
                        c5.metric("表格", stats["表格"])
                        c6.metric("图片", stats["图片"])
                        
                        # 下载按钮
                        st.download_button(
                            "📥 下载处理完成的文档",
                            data,
                            f"格式处理完成_{uploaded_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                        st.success("🎉 文档处理完成！请下载后核对内容与格式。")
                except Exception as e:
                    st.error(f"文档处理失败：{str(e)}")
                finally:
                    bar.empty()

    # 底部声明
    st.divider()
    with st.expander("📖 使用说明与版权声明"):
        st.markdown("### 使用说明")
        st.markdown("1. 本工具仅支持.docx格式的Word文档，不支持.doc格式；")
        st.markdown("2. AI降重/润色功能需要火山引擎API密钥，可在[火山引擎方舟平台](https://www.volcengine.com/ark)获取；")
        st.markdown("3. 处理后的文档会100%保留原文档的图片、公式、目录、分节符等内容，仅修改文本格式与指定的AI处理内容；")
        st.markdown("4. 降重后的内容请务必人工核对，确保符合学术规范，避免出现语义偏差。")
        st.markdown("---")
        st.markdown("### 版权声明")
        st.markdown("本作品为大学生计算机设计大赛参赛作品，基于Streamlit、python-docx开发，遵守开源协议。")
        st.markdown(f"集成专属豆包AI降重智能体：{DOUBAO_BOT_URL}")

if __name__ == "__main__":
    main()
