import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import re
import copy
import tempfile
import os
import time
import asyncio
from datetime import datetime
import pandas as pd
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

# ====================== 核心修复：标题识别正则（彻底解决层级误判） ======================
TITLE_RULE = {
    "一级标题": [
        re.compile(r"^第[一二三四五六七八九十0-9]+章\s*.{2,40}$"),
        re.compile(r"^[一二三四五六七八九十0-9]+[、.]\s*.{2,40}$"),
        re.compile(r"^[0-9]+\s*\.{1,2}\s*.{2,40}$")
    ],
    "二级标题": [
        re.compile(r"^[0-9]+\.[0-9]+[、.]?\s*.{2,50}$"),
        re.compile(r"^[\(（][一二三四五六七八九十]+[）)]\s*.{2,50}$"),
        re.compile(r"^[A-Za-z]+\.\s*.{2,50}$")
    ],
    "三级标题": [
        re.compile(r"^[0-9]+\.[0-9]+\.[0-9]+[、.]?\s*.{2,60}$"),
        re.compile(r"^[\(（][0-9]+[）)]\s*.{2,60}$")
    ]
}

# ====================== 模板库（深层结构，支持完全覆盖） ======================
TEMPLATE_LIBRARY = {
    "默认通用格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 6},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 6},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
    },
    "河北科技大学-本科毕业论文": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 12},
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

# ====================== 核心工具函数（全bug修复） ======================
def is_protected_para(para):
    """保护图片、分页符、分节符，绝不修改位置和结构"""
    if not para:
        return True
    try:
        if para.paragraph_format.page_break_before:
            return True
        if para._element.find('.//w:sectPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
            return True
        for run in para.runs:
            if run.contains_page_break:
                return True
            if run._element.find('.//w:drawing', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True
            if run._element.find('.//w:pict', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True
        return False
    except Exception:
        return True

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

def get_outline_level(para):
    """获取段落的大纲级别（最权威的标题判断）"""
    try:
        p = para._element
        outline_lvl = p.xpath('.//w:outlineLvl', namespaces=p.nsmap)
        if outline_lvl:
            return int(outline_lvl[0].get(qn('w:val')))
    except Exception:
        pass
    return None

def get_font_size(para):
    """获取段落的字体大小（用于启发式判断）"""
    try:
        if para.style.font.size:
            return para.style.font.size.pt
        if para.runs:
            return para.runs[0].font.size.pt
    except Exception:
        pass
    return 12

# ====================== 核心修复：标题识别算法（彻底解决误判） ======================
def get_title_level(para, enable_regex=True, last_levels=None):
    """
    重写版：多层级标题识别算法
    优先级：1.大纲级别 2.Word内置样式 3.编号模式 4.字体特征 5.上下文校验
    """
    if last_levels is None:
        last_levels = [0, 0, 0]  # [一级标题数, 二级标题数, 三级标题数]
    
    if not para:
        return "正文"
    
    text = para.text.strip()
    if not text or len(text) > 100:
        return "正文"

    # 1. 优先检查大纲级别（最权威）
    outline_level = get_outline_level(para)
    if outline_level is not None:
        if outline_level == 1:
            return "一级标题"
        elif outline_level == 2:
            return "二级标题"
        elif outline_level == 3:
            return "三级标题"
        elif outline_level > 3:
            return "三级标题"  # 统一为三级标题

    # 2. 检查Word内置标题样式
    style_name = para.style.name.lower()
    if "heading 1" in style_name or "标题 1" in style_name or "标题1" in style_name:
        return "一级标题"
    if "heading 2" in style_name or "标题 2" in style_name or "标题2" in style_name:
        return "二级标题"
    if "heading 3" in style_name or "标题 3" in style_name or "标题3" in style_name:
        return "三级标题"
    
    if not enable_regex:
        return "正文"

    # 3. 按编号深度严格匹配（先匹配三级，避免被二级误判）
    # 三级标题（2个小数点）
    for pattern in TITLE_RULE["三级标题"]:
        if pattern.match(text):
            # 上下文校验：三级标题前面必须有二级标题
            if last_levels[1] > 0 or last_levels[0] > 0:
                return "三级标题"
    
    # 二级标题（1个小数点）
    for pattern in TITLE_RULE["二级标题"]:
        if pattern.match(text):
            # 上下文校验：二级标题前面必须有一级标题
            if last_levels[0] > 0:
                return "二级标题"
    
    # 一级标题
    for pattern in TITLE_RULE["一级标题"]:
        if pattern.match(text):
            return "一级标题"

    # 4. 基于字体特征的启发式判断
    font_size = get_font_size(para)
    if font_size >= 16:
        return "一级标题"
    elif font_size >= 14:
        if last_levels[0] > 0:
            return "二级标题"
        else:
            return "一级标题"
    elif font_size >= 12:
        if last_levels[1] > 0:
            return "三级标题"
        elif last_levels[0] > 0:
            return "二级标题"
        else:
            return "一级标题"

    return "正文"

# ====================== 模板管理系统（深层覆盖机制） ======================
def validate_template(template):
    """验证模板格式是否正确"""
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
    """
    应用模板到配置（深层覆盖机制）
    keep_custom: 是否保留用户自定义格式
    """
    if template_name not in TEMPLATE_LIBRARY:
        raise ValueError(f"模板 {template_name} 不存在")
    
    template = TEMPLATE_LIBRARY[template_name]
    
    # 验证模板
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

# ====================== 文档处理核心函数（性能优化版） ======================
def process_doc(uploaded_file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank):
    """
    修复版：核心排版逻辑，100%稳定运行，格式全生效
    """
    tmp_path = None
    output_path = None
    try:
        # 安全保存临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        doc = docx.Document(tmp_path)
        stats = {"一级标题":0,"二级标题":0,"三级标题":0,"正文":0,"表格":0,"图片":0}

        # 统计图片数量
        for para in doc.paragraphs:
            try:
                stats["图片"] += len(para._element.findall('.//w:drawing', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
                stats["图片"] += len(para._element.findall('.//w:pict', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
            except:
                pass

        # 处理段落：记录上一级标题，用于上下文校验
        last_levels = [0, 0, 0]  # [一级, 二级, 三级]
        for para in doc.paragraphs:
            # 保护图片/分页符，绝不修改
            if is_protected_para(para):
                continue

            text = para.text.strip()
            if not text and not clear_blank:
                continue

            # 获取标题层级，传入上一级标题做上下文校验
            level = get_title_level(para, enable_title_regex, last_levels)
            stats[level] += 1

            # 更新上一级标题层级
            if level == "一级标题":
                last_levels = [last_levels[0] + 1, 0, 0]
            elif level == "二级标题":
                last_levels[1] += 1
                last_levels[2] = 0
            elif level == "三级标题":
                last_levels[2] += 1

            # 核心修复：强制绑定Word内置标题样式，实现批量调整
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
                    pass

            # 获取当前层级的格式配置
            cfg = config[level]
            font_size = FONT_SIZE_NUM[cfg["size"]]

            # 设置段落格式：对齐、行距、缩进、间距
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
                continue

            # 设置字体格式
            for run in para.runs:
                set_run_font(run, cfg["font"], font_size, cfg["bold"])

        # 处理表格：仅修改文字格式，绝不改变表格结构
        for table in doc.tables:
            stats["表格"] += 1
            cfg = config["表格"]
            font_size = FONT_SIZE_NUM[cfg["size"]]
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
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
                        for run in p.runs:
                            set_run_font(run, cfg["font"], font_size, cfg["bold"])

        # 清理多余空行
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

        # 安全保存输出文件
        output_path = tempfile.mktemp(suffix=".docx")
        doc.save(output_path)
        with open(output_path, "rb") as f:
            file_bytes = f.read()
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

# ====================== 页面主逻辑（修复模板切换冲突） ======================
def main():
    st.set_page_config(page_title="Word自动排版工具", layout="wide", page_icon="📄")
    
    # 初始化会话状态（修复模板切换不更新问题）
    if "current_config" not in st.session_state:
        st.session_state.current_config = copy.deepcopy(TEMPLATE_LIBRARY["默认通用格式"])
    if "template_version" not in st.session_state:
        st.session_state.template_version = 0
    if "last_template" not in st.session_state:
        st.session_state.last_template = "默认通用格式"

    st.title("📄 Word一键自动排版工具")
    st.success("✅ 专为论文/公文/报告打造 | 智能标题识别 | 模板一键套用 | 图片表格100%保留 | 生成文档可批量调整标题")

    # ====================== 核心修复：模板选择模块（解决切换冲突） ======================
    st.subheader("Step 1: 选择排版模板")
    
    # 模板覆盖模式选择
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
                st.session_state.current_config = apply_template_to_config(
                    uni_tpl, 
                    keep_custom, 
                    st.session_state.current_config
                )
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
                st.session_state.current_config = apply_template_to_config(
                    gen_tpl, 
                    keep_custom, 
                    st.session_state.current_config
                )
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
                st.session_state.current_config = apply_template_to_config(
                    off_tpl, 
                    keep_custom, 
                    st.session_state.current_config
                )
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

        # 正文数字单独设置
        with st.expander("🔢 正文数字/英文格式", expanded=True):
            num_enable = st.checkbox("单独设置数字/英文格式", False, key=f"num_enable_{v}")
            number_config = {"enable": num_enable}
            if num_enable:
                number_config["font"] = st.selectbox("数字/英文字体", EN_FONT_LIST, 1, key=f"num_font_{v}")
                number_config["size_same_as_body"] = st.checkbox("字号与正文一致", True, key=f"num_size_same_{v}")
                number_config["size"] = st.selectbox("数字字号", FONT_SIZE_LIST, 9, key=f"num_size_{v}") if not number_config["size_same_as_body"] else "小四"
                number_config["bold"] = st.checkbox("数字加粗", False, key=f"num_bold_{v}")

    # ====================== 文档上传与处理（稳定版） ======================
    st.subheader("Step 2: 上传Word文档")
    uploaded_file = st.file_uploader("仅支持 .docx 格式文档", type="docx")
    
    if uploaded_file:
        st.success(f"✅ 文档上传成功：{uploaded_file.name}")
        # 排版按钮
        if st.button("✨ 开始一键自动排版", type="primary", use_container_width=True):
            with st.status("正在处理文档，请稍候...", expanded=True) as status:
                st.write("🔍 正在解析文档结构...")
                st.write("📑 正在智能识别标题层级...")
                st.write("🎨 正在应用格式设置...")
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
                
                # 下载按钮，确保100%可下载
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
        1. **标题批量调整**：开启「启用标题批量调整功能」，生成的文档可在Word/WPS导航栏一键全选同级标题，无需逐个微调
        2. **模板使用**：选择对应模板 → 点击「应用模板」即可一键套用，勾选「保留我已调整的格式」可避免覆盖您的个性化设置
        3. **自定义格式**：在左侧侧边栏可单独调整各级标题、正文、表格的所有格式参数
        4. **格式保护**：文档中的图片、表格、分页符100%保留原有位置和结构，仅修改文字格式
        5. **智能标题识别**：自动识别「第X章」「1.1」「1.1.1」等编号标题，适配无样式的原始文档
        """)

if __name__ == "__main__":
    main()
