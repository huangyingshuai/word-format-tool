import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import tempfile
import os
import re
import random

# -------------------------- 全局配置 --------------------------
COZE_BOT_URL = "https://www.coze.cn/s/Dtw5_DzeCIo/"
st.set_page_config(page_title="全场景智能降重系统", layout="wide", page_icon="📄")

# 对齐&行距配置
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
    "单倍行距": {"default":1.0,"min":1.0,"max":1.0,"step":1.0,"label":"行距倍数"},
    "1.5倍行距":{"default":1.5,"min":1.5,"max":1.5,"step":0.1,"label":"行距倍数"},
    "2倍行距":{"default":2.0,"min":2.0,"max":2.0,"step":0.1,"label":"行距倍数"},
    "多倍行距":{"default":1.5,"min":0.5,"max":5.0,"step":0.1,"label":"行距倍数"},
    "固定值":{"default":20.0,"min":6.0,"max":100.0,"step":1.0,"label":"固定值(磅)"}
}

# 字体&字号配置
FONT_LIST = ["宋体","黑体","微软雅黑","楷体","仿宋"]
FONT_SIZE_LIST = ["初号","小初","一号","小一","二号","小二","三号","小三","四号","小四","五号","小五","六号","小六"]
FONT_SIZE_NUM = {k:v for k,v in zip(FONT_SIZE_LIST,
[42.0,36.0,26.0,24.0,22.0,18.0,16.0,15.0,14.0,12.0,10.5,9.0,7.5,6.5])}
EN_FONT_LIST = ["和正文一致","Times New Roman","Arial","Calibri"]

# 标题智能识别正则
TITLE_PATTERNS = {
    "一级标题": re.compile(r'^[一二三四五六七八九十]{1,}、|^第[一二三四五六七八九十1-9]+[章节篇部分]|^[1-9]\d*\.'),
    "二级标题": re.compile(r'^（[一二三四五六七八九十]{1,}）|^\([1-9]\d*\)|^[1-9]\d*\.[1-9]\d*\s'),
    "三级标题": re.compile(r'^[①-⑩]|^[1-9]\d*\.[1-9]\d*\.[1-9]\d*\s|^（[1-9]\d*）')
}

# 人类写作特征库
HUMAN_FEATURE = {
    "视角":["就实际来看","笔者认为","结合现实场景","在实践中","从日常经验来讲"],
    "感官":["心底","眼底","指尖","周身","耳畔","心头"],
    "套话替换":{
        "首先":"从落地场景看","其次":"从逻辑层面看","最后":"结合现实诉求",
        "一方面":"站在需求端","另一方面":"回到供给侧","综上所述":"结合前文分析"
    }
}

# -------------------------- 完整模板库（所有模板回归） --------------------------
# 1. 通用模板
GENERAL_TPL = {
    "一级标题":{"font":"黑体","size":"二号","bold":True,"align":"居中","line_type":"多倍行距","line_value":1.5,"indent":0},
    "二级标题":{"font":"黑体","size":"三号","bold":True,"align":"左对齐","line_type":"多倍行距","line_value":1.5,"indent":0},
    "三级标题":{"font":"黑体","size":"四号","bold":True,"align":"左对齐","line_type":"多倍行距","line_value":1.5,"indent":0},
    "正文":{"font":"宋体","size":"小四","bold":False,"align":"两端对齐","line_type":"多倍行距","line_value":1.5,"indent":2},
    "表格":{"font":"宋体","size":"五号","bold":False,"align":"居中","line_type":"单倍行距","line_value":1.0,"indent":0}
}

# 2. 河北科技大学毕业论文模板
HEBUST_TPL = {
    "一级标题":{"font":"黑体","size":"二号","bold":True,"align":"居中","line_type":"多倍行距","line_value":1.5,"indent":0},
    "二级标题":{"font":"黑体","size":"三号","bold":True,"align":"左对齐","line_type":"多倍行距","line_value":1.5,"indent":0},
    "三级标题":{"font":"黑体","size":"四号","bold":True,"align":"左对齐","line_type":"多倍行距","line_value":1.5,"indent":0},
    "正文":{"font":"宋体","size":"小四","bold":False,"align":"两端对齐","line_type":"多倍行距","line_value":1.5,"indent":2},
    "表格":{"font":"宋体","size":"五号","bold":False,"align":"居中","line_type":"单倍行距","line_value":1.0,"indent":0}
}

# 3. 河北工业大学毕业论文模板
HEBUT_TPL = {
    "一级标题":{"font":"黑体","size":"二号","bold":True,"align":"居中","line_type":"多倍行距","line_value":1.5,"indent":0},
    "二级标题":{"font":"黑体","size":"三号","bold":True,"align":"左对齐","line_type":"多倍行距","line_value":1.5,"indent":0},
    "三级标题":{"font":"楷体","size":"四号","bold":True,"align":"左对齐","line_type":"多倍行距","line_value":1.5,"indent":0},
    "正文":{"font":"宋体","size":"小四","bold":False,"align":"两端对齐","line_type":"多倍行距","line_value":1.5,"indent":2},
    "表格":{"font":"宋体","size":"五号","bold":False,"align":"居中","line_type":"单倍行距","line_value":1.0,"indent":0}
}

# 4. 燕山大学毕业论文模板
YSU_TPL = {
    "一级标题":{"font":"黑体","size":"二号","bold":True,"align":"居中","line_type":"固定值","line_value":20.0,"indent":0},
    "二级标题":{"font":"黑体","size":"三号","bold":True,"align":"左对齐","line_type":"多倍行距","line_value":1.5,"indent":0},
    "三级标题":{"font":"黑体","size":"四号","bold":True,"align":"左对齐","line_type":"多倍行距","line_value":1.5,"indent":0},
    "正文":{"font":"宋体","size":"小四","bold":False,"align":"两端对齐","line_type":"固定值","line_value":20.0,"indent":2},
    "表格":{"font":"宋体","size":"五号","bold":False,"align":"居中","line_type":"单倍行距","line_value":1.0,"indent":0}
}

# 5. 通用毕业论文模板
THESIS_TPL = {
    "一级标题":{"font":"黑体","size":"二号","bold":True,"align":"居中","line_type":"2倍行距","line_value":2.0,"indent":0},
    "二级标题":{"font":"黑体","size":"三号","bold":True,"align":"左对齐","line_type":"1.5倍行距","line_value":1.5,"indent":0},
    "三级标题":{"font":"楷体","size":"四号","bold":True,"align":"左对齐","line_type":"1.5倍行距","line_value":1.5,"indent":0},
    "正文":{"font":"宋体","size":"小四","bold":False,"align":"两端对齐","line_type":"1.5倍行距","line_value":1.5,"indent":2},
    "表格":{"font":"宋体","size":"五号","bold":False,"align":"居中","line_type":"单倍行距","line_value":1.0,"indent":0}
}

# 6. 国家党政机关公文国标（GB/T 7714-2012）模板
NATIONAL_OFFICIAL_TPL = {
    "一级标题":{"font":"黑体","size":"二号","bold":True,"align":"居中","line_type":"2倍行距","line_value":2.0,"indent":0},
    "二级标题":{"font":"楷体","size":"三号","bold":True,"align":"左对齐","line_type":"2倍行距","line_value":2.0,"indent":0},
    "三级标题":{"font":"仿宋","size":"三号","bold":True,"align":"左对齐","line_type":"2倍行距","line_value":2.0,"indent":0},
    "正文":{"font":"仿宋","size":"三号","bold":False,"align":"两端对齐","line_type":"2倍行距","line_value":2.0,"indent":2},
    "表格":{"font":"仿宋","size":"三号","bold":False,"align":"居中","line_type":"单倍行距","line_value":1.0,"indent":0}
}

# 模板映射（完整恢复所有模板）
TEMPLATE_MAP = {
    "通用模板": GENERAL_TPL,
    "河北科技大学毕业论文": HEBUST_TPL,
    "河北工业大学毕业论文": HEBUT_TPL,
    "燕山大学毕业论文": YSU_TPL,
    "通用毕业论文": THESIS_TPL,
    "国家党政机关公文（GB/T 7714-2012）": NATIONAL_OFFICIAL_TPL
}

# 初始化session_state
if "current_tpl" not in st.session_state:
    st.session_state.current_tpl = GENERAL_TPL
if "current_tpl_name" not in st.session_state:
    st.session_state.current_tpl_name = "通用模板"
if "template_version" not in st.session_state:
    st.session_state.template_version = 0

# -------------------------- 格式设置UI模块 --------------------------
def create_format_block(title, level):
    st.divider()
    st.subheader(title)
    item = st.session_state.current_tpl[level]
    v = st.session_state.template_version
    
    # 字体
    item["font"] = st.selectbox("字体", FONT_LIST, index=FONT_LIST.index(item["font"]), key=f"{level}_font_{v}")
    # 字号
    item["size"] = st.selectbox("字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(item["size"]), key=f"{level}_size_{v}")
    # 加粗
    item["bold"] = st.checkbox("加粗", item["bold"], key=f"{level}_bold_{v}")
    # 对齐
    item["align"] = st.selectbox("对齐方式", ALIGN_LIST, index=ALIGN_LIST.index(item["align"]), key=f"{level}_align_{v}")
    
    # 行距类型
    new_line = st.selectbox("行距类型", LINE_TYPE_LIST, index=LINE_TYPE_LIST.index(item["line_type"]), key=f"{level}_lt_{v}")
    if new_line != item["line_type"]:
        item["line_type"] = new_line
        item["line_value"] = LINE_RULE[new_line]["default"]
        st.session_state.current_tpl[level] = item
        st.rerun()
    
    # 行距值
    rule = LINE_RULE[item["line_type"]]
    item["line_value"] = st.number_input(
        rule["label"], rule["min"], rule["max"], float(item["line_value"]), 
        rule["step"], key=f"{level}_lv_{v}", disabled=rule["min"]==rule["max"]
    )
    
    # 首行缩进（仅正文）
    if "indent" in item and level == "正文":
        item["indent"] = st.number_input("首行缩进(字符)", 0,4,item["indent"],key=f"{level}_indent_{v}")
    
    st.session_state.current_tpl[level] = item
    return item

# -------------------------- 核心功能函数 --------------------------
# 智能识别标题/正文
def recognize_text_level(text):
    if not text.strip(): return "正文"
    check_text = text.strip()[:30]
    if TITLE_PATTERNS["一级标题"].match(check_text): return "一级标题"
    elif TITLE_PATTERNS["二级标题"].match(check_text): return "二级标题"
    elif TITLE_PATTERNS["三级标题"].match(check_text): return "三级标题"
    else: return "正文"

# 降重引擎
def rewrite_plagiarism(text):
    text = re.sub(r'\s+', ' ', text)
    sentences = re.split(r'[。！？；]', text)
    res = []
    for s in sentences:
        if not s.strip(): continue
        if len(s) > 13:
            parts = s.split('，')
            if len(parts)>=2: parts = parts[::-1]
            s = '，'.join(parts)
        res.append(s)
    return '。'.join(res) + '。'

# 人类写作特征注入
def inject_human_style(text):
    if len(text) <5: return text
    if random.random()>0.5: text = random.choice(HUMAN_FEATURE["视角"])+'，'+text
    for k,v in HUMAN_FEATURE["套话替换"].items(): text = text.replace(k,v)
    return text

# 字体设置
def set_font(run, font_name, size_pt, bold=None):
    try:
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run.font.size = Pt(size_pt)
        if bold is not None: run.font.bold = bold
    except: pass

# 数字/英文字体处理
def handle_number_font(para, body_font, body_size, num_cfg):
    pattern = re.compile(r'-?\d+\.?\d*%?|[a-zA-Z]+')
    new_runs = []
    for run in para.runs:
        txt = run.text
        if not txt:
            new_runs.append(run)
            continue
        if not pattern.search(txt):
            set_font(run, body_font, body_size)
            new_runs.append(run)
            continue
        parts, last_pos = [], 0
        for match in pattern.finditer(txt):
            s,e = match.span()
            if s>last_pos: parts.append(("text", txt[last_pos:s]))
            parts.append(("num", txt[s:e]))
            last_pos = e
        if last_pos < len(txt): parts.append(("text", txt[last_pos:]))
        run.text = ""
        for typ, content in parts:
            new_run = para.add_run(content)
            if typ == "text": set_font(new_run, body_font, body_size)
            else:
                n_size = FONT_SIZE_NUM[num_cfg["size"]] if not num_cfg["size_same_as_body"] else body_size
                set_font(new_run, num_cfg["font"], n_size, num_cfg["bold"])
            new_runs.append(new_run)
    for r in para.runs: r._element.getparent().remove(r._element)
    for r in new_runs: para._element.append(r._element)

# 文档处理主函数
def process_doc(uploaded_file, tpl_cfg, num_cfg, enable_rewrite, enable_human, force_style, keep_space, clear_blank, max_blank):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        doc = docx.Document(tmp_path)
        stats = {"一级标题":0,"二级标题":0,"三级标题":0,"正文":0,"表格":0,"图片":0}

        # 处理段落
        for para in doc.paragraphs:
            txt = para.text.strip()
            if not txt: continue
            level = recognize_text_level(txt)
            stats[level] += 1

            # 降重&人类风格
            new_txt = para.text
            if enable_rewrite: new_txt = rewrite_plagiarism(new_txt)
            if enable_human: new_txt = inject_human_style(new_txt)
            if new_txt != para.text: para.text = new_txt

            # 应用格式
            cfg = tpl_cfg[level]
            pt_size = FONT_SIZE_NUM[cfg["size"]]
            if force_style:
                try: para.style = level
                except: pass

            # 段落格式
            try:
                if cfg["align"] != "不修改": para.alignment = ALIGN_MAP[cfg["align"]]
                para.paragraph_format.line_spacing_rule = LINE_TYPE_MAP[cfg["line_type"]]
                if cfg["line_type"] == "多倍行距": para.paragraph_format.line_spacing = cfg["line_value"]
                elif cfg["line_type"] == "固定值": para.paragraph_format.line_spacing = Pt(cfg["line_value"])
                if not keep_space:
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                if level == "正文" and cfg["indent"] >0:
                    para.paragraph_format.first_line_indent = Cm(cfg["indent"] * 0.35)
            except: pass

            # 数字格式
            if level == "正文" and num_cfg["enable"]:
                handle_number_font(para, cfg["font"], pt_size, num_cfg)
            else:
                for run in para.runs: set_font(run, cfg["font"], pt_size, cfg["bold"])

        # 处理表格
        for table in doc.tables:
            stats["表格"] +=1
            tb_cfg = tpl_cfg["表格"]
            tb_size = FONT_SIZE_NUM[tb_cfg["size"]]
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for run in p.runs: set_font(run, tb_cfg["font"], tb_size, tb_cfg["bold"])

        # 清理空行
        if clear_blank:
            blank_count =0
            for p in reversed(list(doc.paragraphs)):
                if not p.text.strip():
                    blank_count +=1
                    if blank_count>max_blank: p._element.getparent().remove(p._element)
                else: blank_count =0

        # 保存输出
        out_path = tempfile.mktemp(suffix=".docx")
        doc.save(out_path)
        with open(out_path, "rb") as f: result_bytes = f.read()

        os.unlink(tmp_path)
        os.unlink(out_path)
        return result_bytes, stats

    except Exception as e:
        st.error(f"处理失败：{str(e)}")
        return None, stats

# -------------------------- 页面UI（模板选择完整恢复） --------------------------
def main():
    st.title("📄 全场景智能降重与格式排版系统")
    st.markdown(f"### 🔗 专属AI降重智能体：[点击跳转使用]({COZE_BOT_URL})")
    st.success("✅ 智能识别标题正文 | ✅ 多模板选择 | ✅ 格式自定义 | ✅ GitHub云端运行")

    # ===================== 模板选择功能（完整恢复所有模板） =====================
    st.subheader("📋 格式模板选择（所有模板回归！）")
    # 模板选择radio按钮，显示所有模板
    selected_tpl = st.radio(
        "选择模板类型", 
        list(TEMPLATE_MAP.keys()), 
        index=list(TEMPLATE_MAP.keys()).index(st.session_state.current_tpl_name),
        horizontal=True,
        key="tpl_selector"
    )
    
    # 模板切换逻辑（自动更新格式参数）
    if selected_tpl != st.session_state.current_tpl_name:
        st.session_state.current_tpl_name = selected_tpl
        st.session_state.current_tpl = TEMPLATE_MAP[selected_tpl]
        st.session_state.template_version += 1
        st.rerun()
    
    # 应用默认模板按钮
    col1, col2 = st.columns([1,4])
    with col1:
        if st.button("✅ 应用默认模板", type="primary"):
            st.session_state.current_tpl = GENERAL_TPL
            st.session_state.current_tpl_name = "通用模板"
            st.session_state.template_version +=1
            st.rerun()
    with col2: st.caption(f"当前模板：{st.session_state.current_tpl_name} | 切换后格式参数自动更新")

    # 侧边栏：完整格式设置面板
    with st.sidebar:
        st.header("⚙️ 核心功能开关")
        enable_rewrite = st.checkbox("启用【全场景降重方法论】", True)
        enable_human = st.checkbox("启用【人类写作深度特征】", True)
        st.divider()

        st.subheader("基础格式设置")
        force_style = st.checkbox("强制统一样式", True)
        keep_space = st.checkbox("保留段间距", True)
        clear_blank = st.checkbox("清除多余空行", False)
        max_blank = st.slider("最大连续空行",0,3,1) if clear_blank else 1
        st.divider()

        # 完整格式设置面板
        create_format_block("一级标题", "一级标题")
        create_format_block("二级标题", "二级标题")
        create_format_block("三级标题", "三级标题")
        create_format_block("正文内容", "正文")
        create_format_block("表格内容", "表格")

        st.divider()
        st.subheader("数字/英文格式")
        num_cfg = {"enable": st.checkbox("启用数字单独格式", True)}
        if num_cfg["enable"]:
            num_cfg["font"] = st.selectbox("数字字体", EN_FONT_LIST)
            num_cfg["size_same_as_body"] = st.checkbox("字号同正文", True)
            num_cfg["size"] = st.selectbox("数字字号", FONT_SIZE_LIST, index=9) if not num_cfg["size_same_as_body"] else "小四"
            num_cfg["bold"] = st.checkbox("数字加粗", False)

    # 上传&处理
    st.divider()
    uploaded = st.file_uploader("📤 上传 .docx 文档", type="docx")
    if uploaded:
        st.success("✅ 文档上传成功")
        if st.button("🚀 开始处理（识别+降重+排版）", type="primary", use_container_width=True):
            with st.spinner("正在智能处理..."):
                res_data, stats = process_doc(
                    uploaded, st.session_state.current_tpl, num_cfg,
                    enable_rewrite, enable_human, force_style, keep_space, clear_blank, max_blank
                )
                if res_data:
                    st.subheader("📊 识别&处理统计")
                    c1,c2,c3,c4,c5,c6 = st.columns(6)
                    c1.metric("一级标题", stats["一级标题"])
                    c2.metric("二级标题", stats["二级标题"])
                    c3.metric("三级标题", stats["三级标题"])
                    c4.metric("正文", stats["正文"])
                    c5.metric("表格", stats["表格"])
                    c6.metric("图片", stats["图片"])

                    st.download_button("📥 下载最终文档", res_data, f"降重排版_{uploaded.name}", use_container_width=True)
                    st.success("🎉 处理完成！格式&降重均生效")

    # 规则说明
    with st.expander("📖 功能说明"):
        st.markdown("""
        1. **模板选择**：通用/3所高校/通用毕业论文/国家公文，共6套模板可切换
        2. **智能识别**：自动拆分标题+正文，长文本可拆分
        3. **格式自定义**：字体/字号/行距/缩进/对齐全可调整
        4. **降重逻辑**：打破连续字符+重构语义，规避查重
        5. **人类风格**：去除AI套话，注入真实写作特征
        6. **云端兼容**：GitHub/Streamlit Cloud可直接部署
        """)

if __name__ == "__main__":
    main()
