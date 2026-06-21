import streamlit as st
from openai import OpenAI
from docxtpl import DocxTemplate
from docx.oxml.ns import qn
import io
import json
import re

# --- 辅助函数 ---
def count_chars(text):
    if not text: return 0, 0
    total = len(text)
    zh = len(re.findall(r'[\u4e00-\u9fa5]', text))
    return zh, total

def fix_chinese_punctuation(text):
    if not text: return ""
    text = text.replace(',', '，').replace(':', '：').replace(';', '；').replace('?', '？').replace('!', '！').replace('(', '（').replace(')', '）')
    while '"' in text:
        if text.count('"') >= 2:
            text = text.replace('"', '“', 1)
            text = text.replace('"', '”', 1)
        else:
            text = text.replace('"', '”', 1)
    while "'" in text:
        if text.count("'") >= 2:
            text = text.replace("'", '‘', 1)
            text = text.replace("'", '’', 1)
        else:
            text = text.replace("'", '’', 1)
    return text

def create_subdoc(doc, text):
    subdoc = doc.new_subdoc()
    lines = text.split('\n')
    for line in lines:
        line_clean = line.strip()
        if line_clean:
            p = subdoc.add_paragraph()
            run = p.add_run(line_clean)
            run.font.name = 'Times New Roman'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    return subdoc

def call_llm(api_key, model_id, system_prompt, user_prompt, is_json=False, max_tokens=2000):
    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
    kwargs = {
        "model": model_id,
        "max_tokens": max_tokens,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
    }
    if is_json:
        kwargs["response_format"] = {"type": "json_object"}
        
    response = client.chat.completions.create(**kwargs)
    output = response.choices[0].message.content.strip()
    if is_json:
        if output.startswith("```json"): output = output[7:]
        if output.endswith("```"): output = output[:-3]
        output = output.strip()
    return output

# --- 初始化 Session State ---
if "assignments" not in st.session_state:
    st.session_state.assignments = []

# --- 页面配置 ---
st.set_page_config(page_title="真实问题表格自动生成器", layout="wide", page_icon="📝")
st.title("📝 真实问题表格自动生成器")

# --- 侧边栏 ---
PRESET_MODELS = {
    "Qwen 2.5 72B (推荐)": "qwen/qwen-2.5-72b-instruct",
    "DeepSeek V3": "deepseek/deepseek-chat",
    "Google Gemini 2.5 Flash Lite": "google/gemini-2.5-flash-lite-preview",
}

with st.sidebar:
    st.header("⚙️ 配置")
    
    # 从 Secrets 中读取 API Key（彻底不提供 UI 输入选项）
    api_key = st.secrets["OPENROUTER_API_KEY"] if "OPENROUTER_API_KEY" in st.secrets else None
    
    if not api_key:
        st.error("⚠️ 未检测到 API Key！请在 Streamlit Cloud 后台的 Settings -> Secrets 中配置 `OPENROUTER_API_KEY`。")
    else:
        st.success("✅ API Key 已安全载入")
    
    selected_label = st.selectbox("选择 AI 模型", list(PRESET_MODELS.keys()))
    model_id = PRESET_MODELS[selected_label]
    st.caption(f"当前模型：`{model_id}`")
    
    st.info("📄 已自动加载内置模板：111.docx")

# --- 步骤 1：全局个人信息 ---
st.subheader("👤 步骤 1：全局个人信息")
st.caption("填一次即可，下方所有的作业均会复用此信息生成文档。")
with st.container(border=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        student_name = st.text_input("姓名", "")
        college = st.text_input("学院", "")
    with col2:
        student_id = st.text_input("学号", "")
        grade = st.text_input("年级", "")
    with col3:
        major = st.text_input("专业", "")

st.divider()

# --- 步骤 2：多课程作业工作台 ---
st.subheader("📚 步骤 2：多课程作业工作台")
st.caption("你可以在这里添加无数门课程的作业。每门课程独立编辑，互不干扰！")

if st.button("➕ 添加新课程作业", type="primary"):
    st.session_state.assignments.append({
        "course_name": "",
        "course_id": "",
        "teacher": "Kuma",
        "semester": "2025-2026学年第2学期",
        "assessment_method": "考查",
        "word_count": 2000,
        "topic_keywords": "",
        "generated": False,
        "edit_title": "",
        "edit_scenario": "",
        "edit_problem": "",
        "edit_analysis": ""
    })

# 定义在作业内部渲染审稿区域的函数
def render_assignment_section(assignment, field_key, title, max_zh_words, idx):
    st.markdown(f"**{title}**")
    zh, tot = count_chars(assignment[field_key])
    
    if max_zh_words:
        color = "green" if zh <= max_zh_words else "red"
        st.markdown(f"📝 实时字数：<span style='color:{color}; font-weight:bold;'>{zh} 汉字</span> / {tot} 字符 (要求: 不超 {max_zh_words} 汉字)", unsafe_allow_html=True)
    else:
        target_words = assignment['word_count']
        color = "green" if zh >= target_words else "orange"
        st.markdown(f"📝 实时字数：<span style='color:{color}; font-weight:bold;'>{zh} 汉字</span> / {tot} 字符 (目标: 约 {target_words} 汉字)", unsafe_allow_html=True)
        
    c1, c2, c3, c4 = st.columns(4)
    
    if c1.button("🔄 重写此板块", key=f"regen_{field_key}_{idx}"):
        with st.spinner("正在重新生成..."):
            sys_p = "你是一个优秀的大学生助手。请直接输出文本内容，不要包含任何多余的寒暄和解释。"
            if field_key == "edit_title":
                usr_p = f"为《{assignment['course_name']}》课程的调研报告重新生成一个题目。关键词：{assignment['topic_keywords']}。要求：精准简练，不超过 25 个字，且以问号结束。纯文本输出。"
            elif field_key == "edit_scenario":
                usr_p = f"为《{assignment['course_name']}》的调研报告重新生成一段现实场景描述。关键词：{assignment['topic_keywords']}。要求：用白描手法描述真实的现实场景，引出问题。不超过 200 字。纯文本输出。"
            elif field_key == "edit_problem":
                usr_p = f"为《{assignment['course_name']}》的调研报告重新生成真实问题描述。关键词：{assignment['topic_keywords']}。要求：高度概括核心真实问题。不超过 100 字。纯文本输出。"
            elif field_key == "edit_analysis":
                usr_p = f"为《{assignment['course_name']}》的调研报告重新生成原因分析与对策建议。关键词：{assignment['topic_keywords']}。要求：包含3个原因和3个对策。总字数必须大于 {assignment['word_count']} 字！每一个点都要写成详实的超长自然段，引经据典，深度剖析。纯文本输出。"
            
            res = call_llm(api_key, model_id, sys_p, usr_p, is_json=False, max_tokens=8000)
            assignment[field_key] = fix_chinese_punctuation(res)
            st.rerun()
            
    c2.link_button("🔧 线上免费修复标点", "https://www.mimitool.com/symbol_conversion")
    
    assignment[field_key] = st.text_area("内容编辑", value=assignment[field_key], key=f"ta_{field_key}_{idx}", height=150 if field_key != "edit_analysis" else 400, label_visibility="collapsed")
    st.write("---")

for idx, assignment in enumerate(st.session_state.assignments):
    course_display_name = assignment["course_name"] if assignment["course_name"].strip() else f"未命名作业 {idx+1}"
    with st.expander(f"📝 {course_display_name}", expanded=True):
        col_c1, col_c2, col_c3 = st.columns(3)
        with col_c1:
            assignment["course_name"] = st.text_input("课程名称", value=assignment["course_name"], key=f"cname_{idx}")
            assignment["course_id"] = st.text_input("课程号", value=assignment["course_id"], key=f"cid_{idx}")
        with col_c2:
            assignment["teacher"] = st.text_input("任课老师", value=assignment["teacher"], key=f"teacher_{idx}")
            assignment["semester"] = st.text_input("学年学期", value=assignment["semester"], key=f"semester_{idx}")
        with col_c3:
            assignment["assessment_method"] = st.selectbox("考核方式", ["考查", "考试"], index=0 if assignment["assessment_method"] == "考查" else 1, key=f"assess_{idx}")
            assignment["word_count"] = st.slider("期望[分析与对策]字数", min_value=500, max_value=2000, value=assignment["word_count"], step=100, key=f"wc_{idx}")
            
        assignment["topic_keywords"] = st.text_area("调研方向/关键词", value=assignment["topic_keywords"], placeholder="例如：大学生就业压力、校园外卖安全...", key=f"topic_{idx}")
        
        if st.button("🚀 生成初稿", use_container_width=True, key=f"gen_{idx}"):
            if not api_key: st.error("请确保系统 Secrets 中已配置 API Key！")
            elif not assignment["topic_keywords"].strip(): st.error("请输入关键词！")
            elif not student_name.strip(): st.error("请在最上方填写姓名！")
            else:
                with st.spinner("🧠 正在生成各板块初稿中，请耐心等待..."):
                    try:
                        sys_prompt = "你是一个能够按要求生成规范化 JSON 的 AI 助手。必须严格遵守字数和中文全角标点要求。"
                        user_prompt = f"""
你是一名修读了《{assignment['course_name']}》（任课老师：{assignment['teacher']}）的{college}{major}专业{grade}优秀大学生。
现在你需要完成一份期末考核（{assignment['assessment_method']}）的《真实问题表格》作业。
调研方向是：{assignment['topic_keywords']}

请严格按要求生成，并以 JSON 格式输出四个键：title, scenario, problem, analysis。
1. title：不超过25字，以问号结束。
2. scenario：不超过200字。
3. problem：不超过100字。
4. analysis：【字数底线要求：必须达到或超过 {assignment['word_count']} 字！】必须包含3个原因和3个对策，每一个原因和对策写满超长自然段，引经据典，深度剖析。千万不要只写提纲！
"""
                        result_json_str = call_llm(api_key, model_id, sys_prompt, user_prompt, is_json=True, max_tokens=8000)
                        result_json = json.loads(result_json_str)
                        
                        assignment["edit_title"] = fix_chinese_punctuation(result_json.get("title", ""))
                        assignment["edit_scenario"] = fix_chinese_punctuation(result_json.get("scenario", ""))
                        assignment["edit_problem"] = fix_chinese_punctuation(result_json.get("problem", ""))
                        assignment["edit_analysis"] = fix_chinese_punctuation(result_json.get("analysis", ""))
                        assignment["generated"] = True
                        st.rerun()
                    except Exception as e:
                        st.error(f"初稿生成失败：{e}")
                        
        if assignment["generated"]:
            st.divider()
            st.markdown("#### ✍️ 审稿工作台")
            render_assignment_section(assignment, "edit_title", "📌 1. 调研报告题目", 25, idx)
            render_assignment_section(assignment, "edit_scenario", "📌 2. 现实场景描述", 200, idx)
            render_assignment_section(assignment, "edit_problem", "📌 3. 真实问题描述", 100, idx)
            render_assignment_section(assignment, "edit_analysis", "📌 4. 原因分析与对策建议", None, idx)
            
            st.markdown(f"如果您已经对该作业的内容修改满意，请点击下方按钮将其填入 Word 模板中！")
            if st.button("📥 确认内容并生成文档", type="primary", use_container_width=True, key=f"down_{idx}"):
                try:
                    # 读取内置模板
                    doc = DocxTemplate("111.docx")
                    
                    scenario_subdoc = create_subdoc(doc, assignment["edit_scenario"])
                    problem_subdoc = create_subdoc(doc, assignment["edit_problem"])
                    analysis_subdoc = create_subdoc(doc, assignment["edit_analysis"])

                    context = {
                        "name": student_name, "student_id": student_id, "college": college, "grade": grade,
                        "major": major, "teacher": assignment["teacher"], "course_name": assignment["course_name"], "course_id": assignment["course_id"],
                        "semester": assignment["semester"], "assessment_method": assignment["assessment_method"],
                        "title": assignment["edit_title"],
                        "scenario": scenario_subdoc,
                        "problem": problem_subdoc,
                        "analysis": analysis_subdoc
                    }
                    
                    doc.render(context)
                    bio = io.BytesIO()
                    doc.save(bio)
                    bio.seek(0)
                    
                    st.balloons()
                    st.download_button(
                        label=f"💾 下载《{assignment['course_name']}》最终版",
                        data=bio,
                        file_name=f"真实问题表格_{student_name}_{assignment['course_name']}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_btn_{idx}"
                    )
                except Exception as e:
                    st.error(f"文档生成失败：请确保您的 '111.docx' 模板在同一目录下，错误信息：{str(e)}")
