import streamlit as st
from openai import OpenAI
from docxtpl import DocxTemplate
from docx.oxml.ns import qn
import io
import json
import re

# --- 辅助函数 ---
def count_chars(text):
    """统计汉字数和总字符数"""
    if not text: return 0, 0
    total = len(text)
    zh = len(re.findall(r'[\u4e00-\u9fa5]', text))
    return zh, total

def fix_chinese_punctuation(text):
    """强制将英文标点替换为中文全角标点，并处理成对的引号"""
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
    """转换为真实的 Word 段落 (Hard Returns)，强制中英文字体"""
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
    """通用的大模型调用函数"""
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
if "generated" not in st.session_state:
    st.session_state.generated = False
if "edit_title" not in st.session_state: st.session_state.edit_title = ""
if "edit_scenario" not in st.session_state: st.session_state.edit_scenario = ""
if "edit_problem" not in st.session_state: st.session_state.edit_problem = ""
if "edit_analysis" not in st.session_state: st.session_state.edit_analysis = ""

# --- 页面配置 ---
st.set_page_config(page_title="真实问题表格自动生成器", layout="wide", page_icon="📝")
st.title("📝 真实问题表格自动生成器")

# --- 侧边栏 ---
PRESET_MODELS = {
    "Qwen 2.5 72B (推荐)": "qwen/qwen-2.5-72b-instruct",
    "Claude 3.5 Sonnet": "anthropic/claude-3.5-sonnet",
    "Gemini 1.5 Pro": "google/gemini-1.5-pro",
    "DeepSeek V3": "deepseek/deepseek-chat",
    "⌨️ 手动输入其他模型": "__custom__",
}

with st.sidebar:
    st.header("⚙️ 配置")
    
    # 从 Secrets 中读取 API Key（安全做法）
    api_key = st.secrets["OPENROUTER_API_KEY"] if "OPENROUTER_API_KEY" in st.secrets else None
    if not api_key:
        api_key = st.text_input("OpenRouter API Key", type="password")
        if not api_key:
            st.warning("⚠️ 请输入 OpenRouter API Key（或在部署后台配置 Secrets）")
    else:
        st.success("✅ API Key 已安全载入")
    
    selected_label = st.selectbox("选择 AI 模型", list(PRESET_MODELS.keys()))
    if selected_label == "⌨️ 手动输入其他模型":
        model_id = st.text_input("请输入 OpenRouter 模型 ID", "qwen/qwen-2.5-72b-instruct")
    else:
        model_id = PRESET_MODELS[selected_label]
    st.caption(f"当前模型：`{model_id}`")

# --- 步骤 1：表单输入 ---
st.subheader("📌 步骤 1：填写详细信息与模板")
with st.container(border=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        student_name = st.text_input("姓名", "黄剑韬")
        college = st.text_input("学院", "哲学院")
        course_name = st.text_input("课程名称", "西方伦理思想史")
    with col2:
        student_id = st.text_input("学号", "20230608114")
        grade = st.text_input("年级", "2023级")
        course_id = st.text_input("课程号", "0321512")
    with col3:
        major = st.text_input("专业", "哲学")
        teacher = st.text_input("任课老师", "王宇明")
        semester = st.text_input("学年学期", "2025-2026学年第2学期")
        
    col4, col5 = st.columns(2)
    with col4:
        assessment_method = st.selectbox("考核方式", ["考查", "考试"])
    with col5:
        word_count = st.slider("期望[分析与对策]字数", min_value=500, max_value=5000, value=2000, step=100)
    
    topic_keywords = st.text_area("调研方向/关键词", "例如：大学生就业压力、校园外卖安全...")
    
    # 内置模板
    uploaded_file = "111.docx"
    st.info("📄 已自动加载内置模板：111.docx")

    if st.button("🚀 生成初稿", use_container_width=True):
        if not api_key: st.error("请输入 API Key！")
        elif not topic_keywords.strip(): st.error("请输入关键词！")
        else:
            with st.spinner("🧠 正在生成各板块初稿中，请耐心等待..."):
                try:
                    sys_prompt = "你是一个能够按要求生成规范化 JSON 的 AI 助手。必须严格遵守字数和中文全角标点要求。"
                    user_prompt = f"""
你是一名修读了《{course_name}》（任课老师：{teacher}）的{college}{major}专业{grade}优秀大学生。
现在你需要完成一份期末考核（{assessment_method}）的《真实问题表格》作业。
调研方向是：{topic_keywords}

请严格按要求生成，并以 JSON 格式输出四个键：title, scenario, problem, analysis。
1. title：不超过25字，以问号结束。
2. scenario：不超过200字。
3. problem：不超过100字。
4. analysis：【字数底线要求：必须达到或超过 {word_count} 字！】必须包含3个原因和3个对策，每一个原因和对策写满超长自然段，引经据典，深度剖析。千万不要只写提纲！
"""
                    result_json_str = call_llm(api_key, model_id, sys_prompt, user_prompt, is_json=True, max_tokens=8000)
                    result_json = json.loads(result_json_str)
                    
                    st.session_state.edit_title = fix_chinese_punctuation(result_json.get("title", ""))
                    st.session_state.edit_scenario = fix_chinese_punctuation(result_json.get("scenario", ""))
                    st.session_state.edit_problem = fix_chinese_punctuation(result_json.get("problem", ""))
                    st.session_state.edit_analysis = fix_chinese_punctuation(result_json.get("analysis", ""))
                    st.session_state.generated = True
                except Exception as e:
                    st.error(f"初稿生成失败：{e}")


# --- 步骤 2：审稿与编辑工作台 ---
if st.session_state.generated:
    st.divider()
    st.header("✍️ 步骤 2：审稿工作台")
    st.info("初稿已生成！你可以在下方的文本框中直接手动编辑修改内容。如果不满意，还可以使用下方的工具栏让 AI 单独重新处理某个板块。")

    def render_section(key_name, title, max_zh_words=None):
        st.subheader(title)
        zh, tot = count_chars(st.session_state[key_name])
        
        # 实时字数统计显示
        if max_zh_words:
            color = "green" if zh <= max_zh_words else "red"
            st.markdown(f"📝 实时字数：<span style='color:{color}; font-weight:bold;'>{zh} 汉字</span> / {tot} 字符 (要求: 不超 {max_zh_words} 汉字)", unsafe_allow_html=True)
        else:
            color = "green" if zh >= word_count else "orange"
            st.markdown(f"📝 实时字数：<span style='color:{color}; font-weight:bold;'>{zh} 汉字</span> / {tot} 字符 (目标: 约 {word_count} 汉字)", unsafe_allow_html=True)
            
        # 【核心修复点 1】：工具栏提到 text_area 之前渲染！
        # 这样在点击大模型按钮时，修改状态的动作发生在组件实例化之前，规避 Exception。
        c1, c2, c3, c4 = st.columns(4)
        
        if c1.button("🔄 重写此板块", key=f"regen_{key_name}"):
            with st.spinner("正在重新生成..."):
                sys_p = "你是一个优秀的大学生助手。请直接输出文本内容，不要包含任何多余的寒暄和解释。"
                if key_name == "edit_title":
                    usr_p = f"为《{course_name}》课程的调研报告重新生成一个题目。关键词：{topic_keywords}。要求：精准简练，不超过 25 个字，且以问号结束。纯文本输出。"
                elif key_name == "edit_scenario":
                    usr_p = f"为《{course_name}》的调研报告重新生成一段现实场景描述。关键词：{topic_keywords}。要求：用白描手法描述真实的现实场景，引出问题。不超过 200 字。纯文本输出。"
                elif key_name == "edit_problem":
                    usr_p = f"为《{course_name}》的调研报告重新生成真实问题描述。关键词：{topic_keywords}。要求：高度概括核心真实问题。不超过 100 字。纯文本输出。"
                elif key_name == "edit_analysis":
                    usr_p = f"为《{course_name}》的调研报告重新生成原因分析与对策建议。关键词：{topic_keywords}。要求：包含3个原因和3个对策。总字数必须大于 {word_count} 字！每一个点都要写成详实的超长自然段，引经据典，深度剖析。纯文本输出。"
                
                res = call_llm(api_key, model_id, sys_p, usr_p, is_json=False, max_tokens=8000)
                st.session_state[key_name] = fix_chinese_punctuation(res)
                # 不用 st.rerun()，因为马上就要用新状态渲染下方的 text_area 了
                
        # 将标点修复改为外链跳转
        c2.link_button("🔧 线上免费修复标点", "https://www.mimitool.com/symbol_conversion")
        
        # 经过了上方的逻辑，状态已经安全。此时实例化 text_area。
        st.text_area("内容编辑", key=key_name, height=150 if key_name != "edit_analysis" else 400, label_visibility="collapsed")
            
        st.write("---")

    with st.container(border=True):
        render_section("edit_title", "📌 1. 调研报告题目", 25)
        render_section("edit_scenario", "📌 2. 现实场景描述", 200)
        render_section("edit_problem", "📌 3. 真实问题描述", 100)
        render_section("edit_analysis", "📌 4. 原因分析与对策建议")


# --- 步骤 3：生成文档 ---
if st.session_state.generated:
    st.divider()
    st.subheader("🎉 步骤 3：排版无误，生成 Word 文档")
    st.markdown("如果您已经对上方工作台里的文本修改满意，请点击下方按钮将其填入 Word 模板中！")
    
    if st.button("📥 确认内容并生成文档", type="primary", use_container_width=True):
        try:
            doc = DocxTemplate(uploaded_file)
            
            scenario_subdoc = create_subdoc(doc, st.session_state.edit_scenario)
            problem_subdoc = create_subdoc(doc, st.session_state.edit_problem)
            analysis_subdoc = create_subdoc(doc, st.session_state.edit_analysis)

            context = {
                "name": student_name, "student_id": student_id, "college": college, "grade": grade,
                "major": major, "teacher": teacher, "course_name": course_name, "course_id": course_id,
                "semester": semester, "assessment_method": assessment_method,
                "title": st.session_state.edit_title,
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
                label="💾 点击下载最终版《真实问题表格》",
                data=bio,
                file_name=f"真实问题表格_{student_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"文档生成失败：{str(e)}")
