# app.py - Complete AI Exam Prep with Cerebras API (All Features + .env)
import streamlit as st
import os
import glob
import time
import json
import base64
import hashlib
import io
from datetime import datetime
from collections import defaultdict
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
import requests
import PyPDF2
import docx2txt
from fpdf import FPDF
import pandas as pd
from PIL import Image
from dotenv import load_dotenv

# ------------------ LOAD ENVIRONMENT VARIABLES ------------------
load_dotenv()

# ------------------ CEREBRAS CONFIGURATION ------------------
CEREBRAS_API_KEY = os.getenv("CEREBRAS_API_KEY")
if not CEREBRAS_API_KEY:
    st.error("❌ API key not found! Please add CEREBRAS_API_KEY to .env file")
    st.stop()

CEREBRAS_URL = "https://api.cerebras.ai/v1/chat/completions"
MAX_RETRIES = 3
CACHE_TTL = 3600

# ------------------ API CALL FUNCTION ------------------
@retry(stop=stop_after_attempt(MAX_RETRIES), wait=wait_exponential(multiplier=1, min=2, max=15))
def call_cerebras(prompt, parse_json=False):
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {CEREBRAS_API_KEY}"
    }
    payload = {
        "model": "llama-3.1-8b",
        "messages": [
            {"role": "system", "content": "You are an AI exam preparation assistant. Respond in clean format."},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.3,
        "max_tokens": 4096
    }
    try:
        resp = requests.post(CEREBRAS_URL, headers=headers, json=payload, timeout=60)
        if resp.status_code == 200:
            data = resp.json()
            content = data["choices"][0]["message"]["content"]
            if parse_json:
                if '```json' in content:
                    content = content.split('```json')[1].split('```')[0]
                elif '```' in content:
                    content = content.split('```')[1].split('```')[0]
                return json.loads(content.strip())
            return content
        else:
            st.warning(f"API call failed: {resp.status_code}. Retrying...")
            raise Exception(f"API error: {resp.status_code}")
    except Exception as e:
        st.warning(f"API call failed: {str(e)[:100]}. Retrying...")
        time.sleep(2)
        raise

# ------------------ LANGUAGE SUPPORT ------------------
UI_TEXT = {
    "hindi": {
        "app_title": "📚 AI परीक्षा तैयारी प्रो",
        "bilingual_mode": "🔊 द्विभाषी मोड (हिंदी + अंग्रेज़ी)",
        "pre_loaded_notes": "📖 प्री-लोडेड नोट्स",
        "upload_notes": "📤 अपने नोट्स अपलोड करें",
        "pre_loaded_pyqs": "📚 प्री-लोडेड PYQs",
        "pyq_analysis": "📊 PYQ विश्लेषण",
        "subject_pyqs": "📋 विषय-वार PYQs",
        "explain": "🤔 किसी भी विषय को समझाएं",
        "quiz": "📝 क्विज़ मोड",
        "mock_test": "⏱️ मॉक टेस्ट",
        "adaptive": "🧠 एडाप्टिव लर्निंग",
        "analytics": "📈 डिटेल्ड एनालिटिक्स",
        "doubt": "💬 स्मार्ट डाउट रिज़ॉल्यूशन",
        "generate_notes": "📝 शॉर्ट नोट्स बनाएं",
        "generate_mcqs": "🎯 MCQs बनाएं",
        "topic_filter": "विषय फ़िल्टर",
        "num_mcqs": "MCQs की संख्या",
        "num_questions": "प्रश्नों की संख्या",
        "exam_name": "परीक्षा का नाम",
        "analyze_pyqs": "PYQs विश्लेषण करें",
        "extract_subject_pyqs": "विषय-वार PYQs निकालें",
        "explain_topic": "कोई भी टॉपिक लिखें",
        "explain_btn": "समझाएं",
        "quiz_source": "MCQ स्रोत चुनें",
        "mock_source": "मॉक टेस्ट स्रोत",
        "start_mock": "मॉक टेस्ट शुरू करें",
        "time_minutes": "समय (मिनट)",
        "weak_areas": "कमज़ोर विषय",
        "study_plan": "व्यक्तिगत अध्ययन योजना",
        "ask_doubt": "अपना प्रश्न लिखें",
        "resolve_btn": "समाधान पाएं"
    },
    "english": {
        "app_title": "📚 AI Exam Prep Pro",
        "bilingual_mode": "🔊 Bilingual Mode (Hindi + English)",
        "pre_loaded_notes": "📖 Pre-loaded Notes",
        "upload_notes": "📤 Upload Your Notes",
        "pre_loaded_pyqs": "📚 Pre-loaded PYQs",
        "pyq_analysis": "📊 PYQ Analysis",
        "subject_pyqs": "📋 Subject-wise PYQs",
        "explain": "🤔 Explain Any Topic",
        "quiz": "📝 Quiz Mode",
        "mock_test": "⏱️ Mock Test",
        "adaptive": "🧠 Adaptive Learning",
        "analytics": "📈 Detailed Analytics",
        "doubt": "💬 Smart Doubt Resolution",
        "generate_notes": "📝 Generate Short Notes",
        "generate_mcqs": "🎯 Generate MCQs",
        "topic_filter": "Topic filter",
        "num_mcqs": "Number of MCQs",
        "num_questions": "Number of questions",
        "exam_name": "Exam name",
        "analyze_pyqs": "Analyze PYQs",
        "extract_subject_pyqs": "Extract Subject-wise PYQs",
        "explain_topic": "Enter any topic",
        "explain_btn": "Explain",
        "quiz_source": "Select MCQ source",
        "mock_source": "Select mock test source",
        "start_mock": "Start Mock Test",
        "time_minutes": "Time (minutes)",
        "weak_areas": "Weak Areas",
        "study_plan": "Personalized Study Plan",
        "ask_doubt": "Type your doubt",
        "resolve_btn": "Get Solution"
    }
}

# ------------------ PROMPT BUILDERS ------------------
def get_prompt_lang(bilingual, hindi_full):
    if hindi_full:
        return "purely in Hindi (Devanagari script)"
    elif bilingual:
        return "in Hinglish (mix of Hindi and English)"
    else:
        return "in English only"

def build_notes_prompt(text, topic, bilingual, hindi_full):
    lang = get_prompt_lang(bilingual, hindi_full)
    inst = f"Focus on topic: {topic}." if topic else "Cover all important topics."
    return f"Generate concise exam-oriented notes (500-800 words, bullet points) {lang}. {inst}\n\nMaterial:\n{text[:15000]}"

def build_mcq_prompt(text, num, topic, bilingual, hindi_full):
    lang = get_prompt_lang(bilingual, hindi_full)
    inst = f"Generate ONLY from topic: {topic}." if topic else "Cover all subjects."
    return f"Generate {num} MCQs as JSON array. Each: {{'question': '...', 'options': ['A','B','C','D'], 'correct': 0, 'explanation': '...'}} {lang}. {inst}\n\nMaterial:\n{text[:15000]}"

def build_questions_prompt(text, num, topic, bilingual, hindi_full):
    lang = get_prompt_lang(bilingual, hindi_full)
    inst = f"Focus on {topic}." if topic else "Cover all subjects."
    return f"Generate {num} exam-oriented questions as JSON array. Each: {{'type':'Short Answer','question':'...','model_answer':'...','marks':5,'difficulty':'Medium'}} {lang}. {inst}\n\nMaterial:\n{text[:15000]}"

def build_pyq_prompt(syllabus, pyq_text, exam, bilingual, hindi_full):
    lang = get_prompt_lang(bilingual, hindi_full)
    return f"Analyze exam pattern for {exam}. Give topic weightage (%), trends, predicted questions. {lang}\n\nSyllabus:\n{syllabus[:10000]}\n\nPYQs:\n{pyq_text[:10000]}"

def build_explain_prompt(topic, bilingual, hindi_full):
    lang = get_prompt_lang(bilingual, hindi_full)
    return f"Explain the following topic in simple language with examples. {lang}\n\nTopic: {topic}"

def build_subject_extract_prompt(pyq_text, bilingual, hindi_full):
    lang = get_prompt_lang(bilingual, hindi_full)
    return f"""Extract all questions from the PYQ document and classify by subject. Subjects: History, Polity, Geography, Economy, Science, Environment, Mathematics, Logical Reasoning, English Language, Computer Knowledge, Hindi Language, Current Affairs, Others. Return JSON. {lang}\n\nPYQs:\n{pyq_text[:20000]}"""

def build_doubt_prompt(question, context_text, bilingual, hindi_full):
    lang = get_prompt_lang(bilingual, hindi_full)
    return f"Answer the doubt/question in detail. {lang}\n\nContext:\n{context_text[:5000]}\n\nQuestion: {question}"

def build_study_plan_prompt(weak_topics, exam_name, days_available, bilingual, hindi_full):
    lang = get_prompt_lang(bilingual, hindi_full)
    return f"Create a personalized study plan for {exam_name}. Weak topics: {weak_topics}. Days: {days_available}. {lang}"

# ------------------ TEXT EXTRACTION ------------------
def extract_pdf_text(file_bytes):
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        return text
    except:
        return ""

def extract_docx_text(file_bytes):
    try:
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        text = docx2txt.process(tmp_path)
        os.unlink(tmp_path)
        return text
    except:
        return ""

def extract_txt_text(file_bytes):
    try:
        return file_bytes.decode("utf-8")
    except:
        return ""

def get_text_from_files(uploaded_files):
    text = ""
    for file in uploaded_files:
        ext = file.name.split('.')[-1].lower()
        file_bytes = file.getvalue()
        if ext == 'pdf':
            text += extract_pdf_text(file_bytes) + "\n\n"
        elif ext == 'docx':
            text += extract_docx_text(file_bytes) + "\n\n"
        elif ext == 'txt':
            text += extract_txt_text(file_bytes) + "\n\n"
    return text.strip()

# ------------------ PDF GENERATION ------------------
def create_pdf(content, title):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, txt=title, ln=1, align='C')
    pdf.ln(10)
    pdf.set_font("Arial", size=11)
    for line in content.split('\n'):
        line = line.replace('**', '').replace('*', '')
        try:
            pdf.multi_cell(0, 6, txt=line)
        except:
            safe_line = line.encode('latin-1', 'replace').decode('latin-1')
            pdf.multi_cell(0, 6, txt=safe_line)
    return pdf.output(dest='S').encode('latin1')

def get_download_link(data, filename, label):
    b64 = base64.b64encode(data).decode()
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}" style="background:#4CAF50; color:white; padding:0.3rem 0.8rem; border-radius:5px; text-decoration:none;">📥 {label}</a>'

# ------------------ ANALYTICS CLASS ------------------
class StudentAnalytics:
    def __init__(self):
        if 'quiz_history' not in st.session_state:
            st.session_state.quiz_history = []
        if 'weak_topics' not in st.session_state:
            st.session_state.weak_topics = defaultdict(int)
    def add_quiz_result(self, score, total, topic):
        st.session_state.quiz_history.append({"score": score, "total": total, "topic": topic, "timestamp": datetime.now()})
        if total - score > total/2:
            st.session_state.weak_topics[topic] += 1
    def get_weak_topics(self):
        return dict(st.session_state.weak_topics)
    def get_overall_accuracy(self):
        if not st.session_state.quiz_history:
            return 0
        total_correct = sum(q["score"] for q in st.session_state.quiz_history)
        total_questions = sum(q["total"] for q in st.session_state.quiz_history)
        return (total_correct / total_questions * 100) if total_questions > 0 else 0
    def get_topic_performance(self):
        topic_stats = defaultdict(lambda: {"correct": 0, "total": 0})
        for q in st.session_state.quiz_history:
            topic_stats[q["topic"]]["correct"] += q["score"]
            topic_stats[q["topic"]]["total"] += q["total"]
        return topic_stats

# ------------------ STREAMLIT UI ------------------
st.set_page_config(page_title="AI Exam Prep Pro", page_icon="🎓", layout="wide")

lang_choice = st.radio("🌐 Language / भाषा", ["English", "हिंदी (Hindi)"], horizontal=True)
is_hindi = (lang_choice == "हिंदी (Hindi)")
text = UI_TEXT["hindi"] if is_hindi else UI_TEXT["english"]

bilingual = st.checkbox(text["bilingual_mode"], value=True)

st.markdown(f'<h1 style="text-align:center; color:#1E88E5;">{text["app_title"]}</h1>', unsafe_allow_html=True)

analytics = StudentAnalytics()

tabs = st.tabs([
    text["pre_loaded_notes"], text["upload_notes"], text["pre_loaded_pyqs"],
    text["pyq_analysis"], text["subject_pyqs"], text["explain"],
    text["quiz"], text["mock_test"], "✏️ Short Answer", "📝 Long Answer",
    text["adaptive"], text["analytics"], text["doubt"]
])

# ------------------ TAB 1: PRE-LOADED NOTES ------------------
with tabs[0]:
    st.subheader(text["pre_loaded_notes"])
    preload_dir = "preloaded_notes"
    if not os.path.exists(preload_dir):
        os.makedirs(preload_dir, exist_ok=True)
        st.warning(f"Folder '{preload_dir}' created. Add your PDF notes.")
    pre_files = []
    for ext in ['*.pdf', '*.docx', '*.txt']:
        pre_files.extend(glob.glob(os.path.join(preload_dir, ext)))
    if pre_files:
        file_names = [os.path.basename(f) for f in pre_files]
        selected_file = st.selectbox("Select a note file", file_names, key="pre_note")
        if selected_file:
            with open(os.path.join(preload_dir, selected_file), "rb") as f:
                fake_file = io.BytesIO(f.read())
                fake_file.name = selected_file
                pre_text = get_text_from_files([fake_file])
            st.success(f"✅ Loaded {len(pre_text)} characters")
            topic_filter = st.text_input(text["topic_filter"], key="pre_topic")
            num_mcqs = st.slider(text["num_mcqs"], 5, 100, 20, key="pre_mcq_slider")
            num_ques = st.slider(text["num_questions"], 3, 50, 10, key="pre_ques_slider")
            col1, col2, col3 = st.columns(3)
            with col1:
                if st.button(text["generate_notes"], key="pre_notes_btn") and pre_text:
                    with st.spinner("Generating notes..."):
                        prompt = build_notes_prompt(pre_text, topic_filter, bilingual, is_hindi)
                        notes = call_cerebras(prompt)
                        if notes:
                            st.session_state['pre_notes'] = notes
                            st.success("Notes ready!")
            with col2:
                if st.button(f"{text['generate_mcqs']} ({num_mcqs})", key="pre_mcq_btn") and pre_text:
                    with st.spinner("Generating MCQs..."):
                        prompt = build_mcq_prompt(pre_text, num_mcqs, topic_filter, bilingual, is_hindi)
                        mcqs = call_cerebras(prompt, parse_json=True)
                        if mcqs:
                            st.session_state['pre_mcqs'] = mcqs
                            st.success(f"{len(mcqs)} MCQs generated!")
            with col3:
                if st.button("📝 Generate Short & Long Questions", key="pre_ques_btn") and pre_text:
                    with st.spinner("Generating questions..."):
                        prompt = build_questions_prompt(pre_text, num_ques, topic_filter, bilingual, is_hindi)
                        questions = call_cerebras(prompt, parse_json=True)
                        if questions:
                            st.session_state['pre_questions'] = questions
                            st.session_state['pre_short'] = [q for q in questions if q.get('type') == 'Short Answer']
                            st.session_state['pre_long'] = [q for q in questions if q.get('type') == 'Long Answer']
                            st.success(f"{len(questions)} questions generated!")
            if 'pre_notes' in st.session_state:
                st.markdown(st.session_state['pre_notes'])
                pdf_data = create_pdf(st.session_state['pre_notes'], "Notes")
                st.markdown(get_download_link(pdf_data, "notes.pdf", "Download PDF"), unsafe_allow_html=True)
            if 'pre_mcqs' in st.session_state:
                for i, mcq in enumerate(st.session_state['pre_mcqs'][:20], 1):
                    with st.expander(f"MCQ {i}: {mcq.get('question','')[:100]}"):
                        st.write(mcq.get('question',''))
                        opts = mcq.get('options',[])
                        for j,opt in enumerate(opts):
                            st.write(f"{chr(65+j)}. {opt}")
                        correct = mcq.get('correct',0)
                        if opts:
                            st.success(f"Answer: {opts[correct]}")
                        st.info(f"Explanation: {mcq.get('explanation','')}")
    else:
        st.info("No pre-loaded notes found. Add PDF files to 'preloaded_notes' folder.")

# ------------------ TAB 2: UPLOAD NOTES ------------------
with tabs[1]:
    st.subheader(text["upload_notes"])
    uploaded_files = st.file_uploader("Upload files", type=['pdf','docx','txt'], accept_multiple_files=True, key="student_upload")
    text_input = st.text_area("Or paste text", height=100, key="student_text")
    topic_filter = st.text_input(text["topic_filter"], key="student_topic")
    num_mcqs = st.slider(text["num_mcqs"], 5, 100, 20, key="student_mcq_slider")
    num_ques = st.slider(text["num_questions"], 3, 50, 10, key="student_ques_slider")
    full_text = ""
    if uploaded_files:
        full_text += get_text_from_files(uploaded_files)
    if text_input:
        full_text += "\n\n" + text_input
    if full_text:
        st.success(f"✅ Loaded {len(full_text)} characters")
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button(text["generate_notes"], key="student_notes_btn") and full_text:
            with st.spinner("Generating notes..."):
                prompt = build_notes_prompt(full_text, topic_filter, bilingual, is_hindi)
                notes = call_cerebras(prompt)
                if notes:
                    st.session_state['student_notes'] = notes
                    st.success("Notes ready!")
    with col2:
        if st.button(f"{text['generate_mcqs']} ({num_mcqs})", key="student_mcq_btn") and full_text:
            with st.spinner("Generating MCQs..."):
                prompt = build_mcq_prompt(full_text, num_mcqs, topic_filter, bilingual, is_hindi)
                mcqs = call_cerebras(prompt, parse_json=True)
                if mcqs:
                    st.session_state['student_mcqs'] = mcqs
                    st.success(f"{len(mcqs)} MCQs generated!")
    with col3:
        if st.button("📝 Generate Short & Long Questions", key="student_ques_btn") and full_text:
            with st.spinner("Generating questions..."):
                prompt = build_questions_prompt(full_text, num_ques, topic_filter, bilingual, is_hindi)
                questions = call_cerebras(prompt, parse_json=True)
                if questions:
                    st.session_state['student_questions'] = questions
                    st.session_state['student_short'] = [q for q in questions if q.get('type') == 'Short Answer']
                    st.session_state['student_long'] = [q for q in questions if q.get('type') == 'Long Answer']
                    st.success(f"{len(questions)} questions generated!")
    if 'student_notes' in st.session_state:
        st.markdown(st.session_state['student_notes'])
        pdf_data = create_pdf(st.session_state['student_notes'], "My_Notes")
        st.markdown(get_download_link(pdf_data, "my_notes.pdf", "Download PDF"), unsafe_allow_html=True)
    if 'student_mcqs' in st.session_state:
        for i, mcq in enumerate(st.session_state['student_mcqs'][:20], 1):
            with st.expander(f"MCQ {i}: {mcq.get('question','')[:100]}"):
                st.write(mcq.get('question',''))
                opts = mcq.get('options',[])
                for j,opt in enumerate(opts):
                    st.write(f"{chr(65+j)}. {opt}")
                correct = mcq.get('correct',0)
                if opts:
                    st.success(f"Answer: {opts[correct]}")
                st.info(f"Explanation: {mcq.get('explanation','')}")

# ------------------ TAB 3: PRE-LOADED PYQs ------------------
with tabs[2]:
    st.subheader(text["pre_loaded_pyqs"])
    pyq_dir = "preloaded_pyqs"
    if not os.path.exists(pyq_dir):
        os.makedirs(pyq_dir, exist_ok=True)
        st.warning(f"Folder '{pyq_dir}' created. Add PYQ PDFs.")
    pyq_files = glob.glob(os.path.join(pyq_dir, "*.pdf"))
    if pyq_files:
        pyq_names = [os.path.basename(f) for f in pyq_files]
        selected_pyq = st.selectbox("Select PYQ file", pyq_names, key="pre_pyq")
        if selected_pyq:
            with open(os.path.join(pyq_dir, selected_pyq), "rb") as f:
                fake_file = io.BytesIO(f.read())
                fake_file.name = selected_pyq
                pyq_text = get_text_from_files([fake_file])
            st.success(f"✅ Loaded {len(pyq_text)} characters")
            st.session_state['pre_pyq_text'] = pyq_text
    else:
        st.info("No pre-loaded PYQs found. Add PDF files to 'preloaded_pyqs' folder.")

# ------------------ TAB 4: PYQ ANALYSIS ------------------
with tabs[3]:
    st.subheader(text["pyq_analysis"])
    exam_name = st.text_input(text["exam_name"])
    syllabus_source = st.radio("Syllabus source", [text["pre_loaded_notes"], text["upload_notes"]], key="syllabus_source")
    syllabus_text = ""
    if syllabus_source == text["pre_loaded_notes"]:
        syllabus_text = st.session_state.get('pre_notes', '')
    else:
        syllabus_text = st.session_state.get('student_notes', '')
    if not syllabus_text:
        st.warning("Please generate notes in Tab 1 or Tab 2 first.")
    pyq_source = st.radio("PYQ source", [text["pre_loaded_pyqs"], "Upload PYQ files now"], key="pyq_source")
    pyq_text = ""
    if pyq_source == text["pre_loaded_pyqs"]:
        pyq_text = st.session_state.get('pre_pyq_text', '')
    else:
        pyq_upload = st.file_uploader("Upload PYQ PDFs", type=['pdf'], accept_multiple_files=True, key="pyq_upload")
        if pyq_upload:
            pyq_text = get_text_from_files(pyq_upload)
    if exam_name and syllabus_text and pyq_text:
        if st.button(text["analyze_pyqs"]):
            with st.spinner("Analyzing PYQs..."):
                prompt = build_pyq_prompt(syllabus_text, pyq_text, exam_name, bilingual, is_hindi)
                analysis = call_cerebras(prompt)
                if analysis:
                    st.markdown(analysis)
                    pdf_data = create_pdf(analysis, "PYQ_Analysis")
                    st.markdown(get_download_link(pdf_data, "pyq_analysis.pdf", "Download PDF"), unsafe_allow_html=True)
    else:
        st.info("Provide exam name, syllabus notes, and PYQs.")

# ------------------ TAB 5: SUBJECT-WISE PYQs ------------------
with tabs[4]:
    st.subheader(text["subject_pyqs"])
    pyq_source_subj = st.radio("Select PYQ source", [text["pre_loaded_pyqs"], "Upload new PYQ PDFs"], key="subj_pyq_source")
    pyq_text = ""
    if pyq_source_subj == text["pre_loaded_pyqs"]:
        pyq_text = st.session_state.get('pre_pyq_text', '')
    else:
        uploaded_pyqs = st.file_uploader("Upload PYQ PDFs", type=['pdf'], accept_multiple_files=True, key="subj_upload")
        if uploaded_pyqs:
            pyq_text = get_text_from_files(uploaded_pyqs)
    if pyq_text:
        if st.button(text["extract_subject_pyqs"]):
            with st.spinner("Extracting subject-wise PYQs..."):
                prompt = build_subject_extract_prompt(pyq_text, bilingual, is_hindi)
                subject_data = call_cerebras(prompt, parse_json=True)
                if subject_data:
                    st.session_state['subject_pyqs'] = subject_data
                    st.success(f"Extracted {len(subject_data)} subjects.")
    if 'subject_pyqs' in st.session_state:
        subjects = list(st.session_state['subject_pyqs'].keys())
        selected_subject = st.selectbox("Select subject", subjects)
        if selected_subject:
            questions = st.session_state['subject_pyqs'].get(selected_subject, [])
            if questions:
                st.markdown(f"### {selected_subject} - PYQs ({len(questions)})")
                for i, q in enumerate(questions, 1):
                    with st.expander(f"Q{i}. {q[:150]}..."):
                        st.write(q)
                all_text = "\n\n".join([f"{i}. {q}" for i,q in enumerate(questions,1)])
                pdf_data = create_pdf(all_text, f"{selected_subject}_PYQs")
                st.markdown(get_download_link(pdf_data, f"{selected_subject}_pyqs.pdf", "Download PDF"), unsafe_allow_html=True)

# ------------------ TAB 6: EXPLAIN ANY TOPIC ------------------
with tabs[5]:
    st.subheader(text["explain"])
    user_topic = st.text_input(text["explain_topic"])
    if st.button(text["explain_btn"], type="primary"):
        if user_topic:
            with st.spinner("Explaining..."):
                prompt = build_explain_prompt(user_topic, bilingual, is_hindi)
                explanation = call_cerebras(prompt)
                if explanation:
                    st.markdown(explanation)
                    pdf_data = create_pdf(explanation, f"Explain_{user_topic[:30]}")
                    st.markdown(get_download_link(pdf_data, "explanation.pdf", "Download PDF"), unsafe_allow_html=True)

# ------------------ TAB 7: QUIZ MODE ------------------
with tabs[6]:
    st.subheader(text["quiz"])
    source = st.radio(text["quiz_source"], [text["pre_loaded_notes"], text["upload_notes"]], key="quiz_source")
    mcq_list = []
    topic_name = ""
    if source == text["pre_loaded_notes"]:
        mcq_list = st.session_state.get('pre_mcqs', [])
        topic_name = "Pre-loaded Notes"
    else:
        mcq_list = st.session_state.get('student_mcqs', [])
        topic_name = "Uploaded Notes"
    if not mcq_list:
        st.info("No MCQs generated yet. Use Tab 1 or 2 to generate MCQs.")
    else:
        num_quiz = st.slider("Number of questions", 1, len(mcq_list), min(10, len(mcq_list)))
        selected_mcqs = mcq_list[:num_quiz]
        if 'quiz_submitted' not in st.session_state:
            st.session_state.quiz_submitted = False
            st.session_state.quiz_answers = {}
        if not st.session_state.quiz_submitted:
            with st.form(key="quiz_form"):
                user_answers = {}
                for i, mcq in enumerate(selected_mcqs):
                    st.markdown(f"**Q{i+1}. {mcq.get('question', '')}**")
                    opts = mcq.get('options', [])
                    if len(opts) == 4:
                        choice = st.radio("Options", opts, key=f"quiz_{i}", index=None)
                        user_answers[i] = choice
                submitted = st.form_submit_button("Submit Quiz")
                if submitted:
                    st.session_state.quiz_answers = user_answers
                    st.session_state.quiz_submitted = True
                    score = 0
                    for i, mcq in enumerate(selected_mcqs):
                        correct_idx = mcq.get('correct', 0)
                        correct_opt = mcq.get('options', [])[correct_idx] if mcq.get('options') else None
                        if st.session_state.quiz_answers.get(i) == correct_opt:
                            score += 1
                    analytics.add_quiz_result(score, len(selected_mcqs), topic_name)
                    st.rerun()
        else:
            score = 0
            for i, mcq in enumerate(selected_mcqs):
                correct_idx = mcq.get('correct', 0)
                correct_opt = mcq.get('options', [])[correct_idx] if mcq.get('options') else None
                if st.session_state.quiz_answers.get(i) == correct_opt:
                    score += 1
            st.markdown(f"### Score: {score} / {len(selected_mcqs)} ({score/len(selected_mcqs)*100:.1f}%)")
            if st.button("Take Another Quiz"):
                st.session_state.quiz_submitted = False
                st.session_state.quiz_answers = {}
                st.rerun()

# ------------------ TAB 8: MOCK TEST ------------------
with tabs[7]:
    st.subheader(text["mock_test"])
    if 'mock_started' not in st.session_state:
        st.session_state.mock_started = False
        st.session_state.mock_answers = {}
        st.session_state.mock_time_left = 0
        st.session_state.mock_submitted = False
    if not st.session_state.mock_started:
        source_mock = st.radio(text["mock_source"], [text["pre_loaded_notes"], text["upload_notes"]], key="mock_source")
        mcq_list_mock = []
        if source_mock == text["pre_loaded_notes"]:
            mcq_list_mock = st.session_state.get('pre_mcqs', [])
        else:
            mcq_list_mock = st.session_state.get('student_mcqs', [])
        if not mcq_list_mock:
            st.info("No MCQs available. Generate MCQs first.")
        else:
            num_questions = st.slider("Number of questions", 5, min(100, len(mcq_list_mock)), 20)
            time_minutes = st.number_input(text["time_minutes"], min_value=1, max_value=180, value=60)
            if st.button(text["start_mock"]):
                st.session_state.mock_mcqs = mcq_list_mock[:num_questions]
                st.session_state.mock_time_left = time_minutes * 60
                st.session_state.mock_started = True
                st.session_state.mock_answers = {}
                st.session_state.mock_submitted = False
                st.rerun()
    else:
        if not st.session_state.mock_submitted:
            timer_placeholder = st.empty()
            mins, secs = divmod(st.session_state.mock_time_left, 60)
            timer_placeholder.markdown(f"### ⏰ Time Left: {mins:02d}:{secs:02d}")
            if st.session_state.mock_time_left > 0:
                time.sleep(1)
                st.session_state.mock_time_left -= 1
                st.rerun()
            else:
                st.session_state.mock_submitted = True
                st.rerun()
        else:
            score = 0
            for i, mcq in enumerate(st.session_state.mock_mcqs):
                user_ans = st.session_state.mock_answers.get(i)
                correct_opt = mcq.get('options', [])[mcq.get('correct', 0)] if mcq.get('options') else None
                if user_ans == correct_opt:
                    score += 1
            total = len(st.session_state.mock_mcqs)
            st.markdown(f"### Score: {score} / {total} ({score/total*100:.1f}%)")
            analytics