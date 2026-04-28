import streamlit as st
import os
import pandas as pd
from datetime import datetime
from openpyxl import Workbook, load_workbook

# Page Config
st.set_page_config(
    page_title="AI Persona Quiz",
    page_icon="🤖",
    layout="centered",
    initial_sidebar_state="expanded"
)


# Custom CSS
st.markdown("""
<style>
    .main-title { text-align: center; font-size: 2.5rem; font-weight: 800; }
    .sub-title { text-align: center; font-size: 1.1rem; color: #666; margin-bottom: 2rem; }
    .persona-result {
        text-align: center; font-size: 2rem; font-weight: 700;
        padding: 1rem; border-radius: 12px; margin: 1rem 0;
    }
    .persona-desc { text-align: center; font-size: 1.1rem; padding: 0.5rem 1rem; color: #444; }
</style>
""", unsafe_allow_html=True)

# Persona Definitions
PERSONAS = {
    "AI Driver": {
        "emoji": "🚀",
        "color": "#2ecc71",
        "description": "You are an AI Champion! 🚀 You actively seek out AI tools, experiment boldly, and inspire others to embrace the future. You see AI as a superpower that amplifies human potential.",
        "traits": ["Early adopter", "Tech evangelist", "Innovation leader", "Always experimenting", "Loves automation"]
    },
    "Quiet Optimizer": {
        "emoji": "🧠",
        "color": "#3498db",
        "description": "You are a Silent Strategist! 🧠 You use AI smartly behind the scenes to boost your productivity. You do not need to shout about it. Your results speak for themselves.",
        "traits": ["Efficiency-focused", "Low-key power user", "Results-driven", "Practical thinker", "Works smarter"]
    },
    "Cautious Tester": {
        "emoji": "🔍",
        "color": "#f39c12",
        "description": "You are a Curious Explorer! 🔍 You see the potential of AI but want to understand it better before diving in. You ask great questions and prefer to test the waters carefully.",
        "traits": ["Thoughtful evaluator", "Asks good questions", "Open-minded", "Risk-aware", "Steady learner"]
    },
    "Skeptic": {
        "emoji": "🛡️",
        "color": "#e74c3c",
        "description": "You are a Critical Thinker! 🛡️ You value trust, accuracy, and proven methods. You push back on hype and demand that AI proves its worth before you buy in.",
        "traits": ["Values accuracy", "Questions everything", "Trust-focused", "Detail-oriented", "Prefers proven methods"]
    },
    "Unengaged": {
        "emoji": "💎",
        "color": "#9b59b6",
        "description": "You are an Untapped Opportunity! 💎 AI has not clicked for you yet, and that is okay! Once you see how it solves YOUR specific problems, you might just surprise everyone.",
        "traits": ["Waiting for the right moment", "Focused on current tasks", "Potential game-changer", "Needs a personal use case", "Fresh perspective"]
    }
}

# Quiz Questions
QUESTIONS = [
    {
        "question": "1. A new AI tool is introduced at work. What is your first reaction?",
        "options": {
            "Sign me up! I want to try it right away.": "AI Driver",
            "I will explore it quietly on my own time.": "Quiet Optimizer",
            "Interesting. I will watch some demos first.": "Cautious Tester",
            "I will wait and see if it actually works.": "Skeptic",
            "Another tool? I am fine with what I have.": "Unengaged"
        }
    },
    {
        "question": "2. How do you feel about AI writing your emails?",
        "options": {
            "Love it! I already use it for drafts.": "AI Driver",
            "I use it sometimes. It saves me time.": "Quiet Optimizer",
            "I would try it, but I would heavily edit the result.": "Cautious Tester",
            "No thanks. It will not sound like me.": "Skeptic",
            "I do not see why I would need that.": "Unengaged"
        }
    },
    {
        "question": "3. Your team wants to automate a repetitive task with AI. You...",
        "options": {
            "Volunteer to lead the project!": "AI Driver",
            "Quietly suggest the best tool for the job.": "Quiet Optimizer",
            "Ask for a trial period to test it first.": "Cautious Tester",
            "Raise concerns about accuracy and reliability.": "Skeptic",
            "Let the team decide. It does not affect me much.": "Unengaged"
        }
    },
    {
        "question": "4. When you hear 'artificial intelligence', you think...",
        "options": {
            "Endless possibilities and innovation!": "AI Driver",
            "A useful tool when applied correctly.": "Quiet Optimizer",
            "Promising, but I need to learn more.": "Cautious Tester",
            "Overhyped and potentially risky.": "Skeptic",
            "Not really something that affects my day-to-day.": "Unengaged"
        }
    },
    {
        "question": "5. A coworker shows you a cool AI trick. You...",
        "options": {
            "Get excited and ask them to teach you more!": "AI Driver",
            "Take a mental note to try it later.": "Quiet Optimizer",
            "Think it is neat but wonder about the limitations.": "Cautious Tester",
            "Wonder if it is really reliable enough to use.": "Skeptic",
            "Smile and nod. It is not really your thing.": "Unengaged"
        }
    },
    {
        "question": "6. Your company offers free AI training. You...",
        "options": {
            "Sign up immediately and encourage your team to join!": "AI Driver",
            "Enroll quietly. Free skill upgrade, why not?": "Quiet Optimizer",
            "Look at the curriculum first to see if it is relevant.": "Cautious Tester",
            "Pass. I would rather spend time on proven skills.": "Skeptic",
            "Did not notice the announcement.": "Unengaged"
        }
    },
    {
        "question": "7. How often do you use AI tools (ChatGPT, Copilot, etc.)?",
        "options": {
            "Daily. It is part of my workflow.": "AI Driver",
            "A few times a week for specific tasks.": "Quiet Optimizer",
            "Occasionally. I am still figuring them out.": "Cautious Tester",
            "Rarely. I do not fully trust the output.": "Skeptic",
            "Never. I have not felt the need.": "Unengaged"
        }
    },
    {
        "question": "8. AI makes a mistake on a task. Your reaction?",
        "options": {
            "Expected! I will tweak the prompt and try again.": "AI Driver",
            "Note the limitation and adjust my approach.": "Quiet Optimizer",
            "This is why I always double-check the output.": "Cautious Tester",
            "See, this is exactly why I do not trust it.": "Skeptic",
            "I would not know. I do not use it.": "Unengaged"
        }
    },
    {
        "question": "9. If AI could do 50 percent of your job, you would feel...",
        "options": {
            "Thrilled! More time for creative and strategic work.": "AI Driver",
            "Great. I would use the extra time productively.": "Quiet Optimizer",
            "Mixed. Excited but a little nervous.": "Cautious Tester",
            "Worried about job security and quality control.": "Skeptic",
            "Doubtful. AI cannot really do what I do.": "Unengaged"
        }
    },
    {
        "question": "10. Your manager asks for your opinion on adopting AI. You say...",
        "options": {
            "Let us go all in! Here is my proposal.": "AI Driver",
            "I have some ideas. Let me share them privately.": "Quiet Optimizer",
            "I am open to it, but let us start with a pilot.": "Cautious Tester",
            "We need to be very careful about the risks.": "Skeptic",
            "I do not have a strong opinion on this.": "Unengaged"
        }
    },
    {
        "question": "11. How do you describe AI to a friend?",
        "options": {
            "A game-changer that is transforming everything!": "AI Driver",
            "A handy tool, like a smart assistant.": "Quiet Optimizer",
            "Interesting technology, but still a work in progress.": "Cautious Tester",
            "Something to approach with healthy skepticism.": "Skeptic",
            "Tech stuff I do not really follow.": "Unengaged"
        }
    },
    {
        "question": "12. In 5 years, AI in the workplace will be...",
        "options": {
            "Everywhere, and I will be ahead of the curve!": "AI Driver",
            "Very useful for those who know how to leverage it.": "Quiet Optimizer",
            "Mainstream, but we will still need human oversight.": "Cautious Tester",
            "Overpromised. Reality will not match the hype.": "Skeptic",
            "Probably around, but I will cross that bridge later.": "Unengaged"
        }
    }
]

# Excel Save Function
EXCEL_FILE = "ai_persona_results.xlsx"

def save_to_excel(user_name, persona, scores):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "AI Persona Results"
        headers = [
            "Timestamp", "User Name", "AI Persona Result",
            "AI Driver Score", "Quiet Optimizer Score",
            "Cautious Tester Score", "Skeptic Score", "Unengaged Score"
        ]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
    row = [
        timestamp, user_name, persona,
        scores.get("AI Driver", 0), scores.get("Quiet Optimizer", 0),
        scores.get("Cautious Tester", 0), scores.get("Skeptic", 0),
        scores.get("Unengaged", 0)
    ]
    ws.append(row)
    wb.save(EXCEL_FILE)

# Session State Init
if "quiz_submitted" not in st.session_state:
    st.session_state.quiz_submitted = False
if "result_persona" not in st.session_state:
    st.session_state.result_persona = None
if "result_scores" not in st.session_state:
    st.session_state.result_scores = {}

# Sidebar
with st.sidebar:
    st.markdown("## 🤖 AI Persona Quiz")
    st.markdown("---")
    st.markdown("## 👤 About You")
    user_name = st.text_input("Enter your name:", placeholder="e.g., Jade Ryan")
    st.markdown("---")
    st.markdown("### 📊 How It Works")
    st.markdown("1. Answer **12 quick questions**")
    st.markdown("2. Get your **AI Persona** instantly")
    st.markdown("3. Results saved to **Excel** automatically")
    st.markdown("4. Share with your team! 🎉")
    st.markdown("---")
    st.markdown("### 🎭 The 5 AI Personas")
    for name, info in PERSONAS.items():
        st.markdown(f"{info['emoji']} **{name}**")
    st.markdown("---")
    st.markdown("Built with Streamlit")

# Main Content
st.markdown('<div class="main-title">🤖 What is Your AI Persona?</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Answer 12 quick questions to discover your AI personality at work!</div>', unsafe_allow_html=True)

# Show Results
if st.session_state.quiz_submitted and st.session_state.result_persona:
    persona = st.session_state.result_persona
    scores = st.session_state.result_scores
    info = PERSONAS[persona]

    st.markdown("---")
    st.markdown(
        f'<div class="persona-result" style="background: linear-gradient(135deg, {info["color"]}22, {info["color"]}44); border: 2px solid {info["color"]};">'
        f'{info["emoji"]} You are: <strong>{persona}</strong> {info["emoji"]}'
        f'</div>',
        unsafe_allow_html=True
    )
    st.markdown(f'<div class="persona-desc">{info["description"]}</div>', unsafe_allow_html=True)

    st.markdown("")
    st.markdown("#### 🏷️ Your Key Traits:")
    trait_cols = st.columns(len(info["traits"]))
    for i, trait in enumerate(info["traits"]):
        trait_cols[i].markdown(
            f'<div style="text-align:center; padding:8px; background:{info["color"]}22; border-radius:8px; font-size:0.85rem;">{trait}</div>',
            unsafe_allow_html=True
        )

    st.markdown("")
    st.markdown("#### 📊 Your Score Breakdown:")
    score_df = pd.DataFrame({
        "Persona": list(scores.keys()),
        "Score": list(scores.values())
    }).set_index("Persona")
    st.bar_chart(score_df)

    st.success(f"Results saved to {EXCEL_FILE}!")

    st.markdown("")
    if st.button("🔄 Take Quiz Again", use_container_width=True):
        st.session_state.quiz_submitted = False
        st.session_state.result_persona = None
        st.session_state.result_scores = {}
        st.rerun()

# Quiz Form
else:
    if not user_name:
        st.warning("👈 Please enter your **name** in the sidebar to start the quiz.")
        st.stop()

    st.markdown(f"**Welcome, {user_name}!** Let us find out your AI persona. 👇")
    st.markdown("---")

    answers = {}
    all_answered = True

    for i, q in enumerate(QUESTIONS):
        st.markdown(f"**{q['question']}**")
        options = list(q["options"].keys())
        answer = st.radio(
            label=f"Q{i+1}",
            options=options,
            index=None,
            key=f"q_{i}",
            label_visibility="collapsed"
        )
        if answer is None:
            all_answered = False
        else:
            answers[i] = q["options"][answer]
        st.markdown("")

    st.markdown("---")

    if not all_answered:
        st.info("📝 Please answer **all 12 questions** to reveal your AI persona.")

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        submit = st.button(
            "🎯 Get My AI Persona!",
            use_container_width=True,
            disabled=not all_answered
        )

    if submit and all_answered:
        scores = {p: 0 for p in PERSONAS}
        for persona_val in answers.values():
            scores[persona_val] += 1
        winning_persona = max(scores, key=scores.get)
        save_to_excel(user_name, winning_persona, scores)
        st.session_state.quiz_submitted = True
        st.session_state.result_persona = winning_persona
        st.session_state.result_scores = scores
        st.balloons()
        st.rerun()
