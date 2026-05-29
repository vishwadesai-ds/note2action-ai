📋 note2action — AI 

Turn raw meeting notes into structured Minutes of Meeting (MoM) using Vision AI + LLMs.

🚀 Live Demo
Streamlit App: [View->](https://note2action-ai.streamlit.app/)

📌 Problem Statement

Meeting notes are often messy, unstructured, and difficult to convert into actionable tasks.

Teams spend significant time manually creating:

Minutes of Meeting
action trackers
owner assignments
follow-up documentation

note2action automates this workflow using multimodal AI.

🎯 Features

✅ Upload:

Images
PDFs
DOCX files
TXT notes

✅ AI-powered extraction:

Action items
Owners
Deadlines
Priorities
Risks / blockers

✅ Structured output:

Professional MoM tables
Summary metrics
DOCX export

✅ Interactive UI:

Dark mode
File previews
Real-time metrics dashboard
🤖 Tech Stack
Layer	Technology
Frontend	Streamlit
LLM API	OpenRouter AI
Model	Gemini 2.0 Flash Lite
Vision Processing	PIL + PyMuPDF
Document Export	python-docx
Language	Python

🧠 AI Workflow

Meeting Notes -> Document/Image Processing -> Vision + Text Extraction -> LLM Prompting -> Structured MoM Generation -> DOCX Export + Dashboard Metrics

📊 Generated Output

The app automatically creates:

Work Area	Action Item	Owner	Deadline	Priority	Risk	Status

Plus:

total action items
owner assignment metrics
high-priority counts
estimated time saved

⚙️ Run Locally

git clone YOUR_REPO
cd note2action

Install dependencies:

pip install -r requirements.txt

Add environment variable:

OPENROUTER_API_KEY=your_key

Run app:

streamlit run app.py

📦 Requirements
streamlit
openai
python-docx
Pillow
PyMuPDF
python-dotenv

🛠️ Test Files Folder

Where you will find files to test your application [Folder->](TestFiles/)

👤 Author

Vishwa Desai
- LinkedIn: [View->](https://www.linkedin.com/in/vishwa-desai-ds/)
