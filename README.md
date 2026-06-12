# 🧭 CareerIQ — AI Career Mentor

> **Agents League Hackathon 2026 · Reasoning Agents Track**  
> Built on Microsoft Foundry · Powered by gpt-4.1-mini · Azure AI Projects SDK
## 👩‍💻 Author

**Dilshad Shaik**  
MS Computer Science · Auburn University at Montgomery
[LinkedIn](https://linkedin.com/in/dilshad-shaik) · [GitHub](https://github.com/dilshadshaik3031) · [Portfolio](https://dilshadshaik3031.github.io/MyPortfolio)

[![Live Demo](https://img.shields.io/badge/Live%20Demo-careeriq.onrender.com-0078d4?style=flat-square&logo=microsoft-azure)](https://careeriq.onrender.com)
[![Microsoft Foundry](https://img.shields.io/badge/Microsoft-Foundry-0078d4?style=flat-square&logo=microsoft)](https://ai.azure.com)
[![Python](https://img.shields.io/badge/Python-3.11-3776ab?style=flat-square&logo=python)](https://python.org)
[![Flask](https://img.shields.io/badge/Flask-3.0-000000?style=flat-square&logo=flask)](https://flask.palletsprojects.com)

---

## 🎯 What is CareerIQ?

CareerIQ is a **multi-agent AI career mentor** that helps students and professionals navigate career transitions — from identifying skill gaps to landing their dream job — one achievable step at a time.

Unlike generic chatbots, CareerIQ uses **multi-step reasoning agents** built on Microsoft Foundry to deliver personalized, actionable career guidance with real certification links, live roadmaps, and ATS-optimized resume generation.

---

## ✨ Features

### 💬 Career Coach Agent
- **Visible reasoning trace** — watch the agent think step by step in real time
- **Personalized step-by-step roadmap** — from where you are to your dream job
- **Real certification links** — AZ-900, AI-102, AWS SAA, CKA, Terraform, and 15+ more
- **Career roadmaps** — linked directly to roadmap.sh for every role
- **Free resources** — Microsoft Learn, AWS Skill Builder, freeCodeCamp, CS50, Kaggle
- **Market insights** — salary ranges, demand trends, top hiring companies in 2026

### 📄 Resume Studio (4 parallel agents)
| Agent | What it does |
|-------|-------------|
| **Resume Optimizer** | Rewrites resume to score 93+ ATS using 7 proven rules |
| **Cover Letter Writer** | Tailored, specific, under 300 words — never generic |
| **Match Analyzer** | ATS score, missing keywords, path to 93+, quick wins |
| **Interview Prep** | 10 tailored questions + STAR-format model answers |

### 🏆 My Roadmap (Milestone Tracker)
- Add milestones from your career roadmap
- Check off completed steps
- Live progress bar toward your dream job
- Persists across sessions

---

## 🧠 Multi-Step Reasoning Architecture

```
User Input
    │
    ▼
┌─────────────────────────────────────────┐
│  Reasoning Trace (visible to user)      │
│  Step 1: Assessing current skills       │
│  Step 2: Identifying skill gaps         │
│  Step 3: Mapping career path            │
│  Step 4: Sourcing certifications        │
│  Step 5: Building personalized plan     │
└─────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────┐
│         Microsoft Foundry               │
│      Hosted Agent · gpt-4.1-mini        │
│      Azure AI Projects SDK              │
│      Responses API (streaming SSE)      │
└─────────────────────────────────────────┘
    │
    ├──► Career Coach Agent
    ├──► Resume Optimizer Agent (ATS 93+)
    ├──► Cover Letter Agent
    ├──► Match Analysis Agent
    └──► Interview Prep Agent
```

---

## 🏗️ Tech Stack

| Layer | Technology |
|-------|-----------|
| **AI Platform** | Microsoft Foundry · gpt-4.1-mini |
| **Agent SDK** | Azure AI Projects SDK v2.0+ |
| **API Protocol** | OpenAI Responses API · Server-Sent Events (streaming) |
| **Authentication** | Azure DefaultAzureCredential (Managed Identity) |
| **Backend** | Python 3.11 · Flask 3.0 |
| **Frontend** | Microsoft Fluent Design System · Vanilla JS |
| **PDF Generation** | ReportLab · python-docx |
| **File Parsing** | pdfplumber · python-docx |
| **Deployment** | Azure App Service / Render |

---

## 🚀 Quick Start

### Prerequisites
- Python 3.11+
- Azure subscription with Microsoft Foundry project
- Azure CLI (`az login`)

### Local Setup

```bash
# 1. Clone the repo
git clone https://github.com/dilshadshaik3031/careeriq.git
cd careeriq

# 2. Create virtual environment
python -m venv .venv
.venv\Scripts\activate        # Windows
source .venv/bin/activate     # Mac/Linux

# 3. Install dependencies
pip install -r requirements.txt

# 4. Configure environment
cp .env.example .env
# Edit .env with your Foundry endpoint

# 5. Login to Azure
az login --use-device-code

# 6. Run
python app.py
```

Open `http://localhost:5000`

### Environment Variables

```env
FOUNDRY_PROJECT_ENDPOINT=https://<resource>.services.ai.azure.com/api/projects/<project>
AZURE_AI_MODEL_DEPLOYMENT_NAME=gpt-4.1-mini
```

---

## 📁 Project Structure

```
careeriq/
├── app.py                  # Flask app + all agent endpoints
├── templates/
│   └── index.html          # Microsoft Fluent Design UI
├── requirements.txt        # Python dependencies
├── startup.sh              # Azure App Service startup
├── .deployment             # Azure build config
├── .env.example            # Environment template
└── README.md
```

### API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | Main UI |
| `/chat/traced` | POST | Career coach with reasoning trace (SSE) |
| `/upload` | POST | Parse PDF/DOCX/TXT resume or JD |
| `/studio/resume` | POST | ATS 93+ resume optimizer (SSE) |
| `/studio/cover` | POST | Cover letter generator (SSE) |
| `/studio/analysis` | POST | ATS match analyzer (SSE) |
| `/studio/interview` | POST | Interview prep generator (SSE) |
| `/download/resume/pdf` | POST | Download optimized resume as PDF |
| `/download/resume/docx` | POST | Download optimized resume as DOCX |
| `/download/cover/pdf` | POST | Download cover letter as PDF |
| `/download/cover/docx` | POST | Download cover letter as DOCX |
| `/download/analysis/pdf` | POST | Download match analysis report as PDF |

---

## 🎯 Hackathon Track

**Agents League Hackathon 2026 — Reasoning Agents Track**

> *"Create intelligent agents using Microsoft Foundry that solve complex problems through multi-step reasoning."*

CareerIQ qualifies because:
1. **Multi-step reasoning** — 5-step visible thinking chain before every response
2. **Microsoft Foundry native** — built on hosted agents, Azure AI Projects SDK, Responses API
3. **Complex problem** — career transitions require analyzing skills, gaps, market data, certifications, and producing multiple personalized outputs
4. **Real impact** — helps millions of students and career changers worldwide

---

## 👩‍💻 Author

**Dilshad Shaik**  
MS Computer Science · Auburn University at Montgomery  
IT Systems Engineer · AUM  
[LinkedIn](https://linkedin.com/in/dilshad-shaik) · [GitHub](https://github.com/dilshadshaik3031) · [Portfolio](https://dilshadshaik3031.github.io/MyPortfolio)

---

## 📄 License

MIT License — feel free to use and build on this project.

---

*Built with ❤️ on Microsoft Foundry for the Agents League Hackathon 2026*
