# AI Readiness Copilot

AI Readiness Copilot is an AI-powered decision-support platform designed to assess and compare AI readiness across countries using real-world indicators and explainable analytics.

This project currently focuses on **Nigeria 🇳🇬 and Algeria 🇩🇿**, integrating multi-source public datasets to evaluate readiness across key sectors and provide policy insights through an AI assistant.

---

## Overview

Governments and policymakers often struggle to convert large amounts of fragmented data into actionable insights.

AI Readiness Copilot addresses this by:

- Collecting indicators from trusted public sources
- Computing sector-level readiness scores
- Comparing countries side by side
- Providing AI-generated explanations and recommendations
- Simulating investment and policy scenarios

---

## Features

### Dashboard
- Country comparison: Nigeria vs Algeria
- Overall AI readiness score
- Sector-level readiness scores:
  - Education
  - Workforce
  - Healthcare
  - Government
  - Infrastructure
- Visual charts and analytics

### AI Assistant
- Conversational chatbot
- Policy recommendations
- Country comparisons
- Readiness explanations
- Scenario analysis

### Multi-source Data Integration
Data collected from:

- World Bank
- WHO Global Health Observatory
- UNESCO UIS
- ITU
- Mo Ibrahim IIAG (governance indicators)

### Explainable AI
- Confidence scores
- Data source tracking
- Stale-data detection
- Human review guardrails

---

## System Architecture

```text
Public Data Sources
(World Bank / WHO / UNESCO / ITU / IIAG)
                    ↓
            Data Collection Layer
                    ↓
            Data Cleaning & Processing
                    ↓
          AI Readiness Scoring Engine
                    ↓
              Flask Backend API
                    ↓
       Dashboard + AI Chat Interface
                    ↓
               Mistral AI Model
```

---

## Tech Stack

### Frontend
- HTML
- CSS
- JavaScript
- Dashboard UI

### Backend
- Flask
- Flask-CORS

### Data Processing
- Python
- Requests
- OpenPyXL

### AI
- Mistral API

### Deployment
- GitHub
- Netlify

---

## Project Structure

```text
AI_Readiness_Copilot/
│
├── app.py
├── copilot.html
├── scores.xlsx
├── all_data.json
├── all_data.xlsx
├── requirements.txt
│
├── outputs/
│
├── scripts/
│   ├── data_fetcher.py
│   ├── score_calculator.py
│
├── assets/
│
└── README.md
```

---

## Installation

Clone the repository:

```bash
git clone https://github.com/YOUR_USERNAME/AI_Readiness_Copilot.git
```

Move into project folder:

```bash
cd AI_Readiness_Copilot
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Run Flask backend:

```bash
python app.py
```

Open:

```text
http://localhost:5000
```

---

## API Endpoints

### Health Check

```http
GET /api/health
```

Returns system status.

### Scores

```http
GET /api/scores
```

Returns current readiness scores.

### Chat

```http
POST /api/chat
```

Example:

```json
{
  "messages":[
    {
      "role":"user",
      "content":"Compare healthcare readiness in Nigeria and Algeria"
    }
  ]
}
```

---

## Data Notes

Some indicators may be:

- Missing
- Stale (>5 years old)
- Estimated when public data is unavailable

The system explicitly flags low-confidence or outdated information.

---

## Future Improvements

- Add more countries
- Real-time data updates
- Advanced forecasting
- Interactive scenario simulations
- Additional governance metrics
- Authentication and user accounts

---

## Team

Hackathon Team Project

Contributors:

- Mokeddem Halima Saadia
- Kolad Victor

---

## License

Custom License

Commercial use of this project requires explicit permission from the team.

---

## Acknowledgments

Data sources:

- World Bank
- WHO
- UNESCO UIS
- ITU
- Mo Ibrahim Foundation

AI support:

- Mistral AI

---

Built with ❤️ for innovation and evidence-based policymaking.
