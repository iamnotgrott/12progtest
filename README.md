# Assignment Grader

A Streamlit application for automatically grading student Python assignments using AI via OpenRouter.

## Features

- Upload ZIP files containing task descriptions and student submissions
- Automatically analyze tasks and generate grading criteria (if not provided)
- Grade Python submissions using AI
- Generate individual feedback documents for each student
- Create comprehensive teacher reports
- **Support for FREE AI models** - no costs incurred!
- Multiple model options (free and premium)

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Get a free API key from https://openrouter.ai/

## Usage

1. Run the Streamlit app:
```bash
streamlit run app.py
```

2. Prepare your ZIP file with the following structure:
```
assignment.zip
├── task/
│   ├── description.docx (or .pdf, .md, .txt)
│   ├── criteria.json (optional)
│   └── description.md (optional)
└── submissions/
    ├── ivan_ivanov_12v_sau1.py
    ├── maria_petrova_12v_sau1.py
    └── ...
```

3. Enter your OpenRouter API key in the app
4. Select an AI model (Free Models cost nothing!)
5. Upload the ZIP file
6. Click "Grade Submissions"
7. Download the generated feedback files and teacher report

## Available Models

### Free Models (No Cost!)
- Meta Llama 3.1 8B
- Google Gemma 2 9B
- Mistral 7B
- Qwen 2.5 7B

### Premium Models (Paid)
- Claude Sonnet 3.5
- GPT-4 Turbo
- GPT-4o
- Claude Opus 3

## File Naming Convention

Submissions must follow this pattern:
```
name_surname_grade_assessment.py
```

Example: `ivan_ivanov_12v_sau1.py`

## Generated Files

- `{name}_{surname}_{grade}_{assessment}_feedback.docx` - Individual student feedback
- `assessment_grade_report.docx` - Comprehensive teacher report with all grades
