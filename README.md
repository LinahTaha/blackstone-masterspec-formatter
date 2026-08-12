# Blackstone MasterSpec Formatter

A Streamlit tool that takes a messy civil engineering spec document and reformats it to match a MasterSpec style template, using a two call OpenAI pipeline.

## How it works
1. Upload a template doc (defines the target formatting structure) and a messy doc (the one to reformat)
2. First OpenAI call analyzes the template's heading/numbering structure
3. Second OpenAI call extracts and restructures the messy doc's content to match
4. Outputs a clean, formatted Word document ready to download

## Setup
```bash
pip install -r requirements.txt
```

Create a `config.py` file with your OpenAI API key:
```python
import os
os.environ["OPENAI_API_KEY"] = "your-key-here"
```

## Run
```bash
streamlit run app.py
```

## Built with
- Streamlit
- OpenAI API (GPT-4o-mini)
- python-docx
