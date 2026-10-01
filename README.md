# Automation Assistant

Two tools in one for employees with little time or technical knowledge:

1. **Create an Excel macro.** Upload a data file, describe the task in plain English, and get a ready-to-import
   VBA macro (`.bas`), a plain-English explanation, and an optional process document.
2. **Write a process document.** Paste or upload your notes and get a clear, step-by-step process document
   (Word or text). A macro is optional.

## Run it

```bash
# 1. Create and activate a virtual environment (do not upload this folder)
python -m venv .venv
.venv\Scripts\activate            # Windows (PowerShell)
# source .venv/bin/activate       # Mac / Linux

# 2. Install the requirements
pip install -r requirements.txt

# 3. Start the app
streamlit run app.py
```

## Add an API key

The app works with an Anthropic (Claude) key (starts with `sk-ant-`) or an OpenAI key. Use any ONE of these:

| Option | How |
|---|---|
| Sidebar (easiest) | Open **Connection settings** in the sidebar and paste the key. |
| Secrets file | Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and add the key. |
| Environment variable | Set `ANTHROPIC_API_KEY` or `OPENAI_API_KEY` before running. |

Never write a key into `app.py` before uploading the project to GitHub.

## Try it with the samples

The `samples/` folder has:

- `sample_employees.xlsx` / `.csv`: 100 fictional employees for testing the macro tool.
  Try: "Highlight salaries below 30,000 in red and above 80,000 in green."
- `sample_process_document.docx`: an example of the process document layout.

## Using the macro

1. Download `GeneratedMacro.bas` and the data workbook from the app.
2. Open the workbook in Excel and press **Alt + F11**.
3. Choose **File > Import File...** and select `GeneratedMacro.bas`.
4. Press **Alt + F8**, select the macro, and click **Run**.
5. Save as **Excel Macro-Enabled Workbook (.xlsm)** to keep the macro.

Always test on a copy of your file first.

## Privacy

- Macro tool: sends column names, up to 5 sample rows (none in **Privacy mode**) and your instructions to the AI service.
- Process document writer: sends everything you type or upload. Leave out confidential details.
