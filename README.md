
# 🧾 Word Template Auto-Filler

This application allows you to automatically generate Microsoft Word
documents from a prepared template.
The program replaces placeholders in the document with the values
entered by the user.

------------------------------------------------------------------------

## 🚀 How to Use

### 1️⃣ Run the application
Launch app.exe.

### #️⃣ Select a template
Click “Select template” and choose a .docx file that contains
placeholders
(for example: {{NAME}}, {{AGE}}, {{DATE}}).

### 3️⃣ Fill in the fields - Date
- Name
- Gender
- Age
- Output file name (without .docx)

### 4️⃣ Generate the document
Click “Create document” — the program will replace all placeholders in
the template.

### 5️⃣ Result
A new Word file will be created in the same folder under the name you
specified.

------------------------------------------------------------------------

## 📄 Template Format

Placeholders must be written inside double curly brackets:

    {{NAME}}
    {{AGE}}
    {{GENDER}}
    {{DATE}}
    {{KEY}}

They will be replaced with the user-provided values.

------------------------------------------------------------------------

### 🧩 Supported Features

-   ✔ Works with .docx templates
-   ✔ Replaces text in paragraphs
-   ✔ Replaces text inside Word tables

------------------------------------------------------------------------

### ⚠ Notes

-   The template must be in Microsoft Word (.docx) format
-   If any required field is empty, the program will show a warning
-   Output files are saved in the same directory as the application

------------------------------------------------------------------------

## 🛠 Developer Notes

Install dependencies:

    pip install -r requirements.txt

Run from source:

    python app.py

Build executable (PyInstaller):

    python -m PyInstaller --onefile --noconsole app.py

------------------------------------------------------------------------