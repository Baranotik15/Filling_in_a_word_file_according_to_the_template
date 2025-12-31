# 🧾 Word Template Auto-Filler

This application allows you to automatically generate Microsoft Word
documents from a prepared template.  
The program replaces placeholders in the document with the values
entered by the user.

------------------------------------------------------------------------

## 🚀 How to Use

### 1️⃣ Run the application
Launch `app.exe`.

### 2️⃣ Select a template
Click **“Select template”** and choose a `.docx` file that contains placeholders  
(for example: `{{NAME}}`, `{{AGE}}`, `{{DATE}}`, `{{GENDER}}`, `{{KEY}}`, `{{LOGO}}`).

### 3️⃣ Select a logo image
Click **“Select logo”** and choose an image file:

- `.png`
- `.jpg / .jpeg`
- `.bmp`
- `.gif`

The logo is automatically resized and inserted into the template in the location
of the `{{LOGO}}` placeholder.

📌 **Logo placement rule**
- The placeholder `{{LOGO}}` must be inside a **table cell**
- The logo is aligned to the **right side of that cell** with a small margin
- This guarantees a stable position and prevents layout shifting

👉 Recommended structure:

| Text content (left) | `{{LOGO}}` (right) |
|--------------------|--------------------|

To make the header look professional:
- Create a 2-column table
- Left column → descriptive text
- Right column → `{{LOGO}}`
- Remove table borders in Word if needed

### 4️⃣ Fill in the fields

- Date  
- Name  
- Gender  
- Age  
- Output file name (without `.docx`)

### 5️⃣ Generate the document
Click **“Create document”** — the program will replace all placeholders
inside text and inside tables.

### 6️⃣ Result
A new Word file will be created in the same folder under the name you specified.

------------------------------------------------------------------------

## 📄 Template Format

Placeholders must be written inside double curly brackets:

{{NAME}}
{{AGE}}
{{GENDER}}
{{DATE}}
{{KEY}}
{{LOGO}}


They will be replaced with user-provided values.

### 🖼 `{{LOGO}}` placeholder

- Must be located **inside a table cell**
- Image is resized automatically
- Inserted on the **right side of the cell**
- Text in the left column may extend up to the logo

------------------------------------------------------------------------

## 🧩 Supported Features

- ✔ Works with `.docx` templates  
- ✔ Replaces text in paragraphs  
- ✔ Replaces text inside Word tables  
- ✔ Supports logo insertion via `{{LOGO}}`  
- ✔ Automatically resizes and aligns logo  
- ✔ Prevents document layout deformation  

------------------------------------------------------------------------

## ⚠ Notes

- The template must be in **Microsoft Word (.docx)** format  
- If any required field is empty, the program shows a warning  
- Output files are saved in the same directory as the application  
- For best layout results — use a **table-based header block**  

------------------------------------------------------------------------

## 🛠 Developer Notes

Install dependencies:
```bash
pip install -r requirements.txt
```
Run from source:
```bash
python app.py
```

Build executable (PyInstaller):
```bash
python -m PyInstaller --onefile --noconsole app.py
```
------------------------------------------------------------------------
