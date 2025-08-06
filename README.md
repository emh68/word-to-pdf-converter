# Word to PDF Converter
![Python](https://img.shields.io/badge/Python-3.13-blue.svg)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-lightgrey)
![License: CC BY-NC 4.0](https://img.shields.io/badge/License-CC%20BY--NC%204.0-lightgrey.svg)

A simple desktop GUI application to convert Microsoft Word `.docx` files to `.pdf` format using the `docx2pdf` library.  

![](./media/word-to-pdf-converter-demo.gif)

## Features

- Convert one or more Word `.docx` files to `.pdf`.
- Select multiple files and convert them all at once.
- View and manage input and output paths in a list view.
- Choose a custom output folder for each file.
- Remove selected files from the list.
- Log window to track file additions/removals, path settings, and conversion results.
- Clear log output manually by selecting and deleting text.

## How It Works

1. **Add Word Files** – Select one or more `.docx` files to add to the conversion list.
2. **Set Output Path for Selected** – Choose where each converted PDF will be saved.
3. **Remove Selected** – Select the files you want to remove, then click the **Remove Selected** button.
4. **Convert All** – Start the conversion of all added files. 
5. **View Logs** – Check the log window for real-time updates and results.

![]()

## Requirements

- Python 3.x
- `docx2pdf` (installed in step #3 below)
- `tkinter` (usually included with Python installations)


## Getting Started
Follow the steps below to set up and run the word-to-pdf-converter locally.  

### 1. Clone the repository
<small>Open your terminal (Command Prompt, PowerShell, Git Bash, or macOS/Linux Terminal), then run:</small>
```bash
git clone https://github.com/emh68/word-to-pdf-converter.git
cd word-to-pdf-converter
```

### 2. (Optional but recommended) Create and activate a virtual environment
<small>This keeps dependencies isolated to this project.</small>


<img src="https://upload.wikimedia.org/wikipedia/commons/8/87/Windows_logo_-_2021.svg" alt="Windows logo" width="15"/> Windows (command line)
```
python -m venv venv
venv\Scripts\activate
```
<img src="https://upload.wikimedia.org/wikipedia/commons/8/87/Windows_logo_-_2021.svg" alt="Windows logo" width="15"/> Windows (PowerShell)
> <small>If you're using PowerShell and `activate` doesn't work, try:</small>
```bash
venv\Scripts\Activate.ps1
```

🍎 macOS/🐧Linux
```
python3 -m venv venv
source venv/bin/activate
```

### 3. Install the dependencies
<small>Run the command below. This will install the required dependencies (in this case `docx2pdf`). If you get a "pip not found" error, try `python -m pip install -r requirements.txt` instead.</small>

```bash
pip install -r requirements.txt
```

### 4. Run the application
```bash
python main.py 
```
>**Note:** <small>If you get a *"command not found"* error, try:</small>
```bash
python3 main.py
```

## 🔗 Links

- [View the GitHub repo](https://github.com/emh68/word-to-pdf-converter)
- [More of my projects](https://github.com/emh68)