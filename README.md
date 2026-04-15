# Certificate Generator

A Python desktop application for batch-generating personalized PDF certificates from a template and a list of names. It features a live, interactive preview that allows precise visual placement, scaling, and rotation of text directly on the template.

## Installation Guide

### Prerequisites

- **Python 3.7 or later**
- A display server (X11/Wayland on Linux, built-in on Windows/macOS)

### Step 1: Install Python

#### Linux
Most Linux distributions come with Python 3 pre-installed. Verify by running:
```bash
python3 --version
```

If Python is not installed:
- **Ubuntu/Debian**: `sudo apt install python3 python3-pip python3-tk`
- **Fedora/RHEL**: `sudo dnf install python3 python3-pip python3-tkinter`
- **Arch Linux**: `sudo pacman -S python python-pip tk`

#### Windows
1. Download Python from [python.org](https://www.python.org/downloads/)
2. Run the installer
3. **Important**: Check the box "Add Python to PATH" during installation
4. Click "Install Now"
5. Verify installation by opening Command Prompt and running:
```cmd
python --version
```

### Step 2: Clone or Download the Project

```bash
git clone https://github.com/Jopseps/Auto-Certificate-Generator
cd Auto-Certificate-Generator
```

Or download the ZIP file and extract it.

### Step 3: Install Python Dependencies

#### Linux
```bash
python3 -m pip install -r requirements.txt
```

If you encounter permission issues, use:
```bash
python3 -m pip install --user -r requirements.txt
```

#### Windows
Open Command Prompt in the project folder and run:
```cmd
python -m pip install -r requirements.txt
```

### Step 4: Verify Installation

#### Linux
```bash
python3 AutoCert.py
```

#### Windows
```cmd
python AutoCert.py
```

The application window should open with the certificate generator interface.

### Dependencies

The project automatically installs the following packages:
- `PyMuPDF` (fitz) - PDF manipulation
- `pandas` - Excel data handling
- `openpyxl` - Reading .xlsx files
- `Pillow` - Image processing
- `tkinter` - GUI (system library, pre-installed on most systems)

### Troubleshooting

**"Python not found" on Windows**
- Restart your terminal or computer after installation
- Verify Python was added to PATH: Open Command Prompt and type `python --version`

**"tkinter not found" on Linux**
- Ubuntu/Debian: `sudo apt install python3-tk`
- Fedora/RHEL: `sudo dnf install python3-tkinter`
- Arch Linux: `sudo pacman -S tk`

**Permission denied on Linux**
- Use `python3 -m pip install --user -r requirements.txt` instead

**"Module not found" errors**
- Ensure you've run the pip install command in the correct directory
- Try: `python3 -m pip install --upgrade pip` then reinstall requirements

## Features

- **Interactive Preview**: Select the text bounding box in the preview to move, resize, and rotate text directly using the mouse.
- **Batch Export**: Generates individual PDF files for every row in a selected Excel column.
- **Custom Font Support**: Load custom `.otf` or `.ttf` files, or select from installed system fonts.
- **Text Splitting Logic**: Automatically splits long names into multiple lines based on configurable word counts and character thresholds.
- **Configuration Profiles**: Save and load styling and file path configurations as JSON files.

## Usage

Run the application:
```bash
python3 AutoCert.py
```

### 1. File Configuration
- **Template PDF**: The blank background certificate design.
- **Excel File**: A spreadsheet containing the names of the recipients.
- **Output Dir**: The folder where the generated PDFs will be saved.

### 2. Styling Text
- Select your target font either from a local file or the system font dropdown.
- Adjust values using the spinboxes or use the interactive preview.
- To use the interactive preview: click inside the dotted bounding box to activate handles. Drag the center to move, drag corners to resize, and drag the top handle to rotate. Click outside the box to deactivate.
- **Pro-tip:** While the handles are active, you can use the **Arrow Keys** to nudge the text exactly by 1px. Start holding **Shift** to nudge by 5px! 
- **Pro-tip 2:** Use **Ctrl+Z** and **Ctrl+Shift+Z** to Undo and Redo your text manipulation anywhere.
- Hold Shift while rotating to snap to 15-degree increments.

### 3. Data Association
- Select the exact column from the Excel file that contains the names.
- Use the left/right arrow keys or the Next/Prev buttons below the preview to cycle through names and preview edge-cases (e.g., extremely long names).

### 4. Text Splitting
If names are too wide for the template, configure the split thresholds:
- **Auto**: Splits names exceeding the character threshold into two lines.
- **Always**: Forces multi-word names to split into two lines.
- **No Split**: Forces names to stay on a single line regardless of length.

### 5. Export
- Click "Generate All" to process the entire Excel list. The application will output standard PDF files corresponding to each name in the designated output directory.
- Click "Save" in the Settings panel to export the current configuration (files, positions, sizes, and colors) to a JSON file.
