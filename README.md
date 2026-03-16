# Dormitory Form Automator (Batch Fill)

A Python-based automation tool designed to batch-fill student accommodation forms from Excel data into Word templates. It is specifically optimized for templates containing **two forms per page** (e.g., for A4 printing efficiency).

## ✨ Key Features

* **Batch Processing**: Seamlessly handle hundreds of student records in seconds.
* **Smart Pairing**: Automatically groups two students into a single Word document to save paper and printing costs.
* **Data Cleaning**: Built-in Regex logic to extract clean numbers for Buildings, Rooms, and Bed slots from messy input strings.
* **Style Retention**: Preserves table formatting, including font size (10.5 pt) and paragraph alignment (Left).
* **Robust Error Handling**: Automatically skips incomplete records and provides detailed error logs.

## 🛠️ Requirements

Ensure you have Python 3.x installed. You will need the following libraries:

```bash
pip install pandas python-docx openpyxl

```

## 📂 Project Structure

```text
.
├── main.py              # Main execution script
├── data_sample.xlsx     # Example Excel file (Template)
├── doc.docx             # Word template file
└── Output/              # Directory where generated files are saved

```

## 📖 Usage Guide

### 1. Prepare your Excel (`data.xlsx`)

Your Excel file must contain the following columns:

* `学工号` (Student ID)
* `姓名` (Name)
* `楼栋` (Building)
* `房间` (Room)
* `床位` (Bed Number)

### 2. Configure the Word Template (`doc.docx`)

* The template must contain **at least two tables**.
* The script searches for specific keywords within the cells: `姓  名`, `楼  号`, `房间号`, and `床位号`.
* *Note: Ensure the spacing in your template matches the strings in the script.*

### 3. Run the Script

Execute the following command in your terminal:

```bash
python main.py

```

## ⚙️ How it Works

1. **Grouping**: The script reads the Excel rows and groups them in pairs ($i, i+1$).
2. **Cleaning**: Functions like `clean_building_number` use Regular Expressions to strip unnecessary text (e.g., converting "Building No. 5" to "5").
3. **Filling**: It locates the target cells in the Word tables and injects the cleaned data while maintaining the specified font style.

## ⚠️ Privacy & Security

* **Do NOT upload `data.xlsx**` if it contains real student information (IDs, names, etc.).
* It is highly recommended to use the provided `data_sample.xlsx` for testing purposes.
* The `Output/` directory is ignored by default to prevent accidental data leaks.
