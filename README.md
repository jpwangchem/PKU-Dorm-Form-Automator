# 🏢 PKU Dormitory Form Automator

An automation tool designed to batch-process student dormitory data from Excel and generate standardized Word vouchers. Perfect for administrative staff managing dormitory check-ins at Peking university.

---

## 🌟 Key Features

* **Smart Pairing**: Automatically groups two students into a single Word document to optimize A4 printing.
* **Robust Regex Cleaning**:
* **Buildings**: Extracts digits and immediate suffixes (e.g., "燕园1楼" -> "1楼").
* **Rooms/Beds**: Aggressively strips non-numeric characters (e.g., "3号床" -> "3") for a cleaner look on the voucher.


* **Literal Keyword Matching**: Specifically tuned to match template keywords like `姓  名` and `楼  号` (including the double spaces often found in Chinese forms).
* **Intelligent Naming**: Output files are named using the format `Student1ID-Name&Student2ID-Name.docx`, making them easy to search and distribute.

## 🚀 Quick Start

### 1. Prerequisites

Ensure you have Python 3.12 installed:

```bash
pip install pandas python-docx openpyxl

```

### 2. Configuration

* **Excel (`data.xlsx`)**: Must contain columns: `学工号`, `姓名`, `楼栋`, `房间`, `床位`.
* **Template (`doc.docx`)**: Please fill in the `院系`, `年级`, `经办人`, `时间` before run the python code.



### 3. Execution

```bash
python main.py

```

## 📊 Data Mapping & Processing

The tool doesn't just copy-paste; it "cleans" the data using the following logic:

| Source (Excel) | Target (Word) | Regex Logic | Result Example |
| --- | --- | --- | --- |
| **姓名** | `姓  名` | String Strip | " 张三 " → "张三" |
| **楼栋** | `楼  号` | `(\d+[^\d\s]?)` | "燕园1楼" → "1楼" |
| **房间** | `房间号` | `\d+` | "102室" → "102" |
| **床位** | `床位号` | `\d+` | "3号床" → "3" |

## 📂 Project Structure

```text
.
├── main.ipynb           # Main script with Regex and Docx logic
├── data.xlsx            # Your source data (Keep this private!)
├── doc.docx             # The Word template with 2+ tables
└── 住宿凭单/              # Default output directory (Auto-created after run the code)

```

## ⚙️ Technical Implementation Details

* **Pairing Logic**: Uses Python list slicing `[df.iloc[i:i + 2]]` to iterate through the Excel rows in steps of 2.
* **Font Control**: Forcibly sets the injected text to **10.5pt** (Size 5 in Chinese font systems) and **Left Alignment** to ensure the voucher looks professional regardless of Excel's original formatting.
* **Data Validation**: Checks for `NaN` or empty strings in required fields and skips the entire pair if any critical data is missing to prevent half-filled vouchers.
