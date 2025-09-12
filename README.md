# The Unification v6

A Python automation tool for consolidating fireproofing measurement data.
It extracts measurements from Word tables, classifies structural components,
and writes a fully formatted Excel report. A companion Word file is also
created for easy verification.

---

## Features
- **Word table parsing** – reads `.docx` documents that contain tables with
  the keywords `"测点 1"` and `"平均值"`.
- **Component categorization** – supports *steel columns*, *steel beams*,
  *braces*, and *space frames*; unrecognized entries fall back to **Others**.
- **Excel generation** – fills the template `防火excel模板μ.xlsx` and removes
  unused worksheets automatically. Output is named
  `The Unification_报告版.xlsx` and incremented if a file with the same name
  already exists.
- **Instrument auto‑detection** – determines the model (`23-90` or `24-57`)
  from average values and formats the character `μ` using Times New Roman.
- **Multi‑day distribution** – buckets measurements by date with configurable
  strategies and sensible defaults. Overlapping rules favor later days and
  orphaned data can be appended to the last day.
- **Summary Word export** – generates `汇总原始记录.docx` alongside the source
  document for manual cross‑checking.
- **Cross‑platform paths** – works on Windows, macOS, and Linux. Remember to
  close Word and Excel files before running the script.

---

## Repository Structure
```
.
├── The Unification v6.py      # Main Python script
├── eg.docx                    # Example Word input
├── 防火excel模板μ.xlsx        # Excel template
└── LICENSE
```



---

## Installation
1. **Clone the repository**
   ```bash
   git clone https://github.com/your-username/the-unification-of-my-first-code-thinking.git
   cd the-unification-of-my-first-code-thinking
   ```
2. **Install Python 3.6+ (3.8+ recommended)**
3. **Install dependencies**
   ```bash
   pip install openpyxl python-docx
   # If the network is slow in China:
   pip install openpyxl python-docx -i https://pypi.tuna.tsinghua.edu.cn/simple
   ```

---
## Usage
1. Place your Word measurement files and the Excel template in the working
   directory.
2. Run the script:
   ```bash
   python "The Unification v6.py"
   ```
3. Follow the prompts to specify file paths, component type, date buckets,
   and the number of pages to generate.
4. The script outputs `The Unification_报告版.xlsx` and
   `汇总原始记录.docx` in the same folder as the source files.

---
Output: 防火 2 有支撑版.xlsx / 防火 2 无支撑版.xlsx

## License
This project is released under the [MIT License](LICENSE).
