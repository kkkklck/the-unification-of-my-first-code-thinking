# The Unification v6

A Python tool for construction data processing.  
This script extracts measurement data from Word documents, categorizes structural components (steel columns, beams, supports), and generates Excel reports with automatic formatting, instrument detection, and multi-day distribution.  
It is designed to improve efficiency in construction documentation.

---

## Features
- 📄 Extracts data from Word templates  
- 🏗️ Categorizes structural components: **Steel Columns**, **Beams**, **Supports**  
- 📊 Generates Excel reports with preserved formatting  
- 🔍 Detects instrument models automatically  
- 🗓️ Handles floor segmentation and date bucketing  
- ⚡ Improves reporting speed and reduces manual workload

---

## Repository Structure
.
├── The Unification v2.py # Main Python script
├── eg.docx # Example Word input template
├── 防火 2 无支撑版.xlsx # Example Excel output (no supports)
├── 防火 2 有支撑版.xlsx # Example Excel output (with supports)



---

## Installation
1. Clone this repository:
   ```bash
   git clone https://github.com/your-username/the-unification-of-my-first-code-thinking.git
   cd the-unification-of-my-first-code-thinking
Make sure you have Python 3.6+ installed.
Install dependencies:


pip install -r requirements.txt
(If no requirements.txt, install openpyxl and python-docx manually.)


pip install openpyxl python-docx
Usage
Place your Word measurement files in the working directory.

Run the script:


python "The Unification v2.py"
Follow the on-screen prompts:

Select Excel template (with or without supports)

Confirm component type (column, beam, support)

Choose the number of pages to generate

The filled Excel file will be saved automatically with all formatting preserved.

Example
Input: eg.docx

Output: 防火 2 有支撑版.xlsx / 防火 2 无支撑版.xlsx

License
This project is licensed under the MIT License.
