# 🚀 RapidPaste Automation  

**RapidPaste Automation** is a bot designed to automate **text extraction** from PDF and Word documents and format it systematically in a new Word file. This tool is particularly useful for processing **MCQs, structured documents, or text-based datasets**, efficiently pasting extracted content into a formatted table.

---

## 📌 Features  

✔ **Extract text from PDFs** using `PyPDF2` for structured data retrieval.  
✔ **Process Word documents** by interacting with system-level applications (`win32com.client`).  
✔ **Identify text patterns** to determine structured text ranges for MCQs and text extraction.  
✔ **Automated Copy-Paste** – Extracted data is systematically pasted into a formatted Word table.  
✔ **Batch Processing** – Handles multiple documents efficiently.  

---

## 📂 Folder Structure  
RapidPaste │── 1_DocFile/ # Processed documents (Output) │── 2_With_Bullets/ # Original files with bullets │── 3_Bullets_Sequence/ # Bullet-based sequence files │── 4_Clean_Bullets/ # Cleaned bullet points │── 5_Formatted_Bullets/ # Fully formatted bullet files │── 6_Formatted_Sequence/ # Fully formatted sequence files │── 7_Question_Range/ # Identified question text range │── 8_Options_Range/ # Extracted options text range │── 9_Order_Sequence/ # Ordered extracted text │── 10_Arrangement_Done/ # Arranged content for final processing │── 12_Fast_doc/ # Fast processing folder │── 12_Final_doc/ # Final formatted documents │── 14_Reorder_Answer_Key/ # Reordered answers │── main.ipynb # Core automation script │── Total_Options.txt # Extracted options │── Total_Questions.txt # Extracted questions │── Total_Combined.txt # Merged output text │── xoom_answer.ipynb # Additional processing script │── zx_Delete.ipynb # Utility script

yaml
Copy
Edit

---

## 🛠️ Technologies Used  
- **Python** – Core programming language  
- **PyPDF2** – PDF text extraction  
- **win32com.client** – System-level interaction with MS Word  
- **OS Module** – File handling and automation  
- **Pattern Recognition** – Identifying MCQ/question structure for range extraction  

---

## 📖 How It Works  
1️⃣ **Load the document** – PDF or Word file is selected.  
2️⃣ **Extract text** – Uses `PyPDF2` for PDFs and `win32com.client` for Word files.  
3️⃣ **Identify patterns** – Detects MCQs, numbered sequences, or structured content.  
4️⃣ **Define extraction range** – Determines where questions and answers are located.  
5️⃣ **Automate copy-paste** – Extracted text is formatted and pasted into a new Word file.  
6️⃣ **Save and store** – The structured document is saved in the `1_DocFile/` output folder.  

---

## 📥 Installation & Usage  

### 1️⃣ Clone the Repository  
```bash
git clone https://github.com/uttamofficial/RapidPaste-automation.git
cd RapidPaste-automation
2️⃣ Install Dependencies
bash
Copy
Edit
pip install PyPDF2 pywin32
3️⃣ Run the Script
bash
Copy
Edit
python main.py
For Jupyter Notebook users:
Run main.ipynb step by step.

💡 Future Enhancements
✅ GUI Interface – Making it user-friendly with a simple interface.
✅ Enhanced Pattern Recognition – More adaptable to different document formats.
✅ Cloud Storage Support – Save outputs to Google Drive or Dropbox.
✅ Multi-language Support – Extend compatibility for various languages.

📜 License
This project is open-source and available under the MIT License.

📩 Contact & Contributions
💬 Found a bug or have suggestions? Feel free to open an issue or submit a pull request!

📧 Contact: uttamofficial005@gmail.com
🔗 GitHub: [RapidPaste Automation](https://github.com/uttamofficial/RapidPaste-automation.git)

