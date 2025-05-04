# 📊 ExcelBatchMerge

A privacy-respecting batch Excel (.xlsx) combiner tool for automating data aggregation.

## 🧠 What is ExcelBatchMerge?

ExcelBatchMerge is a simple yet effective Python-based tool that scans a directory for Excel `.xlsx` files, then merges them into a single workbook. Users are prompted for basic inputs like the output file name, save location, and whether to append to an existing worksheet or create a new one.

No headers? No problem. (Yet.) This is a basic MVP meant to automate repetitive Excel tasks. Header validation and matching is a planned feature.

## 🚀 Features

* ✅ Batch scan and combine `.xlsx` files from a folder
* ✅ Prompted user interaction for naming and save location
* ✅ Choose between appending or creating a new worksheet
* ✅ No API calls — works entirely offline
* ✅ Fast, local automation — saves hours of manual copying

## 🔐 Privacy First

This tool runs 100% locally. It doesn't make API calls, connect to external services, or collect any user data.

## ⚙️ How It Works

1. **Scan** a folder for `.xlsx` files
2. **Prompt** user for:

   * Output filename
   * Save location
   * Whether to append or use new worksheet
3. **Merge** all found files into a single Excel workbook
4. **Save** the result to the specified location

## 🧰 Requirements

* Python 3.7+
* `openpyxl`
* `tkinter` (for GUI file/folder prompts)

## 🛠️ Future Improvements

* [ ] Add header validation and matching
* [ ] GUI frontend with progress bar
* [ ] Drag-and-drop support

## 📦 Use Cases

* Finance teams consolidating monthly reports
* Analysts combining survey data
* Researchers organizing batch data exports
* Anyone tired of copy-pasting Excel rows all day

## 🤓 Developer Notes

This is a time-saving side project for internal use, inspired by repetitive work. Keep it light, local, and fast.

---

Made with ☕, frustration, and automation love.

> "Let Excel do Excel things. Let Python handle the grunt work."
