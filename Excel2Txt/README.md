# Excel2Txt Tool

[![Language: VBA](https://img.shields.io/badge/Language-VBA-green.svg)](https://learn.microsoft.com/en-us/office/vba/api/overview/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## English

### 📝 Description
A professional Excel VBA tool designed to recursively scan an input directory (including subfolders) and export all Excel sheets into text files. The tool perfectly replicates the original folder structure in the output directory, making it ideal for performing **Grep** searches using text editors like Sakura Editor or VS Code.

### 🚀 Key Features
* **Recursive Processing**: Automatically handles complex subfolder structures.
* **High Performance**: Uses Array-based data processing for fast conversion.
* **Clean Structure**: One text file per Excel sheet, named as `[FileName]_[SheetName].txt`.
* **Robustness**: Built using Form Controls for maximum compatibility with Office 2021/2026/365, avoiding ActiveX issues.

### 🛠 How to Use
1. Open `Excel2Txt.xlsm`.
2. Enter the **Input Folder** and **Output Folder** paths in the "Excel2Txt" worksheet.
3. Click the **Start Conversion** button.
4. Check the output folder for the generated `.txt` files.

---

## 简体中文

### 📝 工具简介
这是一个专业的 Excel VBA 工具，用于递归扫描输入目录（包括子文件夹）并将所有 Excel 工作表导出为文本文件。该工具会在输出目录中完美还原原始文件夹结构，非常适合使用 Sakura Editor 或 VS Code 等文本编辑器进行 **Grep** 关键字检索。

### 🚀 核心功能
* **递归处理**: 自动遍历所有子层级文件夹。
* **高性能**: 采用数组处理技术，大幅提升大批量文件的转换速度。
* **结构清晰**: 每个工作表导出为一个文本文件，命名规则为 `[文件名]_[工作表名].txt`。
* **高兼容性**: 使用窗体控件（Form Controls）代替 ActiveX，全面支持 Office 2021/2026/365，避免安全禁用风险。

### 🛠 使用方法
1. 打开 `Excel2Txt.xlsm`。
2. 在 "Excel2Txt" 工作表的指定位置输入 **输入文件夹** 和 **输出文件夹** 路径。
3. 点击 **开始转换** 按钮。
4. 转换完成后，在输出文件夹中查看生成的 `.txt` 文件。

---

## 日本語

### 📝 概要
入力フォルダ配下（サブフォルダを含む）の全 Excel ファイルをスキャンし、各シートをテキストファイルとして書き出す VBA ツールです。出力先には元のフォルダ構造がそのまま再現されるため、サクラエディタ等のテキストエディタで **Grep 検索** を行う際に非常に便利です。

### 🚀 主な機能
* **再帰処理**: サブフォルダ内のファイルも自动で処理します。
* **高速化**: 配列を利用したデータ処理により、大量のデータも高速に変換します。
* **整理された出力**: `[ファイル名]_[シート名].txt` 形式で出力されます。
* **高い互換性**: ActiveX を排除し、フォームコントロールを採用しているため、最新の Office 2021/2026/365 でも安定して動作します。

### 🛠 使い方
1. `Excel2Txt.xlsm` を開きます。
2. 「Excel2Txt」シートの指定セルに **入力フォルダ** と **出力フォルダ** のパスを入力します。
3. 「変換開始」ボタンを押下します。
4. 出力フォルダにテキストファイルが生成されたことを確認します。

---

## Project Logic


## License
This project is licensed under the MIT License.
