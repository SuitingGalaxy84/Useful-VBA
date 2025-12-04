# Useful-VBA

A collection of useful VBA macros for Microsoft Word document formatting and editing, particularly focused on tables, captions, headings, and table of contents formatting.

## 📋 Table of Contents

- [Scripts Overview](#scripts-overview)
- [Installation](#installation)
- [Usage](#usage)
- [Script Details](#script-details)

## Scripts Overview

| Script | Description |
|--------|-------------|
| **CenterAllTablesOnPage** | Centers all tables on the page |
| **CenterAndResetIndentForCaptions** | Centers and removes indents from figure and table captions |
| **ChangeTextAfterFigureFieldToSimHei** | Formats text after figure fields with SimHei font |
| **ChangeTextAfterTableFieldToSimHei** | Formats text after table fields with SimHei font |
| **ConvertFigureListToText** | Converts figure list numbering to plain text |
| **ConvertTableListToText** | Converts table list numbering to plain text |
| **ForceAllHeadingsToHeiti_ByScan** | Forces all headings to use Heiti/SimHei fonts |
| **ForceTOCStyles** | Forces table of contents styles to use FangSong font |
| **FormatTableFonts** | Formats table fonts (body: SimHei, headers: bold) |
| **RemoveAllIndentsInTables** | Removes all indents from table content |
| **ReplaceSEQIdentifiers** | Replaces Chinese SEQ identifiers with English ones |

## Installation

1. Open Microsoft Word
2. Press `Alt + F11` to open the VBA Editor
3. Go to `File > Import File...`
4. Select the desired `.bas` file from this repository
5. Close the VBA Editor

## Usage

### Running a Macro

1. Press `Alt + F8` to open the Macro dialog
2. Select the macro you want to run
3. Click `Run`

### Adding to Quick Access Toolbar (Recommended)

1. Right-click the Quick Access Toolbar
2. Select `Customize Quick Access Toolbar...`
3. Choose `Macros` from the dropdown
4. Add your frequently used macros

## Script Details

### 🎯 Table Formatting

#### CenterAllTablesOnPage
Centers all tables horizontally on the page by:
- Resetting left indent to 0
- Setting row alignment to center

**Use case:** Quickly center all tables in a document for consistent formatting.

#### RemoveAllIndentsInTables
Removes all paragraph indents within tables:
- Left indent
- Right indent
- First line indent
- Character unit indents

**Use case:** Clean up table formatting when content has unwanted indents.

#### FormatTableFonts
Applies consistent font formatting to all tables:
- **Body cells:** SimHei (Asian), Times New Roman (ASCII), 10.5pt
- **Header row:** SimHei, Times New Roman, 10.5pt, Bold

**Use case:** Standardize table appearance across the entire document.

---

### 🏷️ Caption Formatting

#### CenterAndResetIndentForCaptions
Processes all SEQ fields (Figure/Table captions) and:
- Removes all indents (left, right, first line, character units)
- Centers the caption paragraph

**Use case:** Ensure all figure and table captions are centered without indents.

#### ChangeTextAfterFigureFieldToSimHei
Formats text following figure SEQ fields:
- Asian font: SimHei (黑体)
- ASCII font: Times New Roman
- Font size: 12pt

**Use case:** Apply consistent styling to figure caption text.

#### ChangeTextAfterTableFieldToSimHei
Formats text following table SEQ fields:
- Asian font: SimHei (黑体)
- ASCII font: Times New Roman

**Use case:** Apply consistent styling to table caption text.

---

### 📝 Heading & TOC Formatting

#### ForceAllHeadingsToHeiti_ByScan
Scans all paragraphs with outline levels (headings) and forces font to:
- Asian font: Heiti (黑体)
- ASCII font: SimHei
- Color: Auto (black)

**Use case:** Ensure all document headings use the correct font style.

#### ForceTOCStyles
Updates Table of Contents formatting:
- Refreshes the TOC
- Applies FangSong (Asian) and Times New Roman (ASCII) fonts to TOC entries

**Use case:** Standardize TOC appearance after document updates.

---

### 🔄 List & SEQ Field Utilities

#### ConvertFigureListToText
Converts numbered figure lists (containing "图") to plain text.

**Use case:** Convert dynamic list numbers to static text for final documents.

#### ConvertTableListToText
Converts numbered table lists (containing "表") to plain text.

**Use case:** Convert dynamic list numbers to static text for final documents.

#### ReplaceSEQIdentifiers
Replaces Chinese SEQ field identifiers with English equivalents:
- "SEQ 图" → "SEQ Figure"
- "SEQ 表格" → "SEQ Table"

**Use case:** Standardize caption field codes for cross-language compatibility.

---

## 📌 Notes

- All scripts include screen updating optimization for faster execution
- Scripts are designed to work with Chinese and English text
- Most scripts provide feedback via message box upon completion
- Scripts handle merged cells safely (where applicable)

## 🤝 Contributing

Feel free to submit issues or pull requests to improve these scripts.

## 📄 License

These scripts are provided as-is for personal and professional use.
