# excel_diff
A tool to generate diff reports from two Excel files or two directories containing Excel files

## Dependencies
This tool depends on `openpyxl`. Install it by running:
```bash
pip install openpyxl
```

The tool is available in two versions: console and GUI

## Console version
*Usage:*
```bash
python excel_diff.py [-h] --original ORIGINAL --modified MODIFIED --wspattern WSPATTERN --source SOURCE --target TARGET [--row-offset ROW_OFFSET] [--word] [--char] [--output OUTPUT] [--dir]
```
| Mandatory arguments | Explanation |
|---------------------|-------------|
| `--original` PATH   | Path to original file or directory |
| `--modified` PASS   | Path to modified file or directory |
| `--wspattern` PATTERN | Regex pattern to match worksheet names |
| `--source` LETTER   | Source text column (e.g., B) |
| `--target` LETTER   | Target text column (e.g., C) |
| `--row-offset` INTEGER | Rows to skip (e.g., header rows) |

| Optional agruments   | Explanation |
|----------------------|-------------|
| `--word` \| `--char` | Diff mode by word or by character (default: `word`) |
| `--output` PATH      | Output path for the HTML report file (default: `diff_report.html` in the script dir) |
| `--dir`              | Directory mode - compare folders, otherwise compare a pair of Excel files |
| `-h` \| `--help`     | Show help and exit |

## GUI version
*Usage:*
```bash
python excel_diff_gui.py
```
All of the above options can be set in a simple GUI.
<img width="759" height="574" alt="image" src="https://github.com/user-attachments/assets/12cd6168-5fbb-4908-a433-804d33e20473" />


## Standalone GUI executable
The GUI version can be built into a platform-specific executable that can be run without installing Python or dependencies.
To build, install PyInstaller:
```bash
pip install pyinstaller
```
and then run:

*Windows:*
```bash
pyinstaller ExcelDiffGuiWin.spec
```
*macOS:*
```bash
python -m PyInstaller ExcelDiffGuiMac.spec
```

To download Windows and macOS versions, see the [Releases](https://github.com/capstanlqc/excel_diff/releases) section.
