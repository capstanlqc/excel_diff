# excel_diff
A tool to generate diff reports from two Excel files or two directories containing Excel files

## Dependencies
This tool depends on `openpyxl` and `xlrd`. Install these dependencies by running:
```bash
pip install openpyxl xlrd
```

The tool is available in two versions: console and GUI

## Console version
*Usage:*
```bash
python excel_diff.py [-h] --original ORIGINAL --modified MODIFIED --output OUTPUT --source SOURCE --target TARGET [--extra_column EXTRA_COLUMN] [--extra_header EXTRA_HEADER] [--row-offset ROW_OFFSET] [--realign REALIGN] [--tolerate TOLERATE] [--nocap]
                     [--dir] [--wspattern WSPATTERN] [--omit_identical]
```
| Mandatory arguments | Explanation |
|---------------------|-------------|
| `--original` PATH   | Path to original file or directory |
| `--modified` PATH   | Path to modified file or directory |
| `--source` LETTER   | Source text column (e.g., B) |
| `--target` LETTER   | Target text column (e.g., C) |
| `--output` PATH     | Output path for the HTML report file |

| Optional agruments            | Explanation |
|-------------------------------|-------------|
| `--row-offset` INTEGER        | Rows to skip (e.g., header rows) |
| `--wspattern` PATTERN         | Regex pattern to match worksheet names |
| `--dir`                       | Directory mode - compare folders, otherwise compare a pair of Excel files |
| `--extra_column` LETTER       | Include text from another column in the compared Excel files |
| `--extra_header` STRING       | Extra column header in the HTML report |
| `--realign` INTEGER           | Search for matching source text the specified numbers of rows above and below the current row |
| `--tolerate` INTEGER          | Accept source text up to the specified number (%) different from the current source |
| `--nocap`                     | Removes caps from `realign` (`15`) and `tolerate` (`35`)<br>Without it any value higher than cap is lowered to cap |
| `--omit_identical`            | Do not include misaligned rows or rows with changed source but with identical target |
| `-h` \| `--help`              | Show help and exit |

## GUI version
*Usage:*
```bash
python excel_diff_gui.py
```
GUI version depends on the console version.
All of the above options can be set in a simple GUI.
<img width="833" height="739" alt="image" src="https://github.com/user-attachments/assets/b82534da-6704-43a0-9c26-013509257936" />

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
