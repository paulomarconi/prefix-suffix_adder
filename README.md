# Prefix-Suffix Adder

This Python script allows you to add custom prefixes and suffixes to file names via the Windows context menu. It provides an easy way to rename files directly from the right-click menu in Windows Explorer.

## Features

- Add predefined prefixes or suffixes to file names.
- Install or uninstall the context menu entries for quick access.
- Automatically handles file name conflicts by appending a counter to the new file name.

## Installation

To install the context menu entries, run the following command:

```bash
python prefix-suffix_adder.py install
```

## Uninstallation

To uninstall the context menu entries, run the following command:

```bash
python prefix-suffix_adder.py uninstall
```

## Requirements

- Python 3.x
- Windows 10/11
- Administrator privileges (required for installing/uninstalling context menu entries)

## Notes
- The script automatically elevates privileges when needed for installation or uninstallation.
- File name conflicts are resolved by appending a counter to the new file name (e.g., example.txt → +Book+year+example (1).txt).

## License

This project is licensed under the GNU General Public License v3.0. See the LICENSE file for details.

