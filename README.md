# Prefix-Suffix Adder for contextual menu

This simple Python script allows you to add custom prefixes and suffixes to file names via the Windows context menu. It provides an easy way to rename files directly from the right-click menu in Windows Explorer.

## Features

- Add predefined prefixes or suffixes to file names according to a list of options.
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
Kill and Restart explorer.exe (optional)
```bash
taskkill /f /im explorer.exe && start explorer.exe
```

## Usage

Once installed, you can right-click on a file in Windows Explorer and select "Add Prefix-Suffix". From there, you can choose a prefix or suffix to apply to the selected file.

## Modify options list

Just modify the following list in the code and uninstall and install the script.
```python
self.prefix_options = ["+Book+year+", "+Paper+year+", "+Thesis+year+", "+Report+year+", 
                        "+Slides+year+", "+Presentation+year+", "+Draft+year+"]
self.suffix_options = ["+authors"] 
```

## Requirements

- Python 3.x
- Windows 10/11
- Administrator privileges (required for installing/uninstalling context menu entries)

## Notes
- The script automatically elevates privileges when needed for installation or uninstallation.
- File name conflicts are resolved by appending a counter to the new file name (e.g., example.pdf → +Book+year+example (1).pdf).

## License

This project is licensed under the GNU General Public License v3.0. See the LICENSE file for details.

