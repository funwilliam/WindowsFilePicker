# Windows File Picker

A Python utility to create Windows dialogs for selecting files or folders, supporting single and multiple file selection.

## Features
- Supports selecting files, folders, and multiple files.
- Uses native Windows dialog APIs.
- Handles both file and folder paths with ease.

## Requirements
- Windows OS
- Python 3.7 or above

## Installation
Copy `file_picker.py` into your project.

## Usage
```python
from file_picker import select_items

# Select a folder
selected_folder = select_items(select_mode='folder')
print(selected_folder)

# Select a file
selected_file = select_items(select_mode='file')
print(selected_file)

# Select multiple files
selected_files = select_items(select_mode='multi-files')
print(selected_files)
