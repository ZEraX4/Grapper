# Grapper

Grapper is a modern AI Slop (Antigravity), PyQt6-based fuzzy search GUI application designed for searching text across files and Office documents (Word and Excel). It features a powerful search engine with regex support, syntax highlighting, and a sleek user interface.

![Grapper Icon](icon.png)

## Features

- **Fuzzy Search**: Quickly find files based on partial or approximate matches.
- **Regex Support**: Use regular expressions for complex search queries.
- **Office Document and PDF Search**: Integration with `python-docx`, `openpyxl`, and `pypdf` to search within `.docx`, `.xlsx`, and `.pdf` files.
- **Syntax Highlighting**: Built-in support for multiple syntax highlighting themes using `pygments`.
- **External Editor Integration**: Open search results directly in your favorite text editor.
- **Dark/Light Themes**: Toggle between dark and light modes for optimal viewing.
- **Regex Designer**: Interactive tool for testing and refining your regex patterns.
- **Reveal in Explorer**: Quickly access the containing folder of any file.

## Requirements

- Python 3.x
- PyQt6
- pygments
- rapidfuzz
- python-docx
- openpyxl
- pypdf

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/ZEraX4/Grapper.git
   cd Grapper
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python main.py
   ```
2. Select a source directory using the folder icon.
3. Enter your search query in the sidebar.
4. Use the "Office" or "Regex" checkboxes to customize your search.
5. Click "Search" to view results in the table.
6. Select a result to view its content with syntax highlighting.

## License

This project is licensed under the MIT License - see the LICENSE file for details (if applicable).
