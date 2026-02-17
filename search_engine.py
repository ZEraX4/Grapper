import os
from rapidfuzz import process, fuzz
import re

class SearchEngine:
    def __init__(self):
        self.common_excludes = {'.git', 'node_modules', '__pycache__', 'venv', '.idea', '.vscode', 'dist', 'build'}

    def is_text_file(self, filepath):
        """Simple check to avoid reading binary files."""
        # Check size first (limit to 1MB)
        try:
            if os.path.getsize(filepath) > 1024 * 1024:
                return False
        except OSError:
            return False

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                f.read(1024)
            return True
        except (UnicodeDecodeError, Exception):
            return False

    def _extract_text_from_docx(self, filepath):
        """Extracts text from a .docx file."""
        try:
            import docx
            doc = docx.Document(filepath)
            # Combine all paragraphs into a list of lines
            lines = [p.text for p in doc.paragraphs if p.text.strip()]
            return lines
        except Exception as e:
            print(f"Error extracting Word text: {e}")
            return []

    def _extract_text_from_xlsx(self, filepath):
        """Extracts text from a .xlsx file."""
        try:
            import openpyxl
            wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
            lines = []
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    row_text = " ".join([str(cell) for cell in row if cell is not None])
                    if row_text.strip():
                        lines.append(row_text)
            return lines
        except Exception as e:
            print(f"Error extracting Excel text: {e}")
            return []

    def _extract_text_from_pdf(self, filepath):
        """Extracts text from a .pdf file."""
        try:
            from pypdf import PdfReader
            reader = PdfReader(filepath)
            lines = []
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    # Split into lines and filter out empty ones
                    lines.extend([line for line in text.splitlines() if line.strip()])
            return lines
        except Exception as e:
            print(f"Error extracting PDF text: {e}")
            return []

    def search(self, directory, query, stop_event=None, threshold=60, update_callback=None, use_regex=False, search_office=False, case_sensitive=False, limit_per_file=0):
        """
        Walks the directory and searches for the query in text files.
        Yields results as they are found.
        """
        if not query or not directory:
            return

        # 1. Pre-scan to count total files for progress bar
        total_files = 0
        if update_callback:
            update_callback(("Scanned", 0, 0)) # Signal start
            for root, dirs, files in os.walk(directory):
                if stop_event and stop_event.is_set():
                    return
                dirs[:] = [d for d in dirs if d not in self.common_excludes]
                total_files += len(files)
            update_callback(("Total", 0, total_files))

        file_count = 0
        for root, dirs, files in os.walk(directory):
            # Check for cancellation
            if stop_event and stop_event.is_set():
                break

            # Modify dirs in-place to skip excluded directories
            dirs[:] = [d for d in dirs if d not in self.common_excludes]
            
            for file in files:
                if stop_event and stop_event.is_set():
                    break

                file_count += 1
                if update_callback and (file_count % 5 == 0 or file_count == total_files):
                    update_callback(("Progress", file_count, total_files))

                filepath = os.path.join(root, file)
                
                is_office = False
                lines = []

                if search_office:
                    lower_file = file.lower()
                    if lower_file.endswith('.docx'):
                        lines = self._extract_text_from_docx(filepath)
                        is_office = True
                    elif lower_file.endswith('.xlsx'):
                        lines = self._extract_text_from_xlsx(filepath)
                        is_office = True
                    elif lower_file.endswith('.pdf'):
                        lines = self._extract_text_from_pdf(filepath)
                        is_office = True

                if not is_office:
                    if not self.is_text_file(filepath):
                        continue

                try:
                    if not is_office:
                        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                            lines = f.readlines()
                    
                    if not lines:
                        continue

                    matches = []
                    if use_regex:
                        try:
                            flags = 0 if case_sensitive else re.IGNORECASE
                            pattern = re.compile(query, flags)
                            for i, line in enumerate(lines):
                                for match in pattern.finditer(line):
                                    matches.append((line.strip(), 100, i, match.span()))
                                    if limit_per_file > 0 and len(matches) >= limit_per_file:
                                        break
                                if limit_per_file > 0 and len(matches) >= limit_per_file:
                                    break
                        except re.error:
                            pass # Invalid regex
                    else:
                        processor = None if case_sensitive else lambda x: x.lower()
                        
                        # Set limit based on limit_per_file. If 0, use None (rapidfuzz default is 10)
                        fuzzy_limit = limit_per_file if limit_per_file > 0 else None
                        
                        matches = process.extract(
                            query, 
                            lines, 
                            scorer=fuzz.partial_ratio, 
                            limit=fuzzy_limit,
                            score_cutoff=threshold,
                            processor=processor
                        )
                    
                    if matches:
                        best_score = matches[0][1]
                        yield {
                            'path': filepath,
                            'filename': file,
                            'score': best_score,
                            'matches': matches
                        }

                except Exception as e:
                    print(f"Error reading {filepath}: {e}")
                    continue
