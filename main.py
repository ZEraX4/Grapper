import sys
import os
import threading
import subprocess
import re
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLineEdit,
                             QSplitter, QFileDialog, QLabel, QPlainTextEdit,
                             QMessageBox, QProgressBar, QTextEdit, QSpinBox, QCheckBox,
                             QComboBox, QStyleFactory, QStackedWidget, QTableWidget, 
                             QTableWidgetItem, QHeaderView, QStyle, QDialog, 
                             QFormLayout, QTextBrowser)
from PyQt6.QtGui import QPalette, QColor, QBrush, QFont, QSyntaxHighlighter, QTextCharFormat, QTextCursor, QPainter, QTextFormat, QIcon
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject, QSettings

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


from search_engine import SearchEngine
from pygments import lex
from pygments.lexers import get_lexer_for_filename, TextLexer
from pygments.styles import get_style_by_name, get_all_styles
from pygments.token import Token
from PyQt6.QtGui import QPalette, QColor, QBrush

class PygmentsHighlighter(QSyntaxHighlighter):
    def __init__(self, document, style_name='default'):
        super().__init__(document)
        self.lexer = TextLexer()
        self.style = get_style_by_name(style_name)
        self._cache_formats()

    def _cache_formats(self):
        self.formats = {}
        for token, style in self.style:
            fmt = QTextCharFormat()
            if style['color']:
                fmt.setForeground(QColor(f"#{style['color']}"))
            if style['bgcolor']:
                fmt.setBackground(QColor(f"#{style['bgcolor']}"))
            if style['bold']:
                fmt.setFontWeight(QFont.Weight.Bold)
            if style['italic']:
                fmt.setFontItalic(True)
            if style['underline']:
                fmt.setFontUnderline(True)
            self.formats[token] = fmt

    def set_file(self, filepath):
        try:
            self.lexer = get_lexer_for_filename(filepath)
        except:
            self.lexer = TextLexer()
        self.rehighlight()

    def set_style(self, style_name):
        self.style = get_style_by_name(style_name)
        self._cache_formats()
        self.rehighlight()

    def highlightBlock(self, text):
        for index, token, value in self.lexer.get_tokens_unprocessed(text):
            length = len(value)
            # Find the best matching format for the token
            while token not in self.formats:
                token = token.parent
                if token is None:
                    break
            
            if token in self.formats:
                self.setFormat(index, length, self.formats[token])


class LineNumberArea(QWidget):
    def __init__(self, editor):
        super().__init__(editor)
        self.codeEditor = editor

    def sizeHint(self):
        return self.codeEditor.lineNumberAreaSizeHint()

    def paintEvent(self, event):
        self.codeEditor.lineNumberAreaPaintEvent(event)


class CodeEditor(QPlainTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.lineNumberArea = LineNumberArea(self)

        self.blockCountChanged.connect(self.updateLineNumberAreaWidth)
        self.updateRequest.connect(self.updateLineNumberArea)
        self.cursorPositionChanged.connect(self.highlightCurrentLine)

        self.updateLineNumberAreaWidth(0)

    def lineNumberAreaWidth(self):
        digits = 1
        max_value = max(1, self.blockCount())
        while max_value >= 10:
            max_value //= 10
            digits += 1

        # Use 10px base padding for more generous space
        space = 10 + self.fontMetrics().horizontalAdvance('9') * digits
        return space

    def lineNumberAreaSizeHint(self):
        return self.lineNumberAreaWidth(), 0

    def updateLineNumberAreaWidth(self, _):
        self.setViewportMargins(self.lineNumberAreaWidth(), 0, 0, 0)

    def updateLineNumberArea(self, rect, dy):
        if dy:
            self.lineNumberArea.scroll(0, dy)
        else:
            self.lineNumberArea.update(0, rect.y(), self.lineNumberArea.width(), rect.height())

        if rect.contains(self.viewport().rect()):
            self.updateLineNumberAreaWidth(0)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        cr = self.contentsRect()
        self.lineNumberArea.setGeometry(cr.left(), cr.top(), self.lineNumberAreaWidth(), cr.height())

    def lineNumberAreaPaintEvent(self, event):
        painter = QPainter(self.lineNumberArea)
        # Use AlternateBase for a subtle distinction from the main background
        bg_color = self.palette().color(QPalette.ColorRole.AlternateBase)
        painter.fillRect(event.rect(), bg_color)

        block = self.firstVisibleBlock()
        blockNumber = block.blockNumber()
        top = int(self.blockBoundingGeometry(block).translated(self.contentOffset()).top())
        bottom = top + int(self.blockBoundingRect(block).height())

        # Use Text color for the numbers
        text_color = self.palette().color(QPalette.ColorRole.Text)
        # Make it slightly more subtle
        text_color.setAlpha(160) 
        painter.setPen(text_color)

        while block.isValid() and top <= event.rect().bottom():
            if block.isVisible() and bottom >= event.rect().top():
                number = str(blockNumber + 1)
                # Subtract 5px padding from the right to ensure a gap from the border
                painter.drawText(0, top, self.lineNumberArea.width() - 5, self.fontMetrics().height(),
                                 Qt.AlignmentFlag.AlignRight, number)

            block = block.next()
            top = bottom
            bottom = top + int(self.blockBoundingRect(block).height())
            blockNumber += 1

    def highlightCurrentLine(self):
        extraSelections = []

        if True: # Always highlight current line, even if read-only
            selection = QTextEdit.ExtraSelection()
            # A more eye-pleasing, translucent blue color
            lineColor = QColor(42, 130, 218, 50) 
            selection.format.setFontWeight(QFont.Weight.Bold)
            selection.format.setBackground(lineColor)
            selection.format.setProperty(QTextFormat.Property.FullWidthSelection, True)
            selection.cursor = self.textCursor()
            selection.cursor.clearSelection()
            extraSelections.append(selection)

        self.setExtraSelections(extraSelections)


class SearchWorker(QThread):
    result_found = pyqtSignal(dict)
    finished_search = pyqtSignal()
    progress_update = pyqtSignal(object)
    error_occurred = pyqtSignal(str)

    def __init__(self, search_engine, directory, query, stop_event, threshold=60, use_regex=False, search_office=False):
        super().__init__()
        self.search_engine = search_engine
        self.directory = directory
        self.query = query
        self.stop_event = stop_event
        self.threshold = threshold
        self.use_regex = use_regex
        self.search_office = search_office

    def run(self):
        try:
            # Clear stop event before starting
            self.stop_event.clear()
            
            results = self.search_engine.search(
                self.directory,
                self.query,
                stop_event=self.stop_event,
                threshold=self.threshold,
                update_callback=self.progress_update.emit,
                use_regex=self.use_regex,
                search_office=self.search_office
            )

            for res in results:
                if self.stop_event.is_set():
                    break
                self.result_found.emit(res)
                
        except Exception as e:
            self.error_occurred.emit(str(e))
        finally:
            self.finished_search.emit()


class RegexDesignerDialog(QDialog):
    def __init__(self, parent=None, initial_pattern=""):
        super().__init__(parent)
        self.setWindowTitle("Regex Designer")
        self.setMinimumSize(600, 500)
        self.init_ui(initial_pattern)
        self.test_regex()

    def init_ui(self, initial_pattern):
        layout = QVBoxLayout(self)

        # Pattern Input
        pattern_layout = QHBoxLayout()
        pattern_layout.addWidget(QLabel("Regex Pattern:"))
        self.pattern_input = QLineEdit(initial_pattern)
        self.pattern_input.setPlaceholderText("Enter regex pattern...")
        self.pattern_input.textChanged.connect(self.test_regex)
        pattern_layout.addWidget(self.pattern_input)
        layout.addLayout(pattern_layout)

        # Text Input
        layout.addWidget(QLabel("Test Text:"))
        self.test_text = QTextEdit()
        self.test_text.setPlaceholderText("Enter sample text to test matches...")
        self.test_text.setPlainText("Hello World 123!\nThis is a sample text for regex testing.\nDate: 2026-02-11")
        self.test_text.textChanged.connect(self.test_regex)
        layout.addWidget(self.test_text)

        # Results / Info
        self.lbl_results = QLabel("Matches: 0")
        self.lbl_results.setStyleSheet("font-weight: bold; color: #2a82da;")
        layout.addWidget(self.lbl_results)

        # Cheat Sheet
        cheat_sheet = QTextBrowser()
        cheat_sheet.setOpenExternalLinks(True)
        cheat_sheet.setHtml(r"""
        <b>Quick Cheat Sheet:</b>
        <table border='0' cellspacing='5'>
            <tr><td><code>.</code></td><td>Any character</td><td><code>\d</code></td><td>Digit (0-9)</td></tr>
            <tr><td><code>*</code></td><td>0 or more</td><td><code>\w</code></td><td>Word char (a-z, 0-9)</td></tr>
            <tr><td><code>+</code></td><td>1 or more</td><td><code>\s</code></td><td>Whitespace</td></tr>
            <tr><td><code>?</code></td><td>0 or 1</td><td><code>[abc]</code></td><td>Any of a, b, c</td></tr>
            <tr><td><code>^</code></td><td>Start of line</td><td><code>$</code></td><td>End of line</td></tr>
        </table>
        """)
        cheat_sheet.setMaximumHeight(100)
        layout.addWidget(cheat_sheet)

        # Buttons
        btns_layout = QHBoxLayout()
        self.btn_use = QPushButton("Use This Regex")
        self.btn_use.setStyleSheet("background-color: #2a82da; color: white; font-weight: bold;")
        self.btn_use.clicked.connect(self.accept)
        
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        
        btns_layout.addStretch()
        btns_layout.addWidget(btn_cancel)
        btns_layout.addWidget(self.btn_use)
        layout.addLayout(btns_layout)

    def test_regex(self):
        pattern_str = self.pattern_input.text()
        text = self.test_text.toPlainText()
        
        # Block signals to avoid recursion during formatting updates
        self.test_text.blockSignals(True)
        try:
            # Reset formatting
            cursor = self.test_text.textCursor()
            cursor.select(QTextCursor.SelectionType.Document)
            fmt = QTextCharFormat()
            cursor.setCharFormat(fmt)
            
            if not pattern_str:
                self.lbl_results.setText("Matches: 0")
                return

            try:
                matches = list(re.finditer(pattern_str, text, re.IGNORECASE))
                self.lbl_results.setText(f"Matches: {len(matches)}")
                
                # Highlight matches
                for match in matches:
                    start, end = match.span()
                    cursor.setPosition(start)
                    cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
                    
                    highlight_fmt = QTextCharFormat()
                    highlight_fmt.setBackground(QColor(42, 130, 218, 100))
                    highlight_fmt.setFontWeight(QFont.Weight.Bold)
                    cursor.setCharFormat(highlight_fmt)
                    
            except re.error as e:
                self.lbl_results.setText(f"Invalid Regex: {str(e)}")
                self.lbl_results.setStyleSheet("color: #d32f2f; font-weight: bold;")
            else:
                self.lbl_results.setStyleSheet("color: #2a82da; font-weight: bold;")
        finally:
            self.test_text.blockSignals(False)

    def get_pattern(self):
        return self.pattern_input.text()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Grapper - Fuzzy Search (PyQt6)")
        self.setWindowIcon(QIcon(resource_path("icon.png")))
        self.setGeometry(100, 100, 1000, 700)

        self.search_engine = SearchEngine()
        self.stop_event = threading.Event()
        self.worker = None

        self.current_matches = []
        self.current_match_index = -1

        self.settings = QSettings("Grapper", "GrapperApp")
        # Initialize variables before UI
        self.editor_path = self.settings.value("editor_path", "")

        self.init_ui()
        self.load_settings()

    def init_ui(self):
        # Main Layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        # --- Sidebar ---
        sidebar_widget = QWidget()
        sidebar_layout = QVBoxLayout(sidebar_widget)
        sidebar_layout.setContentsMargins(5, 5, 5, 5)
        sidebar_layout.setSpacing(10)

        # 1. Directory Section
        dir_section = QVBoxLayout()
        lbl_dir_title = QLabel("Source Directory:")
        lbl_dir_title.setStyleSheet("font-weight: bold;")
        dir_section.addWidget(lbl_dir_title)
        
        dir_controls_layout = QHBoxLayout()
        self.lbl_directory = QLabel("No directory selected")
        self.lbl_directory.setWordWrap(True)
        self.lbl_directory.setStyleSheet("color: #aaa; font-style: italic;")
        dir_controls_layout.addWidget(self.lbl_directory, 1)

        btn_select_dir = QPushButton()
        btn_select_dir.setObjectName("btn_select_dir")
        btn_select_dir.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirIcon))
        btn_select_dir.setFixedSize(30, 30)
        btn_select_dir.setToolTip("Select Directory")
        btn_select_dir.clicked.connect(self.select_directory)
        dir_controls_layout.addWidget(btn_select_dir)
        dir_section.addLayout(dir_controls_layout)
        sidebar_layout.addLayout(dir_section)

        # 2. Search Section
        search_section = QVBoxLayout()
        lbl_search_title = QLabel("Search Query:")
        lbl_search_title.setStyleSheet("font-weight: bold;")
        search_section.addWidget(lbl_search_title)
        
        self.entry_search = QLineEdit()
        self.entry_search.setPlaceholderText("Enter terms or regex...")
        self.entry_search.returnPressed.connect(self.start_search)
        search_section.addWidget(self.entry_search)
        sidebar_layout.addLayout(search_section)

        # 3. Match Settings Section
        match_section = QVBoxLayout()
        lbl_match_title = QLabel("Match Settings:")
        lbl_match_title.setStyleSheet("font-weight: bold;")
        match_section.addWidget(lbl_match_title)
        
        match_params_layout = QHBoxLayout()
        lbl_threshold = QLabel("Min %:")
        self.spin_threshold = QSpinBox()
        self.spin_threshold.setRange(0, 100)
        self.spin_threshold.setValue(60)
        match_params_layout.addWidget(lbl_threshold)
        match_params_layout.addWidget(self.spin_threshold)
        self.chk_office = QCheckBox("Office")
        self.chk_regex = QCheckBox("Regex")
        match_params_layout.addWidget(self.chk_office)
        match_params_layout.addWidget(self.chk_regex)
        
        btn_regex_help = QPushButton()
        btn_regex_help.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MessageBoxInformation))
        btn_regex_help.setFixedSize(20, 20)
        btn_regex_help.setFlat(True)
        btn_regex_help.setToolTip("Regex Cheat Sheet")
        btn_regex_help.clicked.connect(self.show_regex_help)
        match_params_layout.addWidget(btn_regex_help)
        match_section.addLayout(match_params_layout)

        sidebar_layout.addLayout(match_section)

        # 4. Appearance Section
        appearance_section = QVBoxLayout()
        lbl_appearance_title = QLabel("Appearance:")
        lbl_appearance_title.setStyleSheet("font-weight: bold;")
        appearance_section.addWidget(lbl_appearance_title)
        
        theme_layout = QHBoxLayout()
        theme_layout.addWidget(QLabel("Theme:"))
        self.combo_app_theme = QComboBox()
        self.combo_app_theme.addItems(["Dark", "Light"])
        self.combo_app_theme.currentTextChanged.connect(self.apply_app_theme)
        theme_layout.addWidget(self.combo_app_theme)
        appearance_section.addLayout(theme_layout)

        syntax_layout = QHBoxLayout()
        syntax_layout.addWidget(QLabel("Syntax:"))
        self.combo_syntax_theme = QComboBox()
        self.styles = list(get_all_styles())
        self.styles.sort()
        self.combo_syntax_theme.addItems(self.styles)
        default_style = 'monokai' if 'monokai' in self.styles else 'default'
        self.combo_syntax_theme.setCurrentText(default_style)
        self.combo_syntax_theme.currentTextChanged.connect(self.change_syntax_theme)
        syntax_layout.addWidget(self.combo_syntax_theme)
        appearance_section.addLayout(syntax_layout)
        sidebar_layout.addLayout(appearance_section)

        # 5. Editor Section
        editor_section = QVBoxLayout()
        lbl_editor_title = QLabel("External Editor:")
        lbl_editor_title.setStyleSheet("font-weight: bold;")
        editor_section.addWidget(lbl_editor_title)
        
        editor_path_layout = QHBoxLayout()
        self.entry_editor_path = QLineEdit()
        self.entry_editor_path.setPlaceholderText("Path to editor...")
        self.entry_editor_path.setText(self.editor_path)
        self.entry_editor_path.setReadOnly(True)
        editor_path_layout.addWidget(self.entry_editor_path)
        
        btn_browse_editor = QPushButton()
        btn_browse_editor.setObjectName("btn_browse_editor")
        btn_browse_editor.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
        btn_browse_editor.setFixedSize(30, 30)
        btn_browse_editor.setToolTip("Browse for Editor")
        btn_browse_editor.clicked.connect(self.select_editor)
        editor_path_layout.addWidget(btn_browse_editor)
        editor_section.addLayout(editor_path_layout)
        sidebar_layout.addLayout(editor_section)

        # 6. Action Buttons
        buttons_layout = QHBoxLayout()
        self.btn_search = QPushButton("Search")
        self.btn_search.setObjectName("btn_search")
        self.btn_search.clicked.connect(self.start_search)
        buttons_layout.addWidget(self.btn_search)

        self.btn_stop = QPushButton("Stop")
        self.btn_stop.setObjectName("btn_stop")
        self.btn_stop.clicked.connect(self.stop_search)
        self.btn_stop.setEnabled(False)
        buttons_layout.addWidget(self.btn_stop)
        sidebar_layout.addLayout(buttons_layout)

        # Progress / Status
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #555;
                border-radius: 5px;
                text-align: center;
                height: 20px;
                background-color: #222;
            }
            QProgressBar::chunk {
                background-color: #2a82da;
                width: 20px;
            }
        """)
        sidebar_layout.addWidget(self.progress_bar)

        self.lbl_status = QLabel("Ready")
        self.lbl_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_status.setStyleSheet("color: #888; font-size: 9pt;")
        sidebar_layout.addWidget(self.lbl_status)

        # Results Table
        self.table_results = QTableWidget()
        self.table_results.setColumnCount(3)
        self.table_results.setHorizontalHeaderLabels(["Name", "Match %", "Path"])
        self.table_results.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.table_results.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        self.table_results.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        self.table_results.horizontalHeader().setStretchLastSection(True)
        self.table_results.setSortingEnabled(True)
        self.table_results.verticalHeader().setVisible(False)
        self.table_results.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table_results.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table_results.itemSelectionChanged.connect(self.load_file_from_selected_row)
        sidebar_layout.addWidget(self.table_results)

        sidebar_widget.setLayout(sidebar_layout)
        splitter.addWidget(sidebar_widget)

        # --- Main Content Area (File Viewer) ---
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(5, 5, 5, 5)

        content_widget.setLayout(content_layout)
        splitter.addWidget(content_widget)

        # --- Content Area Stack ---
        self.stack = QStackedWidget()
        content_layout.addWidget(self.stack)

        # 1. Placeholder Widget
        self.placeholder_widget = QWidget()
        placeholder_layout = QVBoxLayout(self.placeholder_widget)
        lbl_placeholder = QLabel("Select a file to view content")
        lbl_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_placeholder.setStyleSheet("color: gray; font-size: 16px;")
        placeholder_layout.addWidget(lbl_placeholder)
        self.stack.addWidget(self.placeholder_widget)

        # 2. File Viewer Widget (wrap existing controls)
        self.file_viewer_widget = QWidget()
        file_viewer_layout = QVBoxLayout(self.file_viewer_widget)
        file_viewer_layout.setContentsMargins(0, 0, 0, 0)

        self.lbl_filepath = QLabel("")
        file_viewer_layout.addWidget(self.lbl_filepath)

        # File Viewer Controls
        viewer_controls_layout = QHBoxLayout()
        viewer_controls_layout.addStretch()

        self.lbl_match_info = QLabel("")
        viewer_controls_layout.addWidget(self.lbl_match_info)
        
        self.btn_prev_match = QPushButton("Previous Match")
        self.btn_prev_match.setObjectName("btn_prev_match")
        self.btn_prev_match.clicked.connect(self.prev_match)
        self.btn_prev_match.setEnabled(False)
        viewer_controls_layout.addWidget(self.btn_prev_match)

        self.btn_next_match = QPushButton("Next Match")
        self.btn_next_match.setObjectName("btn_next_match")
        self.btn_next_match.clicked.connect(self.next_match)
        self.btn_next_match.setEnabled(False)
        viewer_controls_layout.addWidget(self.btn_next_match)
        
        self.btn_goto_match = QPushButton("Go to Match")
        self.btn_goto_match.setObjectName("btn_goto_match")
        self.btn_goto_match.clicked.connect(self.scroll_to_current_match)
        self.btn_goto_match.setEnabled(False)
        viewer_controls_layout.addWidget(self.btn_goto_match)
        
        self.btn_reveal_explorer = QPushButton("Reveal in Explorer")
        self.btn_reveal_explorer.setObjectName("btn_reveal_explorer")
        self.btn_reveal_explorer.clicked.connect(self.reveal_in_explorer)
        self.btn_reveal_explorer.setEnabled(False)
        viewer_controls_layout.addWidget(self.btn_reveal_explorer)
        
        self.btn_open_external = QPushButton("Open in Editor")
        self.btn_open_external.setObjectName("btn_open_external")
        self.btn_open_external.clicked.connect(self.open_in_external_editor)
        self.btn_open_external.setEnabled(False)
        viewer_controls_layout.addWidget(self.btn_open_external)
        
        file_viewer_layout.addLayout(viewer_controls_layout)

        self.text_editor = CodeEditor()
        self.text_editor.setReadOnly(True)
        font = QFont("Consolas", 10)
        self.text_editor.setFont(font)
        file_viewer_layout.addWidget(self.text_editor)

        # Syntax Highlighter
        self.highlighter = PygmentsHighlighter(self.text_editor.document())
        
        self.stack.addWidget(self.file_viewer_widget)
        
        # Default to placeholder
        self.stack.setCurrentWidget(self.placeholder_widget)

        # Set initial splitter sizes
        splitter.setSizes([300, 700])

    def select_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Directory")
        if directory:
            self.lbl_directory.setText(directory)

    def show_regex_help(self):
        initial_pattern = self.entry_search.text()
        dialog = RegexDesignerDialog(self, initial_pattern)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.entry_search.setText(dialog.get_pattern())
            self.chk_regex.setChecked(True)

    def select_editor(self):
        editor_path, _ = QFileDialog.getOpenFileName(self, "Select Editor Executable", "", "Executables (*.exe);;All Files (*)")
        if editor_path:
            self.editor_path = editor_path
            self.entry_editor_path.setText(editor_path)
            self.settings.setValue("editor_path", editor_path)

    def start_search(self):
        directory = self.lbl_directory.text()
        query = self.entry_search.text().strip()

        if directory == "No directory selected" or not os.path.exists(directory):
            QMessageBox.warning(self, "Warning", "Please select a valid directory.")
            return

        if not query:
            QMessageBox.warning(self, "Warning", "Please enter a search query.")
            return

        self.table_results.setSortingEnabled(False)
        self.table_results.setRowCount(0)
        self.text_editor.clear()
        self.lbl_filepath.setText("")
        self.stack.setCurrentWidget(self.placeholder_widget)
        self.btn_search.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.lbl_status.setText("Searching...")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)

        threshold = self.spin_threshold.value()
        use_regex = self.chk_regex.isChecked()
        search_office = self.chk_office.isChecked()
        self.worker = SearchWorker(self.search_engine, directory, query, self.stop_event, threshold, use_regex, search_office)
        self.worker.result_found.connect(self.add_result)
        self.worker.progress_update.connect(self.update_status)
        self.worker.error_occurred.connect(self.show_error)
        self.worker.finished_search.connect(self.search_finished)
        
        self.worker.start()

    def stop_search(self):
        if self.worker and self.worker.isRunning():
            self.stop_event.set()
            self.lbl_status.setText("Stopping...")

    def update_status(self, msg_data):
        if isinstance(msg_data, tuple):
            tag, current, total = msg_data
            if tag == "Scanned":
                self.lbl_status.setText("Counting files...")
                self.progress_bar.setMaximum(0) # Busy indicator
            elif tag == "Total":
                self.progress_bar.setMaximum(total)
                self.lbl_status.setText(f"Found {total} files. Starting search...")
            elif tag == "Progress":
                self.progress_bar.setValue(current)
                self.lbl_status.setText(f"Scanning {current} of {total}...")
        else:
            self.lbl_status.setText(str(msg_data))

    def add_result(self, result):
        row = self.table_results.rowCount()
        self.table_results.insertRow(row)
        
        name_item = QTableWidgetItem(result['filename'])
        
        # Use a custom QTableWidgetItem for numeric sorting on Match %
        score_item = QTableWidgetItem()
        score_item.setData(Qt.ItemDataRole.DisplayRole, f"{result['score']:.1f}%")
        score_item.setData(Qt.ItemDataRole.EditRole, result['score']) # This helps with sorting
        
        path_item = QTableWidgetItem(os.path.dirname(result['path']))
        
        # Store the result object for selection handling
        name_item.setData(Qt.ItemDataRole.UserRole, result)
        
        self.table_results.setItem(row, 0, name_item)
        self.table_results.setItem(row, 1, score_item)
        self.table_results.setItem(row, 2, path_item)

    def search_finished(self):
        self.btn_search.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.table_results.setSortingEnabled(True)
        self.lbl_status.setText(f"Done. Found {self.table_results.rowCount()} results.")
        self.progress_bar.setVisible(False)

    def show_error(self, msg):
        QMessageBox.critical(self, "Error", msg)

    def load_file_from_selected_row(self):
        selected_items = self.table_results.selectedItems()
        if not selected_items:
            return
        # The data is attached to the first item in the row
        row = selected_items[0].row()
        name_item = self.table_results.item(row, 0)
        result = name_item.data(Qt.ItemDataRole.UserRole)
        if result:
            self.load_file_from_result(result)

    def load_file_from_result(self, result):
        if not result:
            return

        filepath = result['path']
        self.lbl_filepath.setText(filepath)
        self.stack.setCurrentWidget(self.file_viewer_widget)
        self.btn_open_external.setEnabled(True)
        self.btn_reveal_explorer.setEnabled(True)

        try:
            content = ""
            if filepath.lower().endswith('.docx'):
                lines = self.search_engine._extract_text_from_docx(filepath)
                content = "\n".join(lines)
            elif filepath.lower().endswith('.xlsx'):
                lines = self.search_engine._extract_text_from_xlsx(filepath)
                content = "\n".join(lines)
            else:
                with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
            
            self.text_editor.setPlainText(content)
            self.highlighter.set_file(filepath)
            
            # Highlight Matches - simple selection highlight
            # Note: This just highlights matches in file. Actual fuzzy matches might vary.
            # Using the 'matches' from result could help if they have positions.
            # But process.extract returns strings, not positions usually, unless extended.
            # The original code had some highlight logic. Let's try to highlight the query string simply.
            
            # Reset extra selections
            self.text_editor.setExtraSelections([])
            
            # Highlight the exact query if found (fuzzy search finds partials, but this helps)
            # The original fuzzy search logic returns a list of matched lines?
            # Let's see what visual feedback we can improve later.
            
        except Exception as e:
            self.text_editor.setPlainText(f"Error reading file: {e}")

        # Handle matches
        self.current_matches = []
        self.current_match_index = -1
        self.lbl_match_info.setText("")
        self.btn_prev_match.setEnabled(False)
        self.btn_next_match.setEnabled(False)

        if 'matches' in result and result['matches']:
            # Sort matches by line index (item[2])
            matches = sorted(result['matches'], key=lambda x: x[2])
            self.current_matches = matches
            
            # Find the best match index initially
            best_score = -1
            best_idx = 0
            for i, m in enumerate(matches):
                if m[1] > best_score:
                    best_score = m[1]
                    best_idx = i
            
            self.current_match_index = best_idx
            
            self.update_match_buttons()
            self.highlight_current_match()

    def update_match_buttons(self):
        count = len(self.current_matches)
        if count > 1:
            self.btn_prev_match.setEnabled(True)
            self.btn_next_match.setEnabled(True)
            self.btn_goto_match.setEnabled(True)
            self.lbl_match_info.setText(f"Match {self.current_match_index + 1} of {count}")
        elif count == 1:
            self.lbl_match_info.setText("1 Match")
            self.btn_prev_match.setEnabled(False)
            self.btn_next_match.setEnabled(False)
            self.btn_goto_match.setEnabled(True)
        else:
            self.lbl_match_info.setText("No Matches")
            self.btn_prev_match.setEnabled(False)
            self.btn_next_match.setEnabled(False)
            self.btn_goto_match.setEnabled(False)

    def scroll_to_current_match(self):
        self.highlight_current_match()

    def highlight_current_match(self):
        if not self.current_matches or self.current_match_index < 0:
            return

        match = self.current_matches[self.current_match_index]
        line_index = match[2]

        block = self.text_editor.document().findBlockByNumber(line_index)
        if block.isValid():
            cursor = QTextCursor(block)
            self.text_editor.setTextCursor(cursor)
            self.text_editor.centerCursor()
        
        self.update_match_buttons()

    def next_match(self):
        if not self.current_matches:
            return
        self.current_match_index = (self.current_match_index + 1) % len(self.current_matches)
        self.highlight_current_match()

    def prev_match(self):
        if not self.current_matches:
            return
        self.current_match_index = (self.current_match_index - 1) % len(self.current_matches)
        self.highlight_current_match()

    def open_in_external_editor(self):
        filepath = self.lbl_filepath.text()
        if not filepath or not os.path.exists(filepath):
            return
        
        if not self.editor_path:
            QMessageBox.warning(self, "No Editor", "Please select an external editor first.")
            return

        # Get current line number (1-indexed)
        cursor = self.text_editor.textCursor()
        line_number = cursor.blockNumber() + 1

        try:
            editor_name = os.path.basename(self.editor_path).lower()
            args = [self.editor_path]

            # Support common editors with line number arguments
            if "code" in editor_name: # VS Code
                args.extend(["--goto", f"{filepath}:{line_number}"])
            elif "notepad++" in editor_name:
                args.extend([f"-n{line_number}", filepath])
            elif "subl" in editor_name: # Sublime Text
                args.append(f"{filepath}:{line_number}")
            elif "vim" in editor_name or "nvim" in editor_name or "neovide" in editor_name: # Vim or Neovim
                args.extend([f"+{line_number}", filepath])
            else:
                # Default fallback: just open the file
                args.append(filepath)

            subprocess.Popen(args)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open editor: {e}")

    def reveal_in_explorer(self):
        filepath = self.lbl_filepath.text()
        if not filepath or not os.path.exists(filepath):
            return
        
        if sys.platform == 'win32':
            # explorer /select,"path" highlights the file in Windows Explorer
            subprocess.run(['explorer', '/select,', os.path.normpath(filepath)])
        else:
            folder_path = os.path.dirname(os.path.abspath(filepath))
            if os.path.exists(folder_path):
                os.startfile(folder_path)

    def save_settings(self):
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("windowState", self.saveState())
        self.settings.setValue("directory", self.lbl_directory.text())
        self.settings.setValue("search_query", self.entry_search.text())
        self.settings.setValue("threshold", self.spin_threshold.value())
        self.settings.setValue("regex", self.chk_regex.isChecked())
        self.settings.setValue("office", self.chk_office.isChecked())
        self.settings.setValue("app_theme", self.combo_app_theme.currentText())
        self.settings.setValue("syntax_theme", self.combo_syntax_theme.currentText())
        self.settings.setValue("editor_path", self.editor_path)

    def load_settings(self):
        geom = self.settings.value("geometry")
        if geom:
            self.restoreGeometry(geom)
        state = self.settings.value("windowState")
        if state:
            self.restoreState(state)

        directory = self.settings.value("directory", "No directory selected")
        self.lbl_directory.setText(directory)

        query = self.settings.value("search_query", "")
        self.entry_search.setText(query)

        threshold = self.settings.value("threshold", 60, type=int)
        self.spin_threshold.setValue(threshold)

        regex = self.settings.value("regex", False, type=bool)
        self.chk_regex.setChecked(regex)

        office = self.settings.value("office", False, type=bool)
        self.chk_office.setChecked(office)

        app_theme = self.settings.value("app_theme", "Dark")
        self.combo_app_theme.setCurrentText(app_theme)
        self.apply_app_theme(app_theme)

        syntax_theme = self.settings.value("syntax_theme", "monokai" if "monokai" in self.styles else "default")
        self.combo_syntax_theme.setCurrentText(syntax_theme)
        # Note: apply_app_theme already calls Fusion, and the combo change triggers change_syntax_theme

    def closeEvent(self, event):
        self.save_settings()
        self.stop_search()
        event.accept()

    def apply_app_theme(self, theme_name):
        app = QApplication.instance()
        app.setStyle("Fusion")
        
        if theme_name == "Dark":
            palette = QPalette()
            palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
            palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
            palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25))
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
            palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
            palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
            palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
            palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
            palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
            palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
            palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
            palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
            palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.black)
            app.setPalette(palette)
            app.setStyleSheet(self.get_dark_stylesheet())
        else:
            # Light Theme
            palette = QPalette()
            palette.setColor(QPalette.ColorRole.Window, QColor(240, 240, 240))
            palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.black)
            palette.setColor(QPalette.ColorRole.Base, Qt.GlobalColor.white)
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor(233, 231, 227))
            palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
            palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.black)
            palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.black)
            palette.setColor(QPalette.ColorRole.Button, QColor(240, 240, 240))
            palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.black)
            palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
            palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
            palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
            palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)
            app.setPalette(palette)
            app.setStyleSheet(self.get_light_stylesheet())

    def get_dark_stylesheet(self):
        return """
            QWidget {
                font-family: "Segoe UI", sans-serif;
                font-size: 10pt;
            }
            QLineEdit, QSpinBox, QComboBox {
                border: 1px solid #555;
                border-radius: 4px;
                padding: 4px;
                background-color: #333;
                color: white;
                selection-background-color: #2a82da;
            }
            QLineEdit:focus, QSpinBox:focus, QComboBox:focus {
                border: 1px solid #2a82da;
            }
            QPushButton {
                background-color: #3d3d3d;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 5px 15px;
                color: white;
            }
            QPushButton:hover {
                background-color: #4d4d4d;
                border: 1px solid #666;
            }
            QPushButton:pressed {
                background-color: #2d2d2d;
            }
            QPushButton:disabled {
                background-color: #333;
                color: #777;
                border: 1px solid #444;
            }
            QPushButton#btn_search {
                background-color: #2a82da;
                border: 1px solid #1a62aa;
                color: white;
            }
            QPushButton#btn_search:hover {
                background-color: #358ee5;
            }
            QPushButton#btn_search:pressed {
                background-color: #1a62aa;
            }
            QPushButton#btn_stop {
                background-color: #d32f2f;
                border: 1px solid #b71c1c;
                color: white;
            }
            QPushButton#btn_stop:hover {
                background-color: #e53935;
            }
            QPushButton#btn_stop:pressed {
                background-color: #b71c1c;
            }
            QPushButton#btn_stop:disabled {
                background-color: #5a3333;
                color: #888;
            }
            QPushButton#btn_open_external {
                background-color: #00897b;
                border: 1px solid #00695c;
                color: white;
            }
            QPushButton#btn_open_external:hover {
                background-color: #009688;
            }
            QPushButton#btn_open_external:pressed {
                background-color: #00695c;
            }
            QListWidget {
                border: 1px solid #555;
                border-radius: 4px;
                background-color: #252525;
            }
            QListWidget::item {
                padding: 4px;
            }
            QListWidget::item:selected {
                background-color: #2a82da;
                color: black;
                border-radius: 2px;
            }
            QLabel {
                color: #e0e0e0;
            }
            QSplitter::handle {
                background-color: #444;
                width: 2px;
            }
            QPlainTextEdit, QTextEdit {
                border: 1px solid #555;
                border-radius: 4px;
                background-color: #1e1e1e;
            }
        """

    def get_light_stylesheet(self):
        return """
            QWidget {
                font-family: "Segoe UI", sans-serif;
                font-size: 10pt;
            }
            QLineEdit, QSpinBox, QComboBox {
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 4px;
                background-color: white;
                color: black;
                selection-background-color: #2a82da;
            }
            QLineEdit:focus, QSpinBox:focus, QComboBox:focus {
                border: 1px solid #2a82da;
            }
            QPushButton {
                background-color: #f0f0f0;
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 5px 15px;
                color: black;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
                border: 1px solid #bbb;
            }
            QPushButton:pressed {
                background-color: #d0d0d0;
            }
            QPushButton:disabled {
                background-color: #f5f5f5;
                color: #aaa;
                border: 1px solid #ddd;
            }
            QPushButton#btn_search {
                background-color: #2a82da;
                color: white;
                border: 1px solid #1a62aa;
            }
            QPushButton#btn_search:hover {
                background-color: #358ee5;
            }
            QPushButton#btn_search:pressed {
                background-color: #1a62aa;
            }
            QPushButton#btn_stop {
                background-color: #d32f2f;
                color: white;
                border: 1px solid #b71c1c;
            }
            QPushButton#btn_stop:hover {
                background-color: #e53935;
            }
            QPushButton#btn_stop:pressed {
                background-color: #b71c1c;
            }
            QPushButton#btn_stop:disabled {
                background-color: #ef9a9a;
                color: #555;
            }
            QPushButton#btn_open_external {
                background-color: #00897b;
                color: white;
                border: 1px solid #00695c;
            }
            QPushButton#btn_open_external:hover {
                background-color: #009688;
            }
            QPushButton#btn_open_external:pressed {
                background-color: #00695c;
            }
            QListWidget {
                border: 1px solid #ccc;
                border-radius: 4px;
                background-color: white;
            }
            QListWidget::item {
                padding: 4px;
            }
            QListWidget::item:selected {
                background-color: #2a82da;
                color: white;
                border-radius: 2px;
            }
            QLabel {
                color: black;
            }
            QSplitter::handle {
                background-color: #ccc;
                width: 2px;
            }
             QPlainTextEdit, QTextEdit {
                border: 1px solid #ccc;
                border-radius: 4px;
                background-color: white;
            }
        """

    def change_syntax_theme(self, style_name):
        self.highlighter.set_style(style_name)

if __name__ == '__main__':
    # Fix for Windows taskbar icon
    if sys.platform == 'win32':
        import ctypes
        myappid = 'Grapper.FuzzySearch.0.1' # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    app = QApplication(sys.argv)
    
    # Optional: Apply a dark theme or style
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
