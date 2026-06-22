from __future__ import annotations

import os
import sys
import uuid
import ctypes
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction, QIcon
from PySide6.QtPrintSupport import QPrinterInfo
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from autodocs.document_service import (
    SYSTEM_PLACEHOLDERS,
    convert_docx_to_pdf,
    convert_pdf_to_images,
    copy_template_to_library,
    create_excel_sample,
    format_value_for_output,
    read_excel_rows,
    render_filename,
    replace_placeholders,
    scan_docx_placeholders,
    with_system_values,
)
from autodocs.print_service import PrintCancelled, open_printer_preferences, print_files
from autodocs.storage import ExportHistoryRecord, TemplateRecord, TemplateStore


APP_NAME = "Auto Docs"
APP_ID = "AutoDocs.Desktop"
PROJECT_DIR = Path(__file__).resolve().parent.parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", PROJECT_DIR))
ROOT_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else PROJECT_DIR
DATA_DIR = ROOT_DIR / "data"
EXPORT_DIR = ROOT_DIR / "exports"
APP_ICON_PATH = RESOURCE_DIR / "assets" / "autodocs.ico"


class PrintFileList(QListWidget):
    filesDropped = Signal(list)

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            files = []
            for url in event.mimeData().urls():
                path = Path(url.toLocalFile())
                if path.is_file():
                    files.append(str(path))
            if files:
                self.filesDropped.emit(files)
            event.acceptProposedAction()
            return
        super().dropEvent(event)


class BatchPrintWorker(QThread):
    progressChanged = Signal(int, int, str)
    printFinished = Signal(bool, str)

    def __init__(self, files: list[str], printer_name: str, copies: int, delay_seconds: float):
        super().__init__()
        self.files = files
        self.printer_name = printer_name
        self.copies = copies
        self.delay_seconds = delay_seconds

    def run(self) -> None:
        try:
            print_files(
                self.files,
                self.printer_name,
                self.copies,
                self.delay_seconds,
                self.progressChanged.emit,
                self.isInterruptionRequested,
            )
        except PrintCancelled as exc:
            self.printFinished.emit(False, str(exc))
        except Exception as exc:
            self.printFinished.emit(False, str(exc))
        else:
            self.printFinished.emit(True, f"Da gui {len(self.files)} file den may in.")


class AddTemplateDialog(QDialog):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Thêm mẫu Word")
        self.setMinimumWidth(560)

        self.file_path = QLineEdit()
        self.file_path.setPlaceholderText("Chọn file Word .docx")
        self.file_path.setReadOnly(True)

        browse_button = QPushButton("Chọn file")
        browse_button.clicked.connect(self.choose_file)

        file_row = QHBoxLayout()
        file_row.addWidget(self.file_path, 1)
        file_row.addWidget(browse_button)

        self.name = QLineEdit()
        self.name.setPlaceholderText("Ví dụ: Hợp đồng bán hàng")

        self.category = QLineEdit()
        self.category.setPlaceholderText("Ví dụ: Hợp đồng, Nhân sự, Kế toán")

        self.summary = QLabel("Chưa chọn file.")
        self.summary.setObjectName("muted")
        self.summary.setWordWrap(True)

        form = QFormLayout()
        form.addRow("File Word", file_row)
        form.addRow("Tên mẫu", self.name)
        form.addRow("Nhóm", self.category)
        form.addRow("Kết quả quét", self.summary)

        cancel_button = QPushButton("Hủy")
        cancel_button.clicked.connect(self.reject)
        self.add_button = QPushButton("Lưu mẫu")
        self.add_button.setObjectName("primaryButton")
        self.add_button.clicked.connect(self.accept)

        footer = QHBoxLayout()
        footer.addStretch()
        footer.addWidget(cancel_button)
        footer.addWidget(self.add_button)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addLayout(footer)

        self.placeholders: list[str] = []

    def choose_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file Word mẫu",
            str(ROOT_DIR),
            "Word document (*.docx)",
        )
        if not path:
            return

        try:
            placeholders = scan_docx_placeholders(path)
        except Exception as exc:
            QMessageBox.critical(self, "Không đọc được file", str(exc))
            return

        self.file_path.setText(path)
        self.placeholders = placeholders
        if not self.name.text().strip():
            self.name.setText(Path(path).stem)

        if placeholders:
            preview = ", ".join(f"[[{item}]]" for item in placeholders[:8])
            more = "" if len(placeholders) <= 8 else f" và {len(placeholders) - 8} trường khác"
            self.summary.setText(f"Tìm thấy {len(placeholders)} placeholder: {preview}{more}.")
        else:
            self.summary.setText("Không tìm thấy placeholder dạng [[ten_truong]].")

    def accept(self) -> None:
        if not self.file_path.text().strip():
            QMessageBox.warning(self, "Thiếu file", "Vui lòng chọn file Word .docx.")
            return
        if not self.placeholders:
            answer = QMessageBox.question(
                self,
                "Không có placeholder",
                "File này chưa có placeholder. Bạn vẫn muốn lưu làm mẫu?",
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        super().accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.store = TemplateStore(DATA_DIR)
        self.templates: list[TemplateRecord] = []
        self.current_template: TemplateRecord | None = None
        self.manual_inputs: dict[str, QLineEdit] = {}
        self.excel_rows: list[dict[str, object]] = []
        self.print_worker: BatchPrintWorker | None = None

        self.setWindowTitle(APP_NAME)
        self.resize(1120, 720)
        self.setWindowIcon(self.app_icon())

        self.build_ui()
        self.apply_style()
        self.refresh_templates()

    def build_ui(self) -> None:
        toolbar = self.addToolBar("Công cụ")
        toolbar.setMovable(False)
        add_action = QAction(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon), "Thêm mẫu", self)
        add_action.triggered.connect(self.add_template)
        toolbar.addAction(add_action)
        open_action = QAction(self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon), "Mở thư mục xuất", self)
        open_action.triggered.connect(self.open_output_dir)
        toolbar.addAction(open_action)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self.build_sidebar())
        splitter.addWidget(self.build_workspace())
        splitter.setSizes([300, 880])
        self.setCentralWidget(splitter)

    def app_icon(self) -> QIcon:
        if APP_ICON_PATH.exists():
            return QIcon(str(APP_ICON_PATH))
        return self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView)

    def build_sidebar(self) -> QWidget:
        panel = QWidget()
        panel.setObjectName("sidebar")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        title = QLabel("Auto Docs")
        title.setObjectName("appTitle")
        subtitle = QLabel("Tạo tài liệu từ mẫu Word offline")
        subtitle.setObjectName("muted")

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Tìm mẫu...")
        self.search_box.textChanged.connect(self.populate_template_list)

        self.template_list = QListWidget()
        self.template_list.setObjectName("templateList")
        self.template_list.currentItemChanged.connect(self.on_template_changed)

        add_button = QPushButton("Thêm mẫu Word")
        add_button.setObjectName("primaryButton")
        add_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
        add_button.clicked.connect(self.add_template)

        delete_button = QPushButton("Xóa mẫu")
        delete_button.setObjectName("dangerButton")
        delete_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon))
        delete_button.clicked.connect(self.delete_current_template)

        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(self.search_box)
        layout.addWidget(self.template_list, 1)
        layout.addWidget(add_button)
        layout.addWidget(delete_button)
        return panel

    def build_workspace(self) -> QWidget:
        self.stack = QStackedWidget()
        self.empty_page = self.build_empty_page()
        self.detail_page = self.build_detail_page()
        self.stack.addWidget(self.empty_page)
        self.stack.addWidget(self.detail_page)
        return self.stack

    def build_empty_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title = QLabel("Chưa có mẫu nào")
        title.setObjectName("emptyTitle")
        text = QLabel("Thư viện mẫu đang trống.")
        text.setObjectName("muted")
        text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        button = QPushButton("Thêm mẫu Word")
        button.setObjectName("primaryButton")
        button.clicked.connect(self.add_template)
        layout.addWidget(title)
        layout.addWidget(text)
        layout.addSpacing(12)
        layout.addWidget(button, alignment=Qt.AlignmentFlag.AlignCenter)
        return page

    def build_detail_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(20, 18, 20, 20)
        layout.setSpacing(12)

        header = QHBoxLayout()
        header_text = QVBoxLayout()
        self.template_name = QLabel("")
        self.template_name.setObjectName("pageTitle")
        self.template_meta = QLabel("")
        self.template_meta.setObjectName("muted")
        header_text.addWidget(self.template_name)
        header_text.addWidget(self.template_meta)

        self.output_dir = QLineEdit(str(EXPORT_DIR))
        browse_output = QPushButton("Chọn thư mục xuất")
        browse_output.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirIcon))
        browse_output.clicked.connect(self.choose_output_dir)
        open_output = QPushButton("Mở")
        open_output.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
        open_output.clicked.connect(self.open_output_dir)
        output_row = QHBoxLayout()
        output_row.addWidget(QLabel("Xuất vào"))
        output_row.addWidget(self.output_dir, 1)
        output_row.addWidget(browse_output)
        output_row.addWidget(open_output)

        header.addLayout(header_text, 1)
        layout.addLayout(header)
        layout.addLayout(output_row)

        placeholder_frame = QFrame()
        placeholder_frame.setObjectName("panel")
        placeholder_layout = QVBoxLayout(placeholder_frame)
        placeholder_title = QLabel("Placeholder trong mẫu")
        placeholder_title.setObjectName("sectionTitle")
        self.placeholder_table = QTableWidget(0, 2)
        self.placeholder_table.setHorizontalHeaderLabels(["#", "Tên trường"])
        self.placeholder_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.placeholder_table.verticalHeader().setVisible(False)
        self.placeholder_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.placeholder_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.placeholder_table.setMaximumHeight(155)
        placeholder_layout.addWidget(placeholder_title)
        placeholder_layout.addWidget(self.placeholder_table)

        self.tabs = QTabWidget()
        self.tabs.addTab(
            self.build_manual_tab(),
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogContentsView),
            "Nhập thủ công",
        )
        self.tabs.addTab(
            self.build_excel_tab(),
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogListView),
            "Nhập từ Excel",
        )
        self.tabs.addTab(
            self.build_print_tab(),
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView),
            "In hàng loạt",
        )
        self.tabs.addTab(
            self.build_history_tab(),
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogInfoView),
            "Lịch sử",
        )

        layout.addWidget(placeholder_frame)
        layout.addWidget(self.tabs, 1)
        return page

    def build_manual_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 14, 0, 0)

        form_panel = QFrame()
        form_panel.setObjectName("panel")
        form_layout = QVBoxLayout(form_panel)
        form_title = QLabel("Thông tin cần điền")
        form_title.setObjectName("sectionTitle")

        self.manual_form_widget = QWidget()
        self.manual_form = QFormLayout(self.manual_form_widget)
        self.manual_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.manual_form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.manual_form_widget)

        form_layout.addWidget(form_title)
        form_layout.addWidget(scroll, 1)

        export_panel = QFrame()
        export_panel.setObjectName("panel")
        export_layout = QVBoxLayout(export_panel)
        export_title = QLabel("Quy cách xuất file")
        export_title.setObjectName("sectionTitle")
        self.filename_pattern = QLineEdit("[[so_hop_dong]] - [[ten_khach_hang]]")
        self.filename_pattern.setPlaceholderText("Ví dụ: Hop_dong_[[ten_khach_hang]]_[[today]]")

        button_row = QHBoxLayout()
        docx_button = QPushButton("Xuất DOCX")
        docx_button.setObjectName("primaryButton")
        docx_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
        docx_button.clicked.connect(lambda: self.export_manual("docx"))
        pdf_button = QPushButton("Xuất PDF")
        pdf_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        pdf_button.clicked.connect(lambda: self.export_manual("pdf"))
        image_button = QPushButton("Xuất ảnh PNG")
        image_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogInfoView))
        image_button.clicked.connect(lambda: self.export_manual("images"))
        button_row.addWidget(docx_button)
        button_row.addWidget(pdf_button)
        button_row.addWidget(image_button)
        button_row.addStretch()

        export_layout.addWidget(export_title)
        export_layout.addWidget(self.filename_pattern)
        export_layout.addLayout(button_row)

        layout.addWidget(form_panel, 1)
        layout.addWidget(export_panel)
        return page

    def build_excel_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 14, 0, 0)

        top = QFrame()
        top.setObjectName("panel")
        top_layout = QVBoxLayout(top)
        title = QLabel("Xuất hàng loạt từ Excel")
        title.setObjectName("sectionTitle")

        action_row = QHBoxLayout()
        sample_button = QPushButton("Tạo Excel mẫu")
        sample_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
        sample_button.clicked.connect(self.create_excel_template)
        import_button = QPushButton("Nhập Excel")
        import_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))
        import_button.clicked.connect(self.import_excel)
        batch_docx_button = QPushButton("Xuất DOCX hàng loạt")
        batch_docx_button.setObjectName("primaryButton")
        batch_docx_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogApplyButton))
        batch_docx_button.clicked.connect(lambda: self.export_batch("docx"))
        batch_pdf_button = QPushButton("Xuất PDF hàng loạt")
        batch_pdf_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        batch_pdf_button.clicked.connect(lambda: self.export_batch("pdf"))
        action_row.addWidget(sample_button)
        action_row.addWidget(import_button)
        action_row.addWidget(batch_docx_button)
        action_row.addWidget(batch_pdf_button)
        action_row.addStretch()

        self.batch_pattern = QLineEdit("[[so_hop_dong]] - [[ten_khach_hang]]")
        self.batch_pattern.setPlaceholderText("Quy cách tên file cho từng dòng Excel")

        top_layout.addWidget(title)
        top_layout.addWidget(self.batch_pattern)
        top_layout.addLayout(action_row)

        self.excel_summary = QLabel("Chưa nhập dữ liệu Excel.")
        self.excel_summary.setObjectName("muted")

        self.excel_table = QTableWidget(0, 0)
        self.excel_table.setObjectName("panelTable")
        self.excel_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.excel_table.setAlternatingRowColors(True)
        self.excel_table.setWordWrap(False)
        self.excel_table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.excel_table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.excel_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.excel_table.horizontalHeader().setDefaultAlignment(
            Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap
        )
        self.excel_table.horizontalHeader().setMinimumSectionSize(120)
        self.excel_table.horizontalHeader().setDefaultSectionSize(150)
        self.excel_table.horizontalHeader().setStretchLastSection(False)
        self.excel_table.verticalHeader().setDefaultSectionSize(34)

        layout.addWidget(top)
        layout.addWidget(self.excel_summary)
        layout.addWidget(self.excel_table, 1)
        return page

    def build_print_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 14, 0, 0)
        layout.setSpacing(10)

        config_panel = QFrame()
        config_panel.setObjectName("panel")
        config_layout = QVBoxLayout(config_panel)
        title = QLabel("In hàng loạt")
        title.setObjectName("sectionTitle")

        printer_row = QHBoxLayout()
        self.printer_combo = QComboBox()
        self.refresh_printer_list()

        refresh_printers = QPushButton("Làm mới")
        refresh_printers.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        refresh_printers.clicked.connect(self.refresh_printer_list)

        printer_settings = QPushButton("Cấu hình máy in")
        printer_settings.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        printer_settings.clicked.connect(self.configure_selected_printer)

        self.print_copies = QSpinBox()
        self.print_copies.setRange(1, 99)
        self.print_copies.setValue(1)
        self.print_copies.setFixedWidth(74)

        self.print_delay = QDoubleSpinBox()
        self.print_delay.setRange(0.5, 120.0)
        self.print_delay.setDecimals(1)
        self.print_delay.setSingleStep(0.5)
        self.print_delay.setValue(3.0)
        self.print_delay.setSuffix(" giây")
        self.print_delay.setFixedWidth(110)

        printer_row.addWidget(QLabel("Máy in"))
        printer_row.addWidget(self.printer_combo, 1)
        printer_row.addWidget(refresh_printers)
        printer_row.addWidget(printer_settings)
        printer_row.addWidget(QLabel("Số bản"))
        printer_row.addWidget(self.print_copies)
        printer_row.addWidget(QLabel("Nghỉ giữa file"))
        printer_row.addWidget(self.print_delay)

        file_actions = QHBoxLayout()
        add_files = QPushButton("Thêm file")
        add_files.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
        add_files.clicked.connect(self.add_print_files_dialog)
        remove_files = QPushButton("Xóa khỏi danh sách")
        remove_files.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon))
        remove_files.clicked.connect(self.remove_selected_print_files)
        sort_files = QPushButton("Sắp xếp A-Z")
        sort_files.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowUp))
        sort_files.clicked.connect(self.sort_print_files)
        move_up = QPushButton("Lên")
        move_up.clicked.connect(lambda: self.move_print_file(-1))
        move_down = QPushButton("Xuống")
        move_down.clicked.connect(lambda: self.move_print_file(1))
        clear_files = QPushButton("Xóa tất cả")
        clear_files.clicked.connect(self.clear_print_files)

        file_actions.addWidget(add_files)
        file_actions.addWidget(remove_files)
        file_actions.addWidget(sort_files)
        file_actions.addWidget(move_up)
        file_actions.addWidget(move_down)
        file_actions.addWidget(clear_files)
        file_actions.addStretch()

        config_layout.addWidget(title)
        config_layout.addLayout(printer_row)
        config_layout.addLayout(file_actions)

        self.print_file_list = PrintFileList()
        self.print_file_list.setObjectName("dropList")
        self.print_file_list.filesDropped.connect(self.add_print_files)

        self.print_status = QLabel("Kéo file cần in vào danh sách. Có thể kéo thả để đổi thứ tự.")
        self.print_status.setObjectName("muted")

        footer = QHBoxLayout()
        self.print_progress = QProgressBar()
        self.print_progress.setRange(0, 100)
        self.print_progress.setValue(0)

        self.start_print_button = QPushButton("In")
        self.start_print_button.setObjectName("primaryButton")
        self.start_print_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogApplyButton))
        self.start_print_button.clicked.connect(self.start_batch_print)

        self.cancel_print_button = QPushButton("Hủy")
        self.cancel_print_button.setEnabled(False)
        self.cancel_print_button.clicked.connect(self.cancel_batch_print)

        footer.addWidget(self.print_progress, 1)
        footer.addWidget(self.start_print_button)
        footer.addWidget(self.cancel_print_button)

        layout.addWidget(config_panel)
        layout.addWidget(self.print_file_list, 1)
        layout.addWidget(self.print_status)
        layout.addLayout(footer)
        return page

    def build_history_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 14, 0, 0)
        layout.setSpacing(10)

        top = QFrame()
        top.setObjectName("panel")
        top_layout = QVBoxLayout(top)
        title = QLabel("Lịch sử xuất file")
        title.setObjectName("sectionTitle")

        actions = QHBoxLayout()
        refresh_button = QPushButton("Làm mới")
        refresh_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        refresh_button.clicked.connect(self.refresh_history)
        open_folder_button = QPushButton("Mở thư mục")
        open_folder_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
        open_folder_button.clicked.connect(self.open_selected_history_folder)
        clear_button = QPushButton("Xóa lịch sử mẫu này")
        clear_button.setObjectName("dangerButton")
        clear_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon))
        clear_button.clicked.connect(self.clear_current_history)

        actions.addWidget(refresh_button)
        actions.addWidget(open_folder_button)
        actions.addWidget(clear_button)
        actions.addStretch()

        top_layout.addWidget(title)
        top_layout.addLayout(actions)

        self.history_table = QTableWidget(0, 5)
        self.history_table.setObjectName("panelTable")
        self.history_table.setHorizontalHeaderLabels(["Thời gian", "Định dạng", "File", "Số trường", "Tóm tắt"])
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.history_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.history_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.history_table.itemSelectionChanged.connect(self.show_selected_history_detail)
        self.history_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.history_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.history_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.history_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.history_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)

        self.history_detail = QTableWidget(0, 2)
        self.history_detail.setObjectName("panelTable")
        self.history_detail.setHorizontalHeaderLabels(["Trường", "Giá trị đã xuất"])
        self.history_detail.verticalHeader().setVisible(False)
        self.history_detail.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.history_detail.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.history_detail.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.history_detail.setMaximumHeight(220)

        self.history_summary = QLabel("Chưa có lịch sử xuất file.")
        self.history_summary.setObjectName("muted")

        layout.addWidget(top)
        layout.addWidget(self.history_summary)
        layout.addWidget(self.history_table, 1)
        layout.addWidget(self.history_detail)
        return page

    def refresh_templates(self) -> None:
        self.templates = self.store.list_templates()
        self.populate_template_list()
        if self.templates:
            self.template_list.setCurrentRow(0)
        else:
            self.stack.setCurrentWidget(self.empty_page)

    def populate_template_list(self) -> None:
        query = self.search_box.text().strip().casefold() if hasattr(self, "search_box") else ""
        current_id = self.current_template.id if self.current_template else None
        self.template_list.blockSignals(True)
        self.template_list.clear()
        selected_row = 0
        for record in self.templates:
            haystack = f"{record.name} {record.category}".casefold()
            if query and query not in haystack:
                continue
            detail = record.category or "Chưa phân nhóm"
            item = QListWidgetItem(f"{record.name}\n{detail} • {len(record.placeholders)} trường")
            item.setToolTip(f"{detail}\n{len(record.placeholders)} placeholder")
            item.setData(Qt.ItemDataRole.UserRole, record.id)
            self.template_list.addItem(item)
            if record.id == current_id:
                selected_row = self.template_list.count() - 1
        self.template_list.blockSignals(False)
        if self.template_list.count():
            self.template_list.setCurrentRow(selected_row)

    def on_template_changed(self, current: QListWidgetItem | None) -> None:
        if not current:
            return
        template_id = current.data(Qt.ItemDataRole.UserRole)
        record = self.store.get_template(template_id)
        if record:
            self.load_template(record)

    def load_template(self, record: TemplateRecord) -> None:
        self.current_template = record
        self.excel_rows = []
        self.stack.setCurrentWidget(self.detail_page)
        self.template_name.setText(record.name)
        category = record.category or "Chưa phân nhóm"
        self.template_meta.setText(f"{category} • {len(record.placeholders)} placeholder • {record.original_filename}")
        self.fill_placeholder_table(record.placeholders)
        editable = self.editable_placeholders(record)
        suggested_pattern = self.suggest_filename_pattern(record)
        self.filename_pattern.setText(suggested_pattern)
        self.batch_pattern.setText(suggested_pattern)
        self.rebuild_manual_form(editable)
        self.clear_excel_table()
        self.refresh_history()

    def editable_placeholders(self, record: TemplateRecord) -> list[str]:
        return [item for item in record.placeholders if item not in SYSTEM_PLACEHOLDERS]

    def suggest_filename_pattern(self, record: TemplateRecord) -> str:
        editable = self.editable_placeholders(record)
        preferred = [item for item in ("so_hop_dong", "ma_hop_dong", "ten_khach_hang") if item in editable]
        selected = preferred or editable[:2]
        if selected:
            return " - ".join(f"[[{item}]]" for item in selected)
        return f"{record.name}_[[auto_number]]"

    def fill_placeholder_table(self, placeholders: list[str]) -> None:
        self.placeholder_table.setRowCount(len(placeholders))
        for row, name in enumerate(placeholders):
            self.placeholder_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
            self.placeholder_table.setItem(row, 1, QTableWidgetItem(f"[[{name}]]"))
        self.placeholder_table.resizeColumnsToContents()

    def rebuild_manual_form(self, placeholders: list[str]) -> None:
        while self.manual_form.rowCount():
            self.manual_form.removeRow(0)
        self.manual_inputs = {}
        for placeholder in placeholders:
            edit = QLineEdit()
            edit.setPlaceholderText(f"Nhập {placeholder}")
            self.manual_form.addRow(f"[[{placeholder}]]", edit)
            self.manual_inputs[placeholder] = edit

    def add_template(self) -> None:
        dialog = AddTemplateDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        source_path = Path(dialog.file_path.text())
        template_id = uuid.uuid4().hex
        try:
            stored_path = copy_template_to_library(source_path, self.store.templates_dir, template_id)
            record = self.store.add_template(
                name=dialog.name.text(),
                category=dialog.category.text(),
                original_filename=source_path.name,
                stored_path=stored_path,
                placeholders=dialog.placeholders,
                template_id=template_id,
            )
        except Exception as exc:
            QMessageBox.critical(self, "Không lưu được mẫu", str(exc))
            return

        self.templates = self.store.list_templates()
        self.current_template = record
        self.populate_template_list()
        for row in range(self.template_list.count()):
            if self.template_list.item(row).data(Qt.ItemDataRole.UserRole) == record.id:
                self.template_list.setCurrentRow(row)
                break

    def delete_current_template(self) -> None:
        if not self.current_template:
            return
        answer = QMessageBox.question(
            self,
            "Xóa mẫu",
            f"Xóa mẫu '{self.current_template.name}' khỏi thư viện?",
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.store.delete_template(self.current_template.id)
        self.current_template = None
        self.refresh_templates()

    def choose_output_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Chọn thư mục xuất", self.output_dir.text())
        if path:
            self.output_dir.setText(path)

    def current_output_dir(self) -> Path:
        return Path(self.output_dir.text().strip() or EXPORT_DIR)

    def open_output_dir(self) -> None:
        path = self.current_output_dir()
        path.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(path.resolve()))
        except OSError as exc:
            QMessageBox.warning(self, "Không mở được thư mục", str(exc))

    def show_export_done(self, title: str, message: str, folder: Path) -> None:
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information)
        box.setWindowTitle(title)
        box.setText(message)
        open_button = box.addButton("Mở thư mục", QMessageBox.ButtonRole.ActionRole)
        box.addButton("Đóng", QMessageBox.ButtonRole.AcceptRole)
        box.exec()
        if box.clickedButton() == open_button:
            self.output_dir.setText(str(folder))
            self.open_output_dir()

    def collect_manual_values(self) -> dict[str, object]:
        return {name: edit.text().strip() for name, edit in self.manual_inputs.items()}

    def export_manual(self, mode: str) -> None:
        if not self.current_template:
            return
        values = with_system_values(self.collect_manual_values(), auto_number=1)
        pattern = self.filename_pattern.text().strip() or self.current_template.name
        try:
            created = self.export_one(self.current_template, values, pattern, mode, index=1)
        except Exception as exc:
            QMessageBox.critical(self, "Không xuất được file", str(exc))
            return
        folder = created[-1].parent if created else self.current_output_dir()
        if mode == "images" and created:
            folder = created[-1]
        self.refresh_history()
        self.show_export_done("Đã xuất xong", "Tài liệu đã được tạo trong thư mục xuất.", folder)

    def create_excel_template(self) -> None:
        if not self.current_template:
            return
        default_name = f"{self.current_template.name}_mau_excel.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Lưu Excel mẫu",
            str(EXPORT_DIR / default_name),
            "Excel workbook (*.xlsx)",
        )
        if not path:
            return
        try:
            create_excel_sample(path, self.editable_placeholders(self.current_template))
        except Exception as exc:
            QMessageBox.critical(self, "Không tạo được Excel mẫu", str(exc))
            return
        QMessageBox.information(self, "Đã tạo Excel mẫu", f"File mẫu đã được lưu:\n{path}")

    def import_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Chọn file Excel", str(ROOT_DIR), "Excel workbook (*.xlsx)")
        if not path:
            return
        try:
            self.excel_rows = read_excel_rows(path)
        except Exception as exc:
            QMessageBox.critical(self, "Không đọc được Excel", str(exc))
            return
        self.fill_excel_table(self.excel_rows)

    def fill_excel_table(self, rows: list[dict[str, object]]) -> None:
        placeholders = self.editable_placeholders(self.current_template) if self.current_template else []
        self.excel_table.setColumnCount(len(placeholders))
        self.excel_table.setRowCount(len(rows))
        self.excel_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.excel_summary.setText(f"{len(rows)} dòng dữ liệu • {len(placeholders)} cột. Có thể cuộn ngang để xem đủ cột.")
        for col_index, placeholder in enumerate(placeholders):
            header = QTableWidgetItem(self.compact_placeholder_label(placeholder))
            header.setToolTip(f"[[{placeholder}]]")
            self.excel_table.setHorizontalHeaderItem(col_index, header)
        for row_index, row in enumerate(rows):
            for col_index, placeholder in enumerate(placeholders):
                value = row.get(placeholder, "")
                display_value = format_value_for_output(value)
                item = QTableWidgetItem(display_value)
                item.setToolTip(display_value)
                self.excel_table.setItem(row_index, col_index, item)
        self.resize_excel_columns(placeholders, rows)

    def compact_placeholder_label(self, placeholder: str) -> str:
        words = placeholder.replace("_", " ").replace("-", " ").split()
        if not words:
            return placeholder

        lines: list[str] = []
        current = ""
        for word in words:
            candidate = word if not current else f"{current} {word}"
            if len(candidate) <= 13:
                current = candidate
            else:
                if current:
                    lines.append(current)
                current = word
        if current:
            lines.append(current)
        return "\n".join(lines[:3])

    def resize_excel_columns(self, placeholders: list[str], rows: list[dict[str, object]]) -> None:
        for col_index, placeholder in enumerate(placeholders):
            sample_values = [format_value_for_output(row.get(placeholder, "")) for row in rows[:20]]
            longest = max([len(placeholder.replace("_", " "))] + [len(value) for value in sample_values])
            width = max(130, min(230, longest * 8 + 34))
            self.excel_table.setColumnWidth(col_index, width)
        self.excel_table.horizontalHeader().setFixedHeight(56)

    def clear_excel_table(self) -> None:
        self.excel_table.setRowCount(0)
        self.excel_table.setColumnCount(0)
        if hasattr(self, "excel_summary"):
            self.excel_summary.setText("Chưa nhập dữ liệu Excel.")

    def refresh_printer_list(self) -> None:
        if not hasattr(self, "printer_combo"):
            return
        current = self.printer_combo.currentText()
        default_printer = QPrinterInfo.defaultPrinter().printerName()
        printers = [printer.printerName() for printer in QPrinterInfo.availablePrinters()]

        self.printer_combo.clear()
        self.printer_combo.addItems(printers)
        if current in printers:
            self.printer_combo.setCurrentText(current)
        elif default_printer in printers:
            self.printer_combo.setCurrentText(default_printer)
        self.printer_combo.setEnabled(bool(printers))

    def configure_selected_printer(self) -> None:
        printer_name = self.printer_combo.currentText().strip()
        if not printer_name:
            QMessageBox.warning(self, "Chưa chọn máy in", "Không tìm thấy máy in để cấu hình.")
            return
        try:
            open_printer_preferences(printer_name)
        except Exception as exc:
            QMessageBox.warning(self, "Không mở được cấu hình máy in", str(exc))

    def add_print_files_dialog(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Chọn file cần in",
            str(EXPORT_DIR),
            "Tài liệu (*.doc *.docx *.rtf *.pdf *.png *.jpg *.jpeg *.txt);;Tất cả file (*.*)",
        )
        self.add_print_files(paths)

    def add_print_files(self, paths: list[str]) -> None:
        existing = set(self.get_print_files())
        added = 0
        for path_text in paths:
            path = Path(path_text)
            if not path.is_file():
                continue
            resolved = str(path.resolve())
            if resolved in existing:
                continue
            item = QListWidgetItem(path.name)
            item.setToolTip(resolved)
            item.setData(Qt.ItemDataRole.UserRole, resolved)
            item.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
            self.print_file_list.addItem(item)
            existing.add(resolved)
            added += 1
        self.update_print_status(f"Đã thêm {added} file." if added else "Không có file mới để thêm.")

    def get_print_files(self) -> list[str]:
        return [
            self.print_file_list.item(row).data(Qt.ItemDataRole.UserRole)
            for row in range(self.print_file_list.count())
        ]

    def update_print_status(self, message: str | None = None) -> None:
        count = self.print_file_list.count() if hasattr(self, "print_file_list") else 0
        prefix = f"{count} file trong danh sách."
        self.print_status.setText(f"{prefix} {message}" if message else prefix)

    def remove_selected_print_files(self) -> None:
        for item in self.print_file_list.selectedItems():
            row = self.print_file_list.row(item)
            self.print_file_list.takeItem(row)
        self.update_print_status("Đã xóa file đã chọn.")

    def clear_print_files(self) -> None:
        self.print_file_list.clear()
        self.update_print_status("Danh sách đã trống.")

    def sort_print_files(self) -> None:
        paths = sorted(self.get_print_files(), key=lambda item: Path(item).name.casefold())
        self.print_file_list.clear()
        self.add_print_files(paths)
        self.update_print_status("Đã sắp xếp theo tên file.")

    def move_print_file(self, direction: int) -> None:
        selected = self.print_file_list.selectedItems()
        if len(selected) != 1:
            self.update_print_status("Chọn đúng một file để di chuyển.")
            return
        item = selected[0]
        row = self.print_file_list.row(item)
        new_row = row + direction
        if new_row < 0 or new_row >= self.print_file_list.count():
            return
        item = self.print_file_list.takeItem(row)
        self.print_file_list.insertItem(new_row, item)
        self.print_file_list.setCurrentItem(item)
        self.update_print_status("Đã đổi thứ tự in.")

    def start_batch_print(self) -> None:
        files = self.get_print_files()
        printer_name = self.printer_combo.currentText().strip()
        if not files:
            QMessageBox.warning(self, "Chưa có file", "Kéo thả hoặc thêm file cần in trước.")
            return
        if not printer_name:
            QMessageBox.warning(self, "Chưa chọn máy in", "Vui lòng chọn máy in.")
            return

        self.print_progress.setRange(0, len(files))
        self.print_progress.setValue(0)
        self.start_print_button.setEnabled(False)
        self.cancel_print_button.setEnabled(True)
        self.print_status.setText("Đang chuẩn bị in...")

        self.print_worker = BatchPrintWorker(
            files,
            printer_name,
            self.print_copies.value(),
            self.print_delay.value(),
        )
        self.print_worker.progressChanged.connect(self.on_print_progress)
        self.print_worker.printFinished.connect(self.on_print_finished)
        self.print_worker.finished.connect(self.print_worker.deleteLater)
        self.print_worker.start()

    def cancel_batch_print(self) -> None:
        if self.print_worker and self.print_worker.isRunning():
            self.print_worker.requestInterruption()
            self.print_status.setText("Đang hủy sau file hiện tại...")
            self.cancel_print_button.setEnabled(False)

    def on_print_progress(self, done: int, total: int, message: str) -> None:
        self.print_progress.setRange(0, total)
        self.print_progress.setValue(done)
        self.print_status.setText(message)

    def on_print_finished(self, success: bool, message: str) -> None:
        self.start_print_button.setEnabled(True)
        self.cancel_print_button.setEnabled(False)
        if self.print_worker:
            self.print_worker = None
        if success:
            self.print_progress.setValue(self.print_progress.maximum())
            QMessageBox.information(self, "In hàng loạt", message)
        else:
            QMessageBox.warning(self, "In hàng loạt", message)
        self.update_print_status()

    def refresh_history(self) -> None:
        if not hasattr(self, "history_table"):
            return
        self.history_table.setRowCount(0)
        self.history_detail.setRowCount(0)
        if not self.current_template:
            self.history_summary.setText("Chưa chọn mẫu.")
            return

        records = self.store.list_export_history(self.current_template.id)
        self.history_summary.setText(f"{len(records)} lần xuất gần nhất của mẫu này.")
        self.history_table.setRowCount(len(records))
        for row, record in enumerate(records):
            first_path = Path(record.output_paths[0]) if record.output_paths else Path("")
            values_count = len(record.values)
            summary = self.history_values_summary(record.values)
            items = [
                QTableWidgetItem(self.format_history_time(record.created_at)),
                QTableWidgetItem(record.export_format.upper()),
                QTableWidgetItem(first_path.name),
                QTableWidgetItem(str(values_count)),
                QTableWidgetItem(summary),
            ]
            for item in items:
                item.setToolTip(self.history_tooltip(record))
                item.setData(Qt.ItemDataRole.UserRole, record.id)
            self.history_table.setItem(row, 0, items[0])
            self.history_table.setItem(row, 1, items[1])
            self.history_table.setItem(row, 2, items[2])
            self.history_table.setItem(row, 3, items[3])
            self.history_table.setItem(row, 4, items[4])
        self.history_table.resizeRowsToContents()
        if records:
            self.history_table.setCurrentCell(0, 0)

    def current_history_records(self) -> list[ExportHistoryRecord]:
        if not self.current_template:
            return []
        return self.store.list_export_history(self.current_template.id)

    def selected_history_record(self) -> ExportHistoryRecord | None:
        if not self.current_template or not hasattr(self, "history_table"):
            return None
        row = self.history_table.currentRow()
        if row < 0:
            return None
        item = self.history_table.item(row, 0)
        if not item:
            return None
        record_id = item.data(Qt.ItemDataRole.UserRole)
        for record in self.current_history_records():
            if record.id == record_id:
                return record
        return None

    def show_selected_history_detail(self) -> None:
        record = self.selected_history_record()
        self.history_detail.setRowCount(0)
        if not record:
            return
        values = sorted(record.values.items(), key=lambda item: item[0].casefold())
        self.history_detail.setRowCount(len(values))
        for row, (key, value) in enumerate(values):
            name_item = QTableWidgetItem(f"[[{key}]]")
            value_text = format_value_for_output(value)
            value_item = QTableWidgetItem(value_text)
            value_item.setToolTip(value_text)
            self.history_detail.setItem(row, 0, name_item)
            self.history_detail.setItem(row, 1, value_item)
        self.history_detail.resizeRowsToContents()

    def open_selected_history_folder(self) -> None:
        record = self.selected_history_record()
        if not record or not record.output_paths:
            QMessageBox.warning(self, "Chưa chọn lịch sử", "Vui lòng chọn một dòng lịch sử.")
            return
        path = Path(record.output_paths[0])
        folder = path if path.is_dir() else path.parent
        if not folder.exists():
            QMessageBox.warning(self, "Không tìm thấy thư mục", str(folder))
            return
        try:
            os.startfile(str(folder.resolve()))
        except OSError as exc:
            QMessageBox.warning(self, "Không mở được thư mục", str(exc))

    def clear_current_history(self) -> None:
        if not self.current_template:
            return
        answer = QMessageBox.question(
            self,
            "Xóa lịch sử",
            f"Xóa toàn bộ lịch sử xuất của mẫu '{self.current_template.name}'?",
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.store.clear_export_history(self.current_template.id)
        self.refresh_history()

    def history_values_summary(self, values: dict[str, object]) -> str:
        useful = [(key, value) for key, value in values.items() if key not in SYSTEM_PLACEHOLDERS and value not in ("", None)]
        parts = [f"{key}: {format_value_for_output(value)}" for key, value in useful[:4]]
        more = "" if len(useful) <= 4 else f" +{len(useful) - 4} trường"
        return "; ".join(parts) + more

    def history_tooltip(self, record: ExportHistoryRecord) -> str:
        paths = "\n".join(record.output_paths)
        values = "\n".join(
            f"[[{key}]]: {format_value_for_output(value)}"
            for key, value in sorted(record.values.items(), key=lambda item: item[0].casefold())
        )
        return f"File:\n{paths}\n\nDữ liệu:\n{values}"

    def format_history_time(self, value: str) -> str:
        try:
            return datetime.fromisoformat(value).strftime("%d/%m/%Y %H:%M:%S")
        except ValueError:
            return value

    def export_batch(self, mode: str) -> None:
        if not self.current_template:
            return
        if not self.excel_rows:
            QMessageBox.warning(self, "Chưa có dữ liệu", "Vui lòng nhập file Excel trước.")
            return
        pattern = self.batch_pattern.text().strip() or self.current_template.name
        try:
            for index, row in enumerate(self.excel_rows, start=1):
                values = with_system_values(row, auto_number=index)
                self.export_one(self.current_template, values, pattern, mode, index=index)
        except Exception as exc:
            QMessageBox.critical(self, "Xuất hàng loạt bị lỗi", str(exc))
            return
        self.refresh_history()
        self.show_export_done(
            "Đã xuất xong",
            f"Đã tạo {len(self.excel_rows)} tài liệu trong thư mục xuất.",
            self.current_output_dir(),
        )

    def export_one(
        self,
        record: TemplateRecord,
        values: dict[str, object],
        pattern: str,
        mode: str,
        index: int,
    ) -> list[Path]:
        output_dir = self.current_output_dir()
        output_dir.mkdir(parents=True, exist_ok=True)
        stem = render_filename(pattern, values, fallback=f"{record.name}_{index:04d}")
        docx_path = output_dir / f"{stem}.docx"
        pdf_path = output_dir / f"{stem}.pdf"

        replace_placeholders(record.stored_path, docx_path, values)
        if mode == "docx":
            created = [docx_path]
            self.save_export_history(record, mode, created, values)
            return created
        convert_docx_to_pdf(docx_path, pdf_path)
        if mode == "images":
            image_dir = output_dir / f"{stem}_anh"
            convert_pdf_to_images(pdf_path, image_dir, stem)
            created = [docx_path, pdf_path, image_dir]
            self.save_export_history(record, mode, created, values)
            return created
        created = [docx_path, pdf_path]
        self.save_export_history(record, mode, created, values)
        return created

    def save_export_history(
        self,
        record: TemplateRecord,
        mode: str,
        output_paths: list[Path],
        values: dict[str, object],
    ) -> None:
        clean_values = {key: format_value_for_output(value) for key, value in values.items()}
        self.store.add_export_history(
            template_id=record.id,
            template_name=record.name,
            export_format=mode,
            output_paths=output_paths,
            values=clean_values,
        )

    def apply_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow, QWidget {
                background: #F6F7FB;
                color: #172033;
                font-family: "Segoe UI";
                font-size: 10pt;
            }
            QToolBar {
                background: #FFFFFF;
                border: 0;
                border-bottom: 1px solid #E3E7EF;
                padding: 5px;
                spacing: 8px;
            }
            #sidebar {
                background: #FFFFFF;
                border-right: 1px solid #E3E7EF;
            }
            #appTitle {
                color: #102A43;
                font-size: 21pt;
                font-weight: 700;
            }
            #pageTitle {
                color: #102A43;
                font-size: 18pt;
                font-weight: 700;
            }
            #emptyTitle {
                color: #102A43;
                font-size: 20pt;
                font-weight: 700;
            }
            #sectionTitle {
                color: #172033;
                font-size: 12pt;
                font-weight: 700;
            }
            #muted {
                color: #667085;
            }
            QLineEdit {
                background: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 7px;
                padding: 7px 10px;
                selection-background-color: #2563EB;
            }
            QLineEdit:focus {
                border-color: #2563EB;
            }
            QPushButton {
                background: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 7px;
                padding: 7px 11px;
                color: #172033;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #EEF4FF;
                border-color: #93B4F2;
            }
            QPushButton#primaryButton {
                background: #2563EB;
                border-color: #2563EB;
                color: #FFFFFF;
            }
            QPushButton#primaryButton:hover {
                background: #1D4ED8;
            }
            QPushButton#dangerButton:hover {
                background: #FFF1F2;
                border-color: #FDA4AF;
                color: #BE123C;
            }
            #panel {
                background: #FFFFFF;
                border: 1px solid #E3E7EF;
                border-radius: 8px;
            }
            QListWidget#templateList {
                background: #FFFFFF;
                border: 1px solid #E3E7EF;
                border-radius: 8px;
                outline: 0;
            }
            QListWidget#templateList::item {
                padding: 9px;
                border-radius: 6px;
                margin: 3px;
            }
            QListWidget#templateList::item:selected {
                background: #EAF1FF;
                color: #1D4ED8;
            }
            QListWidget#dropList {
                background: #FFFFFF;
                border: 1px dashed #93B4F2;
                border-radius: 8px;
                padding: 8px;
                outline: 0;
            }
            QListWidget#dropList::item {
                padding: 9px;
                border-bottom: 1px solid #EEF2F7;
            }
            QListWidget#dropList::item:selected {
                background: #EAF1FF;
                color: #1D4ED8;
            }
            QTableWidget {
                background: #FFFFFF;
                alternate-background-color: #F8FAFC;
                border: 1px solid #E3E7EF;
                border-radius: 7px;
                gridline-color: #E3E7EF;
            }
            QHeaderView::section {
                background: #EEF2F7;
                color: #344054;
                border: 0;
                border-right: 1px solid #D8DEE9;
                padding: 7px;
                font-weight: 700;
            }
            QTabWidget::pane {
                border: 0;
            }
            QTabBar::tab {
                background: #E9EEF6;
                padding: 9px 14px;
                margin-right: 4px;
                border-top-left-radius: 7px;
                border-top-right-radius: 7px;
                color: #475467;
                font-weight: 600;
            }
            QTabBar::tab:selected {
                background: #FFFFFF;
                color: #1D4ED8;
            }
            QScrollArea {
                border: 0;
                background: transparent;
            }
            """
        )


def main() -> None:
    set_windows_app_id()
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName("Auto Docs")
    app.setDesktopFileName(APP_ID)
    if APP_ICON_PATH.exists():
        app.setWindowIcon(QIcon(str(APP_ICON_PATH)))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


def set_windows_app_id() -> None:
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass
