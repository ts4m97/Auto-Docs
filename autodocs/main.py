from __future__ import annotations

import os
import sys
import uuid
import ctypes
import unicodedata
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QSize, Qt, QThread, Signal
from PySide6.QtGui import QColor, QIcon
from PySide6.QtPrintSupport import QPrinterInfo
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QButtonGroup,
    QCheckBox,
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
    QRadioButton,
    QScrollArea,
    QSizePolicy,
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
from autodocs.sql_import_service import SqlConnectionConfig, import_drivers_from_sql
from autodocs.storage import CustomerRecord, ExportHistoryRecord, TemplateRecord, TemplateStore


APP_NAME = "Auto Docs"
APP_ID = "AutoDocs.Desktop"
PROJECT_DIR = Path(__file__).resolve().parent.parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", PROJECT_DIR))
ROOT_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else PROJECT_DIR
DATA_DIR = ROOT_DIR / "data"
EXPORT_DIR = ROOT_DIR / "exports"
APP_ICON_PATH = RESOURCE_DIR / "assets" / "autodocs.ico"

CUSTOMER_FIELD_ALIASES = {
    "cccd": ["cccd", "so_cccd", "socccd", "cmnd", "so_cmnd", "socmnd", "cmt", "so_cmt", "socmt", "citizen_id", "id_number"],
    "ten_khach_hang": [
        "ten_khach_hang",
        "tenkhachhang",
        "ho_ten",
        "hoten",
        "ho_va_ten",
        "hovaten",
        "name",
        "customer_name",
        "client_name",
        "khach_hang",
    ],
    "so_dien_thoai": ["so_dien_thoai", "sodienthoai", "dien_thoai", "phone", "mobile", "tel", "sdt"],
    "email": ["email", "mail", "e_mail"],
    "dia_chi": ["dia_chi", "diachi", "address", "add", "dc", "noi_o", "noi_ohientai"],
    "ma_so_thue": ["ma_so_thue", "masothue", "mst", "tax_code", "tax_id"],
    "ngay_cap_cccd": ["ngay_cap_cccd", "ngaycapcccd", "ngay_cap", "issue_date"],
    "noi_cap_cccd": ["noi_cap_cccd", "noicapcccd", "noi_cap", "issue_place"],
    "nguoi_dai_dien": ["nguoi_dai_dien", "nguoidaidien", "representative", "dai_dien"],
    "ghi_chu": ["ghi_chu", "ghichu", "note", "notes"],
    "ten_khoa_hoc": ["ten_khoa_hoc", "tenkhoahoc", "ten_khoa", "tenkhoa", "tenkh"],
    "ma_khoa_hoc": ["ma_khoa_hoc", "makhoahoc", "ma_kh", "makh"],
    "hang_dao_tao": ["hang_dao_tao", "hangdaotao", "hang_dt", "hangdt", "hang_gplx", "hanggplx"],
}

DEFAULT_CUSTOMER_FIELDS = [
    "ten_khach_hang",
    "so_dien_thoai",
    "email",
    "dia_chi",
    "ma_so_thue",
    "ngay_cap_cccd",
    "noi_cap_cccd",
    "ghi_chu",
]


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


def normalize_field_key(value: str) -> str:
    text = unicodedata.normalize("NFKD", value)
    text = "".join(char for char in text if not unicodedata.combining(char))
    return "".join(char.casefold() for char in text if char.isalnum())


def customer_alias_lookup(extra_aliases: dict[str, list[str]] | None = None) -> dict[str, str]:
    lookup: dict[str, str] = {}
    merged = {key: list(value) for key, value in CUSTOMER_FIELD_ALIASES.items()}
    for canonical, aliases in (extra_aliases or {}).items():
        merged.setdefault(canonical, [])
        merged[canonical].extend(aliases)

    for canonical, aliases in merged.items():
        lookup[normalize_field_key(canonical)] = canonical
        for alias in aliases:
            lookup[normalize_field_key(alias)] = canonical
    return lookup


def resolve_customer_value(
    placeholder: str,
    customer: CustomerRecord,
    extra_aliases: dict[str, list[str]] | None = None,
) -> str | None:
    fields = {"cccd": customer.cccd, **customer.fields}
    normalized_fields = {normalize_field_key(key): str(value) for key, value in fields.items()}
    placeholder_key = normalize_field_key(placeholder)
    if placeholder_key in normalized_fields:
        return normalized_fields[placeholder_key]

    aliases = {key: list(value) for key, value in CUSTOMER_FIELD_ALIASES.items()}
    for canonical, custom_aliases in (extra_aliases or {}).items():
        aliases.setdefault(canonical, [])
        aliases[canonical].extend(custom_aliases)

    lookup = customer_alias_lookup(extra_aliases)
    canonical = lookup.get(placeholder_key)
    if not canonical:
        return None

    for alias in [canonical, *aliases.get(canonical, [])]:
        value = normalized_fields.get(normalize_field_key(alias))
        if value not in (None, ""):
            return value
    for field_key, value in normalized_fields.items():
        if lookup.get(field_key) == canonical and value:
            return value
    return None


class CustomerDialog(QDialog):
    def __init__(self, customer: CustomerRecord | None = None, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Thông tin khách hàng")
        self.setMinimumWidth(720)
        self.resize(760, 560)

        self.cccd = QLineEdit()
        self.cccd.setPlaceholderText("Số CCCD dùng làm mã khách hàng")
        if customer:
            self.cccd.setText(customer.cccd)
            self.cccd.setReadOnly(True)

        self.field_table = QTableWidget(0, 2)
        self.field_table.setHorizontalHeaderLabels(["Tên cột", "Giá trị"])
        self.field_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.field_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.field_table.verticalHeader().setVisible(False)
        self.field_table.setAlternatingRowColors(True)

        add_field = QPushButton("Thêm cột")
        add_field.clicked.connect(lambda: self.add_field_row("", ""))
        remove_field = QPushButton("Xóa cột đã chọn")
        remove_field.setObjectName("dangerButton")
        remove_field.clicked.connect(self.remove_selected_rows)

        tools = QHBoxLayout()
        tools.addWidget(add_field)
        tools.addWidget(remove_field)
        tools.addStretch()

        form = QFormLayout()
        form.addRow("CCCD", self.cccd)

        cancel_button = QPushButton("Hủy")
        cancel_button.clicked.connect(self.reject)
        save_button = QPushButton("Lưu khách hàng")
        save_button.setObjectName("primaryButton")
        save_button.clicked.connect(self.accept)

        footer = QHBoxLayout()
        footer.addStretch()
        footer.addWidget(cancel_button)
        footer.addWidget(save_button)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(QLabel("Các cột thông tin"))
        layout.addWidget(self.field_table, 1)
        layout.addLayout(tools)
        layout.addLayout(footer)

        if customer:
            for key, value in customer.fields.items():
                if key != "cccd":
                    self.add_field_row(key, value)
        else:
            for key in DEFAULT_CUSTOMER_FIELDS:
                self.add_field_row(key, "")

    def add_field_row(self, key: str, value: str) -> None:
        row = self.field_table.rowCount()
        self.field_table.insertRow(row)
        self.field_table.setItem(row, 0, QTableWidgetItem(key))
        self.field_table.setItem(row, 1, QTableWidgetItem(value))

    def remove_selected_rows(self) -> None:
        rows = sorted({item.row() for item in self.field_table.selectedItems()}, reverse=True)
        for row in rows:
            self.field_table.removeRow(row)

    def fields(self) -> dict[str, str]:
        fields: dict[str, str] = {}
        for row in range(self.field_table.rowCount()):
            key_item = self.field_table.item(row, 0)
            value_item = self.field_table.item(row, 1)
            key = key_item.text().strip() if key_item else ""
            if not key:
                continue
            value = value_item.text().strip() if value_item else ""
            fields[key] = value
        return fields

    def accept(self) -> None:
        if not self.cccd.text().strip():
            QMessageBox.warning(self, "Thiếu CCCD", "Vui lòng nhập số CCCD của khách hàng.")
            return
        super().accept()


class FieldAliasDialog(QDialog):
    def __init__(self, store: TemplateStore, parent: QWidget | None = None):
        super().__init__(parent)
        self.store = store
        self.setWindowTitle("Ánh xạ trường")
        self.setMinimumWidth(640)
        self.resize(680, 480)

        self.canonical = QComboBox()
        self.canonical.setEditable(True)
        self.canonical.addItems(sorted(CUSTOMER_FIELD_ALIASES.keys()))

        self.alias = QLineEdit()
        self.alias.setPlaceholderText("Ví dụ: diachi, address, noi_o_hien_tai")

        add_button = QPushButton("Thêm ánh xạ")
        add_button.setObjectName("primaryButton")
        add_button.clicked.connect(self.add_alias)
        delete_button = QPushButton("Xóa ánh xạ đã chọn")
        delete_button.setObjectName("dangerButton")
        delete_button.clicked.connect(self.delete_selected_alias)

        form = QFormLayout()
        form.addRow("Trường chuẩn", self.canonical)
        form.addRow("Alias", self.alias)

        action_row = QHBoxLayout()
        action_row.addWidget(add_button)
        action_row.addWidget(delete_button)
        action_row.addStretch()

        self.alias_table = QTableWidget(0, 2)
        self.alias_table.setHorizontalHeaderLabels(["Trường chuẩn", "Alias tự thêm"])
        self.alias_table.verticalHeader().setVisible(False)
        self.alias_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.alias_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.alias_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.alias_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.alias_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)

        close_button = QPushButton("Đóng")
        close_button.clicked.connect(self.accept)
        footer = QHBoxLayout()
        footer.addStretch()
        footer.addWidget(close_button)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addLayout(action_row)
        layout.addWidget(self.alias_table, 1)
        layout.addLayout(footer)

        self.refresh_aliases()

    def refresh_aliases(self) -> None:
        aliases = self.store.list_field_aliases()
        rows = [(canonical, alias) for canonical, values in aliases.items() for alias in values]
        self.alias_table.setRowCount(len(rows))
        for row, (canonical, alias) in enumerate(rows):
            self.alias_table.setItem(row, 0, QTableWidgetItem(canonical))
            self.alias_table.setItem(row, 1, QTableWidgetItem(alias))

    def add_alias(self) -> None:
        try:
            self.store.add_field_alias(self.canonical.currentText(), self.alias.text())
        except Exception as exc:
            QMessageBox.warning(self, "Không thêm được ánh xạ", str(exc))
            return
        self.alias.clear()
        self.refresh_aliases()

    def delete_selected_alias(self) -> None:
        row = self.alias_table.currentRow()
        if row < 0:
            return
        canonical_item = self.alias_table.item(row, 0)
        alias_item = self.alias_table.item(row, 1)
        if not canonical_item or not alias_item:
            return
        self.store.delete_field_alias(canonical_item.text(), alias_item.text())
        self.refresh_aliases()


class SqlImportDialog(QDialog):
    SETTING_KEY = "sql_import_config"

    def __init__(self, store: TemplateStore, parent: QWidget | None = None):
        super().__init__(parent)
        self.store = store
        self.setWindowTitle("Nhập khách hàng từ SQL Server")
        self.setMinimumWidth(620)

        config = self.store.get_app_setting(self.SETTING_KEY)

        self.server = QLineEdit(str(config.get("server") or "192.168.1.109\\SQLEXPRESS"))
        self.database = QLineEdit(str(config.get("database") or "GPLX_CDB_CSDT_v2"))
        self.username = QLineEdit(str(config.get("username") or "sa"))
        self.password = QLineEdit(str(config.get("password") or ""))
        self.password.setEchoMode(QLineEdit.EchoMode.Password)

        form = QFormLayout()
        form.addRow("Địa chỉ máy chủ", self.server)
        form.addRow("Tên CSDL", self.database)
        form.addRow("Tài khoản", self.username)
        form.addRow("Mật khẩu", self.password)

        hint = QLabel("Nguồn dữ liệu: dbo.NguoiLX. Nếu tìm được khóa liên kết, app sẽ lấy thêm thông tin từ dbo.KhoaHoc.")
        hint.setObjectName("muted")
        hint.setWordWrap(True)

        save_button = QPushButton("Lưu cấu hình")
        save_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
        save_button.clicked.connect(self.save_config)
        import_button = QPushButton("Nhập khách hàng")
        import_button.setObjectName("primaryButton")
        import_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogApplyButton))
        import_button.clicked.connect(self.import_customers)
        close_button = QPushButton("Thoát")
        close_button.clicked.connect(self.reject)

        buttons = QHBoxLayout()
        buttons.addStretch()
        buttons.addWidget(save_button)
        buttons.addWidget(import_button)
        buttons.addWidget(close_button)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(hint)
        layout.addLayout(buttons)

    def config(self) -> SqlConnectionConfig:
        return SqlConnectionConfig(
            server=self.server.text().strip(),
            database=self.database.text().strip(),
            username=self.username.text().strip(),
            password=self.password.text(),
        )

    def save_config(self) -> None:
        config = self.config()
        if not self.validate_config(config):
            return
        self.store.set_app_setting(
            self.SETTING_KEY,
            {
                "server": config.server,
                "database": config.database,
                "username": config.username,
                "password": config.password,
            },
        )
        QMessageBox.information(self, "Đã lưu", "Cấu hình SQL đã được lưu trên máy này.")

    def import_customers(self) -> None:
        config = self.config()
        if not self.validate_config(config):
            return
        self.store.set_app_setting(
            self.SETTING_KEY,
            {
                "server": config.server,
                "database": config.database,
                "username": config.username,
                "password": config.password,
            },
        )

        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            result = import_drivers_from_sql(config, self.store)
        except Exception as exc:
            QMessageBox.critical(self, "Không nhập được dữ liệu SQL", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()

        QMessageBox.information(
            self,
            "Đã nhập khách hàng",
            (
                f"Đã nhập/cập nhật {result.imported} khách hàng.\n"
                f"Bỏ qua {result.skipped} dòng thiếu CCCD.\n"
                f"Cột CCCD: {result.cccd_column}\n"
                f"Khóa học: {result.course_join}"
            ),
        )
        self.accept()

    def validate_config(self, config: SqlConnectionConfig) -> bool:
        if not config.server or not config.database or not config.username:
            QMessageBox.warning(self, "Thiếu cấu hình", "Vui lòng nhập đủ máy chủ, tên CSDL và tài khoản.")
            return False
        return True


class CustomerExportDialog(QDialog):
    def __init__(self, count: int, default_pattern: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Xuất khách hàng đã chọn")
        self.setMinimumWidth(520)

        self.pattern = QLineEdit(default_pattern)
        self.pattern.setPlaceholderText("Quy cách tên file")

        self.docx_radio = QRadioButton("DOCX")
        self.pdf_radio = QRadioButton("PDF")
        self.docx_radio.setChecked(True)

        mode_row = QHBoxLayout()
        mode_row.addWidget(self.docx_radio)
        mode_row.addWidget(self.pdf_radio)
        mode_row.addStretch()

        summary = QLabel(f"Sẽ xuất {count} khách hàng đang chọn bằng mẫu Word hiện tại.")
        summary.setObjectName("muted")
        summary.setWordWrap(True)

        form = QFormLayout()
        form.addRow("Tên file", self.pattern)
        form.addRow("Định dạng", mode_row)

        cancel_button = QPushButton("Hủy")
        cancel_button.clicked.connect(self.reject)
        export_button = QPushButton("Xuất")
        export_button.setObjectName("primaryButton")
        export_button.clicked.connect(self.accept)

        footer = QHBoxLayout()
        footer.addStretch()
        footer.addWidget(cancel_button)
        footer.addWidget(export_button)

        layout = QVBoxLayout(self)
        layout.addWidget(summary)
        layout.addLayout(form)
        layout.addLayout(footer)

    def export_mode(self) -> str:
        return "pdf" if self.pdf_radio.isChecked() else "docx"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.store = TemplateStore(DATA_DIR)
        self.templates: list[TemplateRecord] = []
        self.customers: list[CustomerRecord] = []
        self.customer_aliases: dict[str, list[str]] = {}
        self.current_template: TemplateRecord | None = None
        self.manual_inputs: dict[str, QLineEdit] = {}
        self.excel_rows: list[dict[str, object]] = []
        self.print_worker: BatchPrintWorker | None = None

        self.setWindowTitle(APP_NAME)
        self.resize(1120, 720)
        self.setWindowIcon(self.app_icon())

        self.build_ui()
        self.apply_style()
        self.refresh_customers()
        self.refresh_templates()

    def build_ui(self) -> None:
        toolbar = self.addToolBar("Công cụ")
        toolbar.setObjectName("topToolbar")
        toolbar.setMovable(False)
        toolbar.setIconSize(QSize(18, 18))

        brand = QLabel("Auto Docs")
        brand.setObjectName("toolbarBrand")
        toolbar.addWidget(brand)

        nav_shell = QFrame()
        nav_shell.setObjectName("toolbarNav")
        nav_layout = QHBoxLayout(nav_shell)
        nav_layout.setContentsMargins(4, 3, 4, 3)
        nav_layout.setSpacing(2)
        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        self.nav_buttons: list[QPushButton] = []
        for index, label in enumerate(("Mẫu", "Khách hàng", "Excel", "In", "Lịch sử")):
            button = QPushButton(label)
            button.setObjectName("navButton")
            button.setCheckable(True)
            button.setMinimumHeight(34)
            button.clicked.connect(lambda _checked=False, tab_index=index: self.on_toolbar_tab_changed(tab_index))
            self.nav_group.addButton(button, index)
            self.nav_buttons.append(button)
            nav_layout.addWidget(button)
        self.nav_buttons[0].setChecked(True)
        toolbar.addWidget(nav_shell)

        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        toolbar.addWidget(spacer)

        add_button = QPushButton("Thêm mẫu")
        add_button.setObjectName("toolbarActionButton")
        add_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
        add_button.clicked.connect(self.add_template)
        toolbar.addWidget(add_button)

        open_button = QPushButton("Mở thư mục xuất")
        open_button.setObjectName("toolbarActionButton")
        open_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
        open_button.clicked.connect(self.open_output_dir)
        toolbar.addWidget(open_button)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self.build_sidebar())
        splitter.addWidget(self.build_workspace())
        splitter.setSizes([300, 880])
        self.setCentralWidget(splitter)

    def on_toolbar_tab_changed(self, index: int) -> None:
        if not hasattr(self, "tabs"):
            return
        self.update_template_context_visibility(index)
        if index == 0:
            self.tabs.setCurrentIndex(0)
            self.stack.setCurrentWidget(self.detail_page if self.current_template else self.empty_page)
            return
        self.stack.setCurrentWidget(self.detail_page)
        self.tabs.setCurrentIndex(index)

    def sync_toolbar_tab(self, index: int) -> None:
        self.update_template_context_visibility(index)
        if not hasattr(self, "nav_buttons") or index < 0 or index >= len(self.nav_buttons):
            return
        self.nav_buttons[index].setChecked(True)

    def update_template_context_visibility(self, tab_index: int) -> None:
        if hasattr(self, "template_context_panel"):
            self.template_context_panel.setVisible(tab_index not in (1, 3))

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

        title = QLabel("Thư viện mẫu")
        title.setObjectName("appTitle")
        subtitle = QLabel("Chọn mẫu Word để tạo tài liệu")
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

        self.template_context_panel = QWidget()
        context_layout = QVBoxLayout(self.template_context_panel)
        context_layout.setContentsMargins(0, 0, 0, 0)
        context_layout.setSpacing(12)
        context_layout.addLayout(header)
        context_layout.addLayout(output_row)

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
        context_layout.addWidget(placeholder_frame)

        self.tabs = QTabWidget()
        self.tabs.addTab(
            self.build_manual_tab(),
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogContentsView),
            "Nhập thủ công",
        )
        self.tabs.addTab(
            self.build_customer_tab(),
            self.style().standardIcon(QStyle.StandardPixmap.SP_DirHomeIcon),
            "Khách hàng",
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
        self.tabs.tabBar().hide()
        self.tabs.currentChanged.connect(self.sync_toolbar_tab)

        layout.addWidget(self.template_context_panel)
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

        customer_row = QHBoxLayout()
        self.manual_customer_combo = QComboBox()
        self.manual_customer_combo.setMinimumWidth(260)
        fill_customer_button = QPushButton("Điền từ khách hàng")
        fill_customer_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogApplyButton))
        fill_customer_button.clicked.connect(self.fill_manual_from_selected_customer)
        manage_customer_button = QPushButton("Quản lý")
        manage_customer_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirHomeIcon))
        manage_customer_button.clicked.connect(lambda: self.tabs.setCurrentWidget(self.customer_tab))
        customer_row.addWidget(QLabel("Khách hàng"))
        customer_row.addWidget(self.manual_customer_combo, 1)
        customer_row.addWidget(fill_customer_button)
        customer_row.addWidget(manage_customer_button)

        self.manual_form_widget = QWidget()
        self.manual_form = QFormLayout(self.manual_form_widget)
        self.manual_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.manual_form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.manual_form_widget)

        form_layout.addWidget(form_title)
        form_layout.addLayout(customer_row)
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

    def build_customer_tab(self) -> QWidget:
        page = QWidget()
        self.customer_tab = page
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 14, 0, 0)
        layout.setSpacing(10)

        top = QFrame()
        top.setObjectName("panel")
        top_layout = QVBoxLayout(top)
        title = QLabel("Danh bạ khách hàng")
        title.setObjectName("sectionTitle")

        self.customer_search = QLineEdit()
        self.customer_search.setPlaceholderText("Tìm theo CCCD, tên, điện thoại, email...")
        self.customer_search.textChanged.connect(self.refresh_customers)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        add_button = QPushButton("Thêm khách hàng")
        add_button.setObjectName("customerPrimaryAction")
        add_button.clicked.connect(self.add_customer)
        sql_import_button = QPushButton("Nhập từ SQL")
        sql_import_button.setObjectName("customerAction")
        sql_import_button.clicked.connect(self.import_customers_from_sql)
        edit_button = QPushButton("Sửa")
        edit_button.setObjectName("customerAction")
        edit_button.clicked.connect(self.edit_selected_customer)
        delete_button = QPushButton("Xóa")
        delete_button.setObjectName("customerDangerAction")
        delete_button.clicked.connect(self.delete_selected_customer)
        select_all_button = QPushButton("Chọn tất cả")
        select_all_button.setObjectName("customerAction")
        select_all_button.clicked.connect(self.check_all_visible_customers)
        clear_selection_button = QPushButton("Bỏ chọn")
        clear_selection_button.setObjectName("customerAction")
        clear_selection_button.clicked.connect(self.clear_checked_customers)
        fill_button = QPushButton("Điền vào mẫu")
        fill_button.setObjectName("customerAction")
        fill_button.clicked.connect(self.fill_manual_from_customer_table)
        export_selected_button = QPushButton("Xuất KH đã chọn")
        export_selected_button.setObjectName("customerPrimaryAction")
        export_selected_button.clicked.connect(self.export_selected_customers)
        alias_button = QPushButton("Ánh xạ")
        alias_button.setObjectName("customerAction")
        alias_button.clicked.connect(self.manage_field_aliases)

        actions.addWidget(add_button)
        actions.addWidget(sql_import_button)
        actions.addWidget(edit_button)
        actions.addWidget(delete_button)
        actions.addWidget(select_all_button)
        actions.addWidget(clear_selection_button)
        actions.addWidget(fill_button)
        actions.addWidget(export_selected_button)
        actions.addWidget(alias_button)
        actions.addStretch()

        top_layout.addWidget(title)
        top_layout.addWidget(self.customer_search)
        top_layout.addLayout(actions)

        self.customer_summary = QLabel("Chưa có khách hàng.")
        self.customer_summary.setObjectName("muted")
        self.customer_selection_summary = QLabel("Chưa chọn khách hàng nào.")
        self.customer_selection_summary.setObjectName("selectionSummary")

        self.customer_table = QTableWidget(0, 6)
        self.customer_table.setObjectName("customerTable")
        self.customer_table.setHorizontalHeaderLabels(["Chọn", "CCCD", "Tên khách hàng", "Điện thoại", "Email", "Số cột"])
        self.customer_table.verticalHeader().setVisible(False)
        self.customer_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.customer_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.customer_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.customer_table.setAlternatingRowColors(True)
        self.customer_table.setWordWrap(False)
        self.customer_table.setShowGrid(False)
        self.customer_table.verticalHeader().setDefaultSectionSize(42)
        self.customer_table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.customer_table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.customer_table.doubleClicked.connect(self.edit_selected_customer)
        self.customer_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.customer_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.customer_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.customer_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.customer_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.customer_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)

        layout.addWidget(top)
        layout.addWidget(self.customer_summary)
        layout.addWidget(self.customer_selection_summary)
        layout.addWidget(self.customer_table, 1)
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

    def refresh_customers(self, *_args) -> None:
        query = self.customer_search.text() if hasattr(self, "customer_search") else ""
        self.customer_aliases = self.store.list_field_aliases()
        self.customers = self.store.list_customers(query)
        self.populate_customer_combo()
        self.populate_customer_table()

    def populate_customer_combo(self) -> None:
        if not hasattr(self, "manual_customer_combo"):
            return
        current_cccd = self.manual_customer_combo.currentData()
        self.manual_customer_combo.blockSignals(True)
        self.manual_customer_combo.clear()
        self.manual_customer_combo.addItem("Chọn khách hàng...", "")
        selected_index = 0
        for index, customer in enumerate(self.store.list_customers(), start=1):
            label = f"{customer.display_name} ({self.mask_cccd(customer.cccd)})"
            self.manual_customer_combo.addItem(label, customer.cccd)
            if customer.cccd == current_cccd:
                selected_index = index
        self.manual_customer_combo.setCurrentIndex(selected_index)
        self.manual_customer_combo.blockSignals(False)

    def populate_customer_table(self) -> None:
        if not hasattr(self, "customer_table"):
            return
        checked = set(self.checked_customer_cccds())
        self.customer_table.blockSignals(True)
        self.customer_table.setRowCount(len(self.customers))
        for row, customer in enumerate(self.customers):
            check_item = QTableWidgetItem()
            check_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            check_item.setData(Qt.ItemDataRole.UserRole, customer.cccd)
            self.customer_table.setItem(row, 0, check_item)
            self.customer_table.setCellWidget(row, 0, self.create_customer_check_cell(customer.cccd, customer.cccd in checked))

            values = [
                customer.cccd,
                customer.display_name,
                resolve_customer_value("so_dien_thoai", customer, self.customer_aliases) or "",
                resolve_customer_value("email", customer, self.customer_aliases) or "",
                str(len(customer.fields)),
            ]
            for column, value in enumerate(values, start=1):
                item = QTableWidgetItem(value)
                item.setData(Qt.ItemDataRole.UserRole, customer.cccd)
                item.setToolTip(self.customer_tooltip(customer))
                self.customer_table.setItem(row, column, item)
            self.set_customer_row_checked(row, customer.cccd in checked)
            self.customer_table.setRowHeight(row, 42)
        self.customer_table.blockSignals(False)
        self.customer_table.setColumnWidth(0, 58)
        if hasattr(self, "customer_summary"):
            self.customer_summary.setText(f"{len(self.customers)} khách hàng trong danh sách.")
        self.update_customer_selection_summary()

    def create_customer_check_cell(self, cccd: str, checked: bool) -> QWidget:
        wrapper = QWidget()
        wrapper.setObjectName("checkCell")
        layout = QHBoxLayout(wrapper)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        checkbox = QCheckBox()
        checkbox.setObjectName("customerCheckBox")
        checkbox.setChecked(checked)
        checkbox.setProperty("cccd", cccd)
        checkbox.stateChanged.connect(self.on_customer_check_changed)
        layout.addWidget(checkbox)
        return wrapper

    def on_customer_check_changed(self, *_args) -> None:
        checkbox = self.sender()
        if not isinstance(checkbox, QCheckBox):
            return
        row = self.customer_row_for_cccd(str(checkbox.property("cccd") or ""))
        if row >= 0:
            self.set_customer_row_checked(row, checkbox.isChecked())
        self.update_customer_selection_summary()

    def customer_row_for_cccd(self, cccd: str) -> int:
        for row in range(self.customer_table.rowCount()):
            item = self.customer_table.item(row, 0)
            if item and item.data(Qt.ItemDataRole.UserRole) == cccd:
                return row
        return -1

    def set_customer_row_checked(self, row: int, checked: bool) -> None:
        color = QColor("#EEF4FF") if checked else QColor("#FFFFFF")
        alternate = QColor("#F8FAFC") if row % 2 else QColor("#FFFFFF")
        background = color if checked else alternate
        for column in range(1, self.customer_table.columnCount()):
            item = self.customer_table.item(row, column)
            if item:
                item.setBackground(background)

    def mask_cccd(self, cccd: str) -> str:
        if len(cccd) <= 6:
            return cccd
        return f"{cccd[:4]}...{cccd[-3:]}"

    def customer_tooltip(self, customer: CustomerRecord) -> str:
        fields = [f"CCCD: {customer.cccd}"]
        for key, value in sorted(customer.fields.items(), key=lambda item: item[0].casefold()):
            if key != "cccd" and value:
                fields.append(f"{key}: {value}")
        return "\n".join(fields)

    def selected_customer_from_combo(self) -> CustomerRecord | None:
        if not hasattr(self, "manual_customer_combo"):
            return None
        cccd = self.manual_customer_combo.currentData()
        return self.store.get_customer(cccd) if cccd else None

    def selected_customer_from_table(self) -> CustomerRecord | None:
        if not hasattr(self, "customer_table"):
            return None
        row = self.customer_table.currentRow()
        if row < 0:
            return None
        item = self.customer_table.item(row, 1) or self.customer_table.item(row, 0)
        if not item:
            return None
        return self.store.get_customer(item.data(Qt.ItemDataRole.UserRole))

    def selected_customers_from_table(self) -> list[CustomerRecord]:
        cccds = self.checked_customer_cccds()
        customers = [self.store.get_customer(cccd) for cccd in cccds]
        return [customer for customer in customers if customer]

    def checked_customer_cccds(self) -> list[str]:
        if not hasattr(self, "customer_table"):
            return []
        cccds: list[str] = []
        for row in range(self.customer_table.rowCount()):
            item = self.customer_table.item(row, 0)
            checkbox = self.customer_checkbox_at(row)
            if not item or not checkbox:
                continue
            cccd = item.data(Qt.ItemDataRole.UserRole)
            if not cccd or not checkbox.isChecked():
                continue
            cccds.append(cccd)
        return cccds

    def customer_checkbox_at(self, row: int) -> QCheckBox | None:
        widget = self.customer_table.cellWidget(row, 0)
        return widget.findChild(QCheckBox) if widget else None

    def update_customer_selection_summary(self) -> None:
        if not hasattr(self, "customer_selection_summary"):
            return
        count = len(self.checked_customer_cccds())
        total = self.customer_table.rowCount() if hasattr(self, "customer_table") else 0
        if count:
            self.customer_selection_summary.setText(f"Đã chọn {count}/{total} khách hàng để xuất.")
        else:
            self.customer_selection_summary.setText("Tick vào cột Chọn để xuất một hoặc nhiều khách hàng.")

    def check_all_visible_customers(self) -> None:
        if not hasattr(self, "customer_table"):
            return
        self.customer_table.blockSignals(True)
        for row in range(self.customer_table.rowCount()):
            checkbox = self.customer_checkbox_at(row)
            if checkbox:
                checkbox.blockSignals(True)
                checkbox.setChecked(True)
                checkbox.blockSignals(False)
                self.set_customer_row_checked(row, True)
        self.customer_table.blockSignals(False)
        self.update_customer_selection_summary()

    def clear_checked_customers(self) -> None:
        if not hasattr(self, "customer_table"):
            return
        self.customer_table.blockSignals(True)
        for row in range(self.customer_table.rowCount()):
            checkbox = self.customer_checkbox_at(row)
            if checkbox:
                checkbox.blockSignals(True)
                checkbox.setChecked(False)
                checkbox.blockSignals(False)
                self.set_customer_row_checked(row, False)
        self.customer_table.blockSignals(False)
        self.update_customer_selection_summary()

    def add_customer(self) -> None:
        dialog = CustomerDialog(parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        cccd = dialog.cccd.text().strip()
        if self.store.get_customer(cccd):
            QMessageBox.warning(self, "CCCD đã tồn tại", "Khách hàng với số CCCD này đã có trong danh bạ.")
            return
        try:
            self.store.upsert_customer(cccd, dialog.fields())
        except Exception as exc:
            QMessageBox.critical(self, "Không lưu được khách hàng", str(exc))
            return
        self.refresh_customers()

    def import_customers_from_sql(self) -> None:
        dialog = SqlImportDialog(self.store, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh_customers()

    def edit_selected_customer(self, *_args) -> None:
        customer = self.selected_customer_from_table()
        if not customer:
            QMessageBox.warning(self, "Chưa chọn khách hàng", "Vui lòng chọn một khách hàng để sửa.")
            return
        dialog = CustomerDialog(customer, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            self.store.upsert_customer(customer.cccd, dialog.fields())
        except Exception as exc:
            QMessageBox.critical(self, "Không lưu được khách hàng", str(exc))
            return
        self.refresh_customers()

    def delete_selected_customer(self) -> None:
        customer = self.selected_customer_from_table()
        if not customer:
            QMessageBox.warning(self, "Chưa chọn khách hàng", "Vui lòng chọn một khách hàng để xóa.")
            return
        answer = QMessageBox.question(
            self,
            "Xóa khách hàng",
            f"Xóa khách hàng '{customer.display_name}' khỏi danh bạ?",
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.store.delete_customer(customer.cccd)
        self.refresh_customers()

    def manage_field_aliases(self) -> None:
        dialog = FieldAliasDialog(self.store, self)
        dialog.exec()
        self.refresh_customers()

    def fill_manual_from_selected_customer(self) -> None:
        customer = self.selected_customer_from_combo()
        if not customer:
            QMessageBox.warning(self, "Chưa chọn khách hàng", "Vui lòng chọn một khách hàng trong danh sách.")
            return
        self.apply_customer_to_manual_form(customer)

    def fill_manual_from_customer_table(self) -> None:
        customer = self.selected_customer_from_table()
        if not customer:
            QMessageBox.warning(self, "Chưa chọn khách hàng", "Vui lòng chọn một khách hàng trong danh sách.")
            return
        if hasattr(self, "manual_customer_combo"):
            index = self.manual_customer_combo.findData(customer.cccd)
            if index >= 0:
                self.manual_customer_combo.setCurrentIndex(index)
        self.apply_customer_to_manual_form(customer)
        self.tabs.setCurrentWidget(self.tabs.widget(0))

    def export_selected_customers(self) -> None:
        if not self.current_template:
            QMessageBox.warning(self, "Chưa chọn mẫu", "Vui lòng chọn mẫu Word trước khi xuất khách hàng.")
            return
        customers = self.selected_customers_from_table()
        if not customers:
            QMessageBox.warning(self, "Chưa chọn khách hàng", "Vui lòng chọn một hoặc nhiều khách hàng trong danh sách.")
            return

        default_pattern = self.suggest_customer_export_pattern(self.current_template)
        dialog = CustomerExportDialog(len(customers), default_pattern, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        pattern = dialog.pattern.text().strip() or default_pattern
        mode = dialog.export_mode()
        try:
            for index, customer in enumerate(customers, start=1):
                values = self.customer_values_for_template(customer, self.current_template, index)
                self.export_one(self.current_template, values, pattern, mode, index=index)
        except Exception as exc:
            QMessageBox.critical(self, "Xuất khách hàng bị lỗi", str(exc))
            return

        self.refresh_history()
        self.show_export_done(
            "Đã xuất xong",
            f"Đã tạo {len(customers)} tài liệu từ khách hàng đã chọn.",
            self.current_output_dir(),
        )

    def suggest_customer_export_pattern(self, record: TemplateRecord) -> str:
        editable = self.editable_placeholders(record)
        preferred = [
            item
            for item in ("ten_khach_hang", "cccd", "so_cmt", "so_cccd", "ten_khoa_hoc", "ma_khoa_hoc")
            if item in editable
        ]
        selected = preferred[:2] or editable[:2]
        if selected:
            return " - ".join(f"[[{item}]]" for item in selected)
        return f"{record.name}_[[auto_number]]"

    def customer_values_for_template(
        self,
        customer: CustomerRecord,
        record: TemplateRecord,
        index: int,
    ) -> dict[str, object]:
        values: dict[str, object] = {}
        for placeholder in self.editable_placeholders(record):
            values[placeholder] = resolve_customer_value(placeholder, customer, self.customer_aliases) or ""
        return with_system_values(values, auto_number=index)

    def apply_customer_to_manual_form(self, customer: CustomerRecord) -> None:
        if not self.manual_inputs:
            QMessageBox.warning(self, "Chưa có mẫu", "Vui lòng chọn mẫu Word trước khi điền khách hàng.")
            return
        filled = 0
        missing: list[str] = []
        for placeholder, edit in self.manual_inputs.items():
            value = resolve_customer_value(placeholder, customer, self.customer_aliases)
            if value not in (None, ""):
                edit.setText(value)
                filled += 1
            else:
                missing.append(placeholder)
        if filled:
            message = f"Đã điền {filled} trường từ khách hàng."
            if missing:
                message += f"\n{len(missing)} trường chưa có dữ liệu hoặc chưa ánh xạ."
            QMessageBox.information(self, "Đã điền dữ liệu", message)
        else:
            QMessageBox.warning(self, "Chưa khớp trường", "Không tìm thấy trường nào khớp với khách hàng này.")

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
                background: #F8FAFC;
                border: 0;
                border-bottom: 1px solid #DDE3EA;
                padding: 9px 12px;
                spacing: 10px;
            }
            #toolbarBrand {
                color: #111827;
                font-size: 14pt;
                font-weight: 800;
                padding: 0 16px 0 2px;
            }
            QFrame#toolbarNav {
                background: #FFFFFF;
                border: 1px solid #D7DEE8;
                border-radius: 9px;
            }
            QPushButton#navButton {
                background: transparent;
                color: #4B5563;
                border: 0;
                border-radius: 6px;
                padding: 7px 15px;
                font-weight: 700;
            }
            QPushButton#navButton:hover {
                background: #F1F5F9;
                color: #111827;
            }
            QPushButton#navButton:checked {
                background: #172033;
                color: #FFFFFF;
            }
            QPushButton#toolbarActionButton {
                background: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 7px;
                padding: 7px 12px;
                color: #172033;
                font-weight: 700;
            }
            QPushButton#toolbarActionButton:hover {
                background: #F1F5F9;
                border-color: #94A3B8;
            }
            #sidebar {
                background: #FFFFFF;
                border-right: 1px solid #E3E7EF;
            }
            #appTitle {
                color: #111827;
                font-size: 15pt;
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
            #selectionSummary {
                color: #334155;
                background: #EEF2F7;
                border: 1px solid #D8DEE9;
                border-radius: 7px;
                padding: 7px 10px;
                font-weight: 600;
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
            QPushButton#customerAction,
            QPushButton#customerDangerAction,
            QPushButton#customerPrimaryAction {
                min-height: 28px;
                padding: 6px 12px;
                border-radius: 7px;
                font-weight: 650;
            }
            QPushButton#customerAction {
                background: #FFFFFF;
                border: 1px solid #D5DCE7;
                color: #1F2937;
            }
            QPushButton#customerAction:hover {
                background: #F8FAFC;
                border-color: #AAB7C8;
            }
            QPushButton#customerPrimaryAction {
                background: #1F5FE0;
                border: 1px solid #1F5FE0;
                color: #FFFFFF;
            }
            QPushButton#customerPrimaryAction:hover {
                background: #174FC0;
                border-color: #174FC0;
            }
            QPushButton#customerDangerAction {
                background: #FFFFFF;
                border: 1px solid #F1C8D0;
                color: #A31535;
            }
            QPushButton#customerDangerAction:hover {
                background: #FFF5F6;
                border-color: #E89AAA;
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
            QTableWidget#customerTable {
                background: #FFFFFF;
                alternate-background-color: #F8FAFC;
                border: 1px solid #DDE5EF;
                border-radius: 8px;
                selection-background-color: #EAF1FF;
                selection-color: #0F172A;
                outline: 0;
            }
            QTableWidget#customerTable::item {
                border-bottom: 1px solid #E8EDF4;
                padding: 8px 6px;
            }
            QTableWidget#customerTable::item:selected {
                background: #EAF1FF;
                color: #0F172A;
            }
            QFrame#checkCell {
                background: transparent;
            }
            QCheckBox#customerCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 5px;
                border: 1px solid #AAB7C8;
                background: #FFFFFF;
            }
            QCheckBox#customerCheckBox::indicator:hover {
                border-color: #1F5FE0;
            }
            QCheckBox#customerCheckBox::indicator:checked {
                background: #1F5FE0;
                border-color: #1F5FE0;
                image: none;
            }
            QHeaderView::section {
                background: #EEF2F7;
                color: #344054;
                border: 0;
                border-right: 1px solid #D8DEE9;
                padding: 7px;
                font-weight: 700;
            }
            QTableWidget#customerTable QHeaderView::section {
                background: #F1F5F9;
                color: #1F2937;
                border: 0;
                border-bottom: 1px solid #DDE5EF;
                padding: 10px 8px;
                font-weight: 750;
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
