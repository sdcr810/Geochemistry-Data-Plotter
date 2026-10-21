import os
import re
import ast
import subprocess
from contextlib import contextmanager
import sys
from io import BytesIO
from pathlib import Path


def running_inside_spyder():
    return (
        "spyder_kernels" in sys.modules
        or "SPYDER_KERNEL_ID" in os.environ
        or any(name.startswith("SPYDER_") for name in os.environ)
    )


def relaunch_outside_spyder_if_needed():
    if not running_inside_spyder():
        return
    if os.environ.get("EXCEL_BIVARIATE_PLOTTER_EXTERNAL") == "1":
        return
    script_path = Path(__file__).resolve()
    child_environment = os.environ.copy()
    child_environment["EXCEL_BIVARIATE_PLOTTER_EXTERNAL"] = "1"
    child_environment["QT_API"] = "pyside6"
    creation_flags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen(
        [sys.executable, str(script_path)],
        cwd=str(script_path.parent),
        env=child_environment,
        creationflags=creation_flags,
    )
    print("The plotter was launched in a separate process to avoid mixing Spyder's Qt binding with PySide6.")
    raise SystemExit


relaunch_outside_spyder_if_needed()
os.environ["QT_API"] = "pyside6"

import pandas as pd
import matplotlib
from PySide6.QtCore import (
    QAbstractTableModel, QByteArray, QMimeData, QModelIndex, QRectF, QSettings, Qt, Signal
)
from PySide6.QtGui import QAction, QActionGroup, QColor, QFontDatabase, QImage, QPageLayout, QPainter, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QColorDialog,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStatusBar,
    QSpinBox,
    QTabWidget,
    QTableView,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtPrintSupport import QPrintDialog, QPrintPreviewDialog, QPrinter
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure
from matplotlib.ticker import FormatStrFormatter, ScalarFormatter


THEMES = {
    "System": None,
    "Light": {
        "ui_background": "#f0f0f0",
        "ui_surface": "#ffffff",
        "plot_background": "#ffffff",
        "text": "#202020",
        "muted_text": "#555555",
        "grid": "#c7c7c7",
        "border": "#a8a8a8",
        "selection": "#2878b5",
        "button": "#f7f7f7",
    },
    "Dark": {
        "ui_background": "#2b2b2b",
        "ui_surface": "#353535",
        "plot_background": "#2b2b2b",
        "text": "#e6e6e6",
        "muted_text": "#b8b8b8",
        "grid": "#626262",
        "border": "#606060",
        "selection": "#2f83c5",
        "button": "#404040",
    },
    "Blue": {
        "ui_background": "#eaf2fa",
        "ui_surface": "#f6fafe",
        "plot_background": "#f6fafe",
        "text": "#1c3f60",
        "muted_text": "#456782",
        "grid": "#a5bbd1",
        "border": "#8da9c1",
        "selection": "#2a78ad",
        "button": "#dcebf7",
    },
}

DEFAULT_CUSTOM_THEME = {
    "ui_background": "#ececec",
    "ui_surface": "#ffffff",
    "plot_background": "#ffffff",
    "text": "#202020",
    "muted_text": "#555555",
    "grid": "#b8b8b8",
    "border": "#9a9a9a",
    "selection": "#2878b5",
    "button": "#f5f5f5",
}

MARKERS = {
    "Circle": "o",
    "Square": "s",
    "Triangle up": "^",
    "Triangle down": "v",
    "Diamond": "D",
    "Pentagon": "p",
    "Hexagon": "h",
    "Star": "*",
    "Plus": "+",
    "X": "x",
}


ELEMENT_SYMBOLS = {
    "H", "He", "Li", "Be", "B", "C", "N", "O", "F", "Ne", "Na", "Mg", "Al", "Si",
    "P", "S", "Cl", "Ar", "K", "Ca", "Sc", "Ti", "V", "Cr", "Mn", "Fe", "Co", "Ni",
    "Cu", "Zn", "Ga", "Ge", "As", "Se", "Br", "Kr", "Rb", "Sr", "Y", "Zr", "Nb", "Mo",
    "Tc", "Ru", "Rh", "Pd", "Ag", "Cd", "In", "Sn", "Sb", "Te", "I", "Xe", "Cs", "Ba",
    "La", "Ce", "Pr", "Nd", "Pm", "Sm", "Eu", "Gd", "Tb", "Dy", "Ho", "Er", "Tm", "Yb",
    "Lu", "Hf", "Ta", "W", "Re", "Os", "Ir", "Pt", "Au", "Hg", "Tl", "Pb", "Bi", "Po",
    "At", "Rn", "Fr", "Ra", "Ac", "Th", "Pa", "U", "Np", "Pu", "Am", "Cm", "Bk", "Cf",
    "Es", "Fm", "Md", "No", "Lr", "Rf", "Db", "Sg", "Bh", "Hs", "Mt", "Ds", "Rg", "Cn",
    "Nh", "Fl", "Mc", "Lv", "Ts", "Og",
}
# Chemical subscripts are rendered with Matplotlib math text rather than
# Unicode subscript characters. This allows the selected plot font to remain
# Arial (or another user-selected family) without requiring Unicode subscript
# glyphs in that font.
CHEMICAL_TOKEN_PATTERN = re.compile(r"(?<![A-Za-z0-9])(?:[A-Z][a-z]?\d*)+(?![A-Za-z0-9])")
CHEMICAL_PART_PATTERN = re.compile(r"([A-Z][a-z]?)(\d*)")

# Published normalization values in ppm. Element order is intentional and is
# used as the x-axis order in the spider diagrams.
# Conventional standard atomic weights used for oxide-to-element conversion.
# Oxygen uses 15.999 and REE values follow current CIAAW/IUPAC recommendations.
SPIDER_ATOMIC_WEIGHTS = {
    "O": 15.999,
    "La": 138.90547, "Ce": 140.116, "Pr": 140.90766, "Nd": 144.242,
    "Sm": 150.36, "Eu": 151.964, "Gd": 157.249, "Tb": 158.925354,
    "Dy": 162.500, "Ho": 164.930329, "Er": 167.259, "Tm": 168.934219,
    "Yb": 173.045, "Lu": 174.96669,
}

SPIDER_NORMALIZATIONS = {
    "REE - CI chondrite (McDonough & Sun, 1995)": {
        "La": 0.237, "Ce": 0.613, "Pr": 0.0928, "Nd": 0.457,
        "Sm": 0.148, "Eu": 0.0563, "Gd": 0.199, "Tb": 0.0361,
        "Dy": 0.246, "Ho": 0.0546, "Er": 0.160, "Tm": 0.0247,
        "Yb": 0.161, "Lu": 0.0246,
    },
    "Multi-element - primitive mantle (Sun & McDonough, 1989)": {
        "Cs": 0.0079, "Rb": 0.635, "Ba": 6.989, "Th": 0.085,
        "U": 0.021, "Nb": 0.713, "Ta": 0.041, "K": 250.0,
        "La": 0.687, "Ce": 1.775, "Pb": 0.071, "Pr": 0.276,
        "Sr": 21.1, "P": 95.0, "Nd": 1.354, "Zr": 11.2,
        "Hf": 0.309, "Sm": 0.444, "Eu": 0.168, "Ti": 1300.0,
        "Gd": 0.596, "Tb": 0.108, "Dy": 0.737, "Y": 4.55,
        "Ho": 0.164, "Er": 0.480, "Yb": 0.493, "Lu": 0.074,
    },
}


class ColumnSelectionButton(QPushButton):
    """Column button whose mouse buttons have completely separate actions.

    The native QPushButton clicked/toggled machinery is intentionally bypassed.
    This guarantees that a right-click can never enter the left-click variable
    selection path, regardless of platform style or frozen-application state.
    """

    leftClicked = Signal(object)
    rightClicked = Signal(object)

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setCheckable(False)

    def mousePressEvent(self, event):
        if event.button() in (Qt.MouseButton.LeftButton, Qt.MouseButton.RightButton):
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if not self.rect().contains(event.position().toPoint()):
            event.accept()
            return
        if event.button() == Qt.MouseButton.LeftButton:
            self.leftClicked.emit(event.modifiers())
            event.accept()
            return
        if event.button() == Qt.MouseButton.RightButton:
            self.rightClicked.emit(event.position().toPoint())
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Space, Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.leftClicked.emit(event.modifiers())
            event.accept()
            return
        super().keyPressEvent(event)

    def contextMenuEvent(self, event):
        event.accept()

class ScrollableButtonRow(QWidget):
    def __init__(self, title, selected_colour, selection_callback, parent=None, multi_select=False, context_callback=None):
        super().__init__(parent)
        self.selected_colour = selected_colour
        self.selection_callback = selection_callback
        self.buttons = {}
        self.selected_column = None
        self.selected_columns = []
        self.multi_select = multi_select
        self.context_callback = context_callback

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(4)

        self.title_label = QLabel(title)
        self.title_label.setObjectName("sectionLabel")
        outer.addWidget(self.title_label)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setFrameShape(QFrame.Shape.StyledPanel)
        self.scroll_area.setMinimumHeight(70)
        self.scroll_area.setMaximumHeight(80)

        self.button_container = QWidget()
        self.button_layout = QHBoxLayout(self.button_container)
        self.button_layout.setContentsMargins(6, 6, 6, 6)
        self.button_layout.setSpacing(6)
        self.button_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.scroll_area.setWidget(self.button_container)
        outer.addWidget(self.scroll_area)

    def set_selection_colour(self, colour):
        self.selected_colour = colour
        self._refresh_styles()

    def set_columns(self, columns):
        while self.button_layout.count():
            item = self.button_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self.buttons.clear()
        self.selected_column = None
        self.selected_columns = []

        for column in columns:
            button = ColumnSelectionButton(str(column))
            button.setProperty("columnButton", True)
            button.setProperty("columnName", column)
            button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
            button.leftClicked.connect(
                lambda modifiers, name=column: self._select(name, modifiers)
            )
            if self.context_callback is not None:
                button.rightClicked.connect(
                    lambda position, source=button: self._open_context_menu_from_button(source, position)
                )
            self.button_layout.addWidget(button)
            self.buttons[column] = button
        self.button_layout.addStretch(1)
        self._refresh_styles()

    def _open_context_menu_from_button(self, source_button, position):
        """Dispatch formatting without invoking the left-click selection path."""
        if self.context_callback is None or source_button is None:
            return
        column = source_button.property("columnName")
        if column is None:
            return
        self.context_callback(column, source_button, position)

    def _select(self, column, modifiers=Qt.KeyboardModifier.NoModifier):
        if self.multi_select:
            shift_pressed = bool(modifiers & Qt.KeyboardModifier.ShiftModifier)
            if shift_pressed:
                # Shift-click toggles one variable while preserving the others.
                if column in self.selected_columns:
                    self.selected_columns.remove(column)
                else:
                    self.selected_columns.append(column)
            else:
                # A normal left-click replaces the current Y selection. Clicking
                # the sole selected variable leaves it selected rather than
                # clearing the plot.
                self.selected_columns = [column]
            self.selected_column = self.selected_columns[0] if self.selected_columns else None
            self._refresh_styles()
        else:
            self.set_selected(column)
        self.selection_callback(column)

    def set_selected(self, column):
        self.selected_column = column if column in self.buttons else None
        self.selected_columns = [self.selected_column] if self.selected_column is not None else []
        self._refresh_styles()

    def set_selected_multiple(self, columns):
        self.selected_columns = [column for column in columns if column in self.buttons]
        self.selected_column = self.selected_columns[0] if self.selected_columns else None
        self._refresh_styles()

    def clear_selection(self):
        self.selected_column = None
        self.selected_columns = []
        self._refresh_styles()

    def _refresh_styles(self):
        for column, button in self.buttons.items():
            selected = column in self.selected_columns if self.multi_select else column == self.selected_column
            button.setProperty("selectedColumn", selected)
            if selected:
                button.setStyleSheet(
                    f"QPushButton {{background-color:{self.selected_colour};color:white;"
                    "font-weight:600;padding:6px 10px;border-radius:4px;}"
                )
            else:
                button.setStyleSheet("")


class DataFrameTableModel(QAbstractTableModel):
    """Read-only, sortable Qt model for the active pandas DataFrame."""
    def __init__(self, dataframe=None, parent=None):
        super().__init__(parent)
        self._dataframe = dataframe.copy() if dataframe is not None else pd.DataFrame()

    def set_dataframe(self, dataframe):
        self.beginResetModel()
        self._dataframe = dataframe.copy() if dataframe is not None else pd.DataFrame()
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._dataframe)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._dataframe.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or role != Qt.ItemDataRole.DisplayRole:
            return None
        value = self._dataframe.iat[index.row(), index.column()]
        if pd.isna(value):
            return ""
        if isinstance(value, float):
            return f"{value:g}"
        return str(value)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return str(self._dataframe.columns[section])
        return str(self._dataframe.index[section])

    def sort(self, column, order=Qt.SortOrder.AscendingOrder):
        if not 0 <= column < len(self._dataframe.columns):
            return
        self.layoutAboutToBeChanged.emit()
        column_name = self._dataframe.columns[column]
        self._dataframe = self._dataframe.sort_values(
            column_name,
            ascending=order == Qt.SortOrder.AscendingOrder,
            kind="mergesort",
            na_position="last",
        )
        self.layoutChanged.emit()


    def source_index_at_row(self, row):
        if 0 <= row < len(self._dataframe):
            return self._dataframe.index[row]
        return None

    def row_for_source_index(self, source_index):
        matches = self._dataframe.index.get_indexer_for([source_index])
        return int(matches[0]) if len(matches) and matches[0] >= 0 else -1

class DataTableView(QTableView):
    """Spreadsheet-style read-only table with rectangular clipboard copying."""
    def keyPressEvent(self, event):
        if event.matches(event.StandardKey.Copy):
            self.copy_selection()
            return
        if event.key() == Qt.Key.Key_Escape:
            window = self.window()
            if hasattr(window, "clear_linked_selection"):
                window.clear_linked_selection()
            return
        super().keyPressEvent(event)

    def copy_selection(self):
        indexes = self.selectedIndexes()
        if not indexes:
            return
        rows = sorted({index.row() for index in indexes})
        columns = sorted({index.column() for index in indexes})
        lookup = {(index.row(), index.column()): index.data() or "" for index in indexes}
        text = "\n".join(
            "\t".join(str(lookup.get((row, column), "")) for column in columns)
            for row in rows
        )
        QApplication.clipboard().setText(text)


class CustomThemeDialog(QDialog):
    def __init__(self, colours, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Custom Colours")
        self.setModal(True)
        self.colours = dict(colours)
        self.swatches = {}

        layout = QVBoxLayout(self)
        grid = QGridLayout()
        labels = [
            ("ui_background", "UI background"),
            ("ui_surface", "UI surface"),
            ("plot_background", "Plot background"),
            ("text", "Text"),
            ("muted_text", "Secondary text"),
            ("grid", "Plot grid"),
            ("border", "Borders"),
            ("selection", "Selection highlight"),
            ("button", "Button background"),
        ]
        for row, (key, label) in enumerate(labels):
            grid.addWidget(QLabel(label), row, 0)
            swatch = QLabel(self.colours[key])
            swatch.setMinimumWidth(100)
            swatch.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.swatches[key] = swatch
            self._update_swatch(key)
            grid.addWidget(swatch, row, 1)
            choose = QPushButton("Choose...")
            choose.clicked.connect(lambda checked=False, name=key: self.choose_colour(name))
            grid.addWidget(choose, row, 2)
        layout.addLayout(grid)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def choose_colour(self, key):
        colour = QColorDialog.getColor(QColor(self.colours[key]), self, "Choose colour")
        if colour.isValid():
            self.colours[key] = colour.name()
            self._update_swatch(key)

    def _update_swatch(self, key):
        colour = QColor(self.colours[key])
        foreground = "#000000" if colour.lightness() > 140 else "#ffffff"
        self.swatches[key].setText(self.colours[key])
        self.swatches[key].setStyleSheet(
            f"background:{self.colours[key]};color:{foreground};padding:5px;border:1px solid #777;"
        )


class ExcelBivariatePlotter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("ScientificPlotTools", "ExcelBivariatePlotterPySide6")
        self.workbook_path = None
        self.excel_file = None
        self.raw_dataframe = None
        self.dataframe = None
        self.filtered_dataframe = None
        self.active_filters = []
        self.calculated_variables = []
        self.qc_horizontal_lines = []
        self.petrological_diagram = None
        self.petrological_columns = {}
        self.spider_configuration = None
        self.x_column = None
        self.y_column = None
        self.y_columns = []
        self.cartesian_y_columns = []
        self.z_column = None
        self.secondary_y_enabled = self.settings.value("axes/secondary_y_enabled", False, type=bool)
        self.secondary_y_column = self.settings.value("axes/secondary_y_column", "")
        self.secondary_axes = None
        self._artist_row_indices = {}
        self._artist_series_columns = {}
        self._plotted_row_coordinates = {}
        self._highlight_artists = []
        self._selection_sync_in_progress = False
        self.custom_x_axis_title = self.settings.value("axes/custom_x_title", "")
        self.custom_left_y_axis_title = self.settings.value("axes/custom_left_y_title", "")
        self.custom_right_y_axis_title = self.settings.value("axes/custom_right_y_title", "")
        self.custom_plot_title = self.settings.value("plot/custom_title", "")
        self.x_range_auto = self.settings.value("axes/x_range_auto", True, type=bool)
        self.left_y_range_auto = self.settings.value("axes/left_y_range_auto", True, type=bool)
        self.right_y_range_auto = self.settings.value("axes/right_y_range_auto", True, type=bool)
        self.x_range_min = float(self.settings.value("axes/x_range_min", 0.0))
        self.x_range_max = float(self.settings.value("axes/x_range_max", 1.0))
        self.left_y_range_min = float(self.settings.value("axes/left_y_range_min", 0.0))
        self.left_y_range_max = float(self.settings.value("axes/left_y_range_max", 1.0))
        self.right_y_range_min = float(self.settings.value("axes/right_y_range_min", 0.0))
        self.right_y_range_max = float(self.settings.value("axes/right_y_range_max", 1.0))
        self.series_line_formats = {}
        self.series_point_formats = {}
        self.aspect_mode = self.settings.value("plot/aspect", "Square")
        self.theme_name = self.settings.value("appearance/theme", "System")
        self.custom_theme = self._load_custom_theme()
        self.plot_colour = self.settings.value("plot/colour", "#1f77b4")
        self.stroke_colour = self.settings.value("plot/stroke_colour", "#202020")
        self.point_opacity = int(self.settings.value("plot/point_opacity", 82))
        self.stroke_width = float(self.settings.value("plot/stroke_width", 0.8))
        self.no_fill = self.settings.value("plot/no_fill", False, type=bool)
        self.no_stroke = self.settings.value("plot/no_stroke", False, type=bool)
        self.line_colour = self.settings.value("line/colour", "#1f77b4")
        self.line_width = float(self.settings.value("line/width", 1.6))
        self.line_style_name = self.settings.value("line/style", "Solid")
        self.line_opacity = int(self.settings.value("line/opacity", 100))
        self.line_show_markers = self.settings.value("line/show_markers", False, type=bool)
        self.line_marker_name = self.settings.value("line/marker_name", "Circle")
        self.line_marker_size = float(self.settings.value("line/marker_size", 6.0))
        self.vector_text_mode = self.settings.value("vector/text_mode", "Editable text")
        self.plot_font_family = self.settings.value("vector/font_family", "Arial")
        self.marker_size = int(self.settings.value("plot/marker_size", 38))
        self.format_chemical_labels = self.settings.value("plot/format_chemical_labels", True, type=bool)
        self.axis_title_size = int(self.settings.value("axes/title_size", 11))
        self.axis_title_bold = self.settings.value("axes/title_bold", False, type=bool)
        self.axis_title_colour = self.settings.value("axes/title_colour", "#202020")
        self.axis_number_size = int(self.settings.value("axes/number_size", 10))
        self.axis_number_colour = self.settings.value("axes/number_colour", "#202020")
        self.axis_number_format = self.settings.value("axes/number_format", "Automatic")
        self.axis_decimal_places = int(self.settings.value("axes/decimal_places", 2))
        self.axis_scientific_limits = int(self.settings.value("axes/scientific_limits", 4))
        self.x_zero_minimum = self.settings.value("axes/x_zero_minimum", False, type=bool)
        self.y_zero_minimum = self.settings.value("axes/y_zero_minimum", False, type=bool)
        self.regression_enabled = self.settings.value("regression/enabled", False, type=bool)
        self.regression_model = self.settings.value("regression/model", "Linear")
        self.regression_force_origin = self.settings.value("regression/force_origin", False, type=bool)
        self.regression_show_equation = self.settings.value("regression/show_equation", True, type=bool)
        self.regression_show_r2 = self.settings.value("regression/show_r2", True, type=bool)
        self.regression_show_n = self.settings.value("regression/show_n", False, type=bool)
        self.regression_colour = self.settings.value("regression/colour", "#d62728")
        self.regression_width = float(self.settings.value("regression/width", 2.0))
        self.regression_style_name = self.settings.value("regression/style", "Solid")
        self.regression_opacity = int(self.settings.value("regression/opacity", 100))
        self.regression_legend_corner = int(self.settings.value("regression/legend_corner", 0)) % 4
        self.data_legend_corner = int(self.settings.value("plot/data_legend_corner", 1)) % 4
        self.histogram_bins = int(self.settings.value("histogram/bins", 20))

        self.setWindowTitle("Geochemistry Data Plotter v11.3")
        self.resize(1280, 820)
        self.setMinimumSize(900, 650)
        self._build_interface()
        self._restore_preferences()
        self._set_data_controls_enabled(False)
        self.copy_cells_button.setEnabled(False)
        self.apply_theme(self.theme_name, save=False)
        self._update_axis_range_button()
        self._draw_placeholder()

    def _build_interface(self):
        self._build_menus()
        central = QWidget()
        self.setCentralWidget(central)
        main = QVBoxLayout(central)
        main.setContentsMargins(10, 8, 10, 8)
        main.setSpacing(8)

        file_row = QHBoxLayout()
        self.open_button = QPushButton("Open Excel File")
        self.open_button.clicked.connect(self.open_excel_file)
        file_row.addWidget(self.open_button)
        self.save_button = QPushButton("Save Plot")
        self.save_button.clicked.connect(self.save_plot)
        file_row.addWidget(self.save_button)
        self.copy_vector_button = QPushButton("Copy Vector")
        self.copy_vector_button.setToolTip("Copy the current plot as SVG vector data with a PNG fallback")
        self.copy_vector_button.clicked.connect(self.copy_plot_as_vector)
        file_row.addWidget(self.copy_vector_button)
        self.copy_png_button = QPushButton("Copy PNG")
        self.copy_png_button.setToolTip("Copy the current plot explicitly as a 300 dpi PNG image")
        self.copy_png_button.clicked.connect(self.copy_plot_as_png)
        file_row.addWidget(self.copy_png_button)
        file_row.addSpacing(18)
        file_row.addWidget(QLabel("Sheet:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(210)
        self.sheet_combo.currentTextChanged.connect(self.load_selected_sheet)
        file_row.addWidget(self.sheet_combo)
        self.clear_button = QPushButton("Clear Selections")
        self.clear_button.clicked.connect(self.clear_selections)
        file_row.addWidget(self.clear_button)
        self.calculations_button = QPushButton("Calculations...")
        self.calculations_button.setToolTip(
            "Create calculated variables such as A + B or (A + B) / C"
        )
        self.calculations_button.clicked.connect(self.open_calculations_manager)
        file_row.addWidget(self.calculations_button)

        self.regression_toggle_button = QPushButton("Show Regression")
        self.regression_toggle_button.setCheckable(True)
        self.regression_toggle_button.setChecked(self.regression_enabled)
        self.regression_toggle_button.setToolTip(
            "Show or hide the currently selected regression model"
        )
        self.regression_toggle_button.toggled.connect(self.toggle_regression)
        self.regression_toggle_button.setText(
            "Hide Regression" if self.regression_enabled else "Show Regression"
        )
        file_row.addWidget(self.regression_toggle_button)

        self.regression_legend_button = QPushButton("Regression Label: Top Left")
        self.regression_legend_button.setToolTip(
            "Cycle the regression equation and statistics through the four plot corners"
        )
        self.regression_legend_button.clicked.connect(self.cycle_regression_legend_corner)
        self.regression_legend_button.setEnabled(self.regression_enabled)
        file_row.addWidget(self.regression_legend_button)
        self._update_regression_legend_button()
        self.data_legend_button = QPushButton("Data Legend: Top Right")
        self.data_legend_button.setToolTip(
            "Cycle the multi-series data legend through the four plot corners"
        )
        self.data_legend_button.clicked.connect(self.cycle_data_legend_corner)
        self.data_legend_button.setEnabled(False)
        file_row.addWidget(self.data_legend_button)
        self._update_data_legend_button()
        file_row.addStretch(1)
        main.addLayout(file_row)

        plot_controls = QHBoxLayout()
        self.plot_type_combo = QComboBox(self)
        self.plot_type_combo.addItems(["Histogram", "Points", "Line", "Ternary"])
        self.plot_type_combo.setCurrentText(self.settings.value("plot/type", "Points"))
        self.plot_type_combo.currentTextChanged.connect(self.plot_style_changed)
        self.plot_type_combo.hide()

        # Point-formatting state widgets are intentionally not shown in the main window.
        # They are edited through Format > Data Point Formatting.
        self.marker_combo = QComboBox(self)
        self.marker_combo.addItems(MARKERS.keys())
        self.marker_combo.setCurrentText(self.settings.value("plot/marker_name", "Circle"))
        self.marker_combo.currentTextChanged.connect(self.plot_style_changed)
        self.marker_combo.hide()

        self.marker_size_spin = QSpinBox(self)
        self.marker_size_spin.setRange(1, 500)
        self.marker_size_spin.setValue(self.marker_size)
        self.marker_size_spin.valueChanged.connect(self.plot_style_changed)
        self.marker_size_spin.hide()

        self.colour_button = QPushButton("Fill colour", self)
        self.colour_button.clicked.connect(self.choose_plot_colour)
        self._update_colour_button()
        self.colour_button.hide()

        self.stroke_colour_button = QPushButton("Stroke colour", self)
        self.stroke_colour_button.clicked.connect(self.choose_stroke_colour)
        self._update_stroke_colour_button()
        self.stroke_colour_button.hide()

        self.opacity_spin = QSpinBox(self)
        self.opacity_spin.setRange(0, 100)
        self.opacity_spin.setSuffix("%")
        self.opacity_spin.setValue(self.point_opacity)
        self.opacity_spin.valueChanged.connect(self.plot_style_changed)
        self.opacity_spin.hide()

        self.stroke_width_spin = QDoubleSpinBox(self)
        self.stroke_width_spin.setRange(0.0, 10.0)
        self.stroke_width_spin.setDecimals(1)
        self.stroke_width_spin.setSingleStep(0.1)
        self.stroke_width_spin.setValue(self.stroke_width)
        self.stroke_width_spin.valueChanged.connect(self.plot_style_changed)
        self.stroke_width_spin.hide()

        self.chemical_format_checkbox = QCheckBox("Format oxide labels")
        self.chemical_format_checkbox.setChecked(self.format_chemical_labels)
        self.chemical_format_checkbox.setToolTip(
            "Display stoichiometric numbers in recognized oxide formulas as Unicode subscripts"
        )
        self.chemical_format_checkbox.toggled.connect(self.chemical_formatting_changed)
        plot_controls.addWidget(self.chemical_format_checkbox)

        # Text-formatting state widgets are edited through Format > Text Formatting.
        self.vector_text_combo = QComboBox(self)
        self.vector_text_combo.addItems(["Editable text", "Text as curves"])
        self.vector_text_combo.setCurrentText(self.vector_text_mode)
        self.vector_text_combo.currentTextChanged.connect(self.vector_settings_changed)
        self.vector_text_combo.hide()

        self.font_combo = QComboBox(self)
        self.font_combo.setEditable(True)
        font_families = QFontDatabase.families()
        self.font_combo.addItems(font_families)
        if self.plot_font_family in font_families:
            self.font_combo.setCurrentText(self.plot_font_family)
        elif "Arial" in font_families:
            self.font_combo.setCurrentText("Arial")
        else:
            self.font_combo.setCurrentText("DejaVu Sans")
        self.font_combo.currentTextChanged.connect(self.vector_settings_changed)
        self.font_combo.hide()

        plot_controls.addSpacing(18)
        plot_controls.addWidget(QLabel("Plot ratio:"))
        self.aspect_button = QPushButton(self.aspect_mode)
        self.aspect_button.clicked.connect(self.toggle_aspect_ratio)
        plot_controls.addWidget(self.aspect_button)
        self.reset_axis_ranges_button = QPushButton("Reset Axes to Auto")
        self.reset_axis_ranges_button.setToolTip(
            "Clear all manual X, left-Y, and right-Y axis limits"
        )
        self.reset_axis_ranges_button.clicked.connect(self.reset_axis_ranges_to_auto)
        plot_controls.addWidget(self.reset_axis_ranges_button)
        self.clear_highlight_button = QPushButton("Clear Highlight")
        self.clear_highlight_button.setToolTip("Clear linked plot and table selections (Esc)")
        self.clear_highlight_button.setShortcut("Esc")
        self.clear_highlight_button.clicked.connect(self.clear_linked_selection)
        self.clear_highlight_button.setEnabled(False)
        plot_controls.addWidget(self.clear_highlight_button)

        self.x_zero_checkbox = QCheckBox("X min = 0")
        self.x_zero_checkbox.setChecked(self.x_zero_minimum)
        self.x_zero_checkbox.setToolTip("Force the lower X-axis limit to zero")
        self.x_zero_checkbox.toggled.connect(self.axis_minimum_changed)
        plot_controls.addWidget(self.x_zero_checkbox)

        self.y_zero_checkbox = QCheckBox("Y min = 0")
        self.y_zero_checkbox.setChecked(self.y_zero_minimum)
        self.y_zero_checkbox.setToolTip("Force the lower Y-axis limit to zero")
        self.y_zero_checkbox.toggled.connect(self.axis_minimum_changed)
        plot_controls.addWidget(self.y_zero_checkbox)
        plot_controls.addStretch(1)
        main.addLayout(plot_controls)

        summary = QHBoxLayout()
        self.x_label = QLabel("X axis: not selected")
        self.x_label.setObjectName("xAxisLabel")
        self.y_label = QLabel("Y axis: not selected")
        self.y_label.setObjectName("yAxisLabel")
        summary.addWidget(self.x_label)
        summary.addSpacing(24)
        summary.addWidget(self.y_label)
        summary.addStretch(1)
        main.addLayout(summary)

        self.x_button_row = ScrollableButtonRow("Choose X-axis column", "#1f77b4", self.select_x)
        self.y_button_row = ScrollableButtonRow(
            "Choose Y-axis column (Shift-click to add or remove)", "#d62728", self.select_y,
            multi_select=True, context_callback=self.open_y_variable_formatting
        )
        self.z_button_row = ScrollableButtonRow("Choose component C", "#2ca02c", self.select_z)
        main.addWidget(self.x_button_row)
        main.addWidget(self.y_button_row)
        main.addWidget(self.z_button_row)
        self.z_button_row.hide()

        self.main_tabs = QTabWidget()
        self.main_tabs.setDocumentMode(True)

        self.plot_tab = QWidget()
        plot_tab_layout = QVBoxLayout(self.plot_tab)
        plot_tab_layout.setContentsMargins(0, 6, 0, 0)
        plot_tab_layout.setSpacing(4)
        self.figure = Figure(dpi=100, constrained_layout=True)
        self.axes = self.figure.add_subplot(111)
        self.canvas = FigureCanvas(self.figure)
        self.canvas.mpl_connect("pick_event", self._on_plot_pick)
        self.canvas.mpl_connect("button_press_event", self._on_plot_mouse_press)
        self.canvas.setMinimumSize(600, 350)
        self.canvas.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        plot_tab_layout.addWidget(self.canvas, stretch=1)
        self.navigation_toolbar = NavigationToolbar(self.canvas, self)
        plot_tab_layout.addWidget(self.navigation_toolbar)
        self.main_tabs.addTab(self.plot_tab, "Plot")

        self.data_tab = QWidget()
        data_tab_layout = QVBoxLayout(self.data_tab)
        data_tab_layout.setContentsMargins(6, 6, 6, 6)
        data_header = QHBoxLayout()
        self.data_table_summary = QLabel("No worksheet loaded")
        self.data_table_summary.setObjectName("sectionLabel")
        data_header.addWidget(self.data_table_summary)
        data_header.addStretch(1)
        self.copy_cells_button = QPushButton("Copy Selected Cells")
        self.copy_cells_button.clicked.connect(lambda: self.data_table.copy_selection())
        data_header.addWidget(self.copy_cells_button)
        data_tab_layout.addLayout(data_header)
        self.data_table_model = DataFrameTableModel()
        self.data_table = DataTableView()
        self.data_table.setModel(self.data_table_model)
        self.data_table.setSortingEnabled(True)
        self.data_table.setAlternatingRowColors(True)
        self.data_table.setSelectionMode(QTableView.SelectionMode.ExtendedSelection)
        self.data_table.setSelectionBehavior(QTableView.SelectionBehavior.SelectItems)
        self.data_table.selectionModel().selectionChanged.connect(self._on_table_selection_changed)
        self.data_table.setWordWrap(False)
        self.data_table.horizontalHeader().setStretchLastSection(False)
        data_tab_layout.addWidget(self.data_table, stretch=1)
        self.main_tabs.addTab(self.data_tab, "Data")
        main.addWidget(self.main_tabs, stretch=1)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Open an Excel workbook to begin.")
        self.marker_combo.setEnabled(self.plot_type_combo.currentText() in ("Points", "Ternary"))
        self._update_plot_type_controls()

    def _build_menus(self):
        MENU_WIDTH = 260

        file_menu = self.menuBar().addMenu("File")
        file_menu.setMinimumWidth(MENU_WIDTH)
        open_action = QAction("Open Excel File", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_excel_file)
        file_menu.addAction(open_action)
        save_action = QAction("Save Plot", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_plot)
        file_menu.addAction(save_action)
        file_menu.addSeparator()
        print_preview_action = QAction("Print Preview...", self)
        print_preview_action.setShortcut("Ctrl+Shift+P")
        print_preview_action.triggered.connect(self.print_preview)
        file_menu.addAction(print_preview_action)
        print_action = QAction("Print...", self)
        print_action.setShortcut("Ctrl+P")
        print_action.triggered.connect(self.print_plot)
        file_menu.addAction(print_action)
        file_menu.addSeparator()
        copy_vector_action = QAction("Copy Plot as Vector (SVG)", self)
        copy_vector_action.setShortcut("Ctrl+Shift+C")
        copy_vector_action.triggered.connect(self.copy_plot_as_vector)
        file_menu.addAction(copy_vector_action)
        copy_png_action = QAction("Copy Plot as PNG", self)
        copy_png_action.setShortcut("Ctrl+Alt+C")
        copy_png_action.triggered.connect(self.copy_plot_as_png)
        file_menu.addAction(copy_png_action)
        file_menu.addSeparator()
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        plot_menu = self.menuBar().addMenu("Plot")
        plot_menu.setMinimumWidth(MENU_WIDTH)
        self.plot_type_actions = {}
        plot_type_group = QActionGroup(self)
        plot_type_group.setExclusive(True)
        saved_plot_type = self.settings.value("plot/type", "Points")
        for plot_name in ("Histogram", "Points", "Line", "Ternary"):
            action = QAction(plot_name, self, checkable=True)
            action.setChecked(plot_name == saved_plot_type)
            action.triggered.connect(
                lambda checked=False, name=plot_name: self.set_plot_type(name)
            )
            plot_type_group.addAction(action)
            plot_menu.addAction(action)
            self.plot_type_actions[plot_name] = action
        plot_menu.addSeparator()
        petrology_menu = plot_menu.addMenu("Petrological Diagrams")
        standard_plot_action = QAction("Standard Plot", self)
        standard_plot_action.triggered.connect(self.clear_petrological_diagram)
        petrology_menu.addAction(standard_plot_action)
        petrology_menu.addSeparator()
        spider_action = QAction("Spider Plot...", self)
        spider_action.triggered.connect(self.configure_spider_plot)
        petrology_menu.addAction(spider_action)
        petrology_menu.addSeparator()
        for diagram_name in (
            "TAS: Total Alkali-Silica",
            "AFM: Alkali-Fe-Mg",
            "QAP: Quartz-Alkali Feldspar-Plagioclase",
            "FAP: Feldspathoid-Alkali Feldspar-Plagioclase",
        ):
            action = QAction(diagram_name, self)
            action.triggered.connect(
                lambda checked=False, name=diagram_name: self.configure_petrological_diagram(name)
            )
            petrology_menu.addAction(action)
        data_menu = self.menuBar().addMenu("Data")
        data_menu.setMinimumWidth(MENU_WIDTH)
        filter_action = QAction("Filters...", self)
        filter_action.setShortcut("Ctrl+F")
        filter_action.triggered.connect(self.open_filter_dialog)
        data_menu.addAction(filter_action)
        reset_filters_action = QAction("Reset Filters", self)
        reset_filters_action.triggered.connect(self.reset_filters)
        data_menu.addAction(reset_filters_action)
        data_menu.addSeparator()
        calculations_action = QAction("Calculated Variables...", self)
        calculations_action.setShortcut("Ctrl+M")
        calculations_action.triggered.connect(self.open_calculations_manager)
        data_menu.addAction(calculations_action)

        qc_menu = self.menuBar().addMenu("QC Checks")
        qc_menu.setMinimumWidth(MENU_WIDTH)
        horizontal_lines_menu = qc_menu.addMenu("Horizontal Lines")
        manage_horizontal_lines_action = QAction("Manage Lines...", self)
        manage_horizontal_lines_action.triggered.connect(self.open_horizontal_lines_dialog)
        horizontal_lines_menu.addAction(manage_horizontal_lines_action)
        horizontal_lines_menu.addSeparator()
        clear_horizontal_lines_action = QAction("Remove All Lines", self)
        clear_horizontal_lines_action.triggered.connect(self.clear_horizontal_qc_lines)
        horizontal_lines_menu.addAction(clear_horizontal_lines_action)

        format_menu = self.menuBar().addMenu("Format")
        format_menu.setMinimumWidth(MENU_WIDTH)
        histogram_format_action = QAction("Histogram Formatting...", self)
        histogram_format_action.triggered.connect(self.open_histogram_formatting_dialog)
        format_menu.addAction(histogram_format_action)
        point_format_action = QAction("Data Point Formatting...", self)
        point_format_action.triggered.connect(self.open_data_point_formatting_dialog)
        format_menu.addAction(point_format_action)

        line_format_action = QAction("Data Line Formatting...", self)
        line_format_action.triggered.connect(self.open_line_formatting_dialog)
        format_menu.addAction(line_format_action)

        text_format_action = QAction("Text Formatting...", self)
        text_format_action.triggered.connect(self.open_text_formatting_dialog)
        format_menu.addAction(text_format_action)

        axis_format_action = QAction("Axis Formatting...", self)
        axis_format_action.triggered.connect(self.open_axis_formatting_dialog)
        format_menu.addAction(axis_format_action)
        axis_ranges_action = QAction("Axis Ranges...", self)
        axis_ranges_action.triggered.connect(self.open_axis_ranges_dialog)
        format_menu.addAction(axis_ranges_action)

        regression_menu = self.menuBar().addMenu("Regression")
        regression_menu.setMinimumWidth(MENU_WIDTH)
        regression_options_action = QAction("Regression Options...", self)
        regression_options_action.triggered.connect(self.open_regression_dialog)
        regression_menu.addAction(regression_options_action)
        regression_menu.addSeparator()
        model_menu = regression_menu.addMenu("Regression Model")
        self.regression_model_actions = {}
        model_group = QActionGroup(self)
        model_group.setExclusive(True)
        for model_name in ("Linear", "Quadratic", "Cubic", "Exponential", "Logarithmic", "Power"):
            action = QAction(model_name, self, checkable=True)
            action.setChecked(model_name == self.regression_model)
            action.triggered.connect(
                lambda checked=False, name=model_name: self.set_regression_model(name)
            )
            model_group.addAction(action)
            model_menu.addAction(action)
            self.regression_model_actions[model_name] = action

        theme_menu = self.menuBar().addMenu("Theme")
        theme_menu.setMinimumWidth(MENU_WIDTH)
        self.theme_actions = {}
        group = QActionGroup(self)
        group.setExclusive(True)
        for name in ("System", "Light", "Dark", "Blue"):
            action = QAction(name, self, checkable=True)
            action.triggered.connect(lambda checked=False, theme=name: self.apply_theme(theme))
            group.addAction(action)
            theme_menu.addAction(action)
            self.theme_actions[name] = action
        theme_menu.addSeparator()
        custom_action = QAction("Custom Colours...", self)
        custom_action.triggered.connect(self.edit_custom_theme)
        theme_menu.addAction(custom_action)

    def _load_custom_theme(self):
        result = dict(DEFAULT_CUSTOM_THEME)
        for key, default in result.items():
            result[key] = self.settings.value(f"appearance/custom/{key}", default)
        return result

    def _restore_preferences(self):
        geometry = self.settings.value("window/geometry")
        if geometry is not None:
            self.restoreGeometry(geometry)

    def closeEvent(self, event):
        self.settings.setValue("window/geometry", self.saveGeometry())
        self.settings.setValue("plot/aspect", self.aspect_mode)
        self.settings.setValue("plot/type", self.plot_type_combo.currentText())
        self.settings.setValue("plot/data_legend_corner", self.data_legend_corner)
        if hasattr(self, "plot_type_actions"):
            for name, action in self.plot_type_actions.items():
                action.setChecked(name == self.plot_type_combo.currentText())
        self.settings.setValue("plot/marker_name", self.marker_combo.currentText())
        self.settings.setValue("plot/marker_size", self.marker_size_spin.value())
        self.settings.setValue("plot/format_chemical_labels", self.chemical_format_checkbox.isChecked())
        self.settings.setValue("plot/colour", self.plot_colour)
        self.settings.setValue("plot/stroke_colour", self.stroke_colour)
        self.settings.setValue("plot/point_opacity", self.opacity_spin.value())
        self.settings.setValue("plot/stroke_width", self.stroke_width_spin.value())
        self.settings.setValue("plot/no_fill", self.no_fill)
        self.settings.setValue("plot/no_stroke", self.no_stroke)
        self.settings.setValue("line/colour", self.line_colour)
        self.settings.setValue("line/width", self.line_width)
        self.settings.setValue("line/style", self.line_style_name)
        self.settings.setValue("line/opacity", self.line_opacity)
        self.settings.setValue("line/show_markers", self.line_show_markers)
        self.settings.setValue("line/marker_name", self.line_marker_name)
        self.settings.setValue("line/marker_size", self.line_marker_size)
        self.settings.setValue("axes/x_zero_minimum", self.x_zero_checkbox.isChecked())
        self.settings.setValue("axes/y_zero_minimum", self.y_zero_checkbox.isChecked())
        self.settings.setValue("axes/secondary_y_enabled", self.secondary_y_enabled)
        self.settings.setValue("axes/secondary_y_column", self.secondary_y_column)
        self.settings.setValue("axes/custom_x_title", self.custom_x_axis_title)
        self.settings.setValue("axes/custom_left_y_title", self.custom_left_y_axis_title)
        self.settings.setValue("axes/custom_right_y_title", self.custom_right_y_axis_title)
        self.settings.setValue("plot/custom_title", self.custom_plot_title)
        self._save_axis_range_settings()
        self.settings.setValue("vector/text_mode", self.vector_text_combo.currentText())
        self.settings.setValue("vector/font_family", self.font_combo.currentText())
        self.settings.setValue("appearance/theme", self.theme_name)
        self._save_axis_format_settings()
        self._save_regression_settings()
        super().closeEvent(event)

    def _system_theme(self):
        palette = QApplication.palette()
        window = palette.color(QPalette.ColorRole.Window)
        surface = palette.color(QPalette.ColorRole.Base)
        text = palette.color(QPalette.ColorRole.WindowText)
        muted = palette.color(QPalette.ColorRole.PlaceholderText)
        border = palette.color(QPalette.ColorRole.Mid)
        selection = palette.color(QPalette.ColorRole.Highlight)
        button = palette.color(QPalette.ColorRole.Button)
        grid = palette.color(QPalette.ColorRole.Mid)
        return {
            "ui_background": window.name(),
            "ui_surface": surface.name(),
            "plot_background": surface.name(),
            "text": text.name(),
            "muted_text": muted.name(),
            "grid": grid.name(),
            "border": border.name(),
            "selection": selection.name(),
            "button": button.name(),
        }

    def apply_theme(self, name, save=True):
        if name not in THEMES and name != "Custom":
            name = "System"
        self.theme_name = name
        colours = self.custom_theme if name == "Custom" else THEMES.get(name)
        if colours is None:
            colours = self._system_theme()

        app = QApplication.instance()
        if name == "System":
            app.setPalette(app.style().standardPalette())
            colours = self._system_theme()
            app.setStyleSheet(
                f"QTabWidget::pane {{background:{colours['ui_surface']};border:1px solid {colours['border']};}}"
                f"QTabBar::tab {{background:{colours['button']};color:{colours['text']};border:1px solid {colours['border']};padding:6px 14px;}}"
                f"QTabBar::tab:selected {{background:{colours['ui_surface']};color:{colours['text']};}}"
                f"QLineEdit,QAbstractSpinBox {{background:{colours['ui_surface']};color:{colours['text']};border:1px solid {colours['border']};border-radius:3px;padding:4px;selection-background-color:{colours['selection']};selection-color:#ffffff;}}"
                f"QLineEdit:disabled,QAbstractSpinBox:disabled {{background:{colours['ui_background']};color:{colours['muted_text']};}}"
                f"QTableView {{background:{colours['ui_surface']};alternate-background-color:{colours['ui_background']};color:{colours['text']};gridline-color:{colours['border']};selection-background-color:{colours['selection']};selection-color:#ffffff;}}"
                f"QHeaderView::section {{background:{colours['button']};color:{colours['text']};border:1px solid {colours['border']};padding:5px;font-weight:600;}}"
            )
        else:
            palette = QPalette()
            palette.setColor(QPalette.ColorRole.Window, QColor(colours["ui_background"]))
            palette.setColor(QPalette.ColorRole.WindowText, QColor(colours["text"]))
            palette.setColor(QPalette.ColorRole.Base, QColor(colours["ui_surface"]))
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor(colours["ui_background"]))
            palette.setColor(QPalette.ColorRole.Text, QColor(colours["text"]))
            palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(colours["muted_text"]))
            palette.setColor(QPalette.ColorRole.BrightText, QColor(colours["text"]))
            palette.setColor(QPalette.ColorRole.Button, QColor(colours["button"]))
            palette.setColor(QPalette.ColorRole.ButtonText, QColor(colours["text"]))
            palette.setColor(QPalette.ColorRole.Highlight, QColor(colours["selection"]))
            palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
            palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(colours["ui_surface"]))
            palette.setColor(QPalette.ColorRole.ToolTipText, QColor(colours["text"]))
            app.setPalette(palette)
            app.setStyleSheet(
                f"QPushButton {{background:{colours['button']};color:{colours['text']};"
                f"border:1px solid {colours['border']};border-radius:4px;padding:5px 9px;}}"
                f"QPushButton:hover {{border:1px solid {colours['selection']};}}"
                f"QComboBox {{background:{colours['ui_surface']};color:{colours['text']};"
                f"border:1px solid {colours['border']};padding:4px;}}"
                f"QLineEdit,QAbstractSpinBox {{background:{colours['ui_surface']};color:{colours['text']};border:1px solid {colours['border']};border-radius:3px;padding:4px;selection-background-color:{colours['selection']};selection-color:#ffffff;}}"
                f"QLineEdit:disabled,QAbstractSpinBox:disabled {{background:{colours['ui_background']};color:{colours['muted_text']};}}"
                f"QMenuBar,QMenu,QStatusBar {{background:{colours['ui_background']};color:{colours['text']};}}"
                f"QScrollArea {{background:{colours['ui_surface']};border:1px solid {colours['border']};}}"
                f"QTabWidget::pane {{background:{colours['ui_surface']};border:1px solid {colours['border']};}}"
                f"QTabBar::tab {{background:{colours['button']};color:{colours['text']};border:1px solid {colours['border']};padding:6px 14px;}}"
                f"QTabBar::tab:selected {{background:{colours['ui_surface']};color:{colours['text']};border-bottom-color:{colours['ui_surface']};}}"
                f"QTableView {{background:{colours['ui_surface']};alternate-background-color:{colours['ui_background']};color:{colours['text']};gridline-color:{colours['border']};selection-background-color:{colours['selection']};selection-color:#ffffff;}}"
                f"QTableView::item {{color:{colours['text']};padding:3px;}}"
                f"QHeaderView::section {{background:{colours['button']};color:{colours['text']};border:1px solid {colours['border']};padding:5px;font-weight:600;}}"
                f"QLabel#sectionLabel {{font-weight:600;color:{colours['text']};}}"
            )

        self.current_theme = colours
        if hasattr(self, "data_table"):
            table_palette = self.data_table.palette()
            table_palette.setColor(QPalette.ColorRole.Base, QColor(colours["ui_surface"]))
            table_palette.setColor(QPalette.ColorRole.AlternateBase, QColor(colours["ui_background"]))
            table_palette.setColor(QPalette.ColorRole.Text, QColor(colours["text"]))
            table_palette.setColor(QPalette.ColorRole.Highlight, QColor(colours["selection"]))
            table_palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
            self.data_table.setPalette(table_palette)
            self.data_table.viewport().setPalette(table_palette)
            self.data_table.viewport().update()
        self.x_button_row.set_selection_colour(colours["selection"])
        self.y_button_row.set_selection_colour("#d65a5a" if name == "Dark" else "#d62728")
        self.z_button_row.set_selection_colour("#55aa66" if name == "Dark" else "#2ca02c")
        self.x_label.setStyleSheet(f"color:{colours['selection']};font-weight:600;")
        self.y_label.setStyleSheet(
            f"color:{'#ff7b7b' if name == 'Dark' else '#d62728'};font-weight:600;"
        )
        self._apply_plot_theme()
        for key, action in self.theme_actions.items():
            action.setChecked(key == name)
        if save:
            self.settings.setValue("appearance/theme", name)

    def edit_custom_theme(self):
        dialog = CustomThemeDialog(self.custom_theme, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.custom_theme = dialog.colours
            for key, value in self.custom_theme.items():
                self.settings.setValue(f"appearance/custom/{key}", value)
            self.apply_theme("Custom")

    def _apply_plot_theme(self):
        if not hasattr(self, "axes") or not hasattr(self, "current_theme"):
            return
        c = self.current_theme
        self.figure.patch.set_facecolor(c["ui_background"])
        self.axes.set_facecolor(c["plot_background"])
        self.axes.tick_params(colors=c["text"])
        self.axes.xaxis.label.set_color(c["text"])
        self.axes.yaxis.label.set_color(c["text"])
        self.axes.title.set_color(c["text"])
        for spine in self.axes.spines.values():
            spine.set_color(c["border"])
        self.axes.grid(True, color=c["grid"], alpha=0.55)
        self._apply_plot_font()
        self._apply_axis_formatting()
        self._apply_custom_plot_title()
        self.canvas.draw_idle()

    def _set_data_controls_enabled(self, enabled):
        self.clear_button.setEnabled(enabled)
        self.save_button.setEnabled(enabled)
        self.copy_vector_button.setEnabled(enabled)
        self.copy_png_button.setEnabled(enabled)
        self.x_button_row.setEnabled(enabled)
        self.y_button_row.setEnabled(enabled)
        self.z_button_row.setEnabled(enabled)
        self.calculations_button.setEnabled(enabled)

    def refresh_data_table(self):
        """Show the active filtered dataframe, including calculated variables."""
        active = self.filtered_dataframe if self.filtered_dataframe is not None else self.dataframe
        self.data_table_model.set_dataframe(active)
        if active is None:
            self.data_table_summary.setText("No worksheet loaded")
            self.copy_cells_button.setEnabled(False)
            return
        self.copy_cells_button.setEnabled(True)
        base_count = len(self.dataframe) if self.dataframe is not None else len(active)
        filter_text = f"Filtered view: {len(active):,} of {base_count:,} rows" if self.active_filters else f"{len(active):,} rows"
        calculated_count = len(self.calculated_variables)
        calculated_text = f" | {calculated_count} calculated variable(s)" if calculated_count else ""
        self.data_table_summary.setText(
            f"{filter_text} | {len(active.columns):,} columns{calculated_text}"
        )
        self.data_table.resizeColumnsToContents()
        for column in range(self.data_table_model.columnCount()):
            if self.data_table.columnWidth(column) > 220:
                self.data_table.setColumnWidth(column, 220)

    def clear_linked_selection(self):
        if not hasattr(self, "data_table"):
            return
        self._selection_sync_in_progress = True
        try:
            self.data_table.clearSelection()
            self.data_table.setCurrentIndex(QModelIndex())
        finally:
            self._selection_sync_in_progress = False
        self._clear_selection_highlights()
        self.clear_highlight_button.setEnabled(False)
        self.canvas.draw_idle()
        self.status_bar.showMessage("Point and table highlights cleared.")

    def _on_plot_mouse_press(self, event):
        if event.button != 1:
            return

        # The title lies partly outside the axes, so test it before requiring
        # event.inaxes. A direct click opens an editable title dialog.
        try:
            title_hit, _details = self.axes.title.contains(event)
        except Exception:
            title_hit = False
        if title_hit:
            self._edit_plot_title()
            return

        # Clicking an axis title, tick label, or the narrow tick-label margin
        # opens the existing Axis Formatting dialog. Data-area clicks retain
        # their normal point-selection and clear-selection behaviour.
        if self._axis_format_hit(event):
            self.open_axis_formatting_dialog()
            return

        if event.inaxes is None:
            return
        for artist in self._artist_row_indices:
            try:
                contains, _details = artist.contains(event)
            except Exception:
                contains = False
            if contains:
                return
        self.clear_linked_selection()

    def _edit_plot_title(self):
        current = self.axes.get_title()
        title, accepted = QInputDialog.getText(
            self, "Plot Title", "Plot title:", QLineEdit.EchoMode.Normal, current
        )
        if not accepted:
            return
        self.custom_plot_title = title.strip()
        self.settings.setValue("plot/custom_title", self.custom_plot_title)
        self.axes.set_title(self.custom_plot_title)
        self.canvas.draw_idle()
        self.status_bar.showMessage(
            "Plot title updated." if self.custom_plot_title else "Plot title cleared."
        )

    def _axis_format_hit(self, event):
        if event.x is None or event.y is None:
            return False
        axes_to_check = [self.axes]
        if self.secondary_axes is not None:
            axes_to_check.append(self.secondary_axes)
        for axes in axes_to_check:
            text_objects = [axes.xaxis.label, axes.yaxis.label]
            text_objects.extend(axes.get_xticklabels())
            text_objects.extend(axes.get_yticklabels())
            for text in text_objects:
                if not text.get_visible() or not text.get_text():
                    continue
                try:
                    hit, _details = text.contains(event)
                except Exception:
                    hit = False
                if hit:
                    return True
        return False

    def _apply_custom_plot_title(self):
        if self.custom_plot_title:
            self.axes.set_title(self.custom_plot_title)

    def _reset_custom_plot_title(self):
        """Return plot titles to the automatic title for the active plot."""
        if not self.custom_plot_title:
            return
        self.custom_plot_title = ""
        self.settings.setValue("plot/custom_title", "")

    def _clear_selection_highlights(self):
        for artist in self._highlight_artists:
            try:
                artist.remove()
            except Exception:
                pass
        self._highlight_artists = []

    def _selected_table_source_indices(self):
        selection_model = self.data_table.selectionModel()
        if selection_model is None:
            return []
        rows = sorted({index.row() for index in selection_model.selectedIndexes()})
        result = []
        for row in rows:
            source_index = self.data_table_model.source_index_at_row(row)
            if source_index is not None:
                result.append(source_index)
        return result

    def _highlight_rows_on_plot(self, source_indices):
        """Highlight selected rows as points or complete spider patterns."""
        import numpy as np

        self._clear_selection_highlights()
        if not source_indices or not self._plotted_row_coordinates:
            self.canvas.draw_idle()
            return

        point_coordinates = {}
        pattern_coordinates = []
        for source_index in source_indices:
            for axes, x_value, y_value in self._plotted_row_coordinates.get(source_index, []):
                x_array = np.asarray(x_value)
                y_array = np.asarray(y_value)
                if x_array.ndim > 0 and x_array.size > 1:
                    pattern_coordinates.append((axes, x_array.astype(float), y_array.astype(float)))
                else:
                    point_coordinates.setdefault(axes, [[], []])
                    point_coordinates[axes][0].append(float(x_array.reshape(-1)[0]))
                    point_coordinates[axes][1].append(float(y_array.reshape(-1)[0]))

        for axes, (x_values, y_values) in point_coordinates.items():
            highlight = axes.scatter(
                x_values, y_values, s=max(80, self.marker_size_spin.value() * 1.8),
                marker="o", facecolors="#fff200", edgecolors="#000000",
                linewidths=1.8, alpha=1.0, zorder=20,
            )
            self._highlight_artists.append(highlight)

        for axes, x_values, y_values in pattern_coordinates:
            highlight, = axes.plot(
                x_values, y_values, color="#fff200", linewidth=3.0,
                marker="o", markersize=7.0, markerfacecolor="#fff200",
                markeredgecolor="#000000", markeredgewidth=1.2,
                alpha=1.0, zorder=20,
            )
            self._highlight_artists.append(highlight)

        self.canvas.draw_idle()

    def _select_table_source_index(self, source_index):
        row = self.data_table_model.row_for_source_index(source_index)
        if row < 0:
            return
        self._selection_sync_in_progress = True
        try:
            self.data_table.clearSelection()
            self.data_table.selectRow(row)
            self.data_table.scrollTo(self.data_table_model.index(row, 0), QTableView.ScrollHint.PositionAtCenter)
        finally:
            self._selection_sync_in_progress = False
        self._highlight_rows_on_plot([source_index])
        self.clear_highlight_button.setEnabled(True)

    def _on_plot_pick(self, event):
        """Handle mouse actions on plotted data artists.

        Left-click selects the corresponding worksheet row. Right-click opens
        formatting for the exact plotted series and must never select a row.
        """
        mouse_event = getattr(event, "mouseevent", None)
        mouse_button = getattr(mouse_event, "button", None)

        if mouse_button == 3:  # Matplotlib right mouse button
            column = self._artist_series_columns.get(event.artist)
            if column is not None:
                self.open_y_variable_formatting(column, None, None)
                self.status_bar.showMessage(
                    f"Opened formatting for plotted series: {column}"
                )
            return

        if mouse_button != 1:  # Only a left-click may select a data row.
            return

        row_indices = self._artist_row_indices.get(event.artist)
        if row_indices is None or not len(event.ind):
            return
        position = int(event.ind[0])
        if 0 <= position < len(row_indices):
            self._select_table_source_index(row_indices[position])
            self.status_bar.showMessage(
                f"Selected worksheet row {row_indices[position]} from the plot."
            )

    def _on_table_selection_changed(self, _selected, _deselected):
        if self._selection_sync_in_progress:
            return
        source_indices = self._selected_table_source_indices()
        self._highlight_rows_on_plot(source_indices)
        self.clear_highlight_button.setEnabled(bool(source_indices))
        if source_indices:
            self.status_bar.showMessage(f"Highlighted {len(source_indices):,} selected table row(s) on the plot.")

    def open_excel_file(self):
        start_dir = self.settings.value("files/last_directory", "")
        path, _ = QFileDialog.getOpenFileName(
            self, "Select an Excel workbook", start_dir,
            "Excel workbooks (*.xlsx *.xlsm *.xls);;All files (*.*)"
        )
        if not path:
            return
        try:
            excel_file = pd.ExcelFile(path)
        except Exception as exc:
            QMessageBox.critical(self, "Unable to open workbook", str(exc))
            return
        if not excel_file.sheet_names:
            QMessageBox.warning(self, "No worksheets", "The workbook has no readable worksheets.")
            return
        self.excel_file = excel_file
        self.workbook_path = Path(path)
        self.settings.setValue("files/last_directory", str(self.workbook_path.parent))
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        self.sheet_combo.addItems(excel_file.sheet_names)
        self.sheet_combo.setCurrentIndex(0)
        self.sheet_combo.blockSignals(False)
        self.load_selected_sheet(excel_file.sheet_names[0])

    def load_selected_sheet(self, sheet_name):
        if self.excel_file is None or not sheet_name:
            return
        try:
            dataframe = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=0)
        except Exception as exc:
            QMessageBox.critical(self, "Unable to read worksheet", str(exc))
            return
        if dataframe.empty or len(dataframe.columns) == 0:
            QMessageBox.warning(self, "Empty worksheet", "The worksheet contains no usable data.")
            return
        dataframe.columns = self._make_unique_column_names(dataframe.columns)
        self._reset_custom_plot_title()
        self.raw_dataframe = dataframe.copy()
        self.calculated_variables = []
        self.dataframe = dataframe.copy()
        self.filtered_dataframe = dataframe.copy()
        self.active_filters = []
        columns = self.dataframe.columns.tolist()
        self.x_button_row.set_columns(columns)
        self.y_button_row.set_columns(columns)
        self.z_button_row.set_columns(columns)
        self.x_column = columns[0]
        self.y_column = None
        self.y_columns = []
        self.cartesian_y_columns = []
        self.z_column = None
        self.x_button_row.set_selected(self.x_column)
        self.y_button_row.clear_selection()
        self.z_button_row.clear_selection()
        self.x_label.setText(f"X axis: {self.x_column}")
        self.y_label.setText("Y axis: not selected")
        self._set_data_controls_enabled(True)
        self.refresh_data_table()
        self._draw_placeholder()
        self.status_bar.showMessage(
            f"Loaded {self.workbook_path.name} | Sheet: {sheet_name} | "
            f"{len(dataframe):,} rows, {len(columns):,} columns | Default X: {self.x_column}"
        )

    @staticmethod
    def _make_unique_column_names(columns):
        names, counts = [], {}
        for index, value in enumerate(columns, start=1):
            base = str(value).strip()
            if not base or base.lower().startswith("unnamed:"):
                base = f"Column {index}"
            counts[base] = counts.get(base, 0) + 1
            names.append(base if counts[base] == 1 else f"{base} ({counts[base]})")
        return names

    @staticmethod
    def _calculated_column_name(name):
        return f"Calc: {name.strip()}"

    @staticmethod
    def _evaluate_calculation_tokens(dataframe, tokens):
        """Evaluate a button-built arithmetic expression without using eval()."""
        aliases = {}
        code_parts = []
        for token_type, value in tokens:
            if token_type == "column":
                alias = f"c{len(aliases)}"
                aliases[alias] = value
                code_parts.append(alias)
            else:
                code_parts.append(value)
        expression = " ".join(code_parts)
        if not expression.strip():
            raise ValueError("The expression is empty.")
        tree = ast.parse(expression, mode="eval")
        allowed_binary = (ast.Add, ast.Sub, ast.Mult, ast.Div)
        def calculate(node):
            if isinstance(node, ast.Expression):
                return calculate(node.body)
            if isinstance(node, ast.Name) and node.id in aliases:
                column = aliases[node.id]
                if column not in dataframe.columns:
                    raise ValueError(f"The referenced variable '{column}' is unavailable.")
                return pd.to_numeric(dataframe[column], errors="coerce")
            if isinstance(node, ast.BinOp) and isinstance(node.op, allowed_binary):
                left = calculate(node.left)
                right = calculate(node.right)
                if isinstance(node.op, ast.Add):
                    return left + right
                if isinstance(node.op, ast.Sub):
                    return left - right
                if isinstance(node.op, ast.Mult):
                    return left * right
                return left / right
            if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
                value = calculate(node.operand)
                return value if isinstance(node.op, ast.UAdd) else -value
            raise ValueError("Use only variable buttons, +, -, *, /, and parentheses.")
        result = calculate(tree)
        if not isinstance(result, pd.Series):
            result = pd.Series(result, index=dataframe.index, dtype=float)
        result = pd.to_numeric(result, errors="coerce")
        return result.replace([float("inf"), float("-inf")], float("nan"))

    @staticmethod
    def _calculation_display(tokens):
        return " ".join(value for _token_type, value in tokens)

    def _rebuild_calculated_dataframe(self):
        if self.raw_dataframe is None:
            return
        rebuilt = self.raw_dataframe.copy()
        for definition in self.calculated_variables:
            column_name = self._calculated_column_name(definition["name"])
            rebuilt[column_name] = self._evaluate_calculation_tokens(
                rebuilt, definition["tokens"]
            )
        self.dataframe = rebuilt
        self._refresh_column_controls()
        self.apply_filters()
        self.refresh_data_table()

    def _refresh_column_controls(self):
        columns = self.dataframe.columns.tolist()
        self.x_button_row.set_columns(columns)
        self.y_button_row.set_columns(columns)
        self.z_button_row.set_columns(columns)
        if self.x_column in columns:
            self.x_button_row.set_selected(self.x_column)
        else:
            self.x_column = columns[0] if columns else None
            self.x_button_row.set_selected(self.x_column)
        self.y_columns = [column for column in self.y_columns if column in columns]
        self.y_column = self.y_columns[0] if self.y_columns else None
        self.y_button_row.set_selected_multiple(self.y_columns)
        if self.z_column in columns:
            self.z_button_row.set_selected(self.z_column)
        else:
            self.z_column = None
            self.z_button_row.clear_selection()
        self._update_plot_type_controls()

    def _calculation_editor(self, existing=None):
        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Calculated Variable" if existing else "New Calculated Variable")
        dialog.resize(820, 520)
        layout = QVBoxLayout(dialog)
        name_row = QHBoxLayout()
        name_row.addWidget(QLabel("Variable name:"))
        name_edit = QLineEdit(existing["name"] if existing else "")
        name_edit.setPlaceholderText("Example: Total alkalis")
        name_row.addWidget(name_edit, 1)
        layout.addLayout(name_row)
        layout.addWidget(QLabel(
            "Build the expression from the buttons below. Calculated variables can be used in "
            "later calculations, for example Calc: A plus B / C or (A + B) / C."
        ))
        expression_edit = QLineEdit()
        expression_edit.setReadOnly(True)
        tokens = list(existing["tokens"]) if existing else []
        def refresh_expression():
            expression_edit.setText(self._calculation_display(tokens))
        refresh_expression()
        layout.addWidget(expression_edit)
        operator_row = QHBoxLayout()
        operator_row.addWidget(QLabel("Math:"))
        for symbol in ("+", "-", "*", "/", "(", ")"):
            button = QPushButton(symbol)
            button.setMinimumWidth(44)
            button.clicked.connect(
                lambda checked=False, value=symbol: (tokens.append(("operator", value)), refresh_expression())
            )
            operator_row.addWidget(button)
        operator_row.addSpacing(12)
        backspace = QPushButton("Backspace")
        def remove_last():
            if tokens:
                tokens.pop()
                refresh_expression()
        backspace.clicked.connect(remove_last)
        operator_row.addWidget(backspace)
        clear_button = QPushButton("Clear")
        def clear_tokens():
            tokens.clear()
            refresh_expression()
        clear_button.clicked.connect(clear_tokens)
        operator_row.addWidget(clear_button)
        operator_row.addStretch(1)
        layout.addLayout(operator_row)
        layout.addWidget(QLabel("Available worksheet and calculated variables:"))
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        holder = QWidget()
        grid = QGridLayout(holder)
        available = self.dataframe.columns.tolist()
        if existing:
            own_name = self._calculated_column_name(existing["name"])
            available = [name for name in available if name != own_name]
        for index, column in enumerate(available):
            button = QPushButton(str(column))
            button.clicked.connect(
                lambda checked=False, value=str(column): (
                    tokens.append(("column", value)), refresh_expression()
                )
            )
            grid.addWidget(button, index // 4, index % 4)
        scroll.setWidget(holder)
        layout.addWidget(scroll, 1)
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.button(QDialogButtonBox.StandardButton.Ok).setText("Save Calculation")
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        while dialog.exec() == QDialog.DialogCode.Accepted:
            name = name_edit.text().strip()
            if not name:
                QMessageBox.warning(dialog, "Missing name", "Enter a name for the calculated variable.")
            elif not tokens:
                QMessageBox.warning(dialog, "Empty expression", "Build a calculation first.")
            else:
                proposed_column = self._calculated_column_name(name)
                other_names = {
                    self._calculated_column_name(item["name"])
                    for item in self.calculated_variables if item is not existing
                }
                if proposed_column in self.raw_dataframe.columns or proposed_column in other_names:
                    QMessageBox.warning(dialog, "Duplicate name", "That calculated-variable name is already in use.")
                else:
                    try:
                        test_frame = self.dataframe.copy()
                        if existing:
                            old_column = self._calculated_column_name(existing["name"])
                            test_frame = test_frame.drop(columns=[old_column], errors="ignore")
                        self._evaluate_calculation_tokens(test_frame, tokens)
                        return {"name": name, "tokens": list(tokens)}
                    except Exception as exc:
                        QMessageBox.warning(dialog, "Invalid calculation", str(exc))
            # Reopen after validation warning while preserving controls.
        return None

    def open_calculations_manager(self):
        if self.raw_dataframe is None:
            QMessageBox.information(self, "No data", "Open an Excel workbook first.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Calculated Variables")
        dialog.resize(760, 430)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(
            "Create reusable row-by-row variables. New variables become normal plotting and "
            "filtering choices and may be used in subsequent calculations."
        ))
        variable_list = QListWidget()
        layout.addWidget(variable_list, 1)
        def refresh_list():
            variable_list.clear()
            for item in self.calculated_variables:
                variable_list.addItem(
                    f"{self._calculated_column_name(item['name'])} = "
                    f"{self._calculation_display(item['tokens'])}"
                )
        refresh_list()
        controls = QHBoxLayout()
        add_button = QPushButton("Add...")
        edit_button = QPushButton("Edit...")
        delete_button = QPushButton("Delete")
        controls.addWidget(add_button)
        controls.addWidget(edit_button)
        controls.addWidget(delete_button)
        controls.addStretch(1)
        layout.addLayout(controls)
        def add_variable():
            definition = self._calculation_editor()
            if definition:
                self.calculated_variables.append(definition)
                try:
                    self._rebuild_calculated_dataframe()
                except Exception as exc:
                    self.calculated_variables.pop()
                    QMessageBox.warning(dialog, "Calculation error", str(exc))
                refresh_list()
        def edit_variable():
            row = variable_list.currentRow()
            if row < 0:
                return
            old = self.calculated_variables[row]
            replacement = self._calculation_editor(old)
            if replacement:
                self.calculated_variables[row] = replacement
                try:
                    self._rebuild_calculated_dataframe()
                except Exception as exc:
                    self.calculated_variables[row] = old
                    self._rebuild_calculated_dataframe()
                    QMessageBox.warning(dialog, "Calculation error", str(exc))
                refresh_list()
        def delete_variable():
            row = variable_list.currentRow()
            if row < 0:
                return
            removed_name = self._calculated_column_name(self.calculated_variables[row]["name"])
            dependants = []
            for index, item in enumerate(self.calculated_variables):
                if index != row and any(t == "column" and v == removed_name for t, v in item["tokens"]):
                    dependants.append(item["name"])
            if dependants:
                QMessageBox.warning(
                    dialog, "Variable is in use",
                    "Delete dependent calculations first: " + ", ".join(dependants)
                )
                return
            self.calculated_variables.pop(row)
            self._rebuild_calculated_dataframe()
            refresh_list()
        add_button.clicked.connect(add_variable)
        edit_button.clicked.connect(edit_variable)
        delete_button.clicked.connect(delete_variable)
        variable_list.itemDoubleClicked.connect(lambda _item: edit_variable())
        close_button = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close_button.rejected.connect(dialog.reject)
        layout.addWidget(close_button)
        dialog.exec()

    def open_filter_dialog(self):
        if self.dataframe is None:
            QMessageBox.information(self, "No data", "Open an Excel workbook first.")
            return
        dialog=QDialog(self)
        dialog.setWindowTitle("Data Filters")
        dialog.resize(620,430)
        layout=QVBoxLayout(dialog)
        controls=QGridLayout()
        column_combo=QComboBox(); column_combo.addItems(self.dataframe.columns.astype(str).tolist())
        operator_combo=QComboBox(); operator_combo.addItems([
            "=", "!=", ">", ">=", "<", "<=", "Contains", "Does not contain",
            "Starts with", "Ends with", "Is blank", "Is not blank"
        ])
        value_edit=QLineEdit(); value_edit.setPlaceholderText("Filter value")
        controls.addWidget(QLabel("Column:"),0,0); controls.addWidget(column_combo,0,1)
        controls.addWidget(QLabel("Condition:"),1,0); controls.addWidget(operator_combo,1,1)
        controls.addWidget(QLabel("Value:"),2,0); controls.addWidget(value_edit,2,1)
        layout.addLayout(controls)
        filter_list=QListWidget()
        layout.addWidget(QLabel("Active filters (combined with AND):"))
        layout.addWidget(filter_list)
        for rule in self.active_filters:
            filter_list.addItem(self._filter_rule_text(rule))
        button_row=QHBoxLayout()
        add_button=QPushButton("Add Filter")
        remove_button=QPushButton("Remove Selected")
        clear_button=QPushButton("Clear All")
        button_row.addWidget(add_button); button_row.addWidget(remove_button); button_row.addWidget(clear_button)
        layout.addLayout(button_row)
        working_filters=list(self.active_filters)
        def update_value_state(operator):
            value_edit.setEnabled(operator not in ("Is blank","Is not blank"))
        operator_combo.currentTextChanged.connect(update_value_state)
        update_value_state(operator_combo.currentText())
        def add_filter():
            operator=operator_combo.currentText(); value=value_edit.text().strip()
            if operator not in ("Is blank","Is not blank") and value=="":
                QMessageBox.warning(dialog,"Missing value","Enter a filter value.")
                return
            rule=(column_combo.currentText(),operator,value)
            working_filters.append(rule); filter_list.addItem(self._filter_rule_text(rule)); value_edit.clear()
        def remove_filter():
            row=filter_list.currentRow()
            if row>=0:
                filter_list.takeItem(row); working_filters.pop(row)
        def clear_filters():
            working_filters.clear(); filter_list.clear()
        add_button.clicked.connect(add_filter); remove_button.clicked.connect(remove_filter); clear_button.clicked.connect(clear_filters)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Apply|QDialogButtonBox.StandardButton.Cancel)
        apply_button=buttons.button(QDialogButtonBox.StandardButton.Apply)
        apply_button.setText("Apply Filters")
        apply_button.clicked.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec()==QDialog.DialogCode.Accepted:
            self.active_filters=working_filters
            self.apply_filters()

    @staticmethod
    def _filter_rule_text(rule):
        column,operator,value=rule
        return f"{column} {operator}" if operator in ("Is blank","Is not blank") else f"{column} {operator} {value}"

    def apply_filters(self):
        if self.dataframe is None:
            return
        mask=pd.Series(True,index=self.dataframe.index)
        errors=[]
        for column,operator,value in self.active_filters:
            series=self.dataframe[column]
            try:
                if operator in (">",">=","<","<="):
                    numeric=pd.to_numeric(series,errors="coerce"); target=float(value)
                    condition={">":numeric>target,">=":numeric>=target,"<":numeric<target,"<=":numeric<=target}[operator]
                elif operator in ("=","!="):
                    numeric=pd.to_numeric(series,errors="coerce")
                    try:
                        target=float(value)
                        condition=numeric.eq(target) if operator=="=" else numeric.ne(target)
                    except ValueError:
                        text=series.astype("string").str.strip()
                        condition=text.eq(value) if operator=="=" else text.ne(value)
                elif operator=="Contains":
                    condition=series.astype("string").str.contains(value,case=False,na=False,regex=False)
                elif operator=="Does not contain":
                    condition=~series.astype("string").str.contains(value,case=False,na=False,regex=False)
                elif operator=="Starts with":
                    condition=series.astype("string").str.startswith(value,na=False)
                elif operator=="Ends with":
                    condition=series.astype("string").str.endswith(value,na=False)
                elif operator=="Is blank":
                    condition=series.isna()|series.astype("string").str.strip().eq("")
                else:
                    condition=series.notna()&series.astype("string").str.strip().ne("")
                mask &= condition.fillna(False)
            except Exception as exc:
                errors.append(f"{column} {operator}: {exc}")
        self.filtered_dataframe=self.dataframe.loc[mask].copy()
        self.refresh_data_table()
        self.update_plot()
        message=f"Filters applied: {len(self.filtered_dataframe):,} of {len(self.dataframe):,} rows retained"
        if self.active_filters: message+=f" using {len(self.active_filters)} filter(s)."
        else: message+="."
        self.status_bar.showMessage(message)
        if errors:
            QMessageBox.warning(self,"Filter warnings","\n".join(errors))

    def reset_filters(self):
        if self.dataframe is None:
            return
        self.active_filters=[]
        self.filtered_dataframe=self.dataframe.copy()
        self.refresh_data_table()
        self.update_plot()
        self.status_bar.showMessage(f"Filters reset: all {len(self.dataframe):,} rows restored.")

    def clear_petrological_diagram(self):
        if self.petrological_diagram is not None or self.spider_configuration is not None:
            self._reset_custom_plot_title()
        self.petrological_diagram = None
        self.petrological_columns = {}
        self.spider_configuration = None
        self._update_plot_type_controls()
        self.update_plot()
        self.status_bar.showMessage("Returned to the standard plot.")

    @staticmethod
    def _compact_element_name(value):
        return re.sub(r"[^a-z0-9]", "", str(value).lower())

    @staticmethod
    def _ree_oxide_stoichiometry(element, column):
        """Return (REE atoms, O atoms) for a matching REE oxide column."""
        compact = re.sub(r"[^a-z0-9]", "", str(column).lower())
        match = re.match(
            rf"^{re.escape(element.lower())}(\d*)o(\d+)(?:wt|wtpct|weightpercent|percent|pct)?$",
            compact,
        )
        if match is None:
            # Permit unit or method suffixes after the oxide formula, for
            # example Ce2O3_wt% or Ce2O3_XRF.
            match = re.match(rf"^{re.escape(element.lower())}(\d*)o(\d+)", compact)
        if match is None:
            return None
        element_atoms = int(match.group(1) or 1)
        oxygen_atoms = int(match.group(2))
        if element_atoms <= 0 or oxygen_atoms <= 0:
            return None
        return element_atoms, oxygen_atoms

    def _find_element_column(self, element, columns):
        """Find elemental ppm/ppb first, then a matching REE oxide wt% column."""
        target = element.lower()
        exact = []
        prefixed = []
        oxide = []
        for column in columns:
            compact = self._compact_element_name(column)
            if compact == target:
                exact.append(column)
            elif compact in (target + "ppm", target + "ppb", target + "ugg", target + "mgg"):
                exact.append(column)
            elif compact.startswith(target) and any(unit in compact for unit in ("ppm", "ppb", "ugg", "mgg")):
                prefixed.append(column)
            elif element in SPIDER_ATOMIC_WEIGHTS and self._ree_oxide_stoichiometry(element, column):
                oxide.append(column)
        return exact[0] if exact else prefixed[0] if prefixed else oxide[0] if oxide else None

    def configure_spider_plot(self):
        if self.dataframe is None:
            QMessageBox.information(self, "No data", "Open an Excel workbook first.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Configure Petrological Spider Plot")
        dialog.resize(760, 620)
        layout = QVBoxLayout(dialog)

        form = QGridLayout()
        form.addWidget(QLabel("Normalization:"), 0, 0)
        normalization_combo = QComboBox()
        normalization_combo.addItems(SPIDER_NORMALIZATIONS.keys())
        form.addWidget(normalization_combo, 0, 1)

        form.addWidget(QLabel("Sample-label column:"), 1, 0)
        label_combo = QComboBox()
        label_combo.addItem("(worksheet row number)")
        label_combo.addItems([str(column) for column in self.dataframe.columns])
        form.addWidget(label_combo, 1, 1)

        units_label = QLabel(
            "Element columns are interpreted as ppm; columns labelled ppb are converted to ppm. "
            "Mapped REE-oxide columns (for example Ce2O3 or CeO2) are assumed to be wt% "
            "and converted stoichiometrically to elemental ppm."
        )
        units_label.setWordWrap(True)
        form.addWidget(units_label, 2, 0, 1, 2)
        layout.addLayout(form)

        layout.addWidget(QLabel("Element-to-column mapping (automatically detected):"))

        element_combos = {}
        mapping_holder = QWidget()
        mapping_grid = QGridLayout(mapping_holder)
        mapping_grid.setContentsMargins(10, 8, 10, 8)
        mapping_grid.setHorizontalSpacing(12)
        mapping_grid.setVerticalSpacing(6)
        mapping_grid.setColumnStretch(1, 1)
        mapping_grid.addWidget(QLabel("Element"), 0, 0)
        mapping_grid.addWidget(QLabel("Worksheet column"), 0, 1)
        mapping_scroll = QScrollArea()
        mapping_scroll.setWidgetResizable(True)
        mapping_scroll.setFrameShape(QFrame.Shape.StyledPanel)
        mapping_scroll.setMinimumHeight(360)
        mapping_scroll.setWidget(mapping_holder)
        layout.addWidget(mapping_scroll, 1)

        columns = [str(column) for column in self.dataframe.columns]

        def rebuild_mapping(_name=None):
            while mapping_grid.count() > 2:
                item = mapping_grid.takeAt(2)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
            element_combos.clear()
            values = SPIDER_NORMALIZATIONS[normalization_combo.currentText()]
            for row, element in enumerate(values, start=1):
                mapping_grid.addWidget(QLabel(element), row, 0)
                combo = QComboBox()
                combo.addItem("(not available)")
                combo.addItems(columns)
                detected = self._find_element_column(element, columns)
                if detected is not None:
                    combo.setCurrentText(str(detected))
                mapping_grid.addWidget(combo, row, 1)
                element_combos[element] = combo
            mapping_grid.setRowStretch(len(values) + 1, 1)

        normalization_combo.currentTextChanged.connect(rebuild_mapping)
        rebuild_mapping()

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Create Spider Plot")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        mapping = {
            element: combo.currentText()
            for element, combo in element_combos.items()
            if combo.currentIndex() > 0
        }
        if len(mapping) < 2:
            QMessageBox.warning(
                self, "Insufficient elements",
                "Map at least two elements to worksheet columns."
            )
            return

        self._reset_custom_plot_title()
        self.petrological_diagram = None
        self.petrological_columns = {}
        self.spider_configuration = {
            "normalization": normalization_combo.currentText(),
            "label_column": None if label_combo.currentIndex() == 0 else label_combo.currentText(),
            "mapping": mapping,
        }
        self.update_plot()

    def _spider_input_ppm(self, dataframe, element, column):
        """Return an elemental concentration in ppm for spider normalization."""
        values = pd.to_numeric(dataframe[column], errors="coerce")
        compact = self._compact_element_name(column)

        oxide_stoichiometry = None
        if element in SPIDER_ATOMIC_WEIGHTS:
            oxide_stoichiometry = self._ree_oxide_stoichiometry(element, column)

        if oxide_stoichiometry is not None:
            element_atoms, oxygen_atoms = oxide_stoichiometry
            element_mass = element_atoms * SPIDER_ATOMIC_WEIGHTS[element]
            oxide_mass = element_mass + oxygen_atoms * SPIDER_ATOMIC_WEIGHTS["O"]
            elemental_mass_fraction = element_mass / oxide_mass
            # wt% oxide -> ppm oxide (x10,000), then retain the elemental
            # mass fraction implied by the oxide stoichiometry.
            return values * 10000.0 * elemental_mass_fraction

        if "ppb" in compact:
            return values / 1000.0
        return values

    def _draw_spider_plot(self, dataframe):
        import numpy as np

        config = self.spider_configuration
        reference = SPIDER_NORMALIZATIONS[config["normalization"]]
        elements = [element for element in reference if element in config["mapping"]]
        columns = [config["mapping"][element] for element in elements]

        normalized = pd.DataFrame(index=dataframe.index)
        for element, column in zip(elements, columns):
            normalized[element] = self._spider_input_ppm(dataframe, element, column) / reference[element]
        normalized = normalized.replace([np.inf, -np.inf], np.nan)
        normalized[normalized <= 0] = np.nan

        self.axes.clear()
        self._restore_cartesian_axes()
        self.axes.set_box_aspect(None)
        self.axes.set_yscale("log")
        x_positions = np.arange(len(elements), dtype=float)

        label_column = config["label_column"]
        plotted = 0
        omitted = 0
        line_artists = []
        for source_index, row in normalized.iterrows():
            valid = row.notna().to_numpy()
            if valid.sum() < 2:
                omitted += 1
                continue
            if label_column is None:
                label = str(source_index)
            else:
                label_value = dataframe.at[source_index, label_column]
                label = str(label_value) if pd.notna(label_value) else str(source_index)
            pattern_x = x_positions[valid]
            pattern_y = row.to_numpy(dtype=float)[valid]
            line, = self.axes.plot(
                pattern_x, pattern_y,
                marker="o", markersize=4.5, linewidth=1.0,
                alpha=0.80, label=label, picker=5,
            )
            # Repeat the source index for each line vertex so any clicked marker
            # or segment resolves to the same worksheet row.
            self._artist_row_indices[line] = np.full(
                len(pattern_x), source_index, dtype=object
            )
            self._plotted_row_coordinates.setdefault(source_index, []).append(
                (self.axes, pattern_x.copy(), pattern_y.copy())
            )
            line_artists.append(line)
            plotted += 1

        self.axes.set_xticks(x_positions)
        self.axes.set_xticklabels(
            elements,
            fontsize=self.axis_number_size,
            color=self.axis_number_colour,
            fontfamily=self.font_combo.currentText() or "Arial",
            rotation=45 if len(elements) > 20 else 0,
            ha="right" if len(elements) > 20 else "center",
        )
        self.axes.tick_params(
            axis="x", which="both", bottom=True, labelbottom=True
        )
        self.axes.set_xlim(-0.35, len(elements) - 0.65)
        self.axes.set_ylabel("Sample / normalization reference")
        self.axes.set_xlabel("")
        self.axes.set_title(config["normalization"])
        self.axes.grid(True, which="major", axis="y", color=self.current_theme["grid"], alpha=0.65)
        self.axes.grid(True, which="minor", axis="y", color=self.current_theme["grid"], alpha=0.25)
        self.axes.grid(False, axis="x")

        if 0 < plotted <= 15:
            _name, location = self._data_legend_position()
            self.axes.legend(loc=location, fontsize=max(7, self.axis_number_size - 1))

        self.figure.patch.set_facecolor(self.current_theme["ui_background"])
        self.axes.set_facecolor(self.current_theme["plot_background"])
        self._apply_plot_theme()
        self.axes.set_yscale("log")
        # Axis formatting installs numeric formatters, so restore the categorical
        # element labels after applying the common theme and axis formatting.
        self.axes.set_xticks(x_positions)
        self.axes.set_xticklabels(
            elements,
            fontsize=self.axis_number_size,
            color=self.axis_number_colour,
            fontfamily=self.font_combo.currentText() or "Arial",
            rotation=45 if len(elements) > 20 else 0,
            ha="right" if len(elements) > 20 else "center",
        )
        self.axes.tick_params(
            axis="x", which="both", bottom=True, labelbottom=True
        )
        self.axes.set_xlabel("")
        # Spider diagrams always default to a 16:9 plotting frame, independent
        # of the standard plot-ratio preference and current window dimensions.
        self.axes.set_box_aspect(9.0 / 16.0)
        self.axes.set_anchor("C")
        self._apply_custom_plot_title()
        selected_source_indices = self._selected_table_source_indices()
        if selected_source_indices:
            self._highlight_rows_on_plot(selected_source_indices)
            self.clear_highlight_button.setEnabled(True)
        self.canvas.draw_idle()
        self.status_bar.showMessage(
            f"Spider plot: {plotted:,} sample pattern(s), {len(elements)} elements; "
            f"omitted {omitted:,} rows with fewer than two positive numeric values."
        )

    def configure_petrological_diagram(self, diagram_name):
        if self.dataframe is None:
            QMessageBox.information(self, "No data", "Open an Excel workbook first.")
            return
        if diagram_name.startswith("TAS"):
            roles = ["SiO2", "Na2O", "K2O"]
            title = "TAS: Total Alkali-Silica"
        elif diagram_name.startswith("AFM"):
            roles = ["A (Na2O + K2O or alkalis)", "F (total iron)", "M (MgO)"]
            title = "AFM: Alkali-Fe-Mg"
        elif diagram_name.startswith("QAP"):
            roles = ["Q (quartz)", "A (alkali feldspar)", "P (plagioclase)"]
            title = "QAP modal projection"
        else:
            roles = ["F (feldspathoid)", "A (alkali feldspar)", "P (plagioclase)"]
            title = "FAP modal projection"
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        layout = QGridLayout(dialog)
        layout.addWidget(QLabel(
            "Assign a worksheet or calculated variable to each required component."
        ), 0, 0, 1, 2)
        combos = {}
        columns = [str(column) for column in self.dataframe.columns]
        def preferred_index(role):
            """Choose the best matching worksheet heading for a diagram role."""
            normalized = role.lower().replace("₂", "2").replace(" ", "")
            candidates = []
            if normalized.startswith("sio2"):
                candidates = ["sio2", "silica"]
            elif normalized.startswith("na2o"):
                candidates = ["na2o"]
            elif normalized.startswith("k2o"):
                candidates = ["k2o"]
            elif role.startswith("M "):
                candidates = ["mgo", "mg"]
            elif role.startswith("Q "):
                candidates = ["quartz", "qtz"]
            elif role.startswith("P "):
                candidates = ["plagioclase", "plag"]
            elif role.startswith("F (feld"):
                candidates = ["feldspathoid", "foid"]
            elif role.startswith("A (alkali feld"):
                # Include the requested AFS abbreviation and the common
                # misspelling 'akali feldspar' in addition to full names.
                candidates = [
                    "alkali feldspar", "akali feldspar", "afs",
                    "k-feldspar", "k feldspar", "orthoclase", "microcline",
                ]
            elif role.startswith("F (total"):
                candidates = ["feo", "fet", "total iron", "fe"]
            elif role.startswith("A (Na"):
                candidates = ["calc: total alkalis", "alkalis", "na2o+k2o"]

            def compact_name(value):
                return re.sub(r"[^a-z0-9]+", "", str(value).lower().replace("₂", "2"))

            def tokens(value):
                return set(re.findall(r"[a-z0-9]+", str(value).lower().replace("₂", "2")))

            best = None
            for index, column in enumerate(columns):
                column_compact = compact_name(column)
                column_tokens = tokens(column)
                for candidate_order, candidate in enumerate(candidates):
                    candidate_compact = compact_name(candidate)
                    candidate_tokens = tokens(candidate)
                    if column_compact == candidate_compact:
                        rank = 0
                    elif column_compact.startswith(candidate_compact):
                        rank = 1
                    elif candidate_tokens and candidate_tokens.issubset(column_tokens):
                        rank = 2
                    elif candidate_compact in column_compact:
                        rank = 3
                    else:
                        continue
                    score = (rank, candidate_order, index)
                    if best is None or score < best[0]:
                        best = (score, index)
            return best[1] if best is not None else 0
        for row, role in enumerate(roles, start=1):
            combo = QComboBox(); combo.addItems(columns)
            if columns: combo.setCurrentIndex(preferred_index(role))
            layout.addWidget(QLabel(role + ":"), row, 0); layout.addWidget(combo, row, 1)
            combos[role] = combo
        if diagram_name.startswith(("QAP", "FAP")):
            warning = QLabel(
                "QAP/FAP classification is conventionally based on normalized modal proportions. "
                "If the selected inputs are wt%, the result is a compositional projection rather than a formal modal classification."
            )
            warning.setWordWrap(True); layout.addWidget(warning, len(roles)+1, 0, 1, 2)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept); buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons, len(roles)+2, 0, 1, 2)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        chosen = [combo.currentText() for combo in combos.values()]
        if len(set(chosen)) != len(chosen):
            QMessageBox.warning(self, "Duplicate components", "Select a different variable for each component.")
            return
        self._reset_custom_plot_title()
        self.spider_configuration = None
        self.petrological_diagram = diagram_name.split(":", 1)[0]
        self.petrological_columns = dict(zip(roles, chosen))
        self.update_plot()

    def _petrology_numeric(self, dataframe, column):
        return pd.to_numeric(dataframe[column], errors="coerce")

    def _draw_qap_diagram(self, dataframe):
        """Draw the standard IUGS plutonic Q-A-P classification framework."""
        import numpy as np

        columns = list(self.petrological_columns.values())
        quartz = self._petrology_numeric(dataframe, columns[0])
        alkali_feldspar = self._petrology_numeric(dataframe, columns[1])
        plagioclase = self._petrology_numeric(dataframe, columns[2])

        valid = quartz.notna() & alkali_feldspar.notna() & plagioclase.notna()
        valid &= (quartz >= 0) & (alkali_feldspar >= 0) & (plagioclase >= 0)
        totals = quartz + alkali_feldspar + plagioclase
        valid &= totals > 0

        self.axes.clear()
        self.axes.set_axis_on()
        self.axes.set_box_aspect(None)
        root = np.sqrt(3.0) / 2.0
        colours = self.current_theme
        boundary_colour = colours["text"]
        field_colour = colours["muted_text"]

        # Conventional QAP orientation: Q at the apex, A lower-left, P lower-right.
        self.axes.plot(
            [0.0, 1.0, 0.5, 0.0], [0.0, 0.0, root, 0.0],
            color=boundary_colour, linewidth=1.35, zorder=2,
        )

        # Quartz boundaries at 5, 20, 60, and 90 modal percent.
        for q_fraction in (0.05, 0.20, 0.60, 0.90):
            self.axes.plot(
                [0.5 * q_fraction, 1.0 - 0.5 * q_fraction],
                [root * q_fraction, root * q_fraction],
                color=field_colour, linewidth=0.9, zorder=1,
            )

        # P/(A+P) boundaries at 10, 35, 65, and 90 percent. These field
        # boundaries extend from the A-P join to the 60% quartz boundary.
        q_limit = 0.60
        for p_ratio in (0.10, 0.35, 0.65, 0.90):
            x_bottom = p_ratio
            x_top = (1.0 - q_limit) * p_ratio + 0.5 * q_limit
            self.axes.plot(
                [x_bottom, x_top], [0.0, root * q_limit],
                color=field_colour, linewidth=0.9, zorder=1,
            )

        def position(q_fraction, p_ratio=0.5):
            return (
                (1.0 - q_fraction) * p_ratio + 0.5 * q_fraction,
                root * q_fraction,
            )

        label_size = max(5.5, self.axis_number_size - 3.0)
        label_kwargs = dict(
            ha="center", va="center", fontsize=label_size,
            color=colours["text"], fontfamily=self.font_combo.currentText() or "Arial",
            zorder=2,
        )

        # Field labels use compact line breaks so the full IUGS framework remains
        # readable at normal window sizes.
        field_labels = [
            (0.95, 0.50, "Quartzolite"),
            (0.75, 0.50, "Quartz-rich\ngranitoid"),
            (0.39, 0.05, "Alkali feldspar\ngranite"),
            (0.39, 0.225, "Syeno-\ngranite"),
            (0.39, 0.50, "Monzo-\ngranite"),
            (0.39, 0.775, "Granodiorite"),
            (0.39, 0.95, "Tonalite"),
            (0.115, 0.05, "Quartz alkali\nfeldspar syenite"),
            (0.115, 0.225, "Quartz\nsyenite"),
            (0.115, 0.50, "Quartz\nmonzonite"),
            (0.115, 0.775, "Quartz monzodiorite\n/ monzogabbro"),
            (0.115, 0.95, "Quartz diorite / gabbro\n/ anorthosite"),
            (0.023, 0.05, "Alkali feldspar\nsyenite"),
            (0.023, 0.225, "Syenite"),
            (0.023, 0.50, "Monzonite"),
            (0.023, 0.775, "Monzodiorite\n/ monzogabbro"),
            (0.023, 0.95, "Diorite / gabbro\n/ anorthosite"),
        ]
        for q_fraction, p_ratio, text in field_labels:
            x_value, y_value = position(q_fraction, p_ratio)
            self.axes.text(x_value, y_value, text, **label_kwargs)

        if valid.any():
            q = quartz[valid].to_numpy(dtype=float)
            a = alkali_feldspar[valid].to_numpy(dtype=float)
            p = plagioclase[valid].to_numpy(dtype=float)
            total = q + a + p
            q, a, p = q / total, a / total, p / total
            plot_x = p + 0.5 * q
            plot_y = root * q
            marker = MARKERS[self.marker_combo.currentText()]
            scatter = self.axes.scatter(
                plot_x, plot_y,
                s=self.marker_size_spin.value(), marker=marker,
                facecolors="none" if self.no_fill else self.plot_colour,
                edgecolors="none" if self.no_stroke else self.stroke_colour,
                linewidths=self.stroke_width_spin.value(),
                alpha=self.opacity_spin.value() / 100.0,
                picker=True, pickradius=5, zorder=5,
            )
            row_indices = dataframe.index[valid].to_numpy()
            self._artist_row_indices[scatter] = row_indices
            self._artist_series_columns[scatter] = "QAP"
            for row_index, x_value, y_value in zip(row_indices, plot_x, plot_y):
                self._plotted_row_coordinates.setdefault(row_index, []).append(
                    (self.axes, float(x_value), float(y_value))
                )

        title_kwargs = dict(
            color=self.axis_title_colour,
            fontsize=self.axis_title_size,
            fontweight="bold" if self.axis_title_bold else "normal",
            fontfamily=self.font_combo.currentText() or "Arial",
            clip_on=False,
            zorder=6,
        )
        self.axes.text(0.5, root + 0.045, "Q", ha="center", va="bottom", **title_kwargs)
        self.axes.text(-0.025, -0.035, "A", ha="right", va="top", **title_kwargs)
        self.axes.text(1.025, -0.035, "P", ha="left", va="top", **title_kwargs)
        self.axes.set_title("IUGS QAP plutonic classification")
        self._apply_custom_plot_title()
        self.axes.set_xlim(-0.10, 1.10)
        self.axes.set_ylim(-0.08, root + 0.10)
        self.axes.set_aspect("equal", adjustable="box")
        self.axes.set_xticks([])
        self.axes.set_yticks([])
        for spine in self.axes.spines.values():
            spine.set_visible(False)
        self.axes.grid(False)
        self.figure.patch.set_facecolor(colours["ui_background"])
        self.axes.set_facecolor(colours["plot_background"])

        selected_source_indices = self._selected_table_source_indices()
        if selected_source_indices:
            self._highlight_rows_on_plot(selected_source_indices)
            self.clear_highlight_button.setEnabled(True)

        self.status_bar.showMessage(
            f"IUGS QAP: plotted {int(valid.sum()):,} normalized row(s); "
            f"omitted {len(dataframe) - int(valid.sum()):,} invalid row(s)."
        )
        self.canvas.draw_idle()

    def _draw_petrological_ternary(self, dataframe, labels, columns, title):
        import numpy as np
        values = [self._petrology_numeric(dataframe, column) for column in columns]
        valid = values[0].notna() & values[1].notna() & values[2].notna()
        valid &= (values[0] >= 0) & (values[1] >= 0) & (values[2] >= 0)
        totals = values[0] + values[1] + values[2]
        valid &= totals > 0
        self.axes.clear(); self.axes.set_axis_on(); self.axes.set_box_aspect(None)
        root = np.sqrt(3.0) / 2.0
        colours = self.current_theme
        self.axes.plot([0,1,0.5,0], [0,0,root,0], color=colours["text"], linewidth=1.3)
        for fraction in (0.2,0.4,0.6,0.8):
            self.axes.plot([1-fraction,0.5*(1-fraction)],[0,root*(1-fraction)],color=colours["grid"],lw=0.7,alpha=0.6)
            self.axes.plot([fraction,0.5+0.5*fraction],[0,root*(1-fraction)],color=colours["grid"],lw=0.7,alpha=0.6)
            self.axes.plot([0.5*fraction,1-0.5*fraction],[root*fraction,root*fraction],color=colours["grid"],lw=0.7,alpha=0.6)
        if valid.any():
            a,b,c = [value[valid].to_numpy(float) for value in values]
            total=a+b+c; a/=total; b/=total; c/=total
            px=b+0.5*c; py=root*c
            self.axes.scatter(px,py,s=self.marker_size_spin.value(),marker=MARKERS[self.marker_combo.currentText()],
                color=self.plot_colour,alpha=self.opacity_spin.value()/100.0,
                edgecolors="none" if self.no_stroke else self.stroke_colour,
                linewidths=self.stroke_width_spin.value(),zorder=3)
        kwargs=dict(color=self.axis_title_colour,fontsize=self.axis_title_size,
                    fontweight="bold" if self.axis_title_bold else "normal",fontfamily=self.font_combo.currentText())
        self.axes.text(-0.03,-0.035,labels[0],ha="right",va="top",**kwargs)
        self.axes.text(1.03,-0.035,labels[1],ha="left",va="top",**kwargs)
        self.axes.text(0.5,root+0.045,labels[2],ha="center",va="bottom",**kwargs)
        self.axes.set_title(title); self.axes.set_xlim(-0.12,1.12); self.axes.set_ylim(-0.10,root+0.12)
        self._apply_custom_plot_title()
        self.axes.set_aspect("equal",adjustable="box"); self.axes.set_axis_off()
        self.status_bar.showMessage(f"{title}: plotted {int(valid.sum()):,} normalized rows.")
        self.canvas.draw()

    def _draw_tas_diagram(self, dataframe):
        """Draw the Le Bas et al. (1986) TAS field framework and sample data."""
        from matplotlib.path import Path as MplPath

        columns = list(self.petrological_columns.values())
        silica = self._petrology_numeric(dataframe, columns[0])
        sodium = self._petrology_numeric(dataframe, columns[1])
        potassium = self._petrology_numeric(dataframe, columns[2])
        alkalis = sodium + potassium
        valid = silica.notna() & alkalis.notna()

        self.axes.clear()
        self.axes.set_axis_on()
        self.axes.set_box_aspect(None)
        self.axes.set_aspect("auto", adjustable="box")

        # TAS boundary network supplied as an ordered coordinate trace.
        # Repeated junctions intentionally retrace to branch into adjacent field
        # boundaries without introducing polygon-closing artefacts.
        tas_boundary_trace = [
            (41.0, 0.0),
            (41.0, 3.0),
            (41.0, 7.0),
            (45.0, 9.4),
            (48.4, 11.5),
            (52.5, 14.0),
            (49.0, 15.5),
            (52.5, 14.0),
            (57.6, 11.7),
            (63.0, 14.56),
            (57.6, 11.7),
            (63.0, 7.0),
            (57.6, 11.7),
            (53.0, 9.3),
            (48.4, 11.5),
            (53.0, 9.3),
            (57.0, 5.9),
            (53.0, 9.3),
            (49.4, 7.3),
            (45.0, 9.4),
            (49.4, 7.3),
            (52.0, 5.0),
            (49.4, 7.3),
            (45.0, 5.0),
            (45.0, 0.0),
            (45.0, 3.0),
            (41.0, 3.0),
            (45.0, 3.0),
            (45.0, 5.0),
            (52.0, 5.0),
            (52.0, 0.0),
            (52.0, 5.0),
            (57.0, 5.9),
            (57.0, 0.0),
            (57.0, 5.9),
            (63.0, 7.0),
            (63.0, 0.0),
            (63.0, 7.0),
            (69.0, 8.0),
            (69.0, 13.0),
            (69.0, 8.0),
            (77.0, 0.0),
        ]
        edge = "#9a9a9a"
        self.axes.plot(
            [point[0] for point in tas_boundary_trace],
            [point[1] for point in tas_boundary_trace],
            color=edge,
            linewidth=1.0,
            zorder=1,
        )

        # Closed polygons are retained only for point-in-field classification.
        # They are not rendered, avoiding artificial closing segments.
        fields = {
            "Pc": ("Picrobasalt", [(41,0),(45,0),(45,3),(41,3)]),
            "B": ("Basalt", [(45,0),(52,0),(52,5),(45,5)]),
            "O1": ("Basaltic andesite", [(52,0),(57,0),(57,5.9),(52,5)]),
            "O2": ("Andesite", [(57,0),(63,0),(63,7),(57,5.9)]),
            "O3": ("Dacite", [(63,0),(77,0),(69,8),(63,7)]),
            "R": ("Rhyolite", [(69,8),(69,16),(80,16),(80,0),(77,0)]),
            "S1": ("Trachybasalt", [(45,5),(49.4,7.3),(52,5)]),
            "S2": ("Basaltic trachyandesite", [(49.4,7.3),(53,9.3),(57,5.9),(52,5)]),
            "S3": ("Trachyandesite", [(53,9.3),(57.6,11.7),(63,7),(57,5.9)]),
            "T": ("Trachyte / trachydacite", [(57.6,11.7),(63,14.56),(69,13),(69,8),(63,7)]),
            "U1": ("Basanite / tephrite", [(41,3),(45,3),(45,5),(49.4,7.3),(45,9.4),(41,7)]),
            "U2": ("Phonotephrite", [(45,9.4),(48.4,11.5),(53,9.3),(49.4,7.3)]),
            "U3": ("Tephriphonolite", [(48.4,11.5),(52.5,14),(57.6,11.7),(53,9.3)]),
            "Ph": ("Phonolite", [(49,15.5),(63,14.56),(57.6,11.7),(52.5,14)]),
            "F": ("Foidite", [(40,7),(40,16),(49,16),(49,15.5),(52.5,14),(48.4,11.5),(45,9.4)]),
        }

        labels = {
            "Pc": (42.7,1.4,"Picro-\nbasalt"),
            "B": (48.5,2.4,"Basalt"),
            "O1": (54.5,3.6,"Basaltic\nandesite"),
            "O2": (59.7,4.3,"Andesite"),
            "O3": (66.9,4.2,"Dacite"),
            "R": (73.0,8.8,"Rhyolite"),
            "S1": (47.8,5.35,"Trachybasalt"),
            "S2": (52.2,6.7,"Basaltic\ntrachyandesite"),
            "S3": (56.8,8.4,"Trachyandesite"),
            "T": (64.1,9.0,"Trachyte /\ntrachydacite"),
            "U1": (43.0,5.2,"Basanite /\ntephrite"),
            "U2": (48.1,9.3,"Phonotephrite"),
            "U3": (52.0,11.8,"Tephriphonolite"),
            "Ph": (57.4,14.0,"Phonolite"),
            "F": (44.8,12.5,"Foidite"),
        }
        for _code, (xpos, ypos, text) in labels.items():
            self.axes.text(xpos, ypos, text, ha="center", va="center",
                           fontsize=max(7, self.axis_number_size-1),
                           color=self.current_theme["text"], zorder=2)

        # Plot samples after field framework so points remain visible.
        self.axes.scatter(
            silica[valid], alkalis[valid],
            s=self.marker_size_spin.value(),
            marker=MARKERS[self.marker_combo.currentText()],
            facecolors="none" if self.no_fill else self.plot_colour,
            edgecolors="none" if self.no_stroke else self.stroke_colour,
            linewidths=self.stroke_width_spin.value(),
            alpha=self.opacity_spin.value()/100.0,
            zorder=4,
        )

        # Classify valid points by polygon for status reporting.
        counts = {code: 0 for code in fields}
        unclassified = 0
        for xvalue, yvalue in zip(silica[valid], alkalis[valid]):
            matched = False
            for code, (_name, polygon) in fields.items():
                if MplPath(polygon).contains_point((float(xvalue), float(yvalue)), radius=1e-9):
                    counts[code] += 1
                    matched = True
                    break
            if not matched:
                unclassified += 1

        self.axes.set_xlim(40, 80)
        self.axes.set_ylim(0, 16)
        # Use the taller plotting-frame proportion of the published TAS chart.
        # This 0.625 height/width box ratio is independent of the application's
        # generic Square/16:9 display preference and prevents horizontal stretching.
        self.axes.set_box_aspect(0.625)
        self.axes.set_anchor("C")
        self.axes.set_xlabel(self.custom_x_axis_title or self.format_display_label("SiO2, wt. %"))
        self.axes.set_ylabel(self.custom_left_y_axis_title or self.format_display_label("Na2O + K2O, wt. %"))
        self.axes.set_title("TAS classification of volcanic rocks")
        self._apply_custom_plot_title()
        self.axes.grid(False)
        self._apply_axis_formatting()
        self.figure.patch.set_facecolor(self.current_theme["ui_background"])
        self.axes.set_facecolor(self.current_theme["plot_background"])
        self.axes.set_box_aspect(0.625)
        self.axes.set_anchor("C")
        self.canvas.draw_idle()
        classified = int(valid.sum()) - unclassified
        self.status_bar.showMessage(
            f"Le Bas TAS: plotted {int(valid.sum()):,} rows; "
            f"{classified:,} within TAS fields, {unclassified:,} outside the displayed framework."
        )

    def _draw_petrological_diagram(self, dataframe):
        diagram=self.petrological_diagram
        columns=list(self.petrological_columns.values())
        if diagram=="TAS": self._draw_tas_diagram(dataframe)
        elif diagram=="AFM": self._draw_petrological_ternary(dataframe,["A","F","M"],columns,"AFM: Alkali-Fe-Mg")
        elif diagram=="QAP": self._draw_qap_diagram(dataframe)
        elif diagram=="FAP": self._draw_petrological_ternary(dataframe,["F","A","P"],columns,"FAP modal projection")

    def select_x(self, column):
        self.x_column = column
        self.x_button_row.set_selected(column)
        self._update_plot_type_controls()
        self.update_plot()

    def open_y_variable_formatting(self, column, source_button, position):
        """Format the right-clicked variable without changing its selection state."""
        plot_type = self.plot_type_combo.currentText()
        if plot_type == "Ternary":
            QMessageBox.information(
                self,
                "Formatting unavailable",
                "Right-click formatting is available for point and line plots."
            )
            return

        # Right-click formatting must not change the plotted-variable
        # selection. Left-click remains the only way to select or deselect a
        # Y variable. Formatting may be prepared for an unselected variable
        # and will be applied if that variable is selected later.

        if plot_type == "Line":
            existing = self.series_line_formats.get(column)
            if existing:
                self.line_colour = existing["colour"]
                self.line_width = existing["width"]
                self.line_style_name = existing["style"]
                self.line_opacity = existing["opacity"]
                self.line_show_markers = existing["show_markers"]
                self.line_marker_name = existing["marker_name"]
                self.line_marker_size = existing["marker_size"]
            self.open_line_formatting_dialog()
            self.series_line_formats[column] = {
                "colour": self.line_colour,
                "width": self.line_width,
                "style": self.line_style_name,
                "opacity": self.line_opacity,
                "show_markers": self.line_show_markers,
                "marker_name": self.line_marker_name,
                "marker_size": self.line_marker_size,
            }
        else:
            # Preserve the global point defaults. The existing formatting dialog
            # edits those global controls, so leaving the edited values in place
            # also changes every series without an explicit per-series format.
            global_defaults = {
                "fill_colour": self.plot_colour,
                "stroke_colour": self.stroke_colour,
                "marker_name": self.marker_combo.currentText(),
                "marker_size": self.marker_size_spin.value(),
                "opacity": self.opacity_spin.value(),
                "stroke_width": self.stroke_width_spin.value(),
                "no_fill": self.no_fill,
                "no_stroke": self.no_stroke,
            }
            existing = self.series_point_formats.get(column, global_defaults)
            widgets = (
                self.marker_combo, self.marker_size_spin, self.opacity_spin,
                self.stroke_width_spin,
            )
            for widget in widgets:
                widget.blockSignals(True)
            self.plot_colour = existing["fill_colour"]
            self.stroke_colour = existing["stroke_colour"]
            self.marker_combo.setCurrentText(existing["marker_name"])
            self.marker_size_spin.setValue(existing["marker_size"])
            self.opacity_spin.setValue(existing["opacity"])
            self.stroke_width_spin.setValue(existing["stroke_width"])
            self.no_fill = existing["no_fill"]
            self.no_stroke = existing["no_stroke"]
            for widget in widgets:
                widget.blockSignals(False)
            self.open_data_point_formatting_dialog()
            self.series_point_formats[column] = {
                "fill_colour": self.plot_colour,
                "stroke_colour": self.stroke_colour,
                "marker_name": self.marker_combo.currentText(),
                "marker_size": self.marker_size_spin.value(),
                "opacity": self.opacity_spin.value(),
                "stroke_width": self.stroke_width_spin.value(),
                "no_fill": self.no_fill,
                "no_stroke": self.no_stroke,
            }
            # Restore the global defaults without emitting intermediate redraws.
            for widget in widgets:
                widget.blockSignals(True)
            self.plot_colour = global_defaults["fill_colour"]
            self.stroke_colour = global_defaults["stroke_colour"]
            self.marker_combo.setCurrentText(global_defaults["marker_name"])
            self.marker_size_spin.setValue(global_defaults["marker_size"])
            self.opacity_spin.setValue(global_defaults["opacity"])
            self.stroke_width_spin.setValue(global_defaults["stroke_width"])
            self.no_fill = global_defaults["no_fill"]
            self.no_stroke = global_defaults["no_stroke"]
            self._update_colour_button()
            self._update_stroke_colour_button()
            for widget in widgets:
                widget.blockSignals(False)
        self.update_plot()

    def select_y(self, column):
        if self.plot_type_combo.currentText() == "Ternary":
            self.y_columns = [column]
            self.y_button_row.set_selected_multiple(self.y_columns)
        else:
            self.y_columns = list(self.y_button_row.selected_columns)
        self.y_column = self.y_columns[0] if self.y_columns else None
        if self.secondary_y_column not in self.y_columns:
            self.secondary_y_column = ""
        self._update_plot_type_controls()
        self.update_plot()

    def select_z(self, column):
        self.z_column = column
        self.z_button_row.set_selected(column)
        self.update_plot()

    def _update_plot_type_controls(self):
        ternary = self.plot_type_combo.currentText() == "Ternary"
        histogram = self.plot_type_combo.currentText() == "Histogram"
        if ternary:
            self.z_button_row.show()
            self.z_button_row.setEnabled(self.dataframe is not None)
            self.z_button_row.raise_()
        else:
            self.z_button_row.hide()
        self.y_button_row.setEnabled(not histogram and self.dataframe is not None)
        self.x_button_row.title_label.setText(
            "Choose component A (bottom-left)" if ternary else "Choose X-axis column"
        )
        self.y_button_row.title_label.setText(
            "Choose component B (bottom-right)" if ternary else
            "Y-axis selection is not used for histograms" if histogram else
            "Choose Y-axis column (Shift-click to add or remove)"
        )
        self.z_button_row.title_label.setText("Choose component C (top)")
        self.x_label.setText(
            f"Component A: {self.x_column or 'not selected'}" if ternary
            else f"X axis: {self.x_column or 'not selected'}"
        )
        self.y_label.setText(
            f"Component B: {self.y_column or 'not selected'}" if ternary
            else f"Y axes: {', '.join(map(str, self.y_columns)) if self.y_columns else 'not selected'}"
        )
        self.regression_toggle_button.setEnabled(not ternary and not histogram)
        self.regression_legend_button.setEnabled(not ternary and not histogram and self.regression_enabled)
        if hasattr(self, "data_legend_button"):
            self.data_legend_button.setEnabled(not ternary and len(self.y_columns) > 1)
        self.aspect_button.setEnabled(not ternary)

    def clear_selections(self):
        if self.dataframe is None:
            return
        self.x_column = self.dataframe.columns[0]
        self.y_column = None
        self.y_columns = []
        self.z_column = None
        self.x_button_row.set_selected(self.x_column)
        self.y_button_row.clear_selection()
        self.z_button_row.clear_selection()
        self.x_label.setText(f"X axis: {self.x_column}")
        self.y_label.setText("Y axis: not selected")
        self._draw_placeholder()
        self.status_bar.showMessage(f"Selections reset. Default X axis: {self.x_column}")

    @staticmethod
    def _is_oxide_formula(token):
        parts = CHEMICAL_PART_PATTERN.findall(token)
        if not parts or "".join(symbol + digits for symbol, digits in parts) != token:
            return False
        symbols = [symbol for symbol, _digits in parts]
        return all(symbol in ELEMENT_SYMBOLS for symbol in symbols) and "O" in symbols

    def format_display_label(self, label):
        """Format recognized oxide formulas using Matplotlib subscripts.

        Formula digits are represented with math-text markup, for example
        ``Na2O`` becomes ``Na$_{2}$O``. No Unicode subscript characters are
        used, so labels remain portable across the selected font families.
        """
        text = str(label)
        if not self.chemical_format_checkbox.isChecked():
            return text

        def replace_formula(match):
            token = match.group(0)
            if not self._is_oxide_formula(token):
                return token
            formatted_parts = []
            for symbol, digits in CHEMICAL_PART_PATTERN.findall(token):
                formatted_parts.append(symbol)
                if digits:
                    formatted_parts.append(f"$_{{{digits}}}$")
            return "".join(formatted_parts)

        return CHEMICAL_TOKEN_PATTERN.sub(replace_formula, text)

    def chemical_formatting_changed(self, enabled):
        self.format_chemical_labels = bool(enabled)
        self.settings.setValue("plot/format_chemical_labels", self.format_chemical_labels)
        if hasattr(self, "axes"):
            self.update_plot()

    def vector_settings_changed(self, *_):
        if not hasattr(self, "vector_text_combo") or not hasattr(self, "font_combo"):
            return
        self.vector_text_mode = self.vector_text_combo.currentText()
        self.plot_font_family = self.font_combo.currentText().strip() or "DejaVu Sans"
        self.settings.setValue("vector/text_mode", self.vector_text_mode)
        self.settings.setValue("vector/font_family", self.plot_font_family)
        if hasattr(self, "axes"):
            self._apply_plot_font()
            self.canvas.draw_idle()

    def _apply_plot_font(self):
        if not hasattr(self, "axes") or not hasattr(self, "font_combo"):
            return
        family = self.font_combo.currentText().strip() or "Arial"
        # Use Matplotlib math text for chemical subscripts and request the same
        # family as the surrounding label. Matplotlib will fall back safely if
        # a particular mathematical glyph is unavailable in the selected font.
        matplotlib.rcParams["mathtext.fontset"] = "custom"
        matplotlib.rcParams["mathtext.rm"] = family
        matplotlib.rcParams["mathtext.it"] = f"{family}:italic"
        matplotlib.rcParams["mathtext.bf"] = f"{family}:bold"
        text_objects = [self.axes.title, self.axes.xaxis.label, self.axes.yaxis.label]
        text_objects.extend(self.axes.get_xticklabels())
        text_objects.extend(self.axes.get_yticklabels())
        for text_object in text_objects:
            text_object.set_fontfamily(family)

    def _svg_fonttype(self):
        return "none" if self.vector_text_combo.currentText() == "Editable text" else "path"

    def set_plot_type(self, plot_type_name):
        previous_type = self.plot_type_combo.currentText() if hasattr(self, "plot_type_combo") else None
        if plot_type_name != previous_type:
            self._reset_custom_plot_title()
        self.petrological_diagram = None
        self.petrological_columns = {}
        self.spider_configuration = None
        previous_type = previous_type or "Points"
        entering_ternary = plot_type_name == "Ternary" and previous_type != "Ternary"
        leaving_ternary = previous_type == "Ternary" and plot_type_name != "Ternary"

        if entering_ternary:
            # Ternary component B must be singular. Preserve the Cartesian
            # multi-Y selection so it can be restored when leaving ternary mode.
            self.cartesian_y_columns = list(self.y_columns)
            available_columns = list(self.dataframe.columns) if self.dataframe is not None else []

            def numeric_candidates(excluded=None):
                excluded = set(excluded or [])
                candidates = []
                for column in available_columns:
                    if column in excluded:
                        continue
                    numeric = pd.to_numeric(self.dataframe[column], errors="coerce")
                    if numeric.notna().any():
                        candidates.append(column)
                return candidates

            component_b = self.y_column or (self.y_columns[0] if self.y_columns else None)
            if component_b not in available_columns:
                choices = numeric_candidates({self.x_column})
                component_b = choices[0] if choices else None
            self.y_columns = [component_b] if component_b is not None else []
            self.y_column = component_b
            if hasattr(self, "y_button_row"):
                self.y_button_row.set_selected_multiple(self.y_columns)

            # Ensure Component C is available and different from A and B.
            if self.z_column not in available_columns or self.z_column in {self.x_column, self.y_column}:
                choices = numeric_candidates({self.x_column, self.y_column})
                self.z_column = choices[0] if choices else None
            if hasattr(self, "z_button_row"):
                if self.z_column is not None:
                    self.z_button_row.set_selected(self.z_column)
                else:
                    self.z_button_row.clear_selection()
        elif leaving_ternary:
            available = set(self.dataframe.columns) if self.dataframe is not None else set()
            restored = [column for column in self.cartesian_y_columns if column in available]
            if not restored and self.y_column in available:
                restored = [self.y_column]
            self.y_columns = restored
            self.y_column = restored[0] if restored else None
            if hasattr(self, "y_button_row"):
                self.y_button_row.set_selected_multiple(restored)

        self.settings.setValue("plot/type", plot_type_name)
        if hasattr(self, "plot_type_combo"):
            self.plot_type_combo.setCurrentText(plot_type_name)
        if hasattr(self, "plot_type_actions"):
            for name, action in self.plot_type_actions.items():
                action.setChecked(name == plot_type_name)
        self._update_plot_type_controls()
        self.update_plot()

    def plot_style_changed(self, *_):
        if not hasattr(self, "plot_type_combo"):
            return
        if self.plot_type_combo.currentText() == "Ternary" and len(self.y_columns) > 1:
            self.cartesian_y_columns = list(self.y_columns)
            self.y_columns = [self.y_columns[0]]
            self.y_column = self.y_columns[0]
            self.y_button_row.set_selected_multiple(self.y_columns)
        self._update_plot_type_controls()
        points = self.plot_type_combo.currentText() in ("Points", "Ternary")
        self.marker_combo.setEnabled(points)
        self.marker_size_spin.setEnabled(points)
        self.colour_button.setEnabled(points)
        self.stroke_colour_button.setEnabled(points)
        self.opacity_spin.setEnabled(points)
        self.stroke_width_spin.setEnabled(points)
        self.settings.setValue("plot/type", self.plot_type_combo.currentText())
        self.settings.setValue("plot/marker_name", self.marker_combo.currentText())
        self.settings.setValue("plot/marker_size", self.marker_size_spin.value())
        self.settings.setValue("plot/point_opacity", self.opacity_spin.value())
        self.settings.setValue("plot/stroke_width", self.stroke_width_spin.value())
        if hasattr(self, "canvas"):
            self.update_plot()

    def open_histogram_formatting_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Histogram Formatting")
        layout = QGridLayout(dialog)
        bins_spin = QSpinBox()
        bins_spin.setRange(1, 500)
        bins_spin.setValue(self.histogram_bins)
        layout.addWidget(QLabel("Number of bins:"), 0, 0)
        layout.addWidget(bins_spin, 0, 1)
        layout.addWidget(QLabel(
            "Histogram fill, stroke, opacity, and line width use the current Data Point Formatting settings."
        ), 1, 0, 1, 2)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons, 2, 0, 1, 2)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.histogram_bins = bins_spin.value()
            self.settings.setValue("histogram/bins", self.histogram_bins)
            if self.plot_type_combo.currentText() == "Histogram":
                self.update_plot()

    def _draw_histogram(self, plot_dataframe):
        if self.x_column is None:
            self._draw_placeholder()
            return
        values = pd.to_numeric(plot_dataframe[self.x_column], errors="coerce").dropna()
        self.axes.clear()
        self._restore_cartesian_axes()
        if values.empty:
            self.axes.set_title("No numeric values to plot")
            self.status_bar.showMessage(
                f"No numeric values are available in '{self.x_column}'."
            )
        else:
            face_colour = "none" if self.no_fill else self.plot_colour
            edge_colour = "none" if self.no_stroke else self.stroke_colour
            self.axes.hist(
                values.to_numpy(dtype=float),
                bins=self.histogram_bins,
                color=face_colour,
                edgecolor=edge_colour,
                linewidth=self.stroke_width_spin.value(),
                alpha=self.opacity_spin.value() / 100.0,
            )
            self.axes.set_title(f"Histogram of {self.format_display_label(self.x_column)}")
            self.status_bar.showMessage(
                f"Histogram: {len(values):,} numeric values in {self.histogram_bins} bins; "
                f"omitted {len(plot_dataframe)-len(values):,} missing or non-numeric rows."
            )
        self.axes.set_xlabel(self.custom_x_axis_title or self.format_display_label(self.x_column))
        self.axes.set_ylabel(self.custom_left_y_axis_title or "Frequency")
        self._draw_horizontal_qc_lines()
        self._apply_zero_minimum_limits()
        self._apply_manual_axis_ranges()
        self._apply_figure_ratio()
        self._apply_plot_theme()

    def open_horizontal_lines_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("QC Horizontal Lines")
        dialog.resize(660, 430)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(
            "Add horizontal QC reference lines at user-defined Y values. "
            "Lines are shown on point and line plots."
        ))
        line_list = QListWidget()
        layout.addWidget(line_list, 1)

        def line_text(item):
            label = item.get("label", "").strip()
            description = f"Y = {item['value']:g}"
            if label:
                description += f"  |  {label}"
            return description

        def refresh_list(selected_row=None):
            line_list.clear()
            for item in self.qc_horizontal_lines:
                line_list.addItem(line_text(item))
            if selected_row is not None and 0 <= selected_row < line_list.count():
                line_list.setCurrentRow(selected_row)

        refresh_list()
        controls = QHBoxLayout()
        add_button = QPushButton("Add...")
        edit_button = QPushButton("Edit...")
        remove_button = QPushButton("Remove")
        controls.addWidget(add_button)
        controls.addWidget(edit_button)
        controls.addWidget(remove_button)
        controls.addStretch(1)
        layout.addLayout(controls)

        def edit_line(existing=None):
            editor = QDialog(dialog)
            editor.setWindowTitle("Edit Horizontal Line" if existing else "Add Horizontal Line")
            form = QGridLayout(editor)
            value_spin = QDoubleSpinBox()
            value_spin.setRange(-1.0e15, 1.0e15)
            value_spin.setDecimals(8)
            value_spin.setValue(float(existing["value"]) if existing else 0.0)
            label_edit = QLineEdit(existing.get("label", "") if existing else "")
            label_edit.setPlaceholderText("Optional label, for example detection limit")
            style_combo = QComboBox()
            style_combo.addItems(["Solid", "Dashed", "Dash-dot", "Dotted"])
            style_combo.setCurrentText(existing.get("style", "Dashed") if existing else "Dashed")
            width_spin = QDoubleSpinBox()
            width_spin.setRange(0.1, 10.0)
            width_spin.setDecimals(1)
            width_spin.setSingleStep(0.1)
            width_spin.setValue(float(existing.get("width", 1.2)) if existing else 1.2)
            colour = {"value": existing.get("colour", "#d62728") if existing else "#d62728"}
            colour_button = QPushButton("Choose line colour...")
            def update_colour_button():
                c = QColor(colour["value"])
                foreground = "#000000" if c.lightness() > 140 else "#ffffff"
                colour_button.setStyleSheet(
                    f"background:{colour['value']};color:{foreground};font-weight:600;"
                )
            def choose_colour():
                selected = QColorDialog.getColor(QColor(colour["value"]), editor, "Choose QC line colour")
                if selected.isValid():
                    colour["value"] = selected.name()
                    update_colour_button()
            colour_button.clicked.connect(choose_colour)
            update_colour_button()
            form.addWidget(QLabel("Y value:"), 0, 0)
            form.addWidget(value_spin, 0, 1)
            form.addWidget(QLabel("Label:"), 1, 0)
            form.addWidget(label_edit, 1, 1)
            form.addWidget(QLabel("Line style:"), 2, 0)
            form.addWidget(style_combo, 2, 1)
            form.addWidget(QLabel("Line width:"), 3, 0)
            form.addWidget(width_spin, 3, 1)
            form.addWidget(colour_button, 4, 0, 1, 2)
            buttons = QDialogButtonBox(
                QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
            )
            buttons.accepted.connect(editor.accept)
            buttons.rejected.connect(editor.reject)
            form.addWidget(buttons, 5, 0, 1, 2)
            if editor.exec() != QDialog.DialogCode.Accepted:
                return None
            return {
                "value": value_spin.value(),
                "label": label_edit.text().strip(),
                "style": style_combo.currentText(),
                "width": width_spin.value(),
                "colour": colour["value"],
            }

        def add_line():
            item = edit_line()
            if item is not None:
                self.qc_horizontal_lines.append(item)
                refresh_list(len(self.qc_horizontal_lines) - 1)
                self.update_plot()

        def edit_selected():
            row = line_list.currentRow()
            if row < 0:
                return
            item = edit_line(self.qc_horizontal_lines[row])
            if item is not None:
                self.qc_horizontal_lines[row] = item
                refresh_list(row)
                self.update_plot()

        def remove_selected():
            row = line_list.currentRow()
            if row < 0:
                return
            self.qc_horizontal_lines.pop(row)
            refresh_list(min(row, len(self.qc_horizontal_lines) - 1))
            self.update_plot()

        add_button.clicked.connect(add_line)
        edit_button.clicked.connect(edit_selected)
        remove_button.clicked.connect(remove_selected)
        line_list.itemDoubleClicked.connect(lambda _item: edit_selected())
        close_buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close_buttons.rejected.connect(dialog.reject)
        layout.addWidget(close_buttons)
        dialog.exec()

    def clear_horizontal_qc_lines(self):
        if not self.qc_horizontal_lines:
            return
        self.qc_horizontal_lines.clear()
        self.update_plot()
        self.status_bar.showMessage("All horizontal QC lines were removed.")

    def _draw_horizontal_qc_lines(self):
        if self.plot_type_combo.currentText() == "Ternary":
            return
        for item in self.qc_horizontal_lines:
            self.axes.axhline(
                y=item["value"],
                color=item.get("colour", "#d62728"),
                linewidth=item.get("width", 1.2),
                linestyle=self._matplotlib_line_style(item.get("style", "Dashed")),
                zorder=4,
            )
            label = item.get("label", "").strip()
            if label:
                self.axes.annotate(
                    label,
                    xy=(1.0, item["value"]),
                    xycoords=("axes fraction", "data"),
                    xytext=(-5, 4),
                    textcoords="offset points",
                    ha="right",
                    va="bottom",
                    color=item.get("colour", "#d62728"),
                    fontsize=max(7, self.axis_number_size),
                    fontfamily=self.font_combo.currentText(),
                    clip_on=True,
                    zorder=5,
                )

    def _save_regression_settings(self):
        self.settings.setValue("regression/enabled", self.regression_enabled)
        self.settings.setValue("regression/model", self.regression_model)
        self.settings.setValue("regression/force_origin", self.regression_force_origin)
        self.settings.setValue("regression/show_equation", self.regression_show_equation)
        self.settings.setValue("regression/show_r2", self.regression_show_r2)
        self.settings.setValue("regression/show_n", self.regression_show_n)
        self.settings.setValue("regression/colour", self.regression_colour)
        self.settings.setValue("regression/width", self.regression_width)
        self.settings.setValue("regression/style", self.regression_style_name)
        self.settings.setValue("regression/opacity", self.regression_opacity)
        self.settings.setValue("regression/legend_corner", self.regression_legend_corner)

    def toggle_regression(self, enabled):
        self.regression_enabled = bool(enabled)
        if hasattr(self, "regression_toggle_button"):
            self.regression_toggle_button.blockSignals(True)
            self.regression_toggle_button.setChecked(self.regression_enabled)
            self.regression_toggle_button.setText(
                "Hide Regression" if self.regression_enabled else "Show Regression"
            )
            self.regression_toggle_button.blockSignals(False)
        if hasattr(self, "regression_legend_button"):
            self.regression_legend_button.setEnabled(self.regression_enabled)
        self._save_regression_settings()
        if hasattr(self, "axes"):
            self.update_plot()

    def set_regression_model(self, model_name):
        self.regression_model = model_name
        self.regression_force_origin = False if model_name != "Linear" else self.regression_force_origin
        self.settings.setValue("regression/model", model_name)
        if hasattr(self, "axes"):
            self.update_plot()

    def _data_legend_position(self):
        positions = [
            ("Top Left", "upper left"),
            ("Top Right", "upper right"),
            ("Bottom Right", "lower right"),
            ("Bottom Left", "lower left"),
        ]
        return positions[self.data_legend_corner % len(positions)]

    def _update_data_legend_button(self):
        if not hasattr(self, "data_legend_button"):
            return
        name, _location = self._data_legend_position()
        self.data_legend_button.setText(f"Data Legend: {name}")

    def cycle_data_legend_corner(self):
        self.data_legend_corner = (self.data_legend_corner + 1) % 4
        self.settings.setValue("plot/data_legend_corner", self.data_legend_corner)
        self._update_data_legend_button()
        self.update_plot()

    def _regression_legend_position(self):
        positions = [
            ("Top Left", 0.02, 0.98, "left", "top"),
            ("Top Right", 0.98, 0.98, "right", "top"),
            ("Bottom Right", 0.98, 0.02, "right", "bottom"),
            ("Bottom Left", 0.02, 0.02, "left", "bottom"),
        ]
        return positions[self.regression_legend_corner % len(positions)]

    def _update_regression_legend_button(self):
        if not hasattr(self, "regression_legend_button"):
            return
        name, _x, _y, _ha, _va = self._regression_legend_position()
        self.regression_legend_button.setText(f"Regression Label: {name}")

    def cycle_regression_legend_corner(self):
        self.regression_legend_corner = (self.regression_legend_corner + 1) % 4
        self.settings.setValue("regression/legend_corner", self.regression_legend_corner)
        self._update_regression_legend_button()
        if self.regression_enabled:
            self.update_plot()

    def open_regression_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Regression Options")
        layout = QGridLayout(dialog)
        enabled_box = QCheckBox("Show regression")
        enabled_box.setChecked(self.regression_enabled)
        model_combo = QComboBox()
        model_combo.addItems(["Linear", "Quadratic", "Cubic", "Exponential", "Logarithmic", "Power"])
        model_combo.setCurrentText(self.regression_model)
        origin_box = QCheckBox("Force linear regression through origin")
        origin_box.setChecked(self.regression_force_origin)
        origin_box.setEnabled(model_combo.currentText()=="Linear")
        model_combo.currentTextChanged.connect(lambda name: origin_box.setEnabled(name=="Linear"))
        equation_box = QCheckBox("Show equation")
        equation_box.setChecked(self.regression_show_equation)
        r2_box = QCheckBox("Show R²")
        r2_box.setChecked(self.regression_show_r2)
        n_box = QCheckBox("Show sample count (n)")
        n_box.setChecked(self.regression_show_n)
        style_combo = QComboBox(); style_combo.addItems(["Solid","Dashed","Dash-dot","Dotted"])
        style_combo.setCurrentText(self.regression_style_name)
        width_spin = QDoubleSpinBox(); width_spin.setRange(0.1,10.0); width_spin.setDecimals(1); width_spin.setValue(self.regression_width)
        opacity_spin = QSpinBox(); opacity_spin.setRange(0,100); opacity_spin.setSuffix("%"); opacity_spin.setValue(self.regression_opacity)
        selected_colour={"value":self.regression_colour}
        colour_button=QPushButton("Choose regression colour...")
        preview_figure=Figure(figsize=(4.0,1.45),dpi=100); preview_canvas=FigureCanvas(preview_figure)
        preview_axes=preview_figure.add_subplot(111); preview_canvas.setMinimumHeight(150)
        def update_preview(*_):
            import numpy as np
            preview_axes.clear()
            px=np.linspace(0.4,3.0,20)
            model=model_combo.currentText()
            if model=="Quadratic": py=0.2*px**2+0.3
            elif model=="Cubic": py=0.08*px**3+0.35
            elif model=="Exponential": py=0.45*np.exp(0.55*px)
            elif model=="Logarithmic": py=0.75*np.log(px)+1.0
            elif model=="Power": py=0.5*px**1.4
            else: py=0.55*px+0.3
            preview_axes.scatter(px[::4],py[::4],s=20,color="#777777")
            preview_axes.plot(px,py,color=selected_colour["value"],linewidth=width_spin.value(),linestyle=self._matplotlib_line_style(style_combo.currentText()),alpha=opacity_spin.value()/100.0)
            preview_axes.set_xticks([]); preview_axes.set_yticks([]); preview_axes.set_title(f"{model} regression preview",fontsize=9)
            preview_figure.tight_layout(pad=0.6); preview_canvas.draw_idle()
        def choose_colour():
            c=QColorDialog.getColor(QColor(selected_colour["value"]),dialog,"Choose regression colour")
            if c.isValid(): selected_colour["value"]=c.name(); update_preview()
        colour_button.clicked.connect(choose_colour)
        for widget in (model_combo,style_combo,width_spin,opacity_spin):
            (widget.currentTextChanged if isinstance(widget,QComboBox) else widget.valueChanged).connect(update_preview)
        update_preview()
        rows=[("Regression model:",model_combo),("Line style:",style_combo),("Line width:",width_spin),("Line opacity:",opacity_spin)]
        layout.addWidget(enabled_box,0,0,1,2); layout.addWidget(origin_box,1,0,1,2)
        layout.addWidget(equation_box,2,0,1,2); layout.addWidget(r2_box,3,0,1,2); layout.addWidget(n_box,4,0,1,2)
        for row,(label,widget) in enumerate(rows,start=5): layout.addWidget(QLabel(label),row,0); layout.addWidget(widget,row,1)
        layout.addWidget(colour_button,9,0,1,2); layout.addWidget(preview_canvas,10,0,1,2)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept); buttons.rejected.connect(dialog.reject); layout.addWidget(buttons,14,0,1,2)
        if dialog.exec()==QDialog.DialogCode.Accepted:
            self.regression_enabled=enabled_box.isChecked(); self.regression_model=model_combo.currentText()
            self.regression_force_origin=origin_box.isChecked() if self.regression_model=="Linear" else False
            self.regression_show_equation=equation_box.isChecked(); self.regression_show_r2=r2_box.isChecked(); self.regression_show_n=n_box.isChecked()
            self.regression_colour=selected_colour["value"]; self.regression_width=width_spin.value(); self.regression_style_name=style_combo.currentText(); self.regression_opacity=opacity_spin.value()
            self.regression_toggle_button.blockSignals(True)
            self.regression_toggle_button.setChecked(self.regression_enabled)
            self.regression_toggle_button.setText(
                "Hide Regression" if self.regression_enabled else "Show Regression"
            )
            self.regression_toggle_button.blockSignals(False)
            for name,action in self.regression_model_actions.items(): action.setChecked(name==self.regression_model)
            self._save_regression_settings(); self.update_plot()

    def _draw_regression(self, x_values, y_values):
        import numpy as np
        x=np.asarray(x_values,dtype=float); y=np.asarray(y_values,dtype=float)
        valid=np.isfinite(x)&np.isfinite(y)
        model=self.regression_model
        if model=="Exponential": valid &= y>0
        elif model=="Logarithmic": valid &= x>0
        elif model=="Power": valid &= (x>0)&(y>0)
        x=x[valid]; y=y[valid]
        minimum={"Linear":2,"Quadratic":3,"Cubic":4,"Exponential":2,"Logarithmic":2,"Power":2}[model]
        if x.size<minimum or np.allclose(x,x[0]):
            self.status_bar.showMessage(f"Insufficient valid data for {model.lower()} regression.")
            return
        if model=="Linear":
            if self.regression_force_origin:
                slope=float(np.dot(x,y)/np.dot(x,x)); intercept=0.0
            else: slope,intercept=map(float,np.polyfit(x,y,1))
            predict=lambda values:slope*values+intercept
            equation=f"y = {slope:.4g}x" if self.regression_force_origin else f"y = {slope:.4g}x {'+' if intercept>=0 else '−'} {abs(intercept):.4g}"
        elif model in ("Quadratic","Cubic"):
            degree=2 if model=="Quadratic" else 3; coefficients=np.polyfit(x,y,degree); predict=lambda values:np.polyval(coefficients,values)
            if degree==2: equation=f"y = {coefficients[0]:.4g}x² + {coefficients[1]:.4g}x + {coefficients[2]:.4g}"
            else: equation=f"y = {coefficients[0]:.4g}x³ + {coefficients[1]:.4g}x² + {coefficients[2]:.4g}x + {coefficients[3]:.4g}"
        elif model=="Exponential":
            b,loga=np.polyfit(x,np.log(y),1); a=float(np.exp(loga)); b=float(b); predict=lambda values:a*np.exp(b*values); equation=f"y = {a:.4g}e^({b:.4g}x)"
        elif model=="Logarithmic":
            a,b=np.polyfit(np.log(x),y,1); a=float(a); b=float(b); predict=lambda values:a*np.log(values)+b; equation=f"y = {a:.4g}ln(x) + {b:.4g}"
        else:
            b,loga=np.polyfit(np.log(x),np.log(y),1); a=float(np.exp(loga)); b=float(b); predict=lambda values:a*values**b; equation=f"y = {a:.4g}x^{b:.4g}"
        yhat=predict(x); rss=float(np.sum((y-yhat)**2)); tss=float(np.sum((y-y.mean())**2)); r2=1-rss/tss if tss>0 else float("nan")
        xline=np.linspace(x.min(),x.max(),300); yline=predict(xline)
        self.axes.plot(xline,yline,color=self.regression_colour,linewidth=self.regression_width,linestyle=self._matplotlib_line_style(self.regression_style_name),alpha=self.regression_opacity/100.0,zorder=5)
        lines=[]
        if self.regression_show_equation: lines.append(equation)
        if self.regression_show_r2 and np.isfinite(r2): lines.append(f"R² = {r2:.4f}")
        if self.regression_show_n: lines.append(f"n = {x.size}")
        if lines:
            c=self.current_theme
            _corner_name, text_x, text_y, horizontal_alignment, vertical_alignment = self._regression_legend_position()
            self.axes.text(
                text_x,
                text_y,
                "\n".join(lines),
                transform=self.axes.transAxes,
                ha=horizontal_alignment,
                va=vertical_alignment,
                color=c["text"],
                fontfamily=self.font_combo.currentText(),
                bbox=dict(
                    facecolor=c["plot_background"],
                    edgecolor=c["border"],
                    alpha=0.85,
                    pad=5,
                ),
                zorder=6,
            )

    def _save_axis_range_settings(self):
        self.settings.setValue("axes/x_range_auto", self.x_range_auto)
        self.settings.setValue("axes/left_y_range_auto", self.left_y_range_auto)
        self.settings.setValue("axes/right_y_range_auto", self.right_y_range_auto)
        self.settings.setValue("axes/x_range_min", self.x_range_min)
        self.settings.setValue("axes/x_range_max", self.x_range_max)
        self.settings.setValue("axes/left_y_range_min", self.left_y_range_min)
        self.settings.setValue("axes/left_y_range_max", self.left_y_range_max)
        self.settings.setValue("axes/right_y_range_min", self.right_y_range_min)
        self.settings.setValue("axes/right_y_range_max", self.right_y_range_max)

    def _update_axis_range_button(self):
        if not hasattr(self, "reset_axis_ranges_button"):
            return
        all_auto = self.x_range_auto and self.left_y_range_auto and self.right_y_range_auto
        self.reset_axis_ranges_button.setEnabled(not all_auto)
        self.reset_axis_ranges_button.setText(
            "Axes: Auto" if all_auto else "Reset Axes to Auto"
        )

    def reset_axis_ranges_to_auto(self):
        self.x_range_auto = True
        self.left_y_range_auto = True
        self.right_y_range_auto = True
        self._save_axis_range_settings()
        self._update_axis_range_button()
        self.update_plot()
        self.status_bar.showMessage("All axis ranges were reset to automatic scaling.")

    def open_axis_ranges_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Axis Ranges")
        dialog.resize(540, 330)
        layout = QGridLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setHorizontalSpacing(12)
        layout.setVerticalSpacing(10)
        layout.addWidget(QLabel("Axis"), 0, 0)
        layout.addWidget(QLabel("Automatic"), 0, 1)
        layout.addWidget(QLabel("Minimum"), 0, 2)
        layout.addWidget(QLabel("Maximum"), 0, 3)

        rows = []
        specifications = [
            ("X axis", self.x_range_auto, self.x_range_min, self.x_range_max),
            ("Left Y axis", self.left_y_range_auto, self.left_y_range_min, self.left_y_range_max),
            ("Right Y axis", self.right_y_range_auto, self.right_y_range_min, self.right_y_range_max),
        ]
        for row_number, (name, automatic, minimum, maximum) in enumerate(specifications, start=1):
            auto_box = QCheckBox()
            auto_box.setChecked(automatic)
            minimum_spin = QDoubleSpinBox()
            maximum_spin = QDoubleSpinBox()
            for spin in (minimum_spin, maximum_spin):
                spin.setRange(-1.0e15, 1.0e15)
                spin.setDecimals(8)
                spin.setKeyboardTracking(False)
            minimum_spin.setValue(minimum)
            maximum_spin.setValue(maximum)
            minimum_spin.setEnabled(not automatic)
            maximum_spin.setEnabled(not automatic)
            auto_box.toggled.connect(lambda checked, widget=minimum_spin: widget.setEnabled(not checked))
            auto_box.toggled.connect(lambda checked, widget=maximum_spin: widget.setEnabled(not checked))
            layout.addWidget(QLabel(name), row_number, 0)
            layout.addWidget(auto_box, row_number, 1, alignment=Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(minimum_spin, row_number, 2)
            layout.addWidget(maximum_spin, row_number, 3)
            rows.append((auto_box, minimum_spin, maximum_spin))

        note = QLabel(
            "Manual limits override the X min = 0 and Y min = 0 controls. "
            "The right-Y range is applied when a secondary Y axis is enabled."
        )
        note.setWordWrap(True)
        layout.addWidget(note, 4, 0, 1, 4)
        reset_button = QPushButton("Reset All to Auto")
        layout.addWidget(reset_button, 5, 0, 1, 2)
        def set_all_auto():
            for auto_box, _minimum, _maximum in rows:
                auto_box.setChecked(True)
        reset_button.clicked.connect(set_all_auto)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons, 5, 2, 1, 2)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        for name, (_auto, minimum, maximum) in zip(
            ("X axis", "Left Y axis", "Right Y axis"), rows
        ):
            if not _auto.isChecked() and minimum.value() >= maximum.value():
                QMessageBox.warning(
                    self, "Invalid axis range",
                    f"{name}: minimum must be less than maximum."
                )
                return
        (x_auto, x_min, x_max), (left_auto, left_min, left_max), (right_auto, right_min, right_max) = rows
        self.x_range_auto = x_auto.isChecked()
        self.left_y_range_auto = left_auto.isChecked()
        self.right_y_range_auto = right_auto.isChecked()
        self.x_range_min, self.x_range_max = x_min.value(), x_max.value()
        self.left_y_range_min, self.left_y_range_max = left_min.value(), left_max.value()
        self.right_y_range_min, self.right_y_range_max = right_min.value(), right_max.value()
        self._save_axis_range_settings()
        self._update_axis_range_button()
        self.update_plot()

    def _apply_manual_axis_ranges(self):
        if self.plot_type_combo.currentText() == "Ternary":
            return
        if not self.x_range_auto:
            self.axes.set_xlim(self.x_range_min, self.x_range_max)
        if not self.left_y_range_auto:
            self.axes.set_ylim(self.left_y_range_min, self.left_y_range_max)
        if self.secondary_axes is not None and not self.right_y_range_auto:
            self.secondary_axes.set_ylim(self.right_y_range_min, self.right_y_range_max)

    def _save_axis_format_settings(self):
        self.settings.setValue("axes/title_size", self.axis_title_size)
        self.settings.setValue("axes/title_bold", self.axis_title_bold)
        self.settings.setValue("axes/title_colour", self.axis_title_colour)
        self.settings.setValue("axes/number_size", self.axis_number_size)
        self.settings.setValue("axes/number_colour", self.axis_number_colour)
        self.settings.setValue("axes/number_format", self.axis_number_format)
        self.settings.setValue("axes/decimal_places", self.axis_decimal_places)
        self.settings.setValue("axes/scientific_limits", self.axis_scientific_limits)

    def open_axis_formatting_dialog(self):
        dialog=QDialog(self)
        dialog.setWindowTitle("Axis Formatting")
        dialog.resize(680, 760)
        dialog.setMinimumWidth(620)
        layout=QGridLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setHorizontalSpacing(12)
        layout.setVerticalSpacing(8)
        layout.setColumnStretch(1, 1)

        title_size=QSpinBox(); title_size.setRange(6,36); title_size.setValue(self.axis_title_size)
        title_bold=QCheckBox("Bold axis titles"); title_bold.setChecked(self.axis_title_bold)
        number_size=QSpinBox(); number_size.setRange(6,30); number_size.setValue(self.axis_number_size)
        number_format=QComboBox(); number_format.addItems(["Automatic","Fixed decimal","Scientific notation"])
        number_format.setCurrentText(self.axis_number_format)
        decimals=QSpinBox(); decimals.setRange(0,10); decimals.setValue(self.axis_decimal_places)
        scientific_limits=QSpinBox(); scientific_limits.setRange(1,12); scientific_limits.setValue(self.axis_scientific_limits)
        scientific_limits.setToolTip("Use scientific notation at values >= 10^n or <= 10^-n in Automatic mode")
        secondary_y_box = QCheckBox("Use a second Y axis on the right")
        secondary_y_box.setChecked(self.secondary_y_enabled)
        secondary_y_combo = QComboBox()
        secondary_y_combo.addItems([str(column) for column in self.y_columns])
        if self.secondary_y_column in self.y_columns:
            secondary_y_combo.setCurrentText(str(self.secondary_y_column))
        secondary_y_combo.setEnabled(secondary_y_box.isChecked())
        secondary_y_box.toggled.connect(secondary_y_combo.setEnabled)
        custom_x_title_edit = QLineEdit(self.custom_x_axis_title)
        custom_x_title_edit.setPlaceholderText("Blank = use selected X variable")
        custom_left_y_title_edit = QLineEdit(self.custom_left_y_axis_title)
        custom_left_y_title_edit.setPlaceholderText("Blank = use selected left-axis Y variable(s)")
        custom_right_y_title_edit = QLineEdit(self.custom_right_y_axis_title)
        custom_right_y_title_edit.setPlaceholderText("Blank = use selected right-axis Y variable")
        custom_right_y_title_edit.setEnabled(secondary_y_box.isChecked())
        secondary_y_box.toggled.connect(custom_right_y_title_edit.setEnabled)
        for title_edit in (custom_x_title_edit, custom_left_y_title_edit, custom_right_y_title_edit):
            title_edit.setMinimumWidth(280)
            title_edit.setClearButtonEnabled(True)

        colours={"title":self.axis_title_colour,"numbers":self.axis_number_colour}
        title_colour_button=QPushButton("Choose axis-title colour...")
        number_colour_button=QPushButton("Choose axis-number colour...")
        def style_colour_button(button,key):
            c=QColor(colours[key]); fg="#000000" if c.lightness()>140 else "#ffffff"
            button.setStyleSheet(f"background:{colours[key]};color:{fg};font-weight:600;")
        def choose_colour(key,button,title):
            c=QColorDialog.getColor(QColor(colours[key]),dialog,title)
            if c.isValid(): colours[key]=c.name(); style_colour_button(button,key); update_preview()
        title_colour_button.clicked.connect(lambda:choose_colour("title",title_colour_button,"Choose axis-title colour"))
        number_colour_button.clicked.connect(lambda:choose_colour("numbers",number_colour_button,"Choose axis-number colour"))
        style_colour_button(title_colour_button,"title"); style_colour_button(number_colour_button,"numbers")

        preview_figure=Figure(figsize=(4.5,2.2),dpi=100)
        preview_canvas=FigureCanvas(preview_figure)
        preview_axes=preview_figure.add_subplot(111)
        preview_canvas.setMinimumHeight(155)
        preview_canvas.setMaximumHeight(180)
        def update_preview(*_):
            preview_axes.clear()
            preview_axes.plot([0,1,2],[0,1234.567,2469.134],color="#777777")
            preview_axes.set_xlabel(r"Si$_{2}$O (wt%)",fontsize=title_size.value(),fontweight="bold" if title_bold.isChecked() else "normal",color=colours["title"])
            preview_axes.set_ylabel(r"Ti$_{2}$O (wt%)",fontsize=title_size.value(),fontweight="bold" if title_bold.isChecked() else "normal",color=colours["title"])
            preview_axes.tick_params(labelsize=number_size.value(),labelcolor=colours["numbers"])
            mode=number_format.currentText()
            if mode=="Fixed decimal":
                formatter=FormatStrFormatter(f"%.{decimals.value()}f")
                preview_axes.xaxis.set_major_formatter(formatter); preview_axes.yaxis.set_major_formatter(formatter)
            elif mode=="Scientific notation":
                formatter=ScalarFormatter(useMathText=True); formatter.set_scientific(True); formatter.set_powerlimits((0,0))
                preview_axes.xaxis.set_major_formatter(formatter); preview_axes.yaxis.set_major_formatter(formatter)
            preview_figure.tight_layout(pad=1.2); preview_canvas.draw_idle()
        for widget in (title_size,title_bold,number_size,number_format,decimals,scientific_limits):
            signal=widget.toggled if isinstance(widget,QCheckBox) else widget.currentTextChanged if isinstance(widget,QComboBox) else widget.valueChanged
            signal.connect(update_preview)
        update_preview()

        layout.addWidget(QLabel("Axis-title size:"),0,0); layout.addWidget(title_size,0,1)
        layout.addWidget(title_bold,1,0,1,2)
        layout.addWidget(title_colour_button,2,0,1,2)
        layout.addWidget(QLabel("Axis-number size:"),3,0); layout.addWidget(number_size,3,1)
        layout.addWidget(number_colour_button,4,0,1,2)
        layout.addWidget(QLabel("Number format:"),5,0); layout.addWidget(number_format,5,1)
        layout.addWidget(QLabel("Decimal places:"),6,0); layout.addWidget(decimals,6,1)
        layout.addWidget(QLabel("Automatic scientific threshold (10^n):"),7,0); layout.addWidget(scientific_limits,7,1)
        layout.addWidget(secondary_y_box,8,0,1,2)
        layout.addWidget(QLabel("Right-axis Y variable:"),9,0); layout.addWidget(secondary_y_combo,9,1)
        layout.addWidget(QLabel("Custom X-axis title:"),10,0); layout.addWidget(custom_x_title_edit,10,1)
        layout.addWidget(QLabel("Custom left Y-axis title:"),11,0); layout.addWidget(custom_left_y_title_edit,11,1)
        layout.addWidget(QLabel("Custom right Y-axis title:"),12,0); layout.addWidget(custom_right_y_title_edit,12,1)
        layout.addWidget(preview_canvas,13,0,1,2)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept); buttons.rejected.connect(dialog.reject); layout.addWidget(buttons,14,0,1,2)
        if dialog.exec()==QDialog.DialogCode.Accepted:
            self.axis_title_size=title_size.value(); self.axis_title_bold=title_bold.isChecked(); self.axis_title_colour=colours["title"]
            self.axis_number_size=number_size.value(); self.axis_number_colour=colours["numbers"]
            self.axis_number_format=number_format.currentText(); self.axis_decimal_places=decimals.value(); self.axis_scientific_limits=scientific_limits.value()
            self.secondary_y_enabled = secondary_y_box.isChecked() and bool(secondary_y_combo.currentText())
            self.secondary_y_column = secondary_y_combo.currentText() if self.secondary_y_enabled else ""
            self.custom_x_axis_title = custom_x_title_edit.text().strip()
            self.custom_left_y_axis_title = custom_left_y_title_edit.text().strip()
            self.custom_right_y_axis_title = custom_right_y_title_edit.text().strip()
            self.settings.setValue("axes/custom_x_title", self.custom_x_axis_title)
            self.settings.setValue("axes/custom_left_y_title", self.custom_left_y_axis_title)
            self.settings.setValue("axes/custom_right_y_title", self.custom_right_y_axis_title)
            self.settings.setValue("axes/secondary_y_enabled", self.secondary_y_enabled)
            self.settings.setValue("axes/secondary_y_column", self.secondary_y_column)
            self._save_axis_format_settings(); self.update_plot()

    def _apply_axis_formatting(self):
        weight="bold" if self.axis_title_bold else "normal"
        self.axes.xaxis.label.set_fontsize(self.axis_title_size)
        self.axes.yaxis.label.set_fontsize(self.axis_title_size)
        self.axes.xaxis.label.set_fontweight(weight)
        self.axes.yaxis.label.set_fontweight(weight)
        self.axes.xaxis.label.set_color(self.axis_title_colour)
        self.axes.yaxis.label.set_color(self.axis_title_colour)
        self.axes.tick_params(axis="both",labelsize=self.axis_number_size,labelcolor=self.axis_number_colour)
        if self.axis_number_format=="Fixed decimal":
            self.axes.xaxis.set_major_formatter(FormatStrFormatter(f"%.{self.axis_decimal_places}f"))
            self.axes.yaxis.set_major_formatter(FormatStrFormatter(f"%.{self.axis_decimal_places}f"))
        else:
            xformatter=ScalarFormatter(useMathText=True); yformatter=ScalarFormatter(useMathText=True)
            xformatter.set_scientific(True); yformatter.set_scientific(True)
            limits=(0,0) if self.axis_number_format=="Scientific notation" else (-self.axis_scientific_limits,self.axis_scientific_limits)
            xformatter.set_powerlimits(limits); yformatter.set_powerlimits(limits)
            self.axes.xaxis.set_major_formatter(xformatter); self.axes.yaxis.set_major_formatter(yformatter)

    def open_text_formatting_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Text Formatting")
        layout = QGridLayout(dialog)

        vector_combo = QComboBox()
        vector_combo.addItems(["Editable text", "Text as curves"])
        vector_combo.setCurrentText(self.vector_text_combo.currentText())

        font_combo = QComboBox()
        font_combo.setEditable(True)
        font_combo.addItems(QFontDatabase.families())
        font_combo.setCurrentText(self.font_combo.currentText())

        preview = QLabel("Al2O3 versus SiO2")
        preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview.setMinimumHeight(70)
        preview.setStyleSheet("border:1px solid #888; padding:12px; font-size:18px;")
        def update_preview(*_):
            preview.setStyleSheet(
                f"border:1px solid #888;padding:12px;font-size:18px;"
                f"font-family:'{font_combo.currentText()}';"
            )
        font_combo.currentTextChanged.connect(update_preview)
        update_preview()

        layout.addWidget(QLabel("SVG vector text:"),0,0)
        layout.addWidget(vector_combo,0,1)
        layout.addWidget(QLabel("Plot font:"),1,0)
        layout.addWidget(font_combo,1,1)
        layout.addWidget(QLabel("Preview:"),2,0)
        layout.addWidget(preview,2,1)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept); buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons,3,0,1,2)
        if dialog.exec()==QDialog.DialogCode.Accepted:
            self.vector_text_combo.setCurrentText(vector_combo.currentText())
            self.font_combo.setCurrentText(font_combo.currentText())
            self.vector_settings_changed()

    def open_line_formatting_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Line Formatting")
        layout = QGridLayout(dialog)

        style_combo = QComboBox()
        style_combo.addItems(["Solid", "Dashed", "Dash-dot", "Dotted"])
        style_combo.setCurrentText(self.line_style_name)

        width_spin = QDoubleSpinBox()
        width_spin.setRange(0.1, 10.0)
        width_spin.setDecimals(1)
        width_spin.setSingleStep(0.1)
        width_spin.setValue(self.line_width)

        opacity_spin = QSpinBox()
        opacity_spin.setRange(0, 100)
        opacity_spin.setSuffix("%")
        opacity_spin.setValue(self.line_opacity)

        marker_box = QCheckBox("Show markers on line")
        marker_box.setChecked(self.line_show_markers)
        marker_combo = QComboBox()
        marker_combo.addItems(MARKERS.keys())
        marker_combo.setCurrentText(self.line_marker_name)
        marker_size_spin = QDoubleSpinBox()
        marker_size_spin.setRange(1.0, 50.0)
        marker_size_spin.setDecimals(1)
        marker_size_spin.setSingleStep(0.5)
        marker_size_spin.setValue(self.line_marker_size)
        marker_combo.setEnabled(marker_box.isChecked())
        marker_size_spin.setEnabled(marker_box.isChecked())
        marker_box.toggled.connect(marker_combo.setEnabled)
        marker_box.toggled.connect(marker_size_spin.setEnabled)

        selected_colour = {"value": self.line_colour}
        colour_button = QPushButton("Choose line colour...")
        def choose_line_colour():
            colour = QColorDialog.getColor(
                QColor(selected_colour["value"]), dialog, "Choose line colour"
            )
            if colour.isValid():
                selected_colour["value"] = colour.name()
                preview = QColor(selected_colour["value"])
                foreground = "#000000" if preview.lightness() > 140 else "#ffffff"
                colour_button.setStyleSheet(
                    f"background:{selected_colour['value']};color:{foreground};font-weight:600;"
                )
        colour_button.clicked.connect(choose_line_colour)
        preview = QColor(selected_colour["value"])
        foreground = "#000000" if preview.lightness() > 140 else "#ffffff"
        colour_button.setStyleSheet(
            f"background:{selected_colour['value']};color:{foreground};font-weight:600;"
        )

        preview_figure = Figure(figsize=(4.0, 1.15), dpi=100)
        preview_canvas = FigureCanvas(preview_figure)
        preview_axes = preview_figure.add_subplot(111)
        preview_canvas.setMinimumHeight(125)
        def update_line_preview(*_):
            preview_axes.clear()
            marker = MARKERS[marker_combo.currentText()] if marker_box.isChecked() else None
            preview_axes.plot([0,1,2,3],[0.25,0.72,0.48,0.82],
                color=selected_colour["value"], linewidth=width_spin.value(),
                linestyle=self._matplotlib_line_style(style_combo.currentText()),
                alpha=opacity_spin.value()/100.0, marker=marker,
                markersize=marker_size_spin.value() if marker_box.isChecked() else None)
            preview_axes.set_xticks([]); preview_axes.set_yticks([])
            preview_axes.set_title("Line preview", fontsize=9)
            preview_figure.tight_layout(pad=0.6)
            preview_canvas.draw_idle()
        for widget in (style_combo,width_spin,opacity_spin,marker_box,marker_combo,marker_size_spin):
            signal = widget.toggled if isinstance(widget,QCheckBox) else widget.currentTextChanged if isinstance(widget,QComboBox) else widget.valueChanged
            signal.connect(update_line_preview)
        colour_button.clicked.connect(update_line_preview)
        update_line_preview()

        layout.addWidget(QLabel("Line style:"), 0, 0)
        layout.addWidget(style_combo, 0, 1)
        layout.addWidget(QLabel("Line width:"), 1, 0)
        layout.addWidget(width_spin, 1, 1)
        layout.addWidget(QLabel("Line opacity:"), 2, 0)
        layout.addWidget(opacity_spin, 2, 1)
        layout.addWidget(colour_button, 3, 0, 1, 2)
        layout.addWidget(marker_box, 4, 0, 1, 2)
        layout.addWidget(QLabel("Marker shape:"), 5, 0)
        layout.addWidget(marker_combo, 5, 1)
        layout.addWidget(QLabel("Marker size:"), 6, 0)
        layout.addWidget(marker_size_spin, 6, 1)
        layout.addWidget(preview_canvas, 7, 0, 1, 2)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons, 8, 0, 1, 2)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.line_colour = selected_colour["value"]
            self.line_width = width_spin.value()
            self.line_style_name = style_combo.currentText()
            self.line_opacity = opacity_spin.value()
            self.line_show_markers = marker_box.isChecked()
            self.line_marker_name = marker_combo.currentText()
            self.line_marker_size = marker_size_spin.value()
            self.settings.setValue("line/colour", self.line_colour)
            self.settings.setValue("line/width", self.line_width)
            self.settings.setValue("line/style", self.line_style_name)
            self.settings.setValue("line/opacity", self.line_opacity)
            self.settings.setValue("line/show_markers", self.line_show_markers)
            self.settings.setValue("line/marker_name", self.line_marker_name)
            self.settings.setValue("line/marker_size", self.line_marker_size)
            if self.plot_type_combo.currentText() == "Line":
                self.update_plot()
            self.status_bar.showMessage("Line formatting updated.")

    @staticmethod
    def _matplotlib_line_style(style_name):
        return {
            "Solid": "-",
            "Dashed": "--",
            "Dash-dot": "-.",
            "Dotted": ":",
        }.get(style_name, "-")

    def open_data_point_formatting_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Data Point Formatting")
        layout = QGridLayout(dialog)

        marker_combo = QComboBox()
        marker_combo.addItems(MARKERS.keys())
        marker_combo.setCurrentText(self.marker_combo.currentText())
        size_spin = QSpinBox()
        size_spin.setRange(1, 500)
        size_spin.setValue(self.marker_size_spin.value())
        opacity_spin = QSpinBox()
        opacity_spin.setRange(0, 100)
        opacity_spin.setSuffix("%")
        opacity_spin.setValue(self.opacity_spin.value())
        width_spin = QDoubleSpinBox()
        width_spin.setRange(0.0, 10.0)
        width_spin.setDecimals(1)
        width_spin.setSingleStep(0.1)
        width_spin.setValue(self.stroke_width_spin.value())
        no_fill_box = QCheckBox("No fill (open symbols)")
        no_fill_box.setChecked(self.no_fill)
        no_stroke_box = QCheckBox("No stroke")
        no_stroke_box.setChecked(self.no_stroke)

        fill_button = QPushButton("Choose fill colour...")
        stroke_button = QPushButton("Choose stroke colour...")
        selected = {"fill": self.plot_colour, "stroke": self.stroke_colour}
        def choose(key, title):
            colour = QColorDialog.getColor(QColor(selected[key]), dialog, title)
            if colour.isValid():
                selected[key] = colour.name()
        fill_button.clicked.connect(lambda: choose("fill", "Choose marker fill colour"))
        stroke_button.clicked.connect(lambda: choose("stroke", "Choose marker stroke colour"))

        preview_figure = Figure(figsize=(4.0, 1.25), dpi=100)
        preview_canvas = FigureCanvas(preview_figure)
        preview_axes = preview_figure.add_subplot(111)
        preview_canvas.setMinimumHeight(135)
        def update_point_preview(*_):
            preview_axes.clear()
            face = "none" if no_fill_box.isChecked() else QColor(selected["fill"])
            if face != "none":
                face.setAlphaF(opacity_spin.value()/100.0)
                face = face.getRgbF()
            edge = "none" if no_stroke_box.isChecked() or width_spin.value()==0 else selected["stroke"]
            preview_axes.scatter([0.7,1.5,2.3],[0.48,0.64,0.42], s=size_spin.value(),
                marker=MARKERS[marker_combo.currentText()], facecolors=face,
                edgecolors=edge, linewidths=width_spin.value())
            preview_axes.set_xlim(0,3); preview_axes.set_ylim(0,1)
            preview_axes.set_xticks([]); preview_axes.set_yticks([])
            preview_axes.set_title("Data point preview", fontsize=9)
            preview_figure.tight_layout(pad=0.6)
            preview_canvas.draw_idle()
        for widget in (marker_combo,size_spin,opacity_spin,width_spin,no_fill_box,no_stroke_box):
            signal=widget.toggled if isinstance(widget,QCheckBox) else widget.currentTextChanged if isinstance(widget,QComboBox) else widget.valueChanged
            signal.connect(update_point_preview)
        fill_button.clicked.connect(update_point_preview); stroke_button.clicked.connect(update_point_preview)
        update_point_preview()

        layout.addWidget(QLabel("Symbol:"),0,0); layout.addWidget(marker_combo,0,1)
        layout.addWidget(QLabel("Marker size:"),1,0); layout.addWidget(size_spin,1,1)
        layout.addWidget(QLabel("Fill opacity:"),2,0); layout.addWidget(opacity_spin,2,1)
        layout.addWidget(QLabel("Stroke width:"),3,0); layout.addWidget(width_spin,3,1)
        layout.addWidget(fill_button,4,0,1,2)
        layout.addWidget(stroke_button,5,0,1,2)
        layout.addWidget(no_fill_box,6,0,1,2)
        layout.addWidget(no_stroke_box,7,0,1,2)
        layout.addWidget(preview_canvas,8,0,1,2)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept); buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons,9,0,1,2)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.marker_combo.setCurrentText(marker_combo.currentText())
            self.marker_size_spin.setValue(size_spin.value())
            self.opacity_spin.setValue(opacity_spin.value())
            self.stroke_width_spin.setValue(width_spin.value())
            self.plot_colour = selected["fill"]
            self.stroke_colour = selected["stroke"]
            self.no_fill = no_fill_box.isChecked()
            self.no_stroke = no_stroke_box.isChecked()
            self._update_colour_button()
            self._update_stroke_colour_button()
            self.settings.setValue("plot/colour", self.plot_colour)
            self.settings.setValue("plot/stroke_colour", self.stroke_colour)
            self.settings.setValue("plot/no_fill", self.no_fill)
            self.settings.setValue("plot/no_stroke", self.no_stroke)
            self.plot_style_changed()

    def choose_stroke_colour(self):
        colour = QColorDialog.getColor(QColor(self.stroke_colour), self, "Choose marker stroke colour")
        if colour.isValid():
            self.stroke_colour = colour.name()
            self.no_stroke = False
            self.settings.setValue("plot/stroke_colour", self.stroke_colour)
            self.settings.setValue("plot/no_stroke", False)
            self._update_stroke_colour_button()
            self.update_plot()

    def _update_stroke_colour_button(self):
        colour = QColor(self.stroke_colour)
        foreground = "#000000" if colour.lightness() > 140 else "#ffffff"
        self.stroke_colour_button.setStyleSheet(
            f"background:{self.stroke_colour};color:{foreground};font-weight:600;padding:5px 10px;"
        )

    def choose_plot_colour(self):
        colour = QColorDialog.getColor(QColor(self.plot_colour), self, "Choose plotted-data colour")
        if colour.isValid():
            self.plot_colour = colour.name()
            self.no_fill = False
            self.settings.setValue("plot/colour", self.plot_colour)
            self.settings.setValue("plot/no_fill", False)
            self._update_colour_button()
            self.update_plot()

    def _update_colour_button(self):
        colour = QColor(self.plot_colour)
        foreground = "#000000" if colour.lightness() > 140 else "#ffffff"
        self.colour_button.setStyleSheet(
            f"background:{self.plot_colour};color:{foreground};font-weight:600;padding:5px 10px;"
        )

    def axis_minimum_changed(self, *_):
        self.x_zero_minimum = self.x_zero_checkbox.isChecked()
        self.y_zero_minimum = self.y_zero_checkbox.isChecked()
        self.settings.setValue("axes/x_zero_minimum", self.x_zero_minimum)
        self.settings.setValue("axes/y_zero_minimum", self.y_zero_minimum)
        if hasattr(self, "axes"):
            self.update_plot()

    def _apply_zero_minimum_limits(self):
        if self.x_zero_checkbox.isChecked():
            _xmin, xmax = self.axes.get_xlim()
            self.axes.set_xlim(left=0.0, right=max(0.0, xmax))
        if self.y_zero_checkbox.isChecked():
            _ymin, ymax = self.axes.get_ylim()
            self.axes.set_ylim(bottom=0.0, top=max(0.0, ymax))

    def toggle_aspect_ratio(self):
        self.aspect_mode = "16:9" if self.aspect_mode == "Square" else "Square"
        self.aspect_button.setText(self.aspect_mode)
        self.settings.setValue("plot/aspect", self.aspect_mode)
        self._apply_figure_ratio()
        self.canvas.draw_idle()

    def _apply_figure_ratio(self):
        if self.plot_type_combo.currentText() == "Ternary":
            return
        self.axes.set_box_aspect(1.0 if self.aspect_mode == "Square" else 9.0 / 16.0)

    def _restore_cartesian_axes(self):
        """Restore ticks, labels, spines, and aspect after ternary mode."""
        self.axes.set_axis_on()
        self.axes.set_aspect("auto", adjustable="box")
        self.axes.set_box_aspect(None)
        self.axes.tick_params(
            axis="both", which="both", bottom=True, top=False,
            left=True, right=False, labelbottom=True, labelleft=True,
        )
        for spine in self.axes.spines.values():
            spine.set_visible(True)
        self.axes.grid(True)

    def _draw_placeholder(self):
        self.axes.clear()
        self._restore_cartesian_axes()
        self.axes.set_title("Select a Y-axis column")
        self.axes.set_xlabel(self.format_display_label(self.x_column or "X"))
        self.axes.set_ylabel(self.format_display_label(self.y_column or "Y"))
        self._apply_figure_ratio()
        self._apply_plot_theme()

    @staticmethod
    def _ternary_numeric(series):
        """Convert numeric and qualifier-prefixed worksheet values to floats.

        Values such as ``<0.01`` and ``>5`` are converted to their stated
        numeric magnitude. Non-numeric text remains NaN.
        """
        if pd.api.types.is_numeric_dtype(series):
            return pd.to_numeric(series, errors="coerce")
        text = series.astype("string").str.strip().str.replace(",", "", regex=False)
        extracted = text.str.extract(
            r"^[<>~=\s]*([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)",
            expand=False,
        )
        return pd.to_numeric(extracted, errors="coerce")

    def _draw_ternary_plot(self, plot_dataframe):
        import numpy as np

        if self.x_column is None or self.y_column is None or self.z_column is None:
            self.axes.clear()
            self.axes.set_axis_on()
            self.axes.set_title("Select components A, B, and C")
            self.axes.set_xticks([])
            self.axes.set_yticks([])
            for spine in self.axes.spines.values():
                spine.set_visible(False)
            self.figure.patch.set_facecolor(self.current_theme["ui_background"])
            self.canvas.draw_idle()
            return

        component_a = self._ternary_numeric(plot_dataframe[self.x_column])
        component_b = self._ternary_numeric(plot_dataframe[self.y_column])
        component_c = self._ternary_numeric(plot_dataframe[self.z_column])

        valid = component_a.notna() & component_b.notna() & component_c.notna()
        valid &= (component_a >= 0) & (component_b >= 0) & (component_c >= 0)
        totals = component_a + component_b + component_c
        valid &= totals > 0

        self.axes.clear()
        self.axes.set_axis_on()
        self.axes.set_box_aspect(None)
        colours = self.current_theme
        root_three_over_two = np.sqrt(3.0) / 2.0

        triangle_x = [0.0, 1.0, 0.5, 0.0]
        triangle_y = [0.0, 0.0, root_three_over_two, 0.0]
        self.axes.plot(
            triangle_x, triangle_y,
            color=colours["text"], linewidth=1.3, zorder=2,
        )

        for fraction in (0.2, 0.4, 0.6, 0.8):
            grid_colour = colours["grid"]
            self.axes.plot(
                [1.0 - fraction, 0.5 * (1.0 - fraction)],
                [0.0, root_three_over_two * (1.0 - fraction)],
                color=grid_colour, linewidth=0.7, alpha=0.6, zorder=1,
            )
            self.axes.plot(
                [fraction, 0.5 + 0.5 * fraction],
                [0.0, root_three_over_two * (1.0 - fraction)],
                color=grid_colour, linewidth=0.7, alpha=0.6, zorder=1,
            )
            self.axes.plot(
                [0.5 * fraction, 1.0 - 0.5 * fraction],
                [root_three_over_two * fraction, root_three_over_two * fraction],
                color=grid_colour, linewidth=0.7, alpha=0.6, zorder=1,
            )

        count = int(valid.sum())
        if count:
            a = component_a[valid].to_numpy(dtype=float)
            b = component_b[valid].to_numpy(dtype=float)
            c = component_c[valid].to_numpy(dtype=float)
            total = a + b + c
            a = a / total
            b = b / total
            c = c / total
            ternary_x = b + 0.5 * c
            ternary_y = root_three_over_two * c

            marker = MARKERS[self.marker_combo.currentText()]
            fill_colour = "none" if self.no_fill else self.plot_colour
            edge_colour = (
                "none" if self.no_stroke or self.stroke_width_spin.value() == 0
                else self.stroke_colour
            )
            ternary_scatter = self.axes.scatter(
                ternary_x, ternary_y,
                s=self.marker_size_spin.value(),
                marker=marker,
                facecolors=fill_colour,
                edgecolors=edge_colour,
                linewidths=self.stroke_width_spin.value(),
                alpha=self.opacity_spin.value() / 100.0,
                picker=True,
                pickradius=5,
                zorder=4,
            )
            row_indices = plot_dataframe.index[valid].to_numpy()
            self._artist_row_indices[ternary_scatter] = row_indices
            self._artist_series_columns[ternary_scatter] = "Ternary"
            for row_index, x_value, y_value in zip(
                row_indices, ternary_x, ternary_y
            ):
                self._plotted_row_coordinates.setdefault(row_index, []).append(
                    (self.axes, float(x_value), float(y_value))
                )

        label_a = self.format_display_label(self.x_column)
        label_b = self.format_display_label(self.y_column)
        label_c = self.format_display_label(self.z_column)
        label_kwargs = dict(
            color=self.axis_title_colour,
            fontsize=self.axis_title_size,
            fontweight="bold" if self.axis_title_bold else "normal",
            fontfamily=self.font_combo.currentText() or "Arial",
            clip_on=False,
            zorder=6,
        )
        self.axes.text(-0.025, -0.035, label_a, ha="right", va="top", **label_kwargs)
        self.axes.text(1.025, -0.035, label_b, ha="left", va="top", **label_kwargs)
        self.axes.text(
            0.5, root_three_over_two + 0.040, label_c,
            ha="center", va="bottom", **label_kwargs,
        )
        self.axes.set_title(
            f"{label_a} - {label_b} - {label_c}",
            color=self.axis_title_colour,
            fontsize=self.axis_title_size,
            fontweight="bold" if self.axis_title_bold else "normal",
            fontfamily=self.font_combo.currentText() or "Arial",
            pad=12,
        )

        # Equal data scaling guarantees an equilateral triangle. The added
        # limits provide room for corner labels without changing geometry.
        self.axes.set_xlim(-0.10, 1.10)
        self.axes.set_ylim(-0.08, root_three_over_two + 0.10)
        self.axes.set_aspect("equal", adjustable="box")
        self.axes.set_xticks([])
        self.axes.set_yticks([])
        self.axes.tick_params(
            left=False, bottom=False, labelleft=False, labelbottom=False
        )
        for spine in self.axes.spines.values():
            spine.set_visible(False)
        self._apply_custom_plot_title()
        self.axes.grid(False)
        self.figure.patch.set_facecolor(colours["ui_background"])
        self.axes.set_facecolor(colours["plot_background"])

        omitted = len(plot_dataframe) - count
        message = f"Ternary plot: {count:,} rows normalized to A + B + C = 100%"
        if omitted:
            message += (
                f"; omitted {omitted:,} rows with missing, non-numeric, "
                "negative, or zero-total components."
            )
        else:
            message += "."
        if count == 0:
            message += " No valid ternary rows are available for the selected components."
        if self.active_filters:
            message += f" Filtered from {len(self.dataframe):,} to {len(plot_dataframe):,} rows."
        selected_source_indices = self._selected_table_source_indices()
        if selected_source_indices:
            self._highlight_rows_on_plot(selected_source_indices)
            self.clear_highlight_button.setEnabled(True)
        self.status_bar.showMessage(message)
        self.canvas.draw_idle()

    def update_plot(self):
        if not hasattr(self, "axes"):
            return
        selected_source_indices = self._selected_table_source_indices() if hasattr(self, "data_table") else []
        self._artist_row_indices = {}
        self._artist_series_columns = {}
        self._plotted_row_coordinates = {}
        self._highlight_artists = []
        if self.secondary_axes is not None:
            try:
                self.secondary_axes.remove()
            except Exception:
                pass
            self.secondary_axes = None
        if self.dataframe is None:
            self._draw_placeholder()
            return
        plot_dataframe = self.filtered_dataframe if self.filtered_dataframe is not None else self.dataframe
        if self.spider_configuration is not None:
            self._draw_spider_plot(plot_dataframe)
            return
        if self.petrological_diagram is not None:
            self._draw_petrological_diagram(plot_dataframe)
            return
        if self.plot_type_combo.currentText() == "Histogram":
            self._draw_histogram(plot_dataframe)
            return
        if self.plot_type_combo.currentText() == "Ternary":
            if len(self.y_columns) > 1:
                self.y_columns = [self.y_columns[0]]
                self.y_column = self.y_columns[0]
                self.y_button_row.set_selected_multiple(self.y_columns)
            self._draw_ternary_plot(plot_dataframe)
            return
        if self.x_column is None or not self.y_columns:
            self._draw_placeholder()
            return
        self.axes.clear()
        self._restore_cartesian_axes()
        x = pd.to_numeric(plot_dataframe[self.x_column], errors="coerce")
        # Preserve the user-selected formatting colour for the first series.
        # Additional Y series use distinct palette colours so they remain identifiable.
        point_colours = [self.plot_colour, "#d62728", "#2ca02c", "#9467bd", "#ff7f0e", "#17becf", "#8c564b", "#e377c2"]
        line_colours = [self.line_colour, "#d62728", "#2ca02c", "#9467bd", "#ff7f0e", "#17becf", "#8c564b", "#e377c2"]
        right_column = self.secondary_y_column if self.secondary_y_enabled and self.secondary_y_column in self.y_columns else None
        if right_column:
            self.secondary_axes = self.axes.twinx()
        total_valid = 0
        plotted_series = 0
        regression_data = None
        for index, column in enumerate(self.y_columns):
            y = pd.to_numeric(plot_dataframe[column], errors="coerce")
            valid = x.notna() & y.notna()
            count = int(valid.sum())
            if count == 0:
                continue
            total_valid = max(total_valid, count)
            plotted_series += 1
            target_axes = self.secondary_axes if column == right_column else self.axes
            is_line_plot = self.plot_type_combo.currentText() == "Line"
            palette = line_colours if is_line_plot else point_colours
            colour = palette[index % len(palette)]
            label = self.format_display_label(column)
            if is_line_plot:
                style = self.series_line_formats.get(column, {})
                line_colour = style.get("colour", colour)
                line_width = style.get("width", self.line_width)
                line_style = style.get("style", self.line_style_name)
                line_opacity = style.get("opacity", self.line_opacity)
                show_markers = style.get("show_markers", self.line_show_markers)
                marker_name = style.get("marker_name", self.line_marker_name)
                marker_size = style.get("marker_size", self.line_marker_size)
                marker = MARKERS[marker_name] if show_markers else None
                plotted_line, = target_axes.plot(
                    x[valid], y[valid], label=label, color=line_colour,
                    linewidth=line_width,
                    linestyle=self._matplotlib_line_style(line_style),
                    alpha=line_opacity / 100.0, marker=marker,
                    markersize=marker_size if show_markers else None,
                    markerfacecolor=line_colour, markeredgecolor=line_colour,
                    picker=5,
                )
                row_indices = plot_dataframe.index[valid].to_numpy()
                self._artist_row_indices[plotted_line] = row_indices
                self._artist_series_columns[plotted_line] = column
            else:
                style = self.series_point_formats.get(column, {})
                marker_name = style.get("marker_name", self.marker_combo.currentText())
                marker_size = style.get("marker_size", self.marker_size_spin.value())
                fill_colour = style.get("fill_colour", colour)
                stroke_colour = style.get("stroke_colour", self.stroke_colour)
                opacity = style.get("opacity", self.opacity_spin.value())
                stroke_width = style.get("stroke_width", self.stroke_width_spin.value())
                no_fill = style.get("no_fill", self.no_fill)
                no_stroke = style.get("no_stroke", self.no_stroke)
                marker = MARKERS[marker_name]
                face = "none" if no_fill else fill_colour
                edge = "none" if no_stroke or stroke_width == 0 else stroke_colour
                plotted_scatter = target_axes.scatter(
                    x[valid], y[valid], label=label, s=marker_size,
                    marker=marker, facecolors=face, edgecolors=edge,
                    linewidths=stroke_width, alpha=opacity/100.0,
                    picker=True, pickradius=5,
                )
                row_indices = plot_dataframe.index[valid].to_numpy()
                self._artist_row_indices[plotted_scatter] = row_indices
                self._artist_series_columns[plotted_scatter] = column
            for row_index, x_value, y_value in zip(
                row_indices, x[valid].to_numpy(dtype=float), y[valid].to_numpy(dtype=float)
            ):
                self._plotted_row_coordinates.setdefault(row_index, []).append((target_axes, x_value, y_value))
            if regression_data is None and column != right_column:
                regression_data = (x[valid], y[valid])
        if hasattr(self, "data_legend_button"):
            self.data_legend_button.setEnabled(plotted_series > 1)
        if plotted_series == 0:
            self.axes.set_title("No paired numeric values to plot")
            self.status_bar.showMessage("No selected Y variable contains paired numeric values with the X variable.")
        else:
            y_names = ", ".join(self.format_display_label(c) for c in self.y_columns)
            self.axes.set_title(f"{y_names} versus {self.format_display_label(self.x_column)}")
            self._show_plot_count(total_valid)
        self._draw_horizontal_qc_lines()
        default_x_title = self.format_display_label(self.x_column)
        self.axes.set_xlabel(self.custom_x_axis_title or default_x_title)
        left_columns = [c for c in self.y_columns if c != right_column]
        default_left_y_title = ", ".join(self.format_display_label(c) for c in left_columns)
        self.axes.set_ylabel(self.custom_left_y_axis_title or default_left_y_title)
        if self.secondary_axes is not None:
            default_right_y_title = self.format_display_label(right_column)
            self.secondary_axes.set_ylabel(self.custom_right_y_axis_title or default_right_y_title)
            self.secondary_axes.tick_params(axis="y", labelsize=self.axis_number_size, labelcolor=self.axis_number_colour)
            self.secondary_axes.yaxis.label.set_fontsize(self.axis_title_size)
            self.secondary_axes.yaxis.label.set_fontweight("bold" if self.axis_title_bold else "normal")
            self.secondary_axes.yaxis.label.set_color(self.axis_title_colour)
            self.secondary_axes.yaxis.label.set_fontfamily(self.font_combo.currentText())
        if plotted_series > 1:
            handles, labels = self.axes.get_legend_handles_labels()
            if self.secondary_axes is not None:
                h2, l2 = self.secondary_axes.get_legend_handles_labels()
                handles += h2; labels += l2
            _legend_name, legend_location = self._data_legend_position()
            self.axes.legend(handles, labels, loc=legend_location)
        if self.regression_enabled and regression_data is not None and len(regression_data[0]) >= 2:
            self._draw_regression(*regression_data)
        self._apply_zero_minimum_limits()
        if self.secondary_axes is not None and self.y_zero_checkbox.isChecked():
            _low, high = self.secondary_axes.get_ylim()
            self.secondary_axes.set_ylim(bottom=0.0, top=max(0.0, high))
        self._apply_manual_axis_ranges()
        self._apply_figure_ratio()
        self._apply_plot_theme()
        if selected_source_indices:
            self._highlight_rows_on_plot(selected_source_indices)
            self.clear_highlight_button.setEnabled(True)
        if self.secondary_axes is not None:
            c = self.current_theme
            self.secondary_axes.set_facecolor("none")
            self.secondary_axes.tick_params(colors=self.axis_number_colour)
            self.secondary_axes.spines["right"].set_color(c["border"])
            self.secondary_axes.grid(False)
            self.canvas.draw_idle()

    def _show_plot_count(self, count):
        plot_dataframe = self.filtered_dataframe if self.filtered_dataframe is not None else self.dataframe
        omitted = len(plot_dataframe) - count
        message = f"Plotted {count:,} paired numeric rows"
        message += f"; omitted {omitted:,} non-numeric or missing rows." if omitted else "."
        if self.active_filters:
            message += f" Filtered from {len(self.dataframe):,} to {len(plot_dataframe):,} rows."
        self.status_bar.showMessage(message)

    @contextmanager
    def _clipboard_export_style(self):
        figure_face = self.figure.get_facecolor()
        axes_face = self.axes.get_facecolor()
        text_objects = list(self.figure.findobj(match=lambda artist: isinstance(artist, matplotlib.text.Text)))
        text_colours = [(text, text.get_color()) for text in text_objects]
        export_font_family = self.font_combo.currentText().strip() or "Arial"
        text_families = [(text, text.get_fontfamily()) for text in text_objects]
        spine_colours = [(spine, spine.get_edgecolor()) for spine in self.axes.spines.values()]
        tick_lines = list(self.axes.get_xticklines()) + list(self.axes.get_yticklines())
        tick_colours = [(line, line.get_color()) for line in tick_lines]
        grid_lines = list(self.axes.get_xgridlines()) + list(self.axes.get_ygridlines())
        grid_styles = [
            (line, line.get_color(), line.get_alpha(), line.get_linewidth())
            for line in grid_lines
        ]
        try:
            self.figure.patch.set_facecolor("none")
            self.figure.patch.set_alpha(0.0)
            self.axes.patch.set_facecolor("none")
            self.axes.patch.set_alpha(0.0)
            for text, _colour in text_colours:
                text.set_color("#000000")
                text.set_fontfamily(export_font_family)
            for spine, _colour in spine_colours:
                spine.set_edgecolor("#000000")
            for line, _colour in tick_colours:
                line.set_color("#000000")
            for line, _colour, _alpha, _width in grid_styles:
                line.set_color("#d0d0d0")
                line.set_alpha(0.75)
            yield
        finally:
            self.figure.patch.set_facecolor(figure_face)
            self.figure.patch.set_alpha(1.0)
            self.axes.patch.set_facecolor(axes_face)
            self.axes.patch.set_alpha(1.0)
            for text, colour in text_colours:
                text.set_color(colour)
            for text, family in text_families:
                text.set_fontfamily(family)
            for spine, colour in spine_colours:
                spine.set_edgecolor(colour)
            for line, colour in tick_colours:
                line.set_color(colour)
            for line, colour, alpha, width in grid_styles:
                line.set_color(colour)
                line.set_alpha(alpha)
                line.set_linewidth(width)
            self.canvas.draw_idle()

    def copy_plot_as_vector(self):
        if self.spider_configuration is None and (self.x_column is None or (
            self.plot_type_combo.currentText() != "Histogram" and not self.y_columns
        )):
            QMessageBox.information(self, "Nothing to copy", "Select the required plot variables first.")
            return

        try:
            self._apply_plot_font()
            svg_buffer = BytesIO()
            with self._clipboard_export_style():
                with matplotlib.rc_context({
                    "svg.fonttype": self._svg_fonttype(),
                    "text.usetex": False,
                    "font.family": self.font_combo.currentText().strip() or "DejaVu Sans",
                }):
                    self.figure.savefig(
                        svg_buffer,
                        format="svg",
                        bbox_inches="tight",
                        transparent=True,
                    )
                png_bytes = self._render_png_bytes(use_clipboard_style=False)
            svg_bytes = svg_buffer.getvalue()
            mime_data = QMimeData()
            mime_data.setData("image/svg+xml", QByteArray(svg_bytes))
            mime_data.setData("image/png", QByteArray(png_bytes))
            mime_data.setHtml(svg_bytes.decode("utf-8"))

            image = QImage.fromData(QByteArray(png_bytes), "PNG")
            if not image.isNull():
                mime_data.setImageData(image)

            QApplication.clipboard().setMimeData(mime_data)
            text_description = (
                "editable text" if self._svg_fonttype() == "none" else "text converted to curves"
            )
            self.status_bar.showMessage(
                f"Plot copied as SVG with {text_description}, plus a PNG fallback."
            )
        except Exception as exc:
            QMessageBox.critical(self, "Unable to copy vector plot", str(exc))

    def _render_png_bytes(self, use_clipboard_style=True):
        png_buffer = BytesIO()
        if use_clipboard_style:
            with self._clipboard_export_style():
                self.figure.savefig(
                    png_buffer,
                    format="png",
                    dpi=300,
                    bbox_inches="tight",
                    transparent=True,
                )
        else:
            self.figure.savefig(
                png_buffer,
                format="png",
                dpi=300,
                bbox_inches="tight",
                transparent=True,
            )
        return png_buffer.getvalue()

    def copy_plot_as_png(self):
        if self.spider_configuration is None and (self.x_column is None or (
            self.plot_type_combo.currentText() != "Histogram" and not self.y_columns
        )):
            QMessageBox.information(self, "Nothing to copy", "Select the required plot variables first.")
            return
        try:
            png_bytes = self._render_png_bytes()
            mime_data = QMimeData()
            mime_data.setData("image/png", QByteArray(png_bytes))
            image = QImage.fromData(QByteArray(png_bytes), "PNG")
            if not image.isNull():
                mime_data.setImageData(image)
            QApplication.clipboard().setMimeData(mime_data)
            self.status_bar.showMessage("Plot copied as a 300 dpi PNG using a light export style and transparent background.")
        except Exception as exc:
            QMessageBox.critical(self, "Unable to copy PNG plot", str(exc))

    def _plot_available_for_output(self):
        return not (
            self.spider_configuration is None
            and (
                self.x_column is None
                or (
                    self.plot_type_combo.currentText() != "Histogram"
                    and not self.y_columns
                )
            )
        )

    def _configured_printer(self):
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setDocName("Geochemistry Data Plot")
        printer.setResolution(300)
        try:
            width, height = self.figure.get_size_inches()
            orientation = (
                QPageLayout.Orientation.Landscape
                if width >= height else QPageLayout.Orientation.Portrait
            )
            printer.setPageOrientation(orientation)
        except Exception:
            pass
        return printer

    def _paint_plot_to_printer(self, printer):
        png_bytes = self._render_png_bytes()
        image = QImage.fromData(QByteArray(png_bytes), "PNG")
        if image.isNull():
            raise RuntimeError("The plot could not be rendered for printing.")

        painter = QPainter()
        if not painter.begin(printer):
            raise RuntimeError("The selected printer could not be opened.")
        try:
            page = printer.pageLayout().paintRectPixels(printer.resolution())
            page_rect = QRectF(page)
            image_size = image.size()
            maximum_scale = min(
                page_rect.width() / image_size.width(),
                page_rect.height() / image_size.height(),
            )
            # Use 65% of the maximum size that fits in the printable area.
            # The target rectangle below keeps the resized plot centred both
            # horizontally and vertically while preserving its aspect ratio.
            scale = maximum_scale * 0.65
            target_width = image_size.width() * scale
            target_height = image_size.height() * scale
            target = QRectF(
                page_rect.left() + (page_rect.width() - target_width) / 2.0,
                page_rect.top() + (page_rect.height() - target_height) / 2.0,
                target_width,
                target_height,
            )
            painter.drawImage(target, image)
        finally:
            painter.end()

    def print_preview(self):
        if not self._plot_available_for_output():
            QMessageBox.information(
                self, "Nothing to print", "Select the required plot variables first."
            )
            return
        try:
            printer = self._configured_printer()
            preview = QPrintPreviewDialog(printer, self)
            preview.setWindowTitle("Print Preview - Geochemistry Data Plotter")
            preview.resize(1100, 800)
            preview.paintRequested.connect(self._paint_plot_to_printer)
            preview.exec()
        except Exception as exc:
            QMessageBox.critical(self, "Unable to preview plot", str(exc))

    def print_plot(self):
        if not self._plot_available_for_output():
            QMessageBox.information(
                self, "Nothing to print", "Select the required plot variables first."
            )
            return
        try:
            printer = self._configured_printer()
            dialog = QPrintDialog(printer, self)
            dialog.setWindowTitle("Print Geochemistry Plot")
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self._paint_plot_to_printer(printer)
                self.status_bar.showMessage("Plot sent to the selected printer.")
        except Exception as exc:
            QMessageBox.critical(self, "Unable to print plot", str(exc))

    def save_plot(self):
        if self.spider_configuration is None and (self.x_column is None or (
            self.plot_type_combo.currentText() != "Histogram" and not self.y_columns
        )):
            QMessageBox.information(self, "Nothing to save", "Select the required plot variables first.")
            return
        safe_x = re.sub(r"[^A-Za-z0-9._-]+", "_", str(self.x_column)).strip("_") or "X"
        if self.spider_configuration is not None:
            default_name = "Petrological_spider_plot.png"
        elif self.plot_type_combo.currentText() == "Histogram":
            default_name = f"Histogram_{safe_x}.png"
        else:
            y_name = "_and_".join(map(str, self.y_columns))
            safe_y = re.sub(r"[^A-Za-z0-9._-]+", "_", y_name).strip("_") or "Y"
            default_name = f"{safe_y}_vs_{safe_x}.png"
        start = str(Path(self.settings.value("files/last_directory", "")) / default_name)
        path, selected_filter = QFileDialog.getSaveFileName(
            self, "Save plot", start,
            "PNG image (*.png);;PDF document (*.pdf);;SVG image (*.svg)"
        )
        if not path:
            return
        output = Path(path)
        if not output.suffix:
            output = output.with_suffix(".pdf" if "PDF" in selected_filter else
                                        ".svg" if "SVG" in selected_filter else ".png")
        try:
            self._apply_plot_font()
            if output.suffix.lower() == ".svg":
                with matplotlib.rc_context({
                    "svg.fonttype": self._svg_fonttype(),
                    "text.usetex": False,
                    "font.family": self.font_combo.currentText().strip() or "DejaVu Sans",
                }):
                    self.figure.savefig(
                        output,
                        bbox_inches="tight",
                        facecolor=self.figure.get_facecolor(),
                    )
            else:
                self.figure.savefig(
                    output,
                    dpi=300,
                    bbox_inches="tight",
                    facecolor=self.figure.get_facecolor(),
                )
            self.settings.setValue("files/last_directory", str(output.parent))
            self.status_bar.showMessage(f"Saved plot to {output}")
        except Exception as exc:
            QMessageBox.critical(self, "Unable to save plot", str(exc))


def main():
    app = QApplication(sys.argv)
    app.setOrganizationName("ScientificPlotTools")
    app.setApplicationName("ExcelBivariatePlotterPySide6")
    window = ExcelBivariatePlotter()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
