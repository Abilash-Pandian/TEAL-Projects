import sys
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QComboBox, QPushButton, QTableView, QHeaderView, QLabel,
                             QLineEdit, QDialog, QFormLayout, QMessageBox)
from PyQt6.QtCore import Qt, QAbstractTableModel


class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return str(self._data.columns[section])
            if orientation == Qt.Orientation.Vertical:
                return str(self._data.index[section])
        return None


class FilterDialog(QDialog):
    def __init__(self, parent=None, columns=None):
        super().__init__(parent)
        self.setWindowTitle("Additional Filters")
        self.setMinimumWidth(300)
        layout = QFormLayout(self)
        self.filters = {}
        for col in columns:
            self.filters[col] = QLineEdit(self)
            layout.addRow(QLabel(col), self.filters[col])
        self.button_box = QPushButton("Apply Filters", self)
        self.button_box.clicked.connect(self.accept)
        layout.addWidget(self.button_box)


class ToolEditorDialog(QDialog):
    def __init__(self, parent=None, columns=None, tool=None):
        super().__init__(parent)
        self.setWindowTitle("Add/Edit Tool")
        self.setMinimumWidth(300)
        layout = QFormLayout(self)
        self.fields = {}
        self.is_edit = tool is not None
        if self.is_edit:
            self.tool = tool
        else:
            self.tool = {}

        for col in columns:
            self.fields[col] = QLineEdit(self)
            if self.is_edit and col in self.tool:
                self.fields[col].setText(str(self.tool[col]))
            layout.addRow(QLabel(col), self.fields[col])

        self.button_box = QPushButton("Save", self)
        self.button_box.clicked.connect(self.accept)
        layout.addWidget(self.button_box)

    def get_data(self):
        return {col: self.fields[col].text() for col in self.fields}


class ToolFinderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Modern Tool Finder")
        self.setGeometry(100, 100, 1200, 800)
        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        self.setStyleSheet("""
            QMainWindow {background-color: #f0f0f0;}
            QLabel {font-size: 14px; color: #333;}
            QComboBox {
                font-size: 14px; 
                padding: 5px;
                border: 1px solid #ccc;
                border-radius: 3px;
            }
            QPushButton {
                font-size: 14px;
                padding: 8px 15px;
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover {background-color: #45a049;}
            QTableView {
                font-size: 12px;
                border: 1px solid #ddd;
                gridline-color: #ddd;
            }
        """)

        matl_group_layout = QHBoxLayout()
        matl_group_label = QLabel("Select Matl Group:")
        self.matl_group_combo = QComboBox()
        self.matl_group_combo.addItems(
            ["ELECTRODE", "TOOL(PUR)", "TOOLS SPL", "TOOLFIX"])
        matl_group_layout.addWidget(matl_group_label)
        matl_group_layout.addWidget(self.matl_group_combo)
        main_layout.addLayout(matl_group_layout)

        tool_layout = QHBoxLayout()
        tool_label = QLabel("Select Tool:")
        self.tool_combo = QComboBox()
        self.tool_combo.addItems([
            "Drills carbide", "Drills HSS", "Special Centre Drills", "Insert type Drills",
            "Drills", "FL Drill", "Countersink Drills", "GUN Drill", "Special Drill",
            "SPAD Drill", "Ball End Mills", "Flat End Mill", "Insert Type Cutter",
            "Toric End Mill", "Taper End Mill", "Slitting Wheel (Carbide)",
            "Chamfer Cutter (Carbide)", "Slitting Wheel (HSS)", "Special End Mill",
            "Batch End Mill", "T Slot Cutters", "Form Cutters", "Thread Form Cutters",
            "Thread Mill Relief Cutters", "Hand Taps - Metric Coarse & Fine Series",
            "Hand Taps UNF Series", "Hand Taps UNJF", "Hand Taps Metric Fine Series",
            "Machine Taps Metric J Series", "Machine Taps UNJC", "UNJF", "UNJEF",
            "Machine Taps UNF & NF Series", "Machine Taps UNEF Series", "Machine Taps UNC Series",
            "Machine Taps UNS Type", "Helicoil Taps Metric & UN Series", "Helicoil Inserts Metric & UN",
            "Spiral Lock Taps", "HeliCoil Machine Taps", "HeliCoil Extracting tools",
            "Helicoil Hand Insert Tool (TANG)", "Helicoil Hand Install w/o TANG",
            "Rep. Insert Blade Kit", "Rep. Punch Tool", "TANG Removal Tool", "Tapping Adaptor",
            "Roll Taps", "Reamers", "Fixtures"
        ])
        tool_layout.addWidget(tool_label)
        tool_layout.addWidget(self.tool_combo)
        main_layout.addLayout(tool_layout)

        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.search_tools)
        main_layout.addWidget(self.search_button)

        self.table_view = QTableView()
        self.table_view.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents)
        self.table_view.horizontalHeader().setStretchLastSection(True)
        self.table_view.setHorizontalScrollMode(
            QTableView.ScrollMode.ScrollPerPixel)
        main_layout.addWidget(self.table_view)

        self.add_button = QPushButton("Add Tool")
        self.add_button.clicked.connect(self.add_tool)
        main_layout.addWidget(self.add_button)

        self.delete_button = QPushButton("Delete Tool")
        self.delete_button.clicked.connect(self.delete_tool)
        main_layout.addWidget(self.delete_button)

        self.update_button = QPushButton("Update Tool")
        self.update_button.clicked.connect(self.update_tool)
        main_layout.addWidget(self.update_button)

    def load_data(self):
        try:
            file_path = "C:\\Users\\Abilash Pandian\\Downloads\\tools.xlsx"
            self.sheets = pd.read_excel(file_path, sheet_name=None)
            self.dataframes = {sheet_name: pd.DataFrame(
                data) for sheet_name, data in self.sheets.items()}
        except Exception as e:
            print(f"Error loading data: {e}")

    def search_tools(self):
        matl_group = self.matl_group_combo.currentText()
        tool = self.tool_combo.currentText()
        df = self.dataframes.get(tool, None)

        if df is not None:
            df_filtered = df.copy()
            columns_to_display = df.columns.tolist()

            if matl_group in df.columns:
                df_filtered = df_filtered[df_filtered[matl_group].notna()]

            dialog = FilterDialog(columns=columns_to_display)
            if dialog.exec():
                for col, line_edit in dialog.filters.items():
                    if line_edit.text():
                        search_value = line_edit.text()
                        try:
                            search_value = float(search_value)
                            tolerance = 0.1
                            df_filtered = df_filtered[
                                df_filtered[col].apply(
                                    lambda x: abs(float(x) - search_value) <= tolerance if isinstance(
                                        x, (int, float)) else str(search_value) in str(x)
                                )
                            ]
                        except ValueError:

                            df_filtered = df_filtered[
                                df_filtered[col].apply(
                                    lambda x: str(search_value) in str(x)
                                )
                            ]

                self.model = PandasModel(df_filtered)
                self.table_view.setModel(self.model)

    def add_tool(self):
        columns = self.dataframes[self.tool_combo.currentText()].columns
        dialog = ToolEditorDialog(columns=columns)
        if dialog.exec():
            new_tool_data = dialog.get_data()
            df = self.dataframes[self.tool_combo.currentText()]
            df = df.append(new_tool_data, ignore_index=True)
            self.dataframes[self.tool_combo.currentText()] = df
            self.save_data()

    def delete_tool(self):
        if not self.table_view.selectionModel().selectedRows():
            QMessageBox.warning(self, "No Selection",
                                "Please select a row to delete.")
            return

        index = self.table_view.selectionModel().selectedRows()[0].row()
        df = self.dataframes[self.tool_combo.currentText()]
        df = df.drop(df.index[index])
        self.dataframes[self.tool_combo.currentText()] = df
        self.save_data()

    def update_tool(self):
        if not self.table_view.selectionModel().selectedRows():
            QMessageBox.warning(self, "No Selection",
                                "Please select a row to update.")
            return

        index = self.table_view.selectionModel().selectedRows()[0].row()
        columns = self.dataframes[self.tool_combo.currentText()].columns
        tool_data = self.dataframes[self.tool_combo.currentText()].iloc[index]

        dialog = ToolEditorDialog(columns=columns, tool=tool_data)
        if dialog.exec():
            updated_data = dialog.get_data()
            df = self.dataframes[self.tool_combo.currentText()]
            df.iloc[index] = updated_data
            self.dataframes[self.tool_combo.currentText()] = df
            self.save_data()

    def save_data(self):
        try:
            with pd.ExcelWriter("C:\\Users\\Abilash Pandian\\Downloads\\tools.xlsx") as writer:
                for sheet_name, df in self.dataframes.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        except Exception as e:
            print(f"Error saving data: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ToolFinderApp()
    window.show()
    sys.exit(app.exec())
