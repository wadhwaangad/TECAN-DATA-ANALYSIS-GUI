import sys
import os
from PyQt6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QWidget, QPushButton, QTableWidget, QTableWidgetItem, 
                             QFileDialog, QLabel, QComboBox, QListWidget, QSplitter,
                             QTabWidget, QInputDialog, QMessageBox, QCheckBox,
                             QSpinBox, QGroupBox, QTextEdit, QScrollArea)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QColor, QFont
import pandas as pd
import numpy as np
from pathlib import Path

class DrugAssignmentDialog(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Assign Drug/Cuboids")
        self.setModal(True)
        self.resize(400, 300)
        
        layout = QVBoxLayout()
        
        # Drug name input
        drug_layout = QHBoxLayout()
        drug_layout.addWidget(QLabel("Drug Name:"))
        self.drug_input = QInputDialog()
        
        # Cuboid count input
        cuboid_layout = QHBoxLayout()
        cuboid_layout.addWidget(QLabel("Number of Cuboids:"))
        self.cuboid_spin = QSpinBox()
        self.cuboid_spin.setMinimum(1)
        self.cuboid_spin.setMaximum(1000)
        self.cuboid_spin.setValue(1)
        cuboid_layout.addWidget(self.cuboid_spin)
        
        # Background checkbox
        self.background_checkbox = QCheckBox("Mark as Background")
        
        # Buttons
        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("OK")
        self.cancel_button = QPushButton("Cancel")
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        
        layout.addLayout(cuboid_layout)
        layout.addWidget(self.background_checkbox)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)

class SelectableTableWidget(QTableWidget):
    assignment_requested = pyqtSignal(list, str, int, bool)  # cells, drug_name, cuboid_count, is_background
    
    def __init__(self):
        super().__init__()
        self.setSelectionMode(QTableWidget.SelectionMode.MultiSelection)
        self.selected_ranges = []
        self.cell_assignments = {}  # {(row, col): {'drug': str, 'cuboids': int, 'is_background': bool}}
        self.removed_cells = set()  # Track cells that have been removed (set to NaN)
        self._saved_selection = set()  # Store selected cells as (row, col)
        
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.RightButton:
            self.show_context_menu()
        else:
            # Check if clicking on empty space to deselect
            item = self.itemAt(event.pos())
            if item is None:
                self.clearSelection()
            super().mousePressEvent(event)
    
    def mouseReleaseEvent(self, event):
        super().mouseReleaseEvent(event)
        self.save_selection()
    
    def save_selection(self):
        self._saved_selection = set((item.row(), item.column()) for item in self.selectedItems())
    
    def restore_selection(self):
        self.clearSelection()
        for row, col in self._saved_selection:
            item = self.item(row, col)
            if item:
                item.setSelected(True)
    
    def show_context_menu(self):
        selected_cells = [(item.row(), item.column()) for item in self.selectedItems()]
        if not selected_cells:
            QMessageBox.warning(self, "Warning", "Please select cells first")
            return
        from PyQt6.QtWidgets import QMenu
        menu = QMenu(self)
        assign_drug_action = menu.addAction("Assign Drug")
        assign_cuboid_action = menu.addAction("Assign Cuboid Count")
        background_action = menu.addAction("Mark as Background")
        remove_action = menu.addAction("Remove Cell (Set to NaN)")
        restore_action = menu.addAction("Restore Cell")
        menu.addSeparator()
        clear_assignment_action = menu.addAction("Clear Assignment")
        has_removed_cells = any((row, col) in self.removed_cells for row, col in selected_cells)
        has_assigned_cells = any((row, col) in self.cell_assignments for row, col in selected_cells)
        restore_action.setEnabled(has_removed_cells)
        clear_assignment_action.setEnabled(has_assigned_cells)
        action = menu.exec(self.mapToGlobal(self.viewport().mapFromGlobal(self.cursor().pos())))
        if action == assign_drug_action:
            self.show_assign_drug_dialog(selected_cells)
        elif action == assign_cuboid_action:
            self.show_assign_cuboid_dialog(selected_cells)
        elif action == background_action:
            self.assign_background(selected_cells)
        elif action == remove_action:
            self.remove_cells(selected_cells)
        elif action == restore_action:
            self.restore_cells(selected_cells)
        elif action == clear_assignment_action:
            self.clear_cell_assignments(selected_cells)

    def show_assign_drug_dialog(self, selected_cells):
        removed_selected = [cell for cell in selected_cells if cell in self.removed_cells]
        if removed_selected:
            QMessageBox.warning(self, "Warning", f"Cannot assign drugs to removed cells. Restore them first.\nRemoved cells: {removed_selected}")
            return
        drug_name, ok = QInputDialog.getText(self, "Drug Assignment", "Enter drug name:")
        if not ok or not drug_name:
            return
        self.assign_cells(selected_cells, drug_name, None, False, assign_type='drug')

    def show_assign_cuboid_dialog(self, selected_cells):
        removed_selected = [cell for cell in selected_cells if cell in self.removed_cells]
        if removed_selected:
            QMessageBox.warning(self, "Warning", f"Cannot assign cuboid count to removed cells. Restore them first.\nRemoved cells: {removed_selected}")
            return
        cuboid_count, ok = QInputDialog.getInt(self, "Cuboid Count", "Number of cuboids:", 1, 1, 1000)
        if not ok:
            return
        self.assign_cells(selected_cells, None, cuboid_count, False, assign_type='cuboid')

    def get_drug_color(self, drug_name):
        # Assign a unique color for each drug, ensuring background color is not used
        if not hasattr(self, '_drug_colors'):
            self._drug_colors = {}
            self._drug_color_list = [
                QColor(255, 200, 200),  # Light red
                QColor(200, 255, 200),  # Light green
                QColor(255, 255, 200),  # Light yellow
                QColor(255, 200, 255),  # Light magenta
                QColor(200, 255, 255),  # Light cyan
                QColor(255, 220, 180),  # Light orange
                QColor(220, 180, 255),  # Light purple
                QColor(255, 180, 180),  # Another light red variant
            ]
            # Note: Removed QColor(200, 200, 255) as it's reserved for background
            self._drug_color_idx = 0
        
        if drug_name and drug_name not in self._drug_colors:
            color = self._drug_color_list[self._drug_color_idx % len(self._drug_color_list)]
            self._drug_colors[drug_name] = color
            self._drug_color_idx += 1
        
        return self._drug_colors.get(drug_name, QColor(255, 255, 255))  # Default to white

    def get_cuboid_color(self, cuboid_count):
        # Assign a unique color for each cuboid count
        if not hasattr(self, '_cuboid_colors'):
            self._cuboid_colors = {}
            self._cuboid_color_list = [QColor(200, 200, 255), QColor(255, 220, 180), QColor(180, 255, 220), QColor(220, 180, 255), QColor(180, 220, 255), QColor(255, 255, 180), QColor(220, 255, 180), QColor(255, 180, 220)]
            self._cuboid_color_idx = 0
        if cuboid_count not in self._cuboid_colors:
            color = self._cuboid_color_list[self._cuboid_color_idx % len(self._cuboid_color_list)]
            self._cuboid_colors[cuboid_count] = color
            self._cuboid_color_idx += 1
        return self._cuboid_colors[cuboid_count]

    def assign_cells(self, cells, drug_name, cuboid_count, is_background, assign_type=None):
        # Find the parent GUI and determine which sheet this table belongs to
        parent_gui = self.parent()
        while parent_gui and not hasattr(parent_gui, 'table_widgets'):
            parent_gui = parent_gui.parent() if hasattr(parent_gui, 'parent') else None
        
        sheet_name = None
        if parent_gui:
            for sname, file_widgets in parent_gui.table_widgets.items():
                if self in file_widgets.values():
                    sheet_name = sname
                    break
        
        # Get all table widgets for this sheet name (across all files)
        widgets = [self]  # Always include current widget
        if parent_gui and sheet_name and sheet_name in parent_gui.table_widgets:
            widgets = list(parent_gui.table_widgets[sheet_name].values())
        
        # Apply changes to all widgets for this sheet name
        for widget in widgets:
            for row, col in cells:
                # Ensure proper initialization of cell_assignments
                if (row, col) not in widget.cell_assignments:
                    widget.cell_assignments[(row, col)] = {
                        'drug': None, 
                        'cuboids': None, 
                        'is_background': False, 
                        'original_value': None
                    }
                
                # Update assignments based on type
                assignment = widget.cell_assignments[(row, col)]
                if assign_type == 'drug':
                    assignment['drug'] = drug_name
                elif assign_type == 'cuboid':
                    assignment['cuboids'] = cuboid_count
                else:
                    # Full assignment
                    assignment['drug'] = drug_name
                    assignment['cuboids'] = cuboid_count
                    assignment['is_background'] = is_background
                
                item = widget.item(row, col)
                if item:
                    # Save original value if not already saved
                    if assignment['original_value'] is None:
                        try:
                            assignment['original_value'] = float(item.text())
                        except ValueError:
                            assignment['original_value'] = 0.0
                    
                    # Remove any dot prefix and reset text/foreground
                    orig_text = item.text()
                    if orig_text.startswith('\u25CF '):
                        orig_text = orig_text[2:]
                    elif orig_text.startswith('● '):
                        orig_text = orig_text[2:]
                    item.setText(orig_text)
                    item.setForeground(QColor(0, 0, 0))
                    
                    # Set background color based on assignment
                    if assignment['is_background']:
                        item.setBackground(QColor(200, 200, 255))  # Light blue for background
                    elif assignment['drug']:
                        drug_color = widget.get_drug_color(assignment['drug'])
                        item.setBackground(drug_color)
                    else:
                        item.setBackground(QColor(255, 255, 255))  # White for unassigned
                    
                    # Set cuboid border if cuboids assigned
                    if assignment['cuboids']:
                        cuboid_color = widget.get_cuboid_color(assignment['cuboids'])
                        item.setData(Qt.ItemDataRole.UserRole + 1, cuboid_color.name())
                    else:
                        item.setData(Qt.ItemDataRole.UserRole + 1, None)
                    
                    # Update tooltip
                    if assignment['is_background']:
                        item.setToolTip("Background cell")
                    else:
                        tooltip_parts = []
                        if assignment['drug']:
                            tooltip_parts.append(f"Drug: {assignment['drug']}")
                        if assignment['cuboids']:
                            tooltip_parts.append(f"Cuboids: {assignment['cuboids']}")
                        tooltip_parts.append(f"Background: {assignment['is_background']}")
                        item.setToolTip("\n".join(tooltip_parts))
        
        if parent_gui and hasattr(parent_gui, 'update_legend'):
            parent_gui.update_legend()

    def paintEvent(self, event):
        # Custom paint to draw cuboid border and selection border
        super().paintEvent(event)
        from PyQt6.QtWidgets import QStyleOptionViewItem, QStyle
        from PyQt6.QtGui import QPainter, QPen
        painter = QPainter(self.viewport())
        for i in range(self.rowCount()):
            for j in range(self.columnCount()):
                item = self.item(i, j)
                if not item:
                    continue
                rect = self.visualItemRect(item)
                # Draw cuboid border if assigned and not selected
                cuboid_color_name = item.data(Qt.ItemDataRole.UserRole + 1)
                if cuboid_color_name and not item.isSelected():
                    pen = QPen(QColor(cuboid_color_name), 3)
                    painter.setPen(pen)
                    painter.drawRect(rect.adjusted(1, 1, -2, -2))
                # Draw selection border (blue, overrides cuboid border)
                if item.isSelected():
                    pen = QPen(QColor(0, 120, 215), 3)  # Windows blue
                    painter.setPen(pen)
                    painter.drawRect(rect.adjusted(0, 0, -1, -1))
        painter.end()

    def remove_cells(self, selected_cells):
        # Find the parent GUI and determine which sheet this table belongs to
        parent_gui = self.parent()
        while parent_gui and not hasattr(parent_gui, 'table_widgets'):
            parent_gui = parent_gui.parent() if hasattr(parent_gui, 'parent') else None
        
        sheet_name = None
        if parent_gui:
            for sname, file_widgets in parent_gui.table_widgets.items():
                if self in file_widgets.values():
                    sheet_name = sname
                    break
        
        # Get all table widgets for this sheet name (across all files)
        widgets = [self]  # Always include current widget
        if parent_gui and sheet_name and sheet_name in parent_gui.table_widgets:
            widgets = list(parent_gui.table_widgets[sheet_name].values())
        
        # Apply cell removal to all widgets for this sheet name
        for widget in widgets:
            for row, col in selected_cells:
                widget.removed_cells.add((row, col))
                item = widget.item(row, col)
                if item:
                    # Save original value before removing
                    if (row, col) not in widget.cell_assignments:
                        widget.cell_assignments[(row, col)] = {
                            'drug': None, 
                            'cuboids': None, 
                            'is_background': False, 
                            'original_value': None
                        }
                    if widget.cell_assignments[(row, col)]['original_value'] is None:
                        try:
                            widget.cell_assignments[(row, col)]['original_value'] = float(item.text())
                        except ValueError:
                            widget.cell_assignments[(row, col)]['original_value'] = 0.0
                    
                    # Set cell to show as removed
                    item.setText("NaN")
                    item.setBackground(QColor(220, 220, 220))  # Gray background
                    item.setForeground(QColor(120, 120, 120))  # Gray text
                    item.setToolTip("Removed cell (NaN)")
                    item.setData(Qt.ItemDataRole.UserRole + 1, None)  # Clear cuboid border

    def restore_cells(self, selected_cells):
        # Find the parent GUI and determine which sheet this table belongs to
        parent_gui = self.parent()
        while parent_gui and not hasattr(parent_gui, 'table_widgets'):
            parent_gui = parent_gui.parent() if hasattr(parent_gui, 'parent') else None
        
        sheet_name = None
        if parent_gui:
            for sname, file_widgets in parent_gui.table_widgets.items():
                if self in file_widgets.values():
                    sheet_name = sname
                    break
        
        # Get all table widgets for this sheet name (across all files)
        widgets = [self]  # Always include current widget
        if parent_gui and sheet_name and sheet_name in parent_gui.table_widgets:
            widgets = list(parent_gui.table_widgets[sheet_name].values())
        
        # Apply cell restoration to all widgets for this sheet name
        for widget in widgets:
            for row, col in selected_cells:
                if (row, col) in widget.removed_cells:
                    widget.removed_cells.remove((row, col))
                    item = widget.item(row, col)
                    if item and (row, col) in widget.cell_assignments:
                        assignment = widget.cell_assignments[(row, col)]
                        # Restore original value
                        if assignment['original_value'] is not None:
                            item.setText(str(assignment['original_value']))
                        else:
                            item.setText("0")
                        
                        # Restore proper formatting based on assignment
                        if assignment['is_background']:
                            item.setBackground(QColor(200, 200, 255))
                            item.setToolTip("Background cell")
                        elif assignment['drug']:
                            drug_color = widget.get_drug_color(assignment['drug'])
                            item.setBackground(drug_color)
                            tooltip_parts = []
                            if assignment['drug']:
                                tooltip_parts.append(f"Drug: {assignment['drug']}")
                            if assignment['cuboids']:
                                tooltip_parts.append(f"Cuboids: {assignment['cuboids']}")
                            tooltip_parts.append(f"Background: {assignment['is_background']}")
                            item.setToolTip("\n".join(tooltip_parts))
                        else:
                            item.setBackground(QColor(255, 255, 255))
                            item.setToolTip("")
                        
                        item.setForeground(QColor(0, 0, 0))
                        
                        # Restore cuboid border if applicable
                        if assignment['cuboids']:
                            cuboid_color = widget.get_cuboid_color(assignment['cuboids'])
                            item.setData(Qt.ItemDataRole.UserRole + 1, cuboid_color.name())

    def clear_cell_assignments(self, selected_cells):
        # Find the parent GUI and determine which sheet this table belongs to
        parent_gui = self.parent()
        while parent_gui and not hasattr(parent_gui, 'table_widgets'):
            parent_gui = parent_gui.parent() if hasattr(parent_gui, 'parent') else None
        
        sheet_name = None
        if parent_gui:
            for sname, file_widgets in parent_gui.table_widgets.items():
                if self in file_widgets.values():
                    sheet_name = sname
                    break
        
        # Get all table widgets for this sheet name (across all files)
        widgets = [self]  # Always include current widget
        if parent_gui and sheet_name and sheet_name in parent_gui.table_widgets:
            widgets = list(parent_gui.table_widgets[sheet_name].values())
        
        # Clear assignments for selected cells across all widgets for this sheet name
        for widget in widgets:
            for row, col in selected_cells:
                if (row, col) in widget.cell_assignments:
                    assignment = widget.cell_assignments[(row, col)]
                    # Reset assignment but keep original value
                    assignment['drug'] = None
                    assignment['cuboids'] = None
                    assignment['is_background'] = False
                    
                    item = widget.item(row, col)
                    if item:
                        # Reset visual appearance
                        item.setBackground(QColor(255, 255, 255))  # White background
                        item.setForeground(QColor(0, 0, 0))  # Black text
                        item.setToolTip("")  # Clear tooltip
                        item.setData(Qt.ItemDataRole.UserRole + 1, None)  # Clear cuboid border
        
        # Update legend after clearing assignments
        if parent_gui and hasattr(parent_gui, 'update_legend'):
            parent_gui.update_legend()

    def assign_background(self, selected_cells):
        # Prevent marking removed cells as background
        removed_selected = [cell for cell in selected_cells if cell in self.removed_cells]
        if removed_selected:
            QMessageBox.warning(self, "Warning", f"Cannot mark removed cells as background. Restore them first.\nRemoved cells: {removed_selected}")
            return
        
        # Find the parent GUI and determine which sheet this table belongs to
        parent_gui = self.parent()
        while parent_gui and not hasattr(parent_gui, 'table_widgets'):
            parent_gui = parent_gui.parent() if hasattr(parent_gui, 'parent') else None
        
        sheet_name = None
        if parent_gui:
            for sname, file_widgets in parent_gui.table_widgets.items():
                if self in file_widgets.values():
                    sheet_name = sname
                    break
        
        # Get all table widgets for this sheet name (across all files)
        widgets = [self]  # Always include current widget
        if parent_gui and sheet_name and sheet_name in parent_gui.table_widgets:
            widgets = list(parent_gui.table_widgets[sheet_name].values())
        
        # Apply background assignment to all widgets for this sheet name
        for widget in widgets:
            for row, col in selected_cells:
                if (row, col) not in widget.cell_assignments:
                    widget.cell_assignments[(row, col)] = {'drug': None, 'cuboids': None, 'is_background': True, 'original_value': None}
                else:
                    widget.cell_assignments[(row, col)]['is_background'] = True
                
                item = widget.item(row, col)
                if item:
                    # Save original value if not already saved
                    if widget.cell_assignments[(row, col)]['original_value'] is None:
                        try:
                            widget.cell_assignments[(row, col)]['original_value'] = float(item.text())
                        except ValueError:
                            widget.cell_assignments[(row, col)]['original_value'] = 0.0
                    
                    # Set background color to a distinct color for background
                    item.setBackground(QColor(200, 200, 255))
                    
                    # Remove any dot prefix and reset text/foreground
                    orig_text = item.text()
                    if orig_text.startswith('\u25CF '):
                        orig_text = orig_text[2:]
                    elif orig_text.startswith('● '):
                        orig_text = orig_text[2:]
                    item.setText(orig_text)
                    item.setForeground(QColor(0, 0, 0))
                    
                    # Remove cuboid border
                    item.setData(Qt.ItemDataRole.UserRole + 1, None)
                    
                    # Tooltip
                    item.setToolTip("Background cell")
        
        # Update legend if parent GUI exists
        if parent_gui and hasattr(parent_gui, 'update_legend'):
            parent_gui.update_legend()

class ExcelAnalyzerGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.excel_files = []
        self.sheet_data = {}  # {sheet_name: {file_path: dataframe}}
        self.current_sheet = None
        self.table_widgets = {}  # {sheet_name: SelectableTableWidget}
        self.selections = {}  # {(sheet_name, file_name): set((row, col))}
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("Excel Drug Analysis Tool")
        self.setGeometry(100, 100, 1400, 800)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # File selection section
        file_section = self.create_file_section()
        main_layout.addWidget(file_section)
        
        # Sheet selection and table display
        content_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left panel - sheet selection and controls
        left_panel = self.create_left_panel()
        content_splitter.addWidget(left_panel)
        
        # Right panel - table display
        self.tab_widget = QTabWidget()
        content_splitter.addWidget(self.tab_widget)
        
        content_splitter.setSizes([300, 1100])
        main_layout.addWidget(content_splitter)
        
        # Analysis section
        analysis_section = self.create_analysis_section()
        main_layout.addWidget(analysis_section)
    
    def create_file_section(self):
        group = QGroupBox("File Selection")
        layout = QHBoxLayout()
        
        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(100)
        
        button_layout = QVBoxLayout()
        self.select_files_btn = QPushButton("Select Excel Files")
        self.select_files_btn.clicked.connect(self.select_files)
        
        self.load_data_btn = QPushButton("Load Data")
        self.load_data_btn.clicked.connect(self.load_data)
        self.load_data_btn.setEnabled(False)
        
        button_layout.addWidget(self.select_files_btn)
        button_layout.addWidget(self.load_data_btn)
        button_layout.addStretch()
        
        layout.addWidget(self.file_list)
        layout.addLayout(button_layout)
        
        group.setLayout(layout)
        return group
    
    def create_left_panel(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Sheet selection
        sheet_group = QGroupBox("Sheet Selection")
        sheet_layout = QVBoxLayout()
        
        self.sheet_list = QListWidget()
        self.sheet_list.itemClicked.connect(self.on_sheet_selected)
        sheet_layout.addWidget(self.sheet_list)
        
        sheet_group.setLayout(sheet_layout)
        layout.addWidget(sheet_group)
        
        # Assignment controls
        assign_group = QGroupBox("Assignment Controls")
        assign_layout = QVBoxLayout()
        
        info_label = QLabel("Right-click on selected cells for options:")
        info_label.setStyleSheet("font-weight: bold; color: #0066cc;")
        assign_layout.addWidget(info_label)
        
        options_label = QLabel("• Assign Drug\n• Remove Cell (NaN)\n• Restore Cell\n• Clear Assignment")
        options_label.setStyleSheet("font-size: 10px; color: #666;")
        assign_layout.addWidget(options_label)
        
        assign_layout.addWidget(QLabel(""))  # Spacer
        
        self.clear_assignments_btn = QPushButton("Clear All Assignments")
        self.clear_assignments_btn.clicked.connect(self.clear_assignments)
        assign_layout.addWidget(self.clear_assignments_btn)
        
        assign_group.setLayout(assign_layout)
        layout.addWidget(assign_group)
        
        # Add legend for drug colors and background
        legend_group = QGroupBox("Legend")
        legend_layout = QVBoxLayout()
        legend_label = QLabel()
        legend_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        legend_layout.addWidget(legend_label)
        legend_group.setLayout(legend_layout)
        layout.addWidget(legend_group)
        
        # Instructions
        instructions = QTextEdit()
        instructions.setMaximumHeight(200)
        instructions.setReadOnly(True)
        instructions.setText("""
Instructions:
1. Select multiple Excel files
2. Load data to see available sheets
3. Select a sheet to view its data
4. Right-click on cells to assign drugs
5. Specify drug name and cuboid count
6. Mark background cells as needed
7. Use analysis tools to process data

Selection Tips:
- Click and drag to select ranges
- Ctrl+click for individual cells
- Shift+click for ranges
- Right-click to assign drugs
        """)
        
        layout.addWidget(QLabel("Instructions:"))
        layout.addWidget(instructions)
        layout.addStretch()
        
        widget.setLayout(layout)
        self.legend_label = legend_label  # Store for later updates
        return widget
    
    def update_legend(self):
        # Collect all drugs and their colors from all tables
        drug_colors = {}
        cuboid_colors = {}
        for sheet_tables in self.table_widgets.values():
            for table_widget in sheet_tables.values():
                if hasattr(table_widget, '_drug_colors'):
                    for drug, color in table_widget._drug_colors.items():
                        drug_colors[drug] = color
                if hasattr(table_widget, '_cuboid_colors'):
                    for cuboid, color in table_widget._cuboid_colors.items():
                        cuboid_colors[cuboid] = color
        
        legend_html = "<b>Background:</b> <span style='background-color: #c8c8ff; padding: 0 8px;'>&nbsp;</span>"
        
        if drug_colors:
            legend_html += "<br><b>Drugs:</b>"
            for drug, color in sorted(drug_colors.items()):
                rgb = color.getRgb()[:3]
                legend_html += f"<br>&nbsp;&nbsp;<b>{drug}:</b> <span style='background-color: rgb{rgb}; padding: 0 8px;'>&nbsp;</span>"
        
        if cuboid_colors:
            legend_html += "<br><b>Cuboid Counts:</b>"  
            for cuboid, color in sorted(cuboid_colors.items()):
                rgb = color.getRgb()[:3]
                legend_html += f"<br>&nbsp;&nbsp;<b>{cuboid}:</b> <span style='border: 3px solid rgb{rgb}; display: inline-block; width: 16px; height: 12px; margin-left: 4px;'></span>"
        
        self.legend_label.setText(legend_html)
    
    def create_analysis_section(self):
        group = QGroupBox("Analysis Tools")
        layout = QHBoxLayout()
        self.calculate_backgrounds_btn = QPushButton("Calculate Background Subtraction")
        self.calculate_backgrounds_btn.clicked.connect(self.calculate_background_subtraction)
        self.export_results_btn = QPushButton("Export Results")
        self.export_results_btn.clicked.connect(self.export_results)
        layout.addWidget(self.calculate_backgrounds_btn)
        layout.addWidget(self.export_results_btn)
        layout.addStretch()
        group.setLayout(layout)
        return group

    def extract_conditions(self):
        # Let user select a sheet name
        sheet_names = list(self.table_widgets.keys())
        if not sheet_names:
            QMessageBox.warning(self, "Warning", "No sheets loaded.")
            return
        from PyQt6.QtWidgets import QInputDialog
        sheet_name, ok = QInputDialog.getItem(self, "Select Sheet", "Sheet name:", sheet_names, 0, False)
        if not ok or not sheet_name:
            return
        # Gather all assignments for this sheet across all files
        all_data = []
        for file_name, table_widget in self.table_widgets[sheet_name].items():
            for (row, col), assignment in table_widget.cell_assignments.items():
                all_data.append({
                    'File': file_name,
                    'Row': row + 1,
                    'Column': col + 1,
                    'Drug': assignment.get('drug'),
                    'Cuboids': assignment.get('cuboids'),
                    'Background': assignment.get('is_background')
                })
        if not all_data:
            QMessageBox.information(self, "No Data", "No assignments found for this sheet.")
            return
        df = pd.DataFrame(all_data)
        # Let user save the extracted conditions
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Extracted Conditions", f"{sheet_name}_conditions.xlsx", "Excel Files (*.xlsx)")
        if not file_path:
            return
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            QMessageBox.information(self, "Success", f"Extracted conditions saved to {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save extracted conditions: {str(e)}")
    
    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "Select Excel Files", 
            "", 
            "Excel Files (*.xlsx *.xls)"
        )
        
        if files:
            self.excel_files = files
            self.file_list.clear()
            for file in files:
                self.file_list.addItem(Path(file).name)
            
            self.load_data_btn.setEnabled(True)
    
    def load_data(self):
        if not self.excel_files:
            QMessageBox.warning(self, "Warning", "Please select Excel files first")
            return
        try:
            self.sheet_data = {}  # {(sheet_name, file_name): DataFrame}
            sheet_display_names = []
            for file_path in self.excel_files:
                file_name = Path(file_path).name
                excel_file = pd.ExcelFile(file_path)
                for sheet_name in excel_file.sheet_names:
                    if sheet_name == "Sheet1":
                        continue
                    display_name = f"{sheet_name} ({file_name})"
                    sheet_display_names.append(display_name)
                    if (sheet_name, file_name) not in self.sheet_data:
                        self.sheet_data[(sheet_name, file_name)] = None
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                        table_start_row = self.find_table_start(df)
                        if table_start_row is not None:
                            table_df = df.iloc[table_start_row:].reset_index(drop=True)
                            # Stop at first completely empty row
                            empty_row_idx = None
                            for idx, row in table_df.iterrows():
                                if all(pd.isna(val) or (isinstance(val, str) and val.strip() == "") for val in row):
                                    empty_row_idx = idx
                                    break
                            if empty_row_idx is not None:
                                table_df = table_df.iloc[:empty_row_idx]
                            # Use first row as header if it contains text
                            if len(table_df) > 0 and (table_df.iloc[0].dtype == 'object' or any(isinstance(val, str) for val in table_df.iloc[0])):
                                table_df.columns = table_df.iloc[0]
                                table_df = table_df.iloc[1:].reset_index(drop=True)
                            self.sheet_data[(sheet_name, file_name)] = table_df
                        else:
                            self.sheet_data[(sheet_name, file_name)] = df
                    except Exception as e:
                        print(f"Error loading {sheet_name} from {file_path}: {e}")
            # Populate sheet list with display names
            self.sheet_list.clear()
            for display_name in sheet_display_names:
                self.sheet_list.addItem(display_name)
            QMessageBox.information(self, "Success", f"Loaded {len(sheet_display_names)} sheets from {len(self.excel_files)} files")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load data: {str(e)}")
    
    def find_table_start(self, df):
        """Find where the actual data table starts: after the row containing 'temperature' (excluding it)"""
        for i in range(len(df)):
            row = df.iloc[i]
            # Check if any cell in the row contains 'temperature' (case-insensitive)
            if any(isinstance(val, str) and 'temperature' in val.lower() for val in row if pd.notna(val)):
                return i + 1  # Start after the 'temperature' row
        return 0  # Default to start if not found
    
    def on_sheet_selected(self, item):
        display_name = item.text()
        # Extract sheet_name and file_name from display_name
        if display_name.endswith(")") and " (" in display_name:
            sheet_name, file_name = display_name.rsplit(" (", 1)
            file_name = file_name[:-1]  # remove trailing )
            self.current_sheet = (sheet_name, file_name)
            self.display_sheet_data(sheet_name)  # Pass only the sheet_name for multi-file operations
    
    def display_sheet_data(self, sheet_name):
        # Find all (sheet_name, file_name) pairs for this sheet_name
        relevant_keys = [(s, f) for (s, f) in self.sheet_data if s == sheet_name]
        if not relevant_keys:
            return
        self.tab_widget.clear()
        for (s, file_name) in relevant_keys:
            df = self.sheet_data[(s, file_name)]
            tab_label = f"{file_name}"
            table_widget = SelectableTableWidget()
            self.populate_table(table_widget, df)
            # Restore previous selection if available
            sel_key = (sheet_name, file_name)
            if sel_key in self.selections:
                table_widget._saved_selection = self.selections[sel_key]
                table_widget.restore_selection()
            # Connect selection change to save_selection
            def save_sel_tw(widget=table_widget, key=sel_key):
                self.selections[key] = set((item.row(), item.column()) for item in widget.selectedItems())
            table_widget.itemSelectionChanged.connect(save_sel_tw)
            scroll = QScrollArea()
            scroll.setWidget(table_widget)
            scroll.setWidgetResizable(True)
            self.tab_widget.addTab(scroll, tab_label)
            if sheet_name not in self.table_widgets:
                self.table_widgets[sheet_name] = {}
            self.table_widgets[sheet_name][file_name] = table_widget
        if hasattr(self, 'legend_label'):
            self.update_legend()
    
    def populate_table(self, table_widget, df):
        # If first column is all letters, use as row labels
        if df.shape[1] > 1 and df.iloc[:, 0].apply(lambda x: isinstance(x, str) and x.isalpha()).all():
            row_labels = df.iloc[:, 0].astype(str).tolist()
            df = df.iloc[:, 1:].reset_index(drop=True)
        else:
            row_labels = [str(i+1) for i in range(len(df))]
        table_widget.setRowCount(len(df))
        table_widget.setColumnCount(len(df.columns))
        table_widget.setHorizontalHeaderLabels([str(col) for col in df.columns])
        table_widget.setVerticalHeaderLabels(row_labels)
        for i in range(len(df)):
            for j in range(len(df.columns)):
                value = df.iloc[i, j]
                if pd.notna(value):
                    item = QTableWidgetItem(str(value))
                    table_widget.setItem(i, j, item)
    
    def clear_assignments(self):
        reply = QMessageBox.question(self, "Confirm Clear", 
                                   "Clear all assignments and removed cells for ALL sheets and files?\n"
                                   "This will reset everything to original state.",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            # Clear assignments across ALL sheets and ALL files
            for sheet_name, file_widgets in self.table_widgets.items():
                for file_path, table_widget in file_widgets.items():
                    table_widget.cell_assignments.clear()
                    table_widget.removed_cells.clear()
                    
                    # Reset cell colors and reload original data
                    if sheet_name in self.sheet_data and file_path in self.sheet_data[sheet_name]:
                        df = self.sheet_data[sheet_name][file_path]
                        self.populate_table(table_widget, df)
            
            # Reset background subtraction flag if it exists
            if hasattr(self, '_background_subtracted'):
                self._background_subtracted = False
            
            QMessageBox.information(self, "Success", "All assignments and modifications cleared for all sheets and files")
    
    def export_results(self):
        # Export all sheets, each as a separate sheet in the output file, combining all files for each sheet name
        if not self.table_widgets:
            QMessageBox.warning(self, "Warning", "No data to export")
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "Export Results", "", "Excel Files (*.xlsx)")
        if not file_path:
            return
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # 1. Main sheets: one per original sheet name (excluding background and NaN)
                all_file_cuboid = {}
                all_file_drug = {}
                all_file_drug_by_digit = {}
                file_cuboid_data = {}  # Fix: ensure this is defined for ratio calculations
                
                for sheet_name, file_widgets in self.table_widgets.items():
                    all_data = []
                    for file_name, table_widget in file_widgets.items():
                        for (row, col), assignment in table_widget.cell_assignments.items():
                            if assignment.get('is_background'):
                                continue
                            item = table_widget.item(row, col)
                            value = item.text() if item else None
                            if isinstance(value, str) and (value.startswith('\u25CF ') or value.startswith('● ')):
                                value = value[2:]
                            if value is None or str(value).strip().lower() == 'nan' or str(value).strip() == '':
                                continue
                            cuboids = assignment.get('cuboids')
                            if cuboids is None:
                                cuboids = 0
                            drug = assignment.get('drug')
                            gui_row = table_widget.verticalHeaderItem(row).text() if table_widget.verticalHeaderItem(row) else str(row + 1)
                            gui_col = table_widget.horizontalHeaderItem(col).text() if table_widget.horizontalHeaderItem(col) else str(col + 1)
                            all_data.append({
                                'File': file_name,
                                'Sheet': sheet_name,
                                'Row': gui_row,
                                'Column': gui_col,
                                'Drug': drug,
                                'Cuboids': cuboids,
                                'Value': value
                            })
                            # --- Drug sheet grouping by file last digit and cuboid ---
                            file_stem = Path(file_name).stem
                            last_digit = None
                            for char in reversed(file_stem):
                                if char.isdigit():
                                    last_digit = int(char)
                                    break
                            # Group by (last_digit, cuboids)
                            if last_digit is not None and drug:
                                key = (last_digit, cuboids)
                                if key not in all_file_drug_by_digit:
                                    all_file_drug_by_digit[key] = []
                                all_file_drug_by_digit[key].append({
                                    'Drug': drug,
                                    'Value': value,
                                    'Row': gui_row,
                                    'Column': gui_col
                                })
                            # For file-cuboid sheets (legacy, for ratio)
                            if file_name and drug:
                                key = (file_name, cuboids)
                                if key not in all_file_cuboid:
                                    all_file_cuboid[key] = []
                                all_file_cuboid[key].append({
                                    'Drug': drug,
                                    'Value': value,
                                    'Row': gui_row,
                                    'Column': gui_col
                                })
                # --- Build file_cuboid_data for ratios as before ---
                for (file_name, cuboids), records in all_file_cuboid.items():
                    file_stem = Path(file_name).stem
                    last_digit = None
                    for char in reversed(file_stem):
                        if char.isdigit():
                            last_digit = int(char)
                            break
                    if last_digit is not None:
                        key = (last_digit, cuboids)
                        # Build drug_values for this (last_digit, cuboids)
                        df_fc = pd.DataFrame(records)
                        drugs = sorted(df_fc['Drug'].dropna().unique())
                        drug_values = {}
                        for drug in drugs:
                            vals = [v[2:] if isinstance(v, str) and (v.startswith('\u25CF ') or v.startswith('● ')) else v for v in df_fc[df_fc['Drug'] == drug]['Value'].tolist()]
                            float_vals = []
                            for val in vals:
                                try:
                                    float_vals.append(float(val))
                                except (ValueError, TypeError):
                                    float_vals.append(0.0)
                            drug_values[drug] = float_vals
                        file_cuboid_data[key] = drug_values
                # Write main sheet after all_data is collected
                if all_data:
                    df = pd.DataFrame(all_data)
                    df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

                # --- Drug sheets by last digit and cuboid (unique for each day/cuboid combo) ---
                for (last_digit, cuboids), records in all_file_drug_by_digit.items():
                    sheet_label = f"Drugs_{last_digit}_Cuboid_{cuboids}"
                    df_drug = pd.DataFrame(records)
                    drugs = sorted(df_drug['Drug'].dropna().unique())
                    drug_values = {}
                    for drug in drugs:
                        vals = [v[2:] if isinstance(v, str) and (v.startswith('\u25CF ') or v.startswith('● ')) else v for v in df_drug[df_drug['Drug'] == drug]['Value'].tolist()]
                        float_vals = []
                        for val in vals:
                            try:
                                float_vals.append(float(val))
                            except (ValueError, TypeError):
                                float_vals.append(0.0)
                        drug_values[drug] = float_vals
                    max_len = max((len(vals) for vals in drug_values.values()), default=0)
                    for drug in drug_values:
                        drug_values[drug] += [0.0] * (max_len - len(drug_values[drug]))
                    export_df = pd.DataFrame(drug_values)
                    export_df.to_excel(writer, sheet_name=sheet_label[:31], index=False)
                
                # 3. Drug-only files (for files without cuboids or as additional sheets)
                file_drug_data = {}  # Store data for ratio calculations without cuboids
                
                for file_name, records in all_file_drug.items():
                    # Check if this file already has cuboid sheets
                    has_cuboid_sheets = any(fn == file_name for (fn, _) in all_file_cuboid.keys())
                    
                    if not has_cuboid_sheets:  # Only create drug-only sheet if no cuboid sheets exist
                        sheet_label = f"File_{Path(file_name).stem}_Drugs"
                        df_fd = pd.DataFrame(records)
                        
                        drugs = sorted(df_fd['Drug'].dropna().unique())
                        drug_values = {}
                        
                        for drug in drugs:
                            vals = [v[2:] if isinstance(v, str) and (v.startswith('\u25CF ') or v.startswith('● ')) else v for v in df_fd[df_fd['Drug'] == drug]['Value'].tolist()]
                            # Convert to float for calculations
                            float_vals = []
                            for val in vals:
                                try:
                                    float_vals.append(float(val))
                                except (ValueError, TypeError):
                                    float_vals.append(0.0)
                            drug_values[drug] = float_vals
                        
                        max_len = max((len(vals) for vals in drug_values.values()), default=0)
                        for drug in drug_values:
                            drug_values[drug] += [0.0] * (max_len - len(drug_values[drug]))
                        
                        export_df = pd.DataFrame(drug_values)
                        export_df.to_excel(writer, sheet_name=sheet_label[:31], index=False)
                        
                        # Store for ratio calculations (cuboid = None or 0)
                        file_stem = Path(file_name).stem
                        last_digit = None
                        for char in reversed(file_stem):
                            if char.isdigit():
                                last_digit = int(char)
                                break
                        
                        if last_digit is not None:
                            key = (last_digit, 0)  # Use 0 for no cuboids
                            if key not in file_drug_data:
                                file_drug_data[key] = {}
                            file_drug_data[key].update(drug_values)
                
                # 4. Create ratio sheets for cuboid data
                cuboid_groups = {}
                for (digit, cuboids), drug_data in file_cuboid_data.items():
                    if cuboids not in cuboid_groups:
                        cuboid_groups[cuboids] = {}
                    cuboid_groups[cuboids][digit] = drug_data
                
                for cuboids, digit_data in cuboid_groups.items():
                    if len(digit_data) < 2:
                        continue  # Need at least 2 files to create ratios
                    
                    min_digit = min(digit_data.keys())
                    baseline_data = digit_data[min_digit]
                    
                    for other_digit in sorted(digit_data.keys()):
                        if other_digit == min_digit:
                            continue
                        
                        comparison_data = digit_data[other_digit]
                        ratio_sheet_name = f"Ratio_{other_digit}_to_{min_digit}_Cuboid_{cuboids}"
                        
                        # Calculate ratios for each drug
                        ratio_data = {}
                        common_drugs = set(baseline_data.keys()) & set(comparison_data.keys())
                        
                        for drug in sorted(common_drugs):
                            baseline_vals = baseline_data[drug]
                            comparison_vals = comparison_data[drug]
                            
                            # Calculate ratios (comparison/baseline), handling division by zero
                            ratios = []
                            max_len = max(len(baseline_vals), len(comparison_vals))
                            
                            for i in range(max_len):
                                baseline_val = baseline_vals[i] if i < len(baseline_vals) else 0.0
                                comparison_val = comparison_vals[i] if i < len(comparison_vals) else 0.0
                                
                                if baseline_val != 0:
                                    ratio = comparison_val / baseline_val
                                else:
                                    ratio = float('inf') if comparison_val != 0 else 1.0
                                
                                ratios.append(ratio)
                            
                            ratio_data[drug] = ratios
                        
                        if ratio_data:
                            ratio_df = pd.DataFrame(ratio_data)
                            ratio_df.to_excel(writer, sheet_name=ratio_sheet_name[:31], index=False)
                
                # 5. Create ratio sheets for drug-only data (if any)
                if file_drug_data:
                    digit_data_drugs = {}
                    for (digit, _), drug_data in file_drug_data.items():
                        digit_data_drugs[digit] = drug_data
                    
                    if len(digit_data_drugs) >= 2:
                        min_digit = min(digit_data_drugs.keys())
                        baseline_data = digit_data_drugs[min_digit]
                        
                        for other_digit in sorted(digit_data_drugs.keys()):
                            if other_digit == min_digit:
                                continue
                            
                            comparison_data = digit_data_drugs[other_digit]
                            ratio_sheet_name = f"Ratio_{other_digit}_to_{min_digit}_Drugs"
                            
                            # Calculate ratios for each drug
                            ratio_data = {}
                            common_drugs = set(baseline_data.keys()) & set(comparison_data.keys())
                            
                            for drug in sorted(common_drugs):
                                baseline_vals = baseline_data[drug]
                                comparison_vals = comparison_data[drug]
                                
                                # Calculate ratios (comparison/baseline), handling division by zero
                                ratios = []
                                max_len = max(len(baseline_vals), len(comparison_vals))
                                
                                for i in range(max_len):
                                    baseline_val = baseline_vals[i] if i < len(baseline_vals) else 0.0
                                    comparison_val = comparison_vals[i] if i < len(comparison_vals) else 0.0
                                    
                                    if baseline_val != 0:
                                        ratio = comparison_val / baseline_val
                                    else:
                                        ratio = float('inf') if comparison_val != 0 else 1.0
                                    
                                    ratios.append(ratio)
                                
                                ratio_data[drug] = ratios
                            
                            if ratio_data:
                                ratio_df = pd.DataFrame(ratio_data)
                                ratio_df.to_excel(writer, sheet_name=ratio_sheet_name[:31], index=False)
            
            QMessageBox.information(self, "Success", f"Results exported to {file_path} (including ratio sheets)")
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "Error", f"Failed to export results: {str(e)}\n{traceback.format_exc()}")
    
    def calculate_background_subtraction(self):
        # Only allow background subtraction once per session
        if hasattr(self, '_background_subtracted') and self._background_subtracted:
            QMessageBox.warning(self, "Warning", "Background subtraction has already been applied.")
            return
        for sheet_name, file_widgets in self.table_widgets.items():
            for file_name, table_widget in file_widgets.items():
                background_values = [assignment['original_value']
                                    for assignment in table_widget.cell_assignments.values()
                                    if assignment['is_background'] and assignment['original_value'] is not None]
                bg_avg = np.mean(background_values) if background_values else 0.0
                for i in range(table_widget.rowCount()):
                    for j in range(table_widget.columnCount()):
                        item = table_widget.item(i, j)
                        if item:
                            try:
                                orig_val = float(item.text())
                                new_val = orig_val - bg_avg
                                item.setText(f"{new_val:.4f}")
                            except ValueError:
                                pass
        self._background_subtracted = True
        QMessageBox.information(self, "Success", "Background subtraction applied to all sheets")
    
    def show_assignment_summary(self):
        # Collect assignments from all tables for the current sheet
        if not self.current_sheet or self.current_sheet not in self.table_widgets:
            QMessageBox.warning(self, "Warning", "No sheet selected or no assignments found")
            return
        
        summary = []
        for table_widget in self.table_widgets[self.current_sheet].values():
            for (row, col), assignment in table_widget.cell_assignments.items():
                if assignment:
                    summary.append({
                        'Row': row + 1,
                        'Column': col + 1,
                        'Drug': assignment['drug'],
                        'Cuboids': assignment['cuboids'],
                        'Background': assignment['is_background']
                    })
        
        if not summary:
            QMessageBox.information(self, "Summary", "No assignments found for the selected sheet")
            return
        
        # Convert to DataFrame for better display
        summary_df = pd.DataFrame(summary)
        self.show_summary_dialog(summary_df)
    
    def show_summary_dialog(self, summary_df):
        from PyQt6.QtWidgets import QDialog, QTableView, QVBoxLayout, QPushButton
        from PyQt6.QtCore import Qt
        dialog = QDialog(self)
        dialog.setWindowTitle("Assignment Summary")
        dialog.resize(800, 600)
        
        layout = QVBoxLayout(dialog)
        
        table_view = QTableView(dialog)
        table_view.setModel(self.pandas_model(summary_df))
        table_view.resizeColumnsToContents()
        table_view.setAlternatingRowColors(True)
        table_view.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        table_view.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        layout.addWidget(table_view)
        
        close_button = QPushButton("Close", dialog)
        close_button.clicked.connect(dialog.accept)
        layout.addWidget(close_button)
        
        dialog.exec()
    
    class pandas_model(pd.DataFrame):
        def __init__(self, df, *args, **kwargs):
            super().__init__(df, *args, **kwargs)
        
        def _get_repr_html_(self):
            return self.to_html(classes='table table-striped table-hover', index=False, border=0)
    
    def export_assignments(self):
        # This method is deprecated and the Export Assignments button is removed from the UI.
        pass
    
    def get_all_table_widgets(self, sheet_name):
        """Return all SelectableTableWidget instances for a given sheet name across all files."""
        if sheet_name in self.table_widgets:
            return list(self.table_widgets[sheet_name].values())
        return []

def main():
    app = QApplication(sys.argv)
    window = ExcelAnalyzerGUI()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
