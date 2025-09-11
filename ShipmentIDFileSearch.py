
# --- Cleaned and reordered code ---
import sys
import os
import configparser
import pyodbc
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox, QDialog, QFormLayout,
    QTableWidget, QTableWidgetItem, QStackedWidget, QHeaderView, QCheckBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPalette, QColor, QIcon

SETTINGS_FILE = 'settings.ini'

def load_settings():
    config = configparser.ConfigParser()
    if os.path.exists(SETTINGS_FILE):
        config.read(SETTINGS_FILE)
    if 'DEFAULT' not in config:
        config['DEFAULT'] = {}
    s = config['DEFAULT']
    return {
        'searchDir': s.get('searchDir') or s.get('searchdir') or r'C:\Program Files\Manhattan Associates\ILS\2020\Interface\Output',
        'fileEditor': s.get('fileEditor') or s.get('fileeditor') or r'C:\Program Files\Notepad++\notepad++.exe',
        'server': s.get('server', ''),
        'database': s.get('database', ''),
        'darkMode': s.get('darkMode', 'true').lower() == 'true'
    }

def save_settings(settings):
    config = configparser.ConfigParser()
    config['DEFAULT'] = {
        'searchDir': settings['searchDir'],
        'fileEditor': settings['fileEditor'],
        'server': settings['server'],
        'database': settings['database'],
        'darkMode': str(settings['darkMode']).lower()
    }
    with open(SETTINGS_FILE, 'w') as f:
        config.write(f)

class SettingsDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Settings')
        self.setMinimumWidth(600)
        self.settings = settings
        layout = QFormLayout()
        self.dirEdit = QLineEdit(settings['searchDir'])
        self.editorEdit = QLineEdit(settings['fileEditor'])
        self.serverEdit = QLineEdit(settings.get('server', ''))
        self.dbEdit = QLineEdit(settings.get('database', ''))
        self.darkModeCheck = QCheckBox('Dark Mode')
        self.darkModeCheck.setChecked(settings.get('darkMode', True))
        layout.addRow('Search Directory:', self.dirEdit)
        layout.addRow('File Editor Path:', self.editorEdit)
        layout.addRow('SQL Server:', self.serverEdit)
        layout.addRow('Database:', self.dbEdit)
        layout.addRow(self.darkModeCheck)
        btnBox = QHBoxLayout()
        saveBtn = QPushButton('Save')
        saveBtn.clicked.connect(self.accept)
        btnBox.addWidget(saveBtn)
        layout.addRow(btnBox)
        self.setLayout(layout)

    def accept(self):
        # Preserve single backslashes in directory path
        search_dir = self.dirEdit.text().strip()
        # Keep single backslashes as-is (don't convert to forward slashes)
        
        self.settings['searchDir'] = search_dir
        self.settings['fileEditor'] = self.editorEdit.text().strip()
        self.settings['server'] = self.serverEdit.text().strip()
        self.settings['database'] = self.dbEdit.text().strip()
        self.settings['darkMode'] = self.darkModeCheck.isChecked()
        save_settings(self.settings)
        super().accept()

class FileViewerDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle('Shipment ID File Search')
        self.setMinimumSize(1000, 700)
        self.setWindowFlags(Qt.Window)  # Ensure it's a proper window with taskbar icon
        self.mapping = {}
        self.file_data = {}
        self.current_page = 0
        self.current_file_path = None
        self.files_searched = 0
        self.map_order = []
        self.initUI()

    def get_mapping(self):
        s = self.settings
        if not s['server'] or not s['database']:
            return {}

        # Use context manager for better resource management
        try:
            with pyodbc.connect(f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={s['server']};DATABASE={s['database']};Trusted_Connection=yes;") as conn:
                with conn.cursor() as cursor:
                    cursor.execute("SELECT MAP_NAME, POSITION, FIELD_NAME FROM INTERFACE_DATA_MAP_DETAIL")
                    mapping = {}
                    for row in cursor.fetchall():
                        map_name, pos, field = row
                        if map_name not in mapping:
                            mapping[map_name] = {}
                        # Convert POSITION to int since SQL Server returns it as Decimal
                        mapping[map_name][int(pos)] = field
                    return mapping
        except Exception as e:
            QMessageBox.warning(self, 'SQL Error', f'Failed to get mapping: {e}')
            return {}

    def parse_file(self, file_path):
        data = {}
        map_order = []
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                for line_num, line in enumerate(f, 1):
                    try:
                        parts = line.strip().split('|')
                        if parts and len(parts) > 0:
                            map_name = parts[0]
                            if map_name not in data:
                                data[map_name] = []
                                map_order.append(map_name)
                            data[map_name].append(parts)
                    except Exception as e:
                        # Skip lines that can't be parsed
                        continue
        except FileNotFoundError:
            QMessageBox.critical(self, 'Error', f'File not found: {file_path}')
        except PermissionError:
            QMessageBox.critical(self, 'Error', f'Permission denied reading file: {file_path}')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to parse file: {e}')
        return data, map_order

    def load_file(self, file_path):
        self.current_file_path = file_path
        self.mapping = self.get_mapping()
        self.file_data, self.map_order = self.parse_file(file_path)
        
        # Clear existing pages
        while self.stacked.count() > 0:
            widget = self.stacked.widget(0)
            self.stacked.removeWidget(widget)
            widget.deleteLater()
        
        self.pages = []
        for map_name in self.map_order:  # Use the order from the file
            page = self.create_page(map_name)
            self.pages.append(page)
            self.stacked.addWidget(page)
        
        self.current_page = 0
        self.update_navigation()
        self.open_btn.setEnabled(True)
        self.setWindowTitle(f'File Viewer - {os.path.basename(file_path)}')

    def initUI(self):
        # Set window icon
        if os.path.exists('shipment.ico'):
            self.setWindowIcon(QIcon('shipment.ico'))
        
        # Apply dark mode if enabled
        if self.settings.get('darkMode', True):
            self.apply_dark_mode()
        
        layout = QVBoxLayout()
        
        # Search section
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel('Shipment ID:'))
        self.shipmentIdEdit = QLineEdit()
        self.shipmentIdEdit.returnPressed.connect(self.on_search)
        search_layout.addWidget(self.shipmentIdEdit)
        
        self.searchBtn = QPushButton('Search')
        self.searchBtn.clicked.connect(self.on_search)
        search_layout.addWidget(self.searchBtn)
        
        self.settingsBtn = QPushButton('Settings')
        self.settingsBtn.clicked.connect(self.open_settings)
        search_layout.addWidget(self.settingsBtn)
        
        layout.addLayout(search_layout)
        
        # Status label
        self.status_label = QLabel('')
        layout.addWidget(self.status_label)
        
        # Navigation buttons
        nav_layout = QHBoxLayout()
        self.prev_btn = QPushButton('← Previous')
        self.prev_btn.clicked.connect(self.prev_page)
        nav_layout.addWidget(self.prev_btn)
        
        # Add stretch to center the page label
        nav_layout.addStretch()
        
        self.page_label = QLabel('No file loaded')
        self.page_label.setAlignment(Qt.AlignCenter)
        nav_layout.addWidget(self.page_label)
        
        # Add stretch to center the page label
        nav_layout.addStretch()
        
        self.next_btn = QPushButton('Next →')
        self.next_btn.clicked.connect(self.next_page)
        nav_layout.addWidget(self.next_btn)
        
        layout.addLayout(nav_layout)
        
        # Stacked widget for pages
        self.stacked = QStackedWidget()
        layout.addWidget(self.stacked)
        
        # Bottom buttons
        btn_layout = QHBoxLayout()
        self.open_btn = QPushButton('Open in Notepad')
        self.open_btn.clicked.connect(self.open_in_notepad)
        self.open_btn.setEnabled(False)
        btn_layout.addWidget(self.open_btn)
        
        self.close_btn = QPushButton('Close')
        self.close_btn.clicked.connect(self.close)
        btn_layout.addWidget(self.close_btn)
        
        layout.addLayout(btn_layout)
        self.setLayout(layout)
        
        # Initialize navigation state
        self.update_navigation()

    def apply_dark_mode(self):
        dark_stylesheet = """
            QWidget { background-color: #232629; color: #F0F0F0; }
            QLineEdit, QTextEdit, QPlainTextEdit { background-color: #2b2b2b; color: #F0F0F0; border: 1px solid #444; }
            QPushButton { background-color: #444; color: #F0F0F0; border: 1px solid #666; padding: 4px; }
            QPushButton:hover { background-color: #555; }
            QLabel { color: #F0F0F0; }
            QDialog { background-color: #232629; color: #F0F0F0; }
            QTableWidget { background-color: #2b2b2b; color: #F0F0F0; gridline-color: #444; }
            QTableWidget::item { background-color: #2b2b2b; color: #F0F0F0; }
            QTableWidget::item:selected { background-color: #444; }
            QHeaderView::section { background-color: #444; color: #F0F0F0; border: 1px solid #666; }
            QCheckBox { color: #F0F0F0; }
        """
        self.setStyleSheet(dark_stylesheet)

    def closeEvent(self, event):
        """Ensure proper cleanup when dialog is closed"""
        # Clear any large data structures
        self.mapping.clear()
        self.file_data.clear()
        self.map_order.clear()

        # Clear the stacked widget
        while self.stacked.count() > 0:
            widget = self.stacked.widget(0)
            self.stacked.removeWidget(widget)
            widget.deleteLater()

        # Clear pages list
        if hasattr(self, 'pages'):
            self.pages.clear()

        # Accept the close event
        event.accept()

        # Ensure the application quits when the dialog is closed
        QApplication.quit()

    def create_page(self, map_name):
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Make the map label larger and centered
        map_label = QLabel(f'Map: {map_name}')
        map_label.setStyleSheet('font-size: 16px; font-weight: bold;')
        map_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(map_label)
        
        table = QTableWidget()
        
        # Get the data rows for this map
        if map_name not in self.file_data or not self.file_data[map_name]:
            # No data for this map
            table.setColumnCount(1)
            table.setHorizontalHeaderLabels(['No Data'])
            table.setRowCount(1)
            table.setItem(0, 0, QTableWidgetItem('No data found for this map'))
        else:
            data_rows = self.file_data[map_name]
            
            # Get field names from mapping for this map
            map_fields = self.mapping.get(map_name, {})
            if not map_fields:
                # No mapping found, create generic columns
                max_cols = max(len(row) for row in data_rows) if data_rows else 1
                table.setColumnCount(max_cols)
                table.setHorizontalHeaderLabels([f'Field_{i+1}' for i in range(max_cols)])
            else:
                # Use mapped field names as columns
                max_pos = max(map_fields.keys()) if map_fields else 1
                table.setColumnCount(max_pos)
                headers = []
                for pos in range(1, max_pos + 1):
                    field_name = map_fields.get(pos, f'Field_{pos}')
                    headers.append(field_name)
                table.setHorizontalHeaderLabels(headers)
            
            # Set row count to number of data rows
            table.setRowCount(len(data_rows))
            
            # Populate the table
            for row_idx, row_parts in enumerate(data_rows):
                for col_idx, value in enumerate(row_parts):
                    if col_idx < table.columnCount():
                        table.setItem(row_idx, col_idx, QTableWidgetItem(value))
        
        # Set table properties
        table.resizeColumnsToContents()
        table.horizontalHeader().setStretchLastSection(True)
        
        layout.addWidget(table)
        widget.setLayout(layout)
        return widget

    def update_navigation(self):
        if not hasattr(self, 'pages') or not self.pages:
            self.page_label.setText('No file loaded')
            self.prev_btn.setEnabled(False)
            self.next_btn.setEnabled(False)
            return
            
        map_names = self.map_order
        if self.current_page >= len(map_names):
            return
            
        current_map = map_names[self.current_page]
        
        prev_map = map_names[self.current_page - 1] if self.current_page > 0 else ''
        next_map = map_names[self.current_page + 1] if self.current_page < len(map_names) - 1 else ''
        
        self.page_label.setText(f'Page {self.current_page + 1} of {len(self.pages)}')
        
        self.prev_btn.setText(f'← {prev_map}' if prev_map else '← Previous')
        self.next_btn.setText(f'{next_map} →' if next_map else 'Next →')
        
        self.prev_btn.setEnabled(self.current_page > 0)
        self.next_btn.setEnabled(self.current_page < len(self.pages) - 1)
        self.stacked.setCurrentIndex(self.current_page)

    def prev_page(self):
        if not hasattr(self, 'pages') or not self.pages:
            return
        if self.current_page > 0:
            self.current_page -= 1
            self.update_navigation()

    def next_page(self):
        if not hasattr(self, 'pages') or not self.pages:
            return
        if self.current_page < len(self.pages) - 1:
            self.current_page += 1
            self.update_navigation()

    def open_in_notepad(self):
        if not self.current_file_path:
            return
        editor = self.settings['fileEditor']
        if not os.path.isfile(editor):
            QMessageBox.critical(self, 'Error', f'Configured file editor not found:\n{editor}')
        else:
            import subprocess
            try:
                subprocess.run([editor, self.current_file_path], check=True)
            except subprocess.CalledProcessError as e:
                QMessageBox.critical(self, 'Error', f'Failed to open file in editor:\n{e}')

    def open_settings(self):
        dlg = SettingsDialog(self.settings, self)
        dlg.exec_()

    def on_search(self):
        shipment_id = self.shipmentIdEdit.text().strip()
        if not shipment_id:
            QMessageBox.warning(self, 'Error', 'Please enter a Shipment ID.')
            return
        dt = self.get_datetime_from_sql(shipment_id)
        if not dt:
            QMessageBox.warning(self, 'Error', 'Could not retrieve creation datetime from SQL.')
            return
        self.search_files(shipment_id, dt)

    def get_datetime_from_sql(self, shipment_id):
        s = self.settings
        if not s['server'] or not s['database']:
            QMessageBox.warning(self, 'Error', 'SQL Server or Database not configured.')
            return None

        # Use context manager for better resource management
        try:
            with pyodbc.connect(f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={s['server']};DATABASE={s['database']};Trusted_Connection=yes;") as conn:
                with conn.cursor() as cursor:
                    cursor.execute(
                        """
                        SELECT CREATION_DATE_TIME_STAMP = DATEADD(HH, -5, SH.CREATION_DATE_TIME_STAMP)
                        FROM SHIPMENT_HEADER SH
                        WHERE SH.SHIPMENT_ID = ?
                        """, shipment_id)
                    row = cursor.fetchone()
                    if row and row[0]:
                        # Format without microseconds to match AHK behavior
                        formatted_date = row[0].strftime('%Y-%m-%d %H:%M:%S')
                        return formatted_date
                    else:
                        QMessageBox.warning(self, 'SQL Error', f'No shipment found with ID: {shipment_id}')
                        return None
        except Exception as e:
            QMessageBox.warning(self, 'SQL Error', f'Database connection failed: {e}')
            return None

    def search_files(self, shipment_id, user_dt):
        try:
            dt = datetime.strptime(user_dt, '%Y-%m-%d %H:%M:%S')
        except ValueError:
            QMessageBox.warning(self, 'Error', 'Invalid date format from SQL.')
            return

        yy = dt.strftime('%y')
        mm = dt.strftime('%m')
        dd = dt.strftime('%d')
        pattern = f"shp-{mm}{dd}{yy}"
        search_dir = self.settings['searchDir']

        if not os.path.isdir(search_dir):
            QMessageBox.critical(self, 'Error', f'Configured search directory not found:\n{search_dir}')
            return

        found = False
        count = 0

        try:
            # Use a more efficient approach with list comprehension
            files = [f for f in os.listdir(search_dir)
                    if f.startswith(pattern) and f.endswith('.sh.proc')]
        except (OSError, PermissionError) as e:
            QMessageBox.critical(self, 'Error', f'Error reading directory:\n{e}')
            return

        # Update status to show we're searching
        self.status_label.setText('Searching files...')

        for fname in files:
            count += 1
            fpath = os.path.join(search_dir, fname)

            # Update progress every 10 files
            if count % 10 == 0:
                self.status_label.setText(f'Searching... checked {count} files')

            try:
                with open(fpath, 'r', encoding='utf-8', errors='ignore') as f:
                    contents = f.read()
                if shipment_id in contents:
                    self.files_searched = count
                    self.status_label.setText(f'Found shipment in file: {fname} (searched {count} files)')
                    self.load_file(fpath)
                    found = True
                    break
            except (OSError, PermissionError):
                # Skip files we can't read
                continue
            except Exception:
                # Skip any other file reading errors
                continue
        
        if not found:
            self.status_label.setText(f'Done searching through {count} files. Shipment not found.')
            self.files_searched = count

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.settings = load_settings()
        # Just open the file viewer dialog directly
        self.open_file_viewer()

    def open_file_viewer(self):
        # Create dialog without parent to avoid cleanup issues
        dialog = FileViewerDialog(self.settings)

        # Connect dialog signals to ensure application quits properly
        dialog.finished.connect(self.on_dialog_finished)
        dialog.rejected.connect(self.on_dialog_finished)

        result = dialog.exec_()
        return result

    def on_dialog_finished(self):
        # Close the main window and quit the application when dialog closes
        self.close()
        QApplication.quit()

    def closeEvent(self, event):
        # Ensure the application quits when the main window is closed
        QApplication.quit()
        event.accept()

    def closeEvent(self, event):
        # Ensure proper cleanup
        super().closeEvent(event)

if __name__ == '__main__':
    app = QApplication(sys.argv)

    # Set application properties for better cleanup
    app.setApplicationName('Shipment File Search')
    app.setApplicationVersion('1.0')
    app.setOrganizationName('JASCO')

    # Create main window which will handle the dialog
    mw = MainWindow()

    # Start the event loop
    exit_code = app.exec_()

    # Ensure clean shutdown
    app.quit()
    sys.exit(exit_code)
