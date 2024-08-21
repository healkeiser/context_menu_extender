#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Built-in
import os
import sys
import winreg

# Third-party
from qtpy import QtWidgets, QtGui, QtCore
from qtpy.QtGui import QIcon


class ContextMenuExtenderApp(QtWidgets.QMainWindow):
    def __init__(self):
        super(ContextMenuExtenderApp, self).__init__()

        self.setup_ui()
        self.update_button_text()

    def setup_ui(self):
        self.setWindowTitle("Context Menu Extender")
        self.setWindowIcon(
            QIcon(
                os.path.join(
                    os.path.dirname(os.path.abspath(__file__)),
                    "icons",
                    "index.svg",
                )
            )
        )
        self.setGeometry(100, 100, 400, 400)

        self.statusbar = self.statusBar()

        central_widget = QtWidgets.QWidget(self)
        self.setCentralWidget(central_widget)
        self.layout = QtWidgets.QVBoxLayout(central_widget)
        central_widget.setLayout(self.layout)

        self.filter_line_edit = QtWidgets.QLineEdit()
        self.filter_line_edit.setPlaceholderText("Filter...")
        self.filter_line_edit.textChanged.connect(self.filter_extensions)
        self.layout.addWidget(self.filter_line_edit)

        self.scroll_area = QtWidgets.QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_widget = QtWidgets.QWidget()
        self.scroll_layout = QtWidgets.QVBoxLayout(self.scroll_widget)

        self.extension_vars = {}

        # Enumerate subkeys under HKEY_CLASSES_ROOT to find file extensions
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "") as key:
            for i in range(winreg.QueryInfoKey(key)[0]):
                subkey_name = winreg.EnumKey(key, i)
                if subkey_name.startswith("."):
                    ext = subkey_name[1:]
                    var = QtWidgets.QCheckBox(ext)
                    var.stateChanged.connect(self.update_button_text)
                    self.extension_vars[ext] = var
                    self.scroll_layout.addWidget(var)

        self.scroll_layout.addStretch()
        self.scroll_area.setWidget(self.scroll_widget)
        self.layout.addWidget(self.scroll_area)

        self.button_layout = QtWidgets.QHBoxLayout()

        self.add_button = QtWidgets.QPushButton("Add Selected Extensions (0)")
        self.add_button.clicked.connect(self.add_to_context_menu)
        self.button_layout.addWidget(self.add_button)

        self.unselect_all_button = QtWidgets.QPushButton("Unselect All")
        self.unselect_all_button.clicked.connect(self.unselect_all)
        self.button_layout.addWidget(self.unselect_all_button)

        cancel_button = QtWidgets.QPushButton("Cancel")
        cancel_button.clicked.connect(self.close)
        self.button_layout.addWidget(cancel_button)

        self.layout.addLayout(self.button_layout)

        self.result_label = QtWidgets.QLabel()
        self.layout.addWidget(self.result_label)

        self.setLayout(self.layout)

    def add_to_context_menu(self):
        selected_items = [
            extension
            for extension, var in self.extension_vars.items()
            if var.isChecked()
        ]
        num_selected = len(selected_items)

        for ext in selected_items:
            key_path = f".{ext}"
            try:
                key = winreg.OpenKey(
                    winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_WRITE
                )
                winreg.CreateKey(key, "ShellNew")
                winreg.SetValue(key, "ShellNew", winreg.REG_SZ, "")
                winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "1")
                winreg.CloseKey(key)
                self.statusbar.showMessage(
                    f"Successfully added {num_selected} extension(s) to the context menu"
                )
            except Exception as e:
                self.statusbar.showMessage(f"Error adding {ext}: {e}")

    def filter_extensions(self):
        filter_text = self.filter_line_edit.text()
        for ext, var in self.extension_vars.items():
            var.setVisible(filter_text.lower() in ext.lower())

    def unselect_all(self):
        """Uncheck all checkboxes."""

        for var in self.extension_vars.values():
            var.setChecked(False)

        # Update the add button text
        self.update_button_text()

    def update_button_text(self):
        """Update the add button text with the number of selected extensions and enable/disable the button."""

        num_selected = sum(
            var.isChecked() for var in self.extension_vars.values()
        )
        self.add_button.setText(f"Add Selected Extensions ({num_selected})")
        self.add_button.setEnabled(num_selected > 0)
        self.unselect_all_button.setEnabled(num_selected > 0)


def dark_palette():
    dark_palette = QtGui.QPalette()
    dark_palette.setColor(QtGui.QPalette.Window, QtGui.QColor(53, 53, 53))
    dark_palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.Base, QtGui.QColor(25, 25, 25))
    dark_palette.setColor(
        QtGui.QPalette.AlternateBase, QtGui.QColor(53, 53, 53)
    )
    dark_palette.setColor(QtGui.QPalette.ToolTipBase, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.ToolTipText, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.Text, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.Button, QtGui.QColor(53, 53, 53))
    dark_palette.setColor(QtGui.QPalette.ButtonText, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.BrightText, QtCore.Qt.red)
    dark_palette.setColor(QtGui.QPalette.Link, QtGui.QColor(42, 130, 218))
    dark_palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(42, 130, 218))
    dark_palette.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.black)
    return dark_palette


if __name__ == "__main__":
    app = QtWidgets.QApplication([])

    # Set style
    app.setStyle("Fusion")
    app.setPalette(dark_palette())

    # Launch
    main_app = ContextMenuExtenderApp()
    main_app.show()
    sys.exit(app.exec_())
