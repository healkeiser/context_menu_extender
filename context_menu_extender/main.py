#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Built-in
import sys
import winreg

# Third-party
from PySide2 import QtWidgets

# Metadatas
__author__ = "Valentin Beaumont"
__email__ = "valentin.onze@gmail.com"


###### CODE ####################################################################


def add_to_context_menu():
    """_summary_"""

    selected_items = [
        extension
        for extension, var in extension_vars.items()
        if var.isChecked()
    ]

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
        except Exception as e:
            result_label.setText(f"Error adding {ext}: {e}")
        else:
            result_label.setText(
                f"{ext} added successfully to the context menu"
            )


def filter_extensions():
    """_summary_"""

    filter_text = filter_line_edit.text()
    for ext, var in extension_vars.items():
        var.setVisible(filter_text.lower() in ext.lower())


app = QtWidgets.QApplication([])

window = QtWidgets.QWidget()
window.setWindowTitle("Context Menu Extender")
window.setGeometry(100, 100, 400, 400)

layout = QtWidgets.QVBoxLayout()

filter_line_edit = QtWidgets.QLineEdit()
filter_line_edit.setPlaceholderText("Filter...")
filter_line_edit.textChanged.connect(filter_extensions)
layout.addWidget(filter_line_edit)

scroll_area = QtWidgets.QScrollArea()
scroll_area.setWidgetResizable(True)
scroll_widget = QtWidgets.QWidget()
scroll_layout = QtWidgets.QVBoxLayout(scroll_widget)

extension_vars = {}

# Enumerate subkeys under HKEY_CLASSES_ROOT to find file extensions
with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "") as key:
    for i in range(winreg.QueryInfoKey(key)[0]):
        subkey_name = winreg.EnumKey(key, i)
        if subkey_name.startswith("."):
            ext = subkey_name[1:]
            var = QtWidgets.QCheckBox(ext)
            extension_vars[ext] = var
            scroll_layout.addWidget(var)

scroll_layout.addStretch()
scroll_area.setWidget(scroll_widget)
layout.addWidget(scroll_area)

button_layout = QtWidgets.QHBoxLayout()

add_button = QtWidgets.QPushButton("Add Selected Extensions")
add_button.clicked.connect(add_to_context_menu)
button_layout.addWidget(add_button)

cancel_button = QtWidgets.QPushButton("Cancel")
cancel_button.clicked.connect(window.close)
button_layout.addWidget(cancel_button)

layout.addLayout(button_layout)

result_label = QtWidgets.QLabel()
layout.addWidget(result_label)

window.setLayout(layout)
window.show()

sys.exit(app.exec_())
