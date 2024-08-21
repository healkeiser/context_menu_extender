"""PyPI setup script."""

# Built-in
from pathlib import Path
from setuptools import setup, find_packages
import sys

# Metadata
__author__ = "Valentin Beaumont"
__email__ = "valentin.onze@gmail.com"


# Add `README.md` as project long description
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text()

# Required dependencies
install_requires = ["qtpy", "QtAwesome"]
if sys.version_info < (3, 11):
    install_requires.append("PySide2")
else:
    install_requires.append("PySide6")

setup(
    name="fxgui",
    use_scm_version=True,
    setup_requires=["setuptools_scm"],
    long_description=long_description,
    long_description_content_type="text/markdown",
    description="A simple Python application using Qt that allows you to add selected file extensions to the Windows context menu effortlessly. Filter and select the extensions you want to add with ease.",
    url="https://github.com/healkeiser/context_menu_extender",
    author="Valentin Beaumont",
    author_email="valentin.onze@gmail.com",
    license="MIT",
    keywords="Qt PySide2 VFX DCC UI",
    packages=find_packages(),
    install_requires=install_requires,
    include_package_data=True,
)
