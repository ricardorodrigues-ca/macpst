"""
Mac PST Converter - Convert Outlook PST files to various formats on macOS
"""

__version__ = "1.0.0"
__author__ = "Ricardo Rodrigues"
__email__ = "ricardo@example.com"

from .core.pst_parser import PSTParser
from .core.converter import Converter
from .gui.main_window import MainWindow

__all__ = ["PSTParser", "Converter", "MainWindow"]