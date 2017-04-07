"""
Worksheet class file.

This class contains the full worksheets (like a Excel file).
"""

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-07T08:52:22+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: Worksheet.py
# @Last modified by:   gysco
# @Last modified time: 2017-04-07T09:30:45+02:00


class Worksheet:
    """Worksheet class."""

    def __init__(self, cells=list(), name=""):
        """Initialization for cells."""
        self.Cells = cells
        self.Name = name
