"""
Worksheet class file.

This class contains the full worksheets (like a Excel file).
"""

# !/usr/bin/env python

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-07T08:52:22+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: Worksheet.py
# @Last modified by:   gysco
# @Last modified time: 2017-05-05T10:01:45+02:00

import pandas


class Worksheet:
    """Worksheet class."""

    def __repr__(self):
        """Return the str."""
        ret = self.Cells.to_csv()
        return (ret)

    def __init__(self, cells=pandas.DataFrame(), name=""):
        """Initialization for cells."""
        self.Cells = cells
        self.Name = name
