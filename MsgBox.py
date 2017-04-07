"""Alert box."""

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-07T09:16:15+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: Mbox.py
# @Last modified by:   gysco
# @Last modified time: 2017-04-07T09:17:44+02:00


import ctypes


def MsgBox(title, text, style):
    """Used to create alert message."""
    ctypes.windll.user32.MessageBoxW(0, text, title, style)
