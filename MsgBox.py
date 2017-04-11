"""Alert box."""

# !/usr/bin/env python

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-07T09:16:15+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: Mbox.py
# @Last modified by:   gysco
# @Last modified time: 2017-04-11T14:52:50+02:00

import ctypes


def MsgBox(title, text, style):
    """
    Used to create alert message.

    Styles:
    0 : OK
    1 : OK | Cancel
    2 : Abort | Retry | Ignore
    3 : Yes | No | Cancel
    4 : Yes | No
    5 : Retry | No 
    6 : Cancel | Try Again | Continue
    """
    ctypes.windll.user32.MessageBoxW(0, text, title, style)
