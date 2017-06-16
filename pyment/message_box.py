# coding=utf-8
"""Alert box."""

# !/usr/bin/env python

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-07T09:16:15+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: message_box.py
# @Last modified by:   gysco
# @Last modified time: 2017-06-16T11:35:46+02:00

import ctypes


def message_box(title, text, style):
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
