# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-06-16T13:20:47+02:00
# @Email:  zackary.b@live.fr
# @Project: PyMENT-SSWD
# @Filename: message_box.py
# @Last modified by:   gysco
# @Last modified time: 2017-06-16T13:21:47+02:00

import wx


def message_box(title, text, style):
    """
    Used to create alert message.

    Styles:
    0 : OK
    1 : OK | Cancel
    2 : Yes | No | Cancel
    3 : Yes | No
    """
    wx_style = [wx.OK, wx.OK | wx.CANCEL, wx.YES_NO | wx.CANCEL, wx.YES_NO]
    wx_end_style = wx.WARNING | wx.CENTER
    wx.MessageBoxW(text, title, wx_style[style] | wx_end_style)
