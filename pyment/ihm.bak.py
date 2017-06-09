# coding=utf-8
"""IHM file."""

# -*- coding: utf-8 -*-
###########################################################################
#  Python code generated with wxFormBuilder (version Sep 12 2010)
#  http://www.wxformbuilder.org/
#
#  PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################
# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-05-29T12:32:17+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: ihm.py
# @Last modified by:   gysco
# @Last modified time: 2017-06-01T16:22:31+02:00

from os.path import basename, splitext
from sys import exit as sysexit

import wx
from common import parse_file
from ihm_functions import charger_parametres
from pandas import read_csv, read_excel


class mainFrame(wx.Frame):
    """IHM class."""

    def __init__(self, parent):
        """__init__ of IHM."""
        wx.Frame.__init__(
            self,
            parent,
            id=wx.ID_ANY,
            title="PyGME[N]T",
            pos=wx.DefaultPosition,
            size=wx.Size(-1, -1),
            style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL ^ wx.RESIZE_BORDER)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.filename = None
        fgSizer1 = wx.FlexGridSizer(12, 2, 9, 25)
        fgSizer1.SetFlexibleDirection(wx.BOTH)
        fgSizer1.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        self.m_staticText1 = wx.StaticText(self, wx.ID_ANY, "Input file:",
                                           wx.DefaultPosition, wx.DefaultSize,
                                           0)
        self.m_staticText1.Wrap(-1)
        fgSizer1.Add(self.m_staticText1, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL,
                     5)

        self.m_filePicker1 = wx.FilePickerCtrl(
            self, wx.ID_ANY, wx.EmptyString, "Select a file",
            "Data files (*.csv, *.xls, *.xlsx)|*.csv;*.xls;*.xlsx",
            wx.DefaultPosition, wx.Size(-1, -1), wx.FLP_DEFAULT_STYLE)
        fgSizer1.Add(self.m_filePicker1, 1, wx.ALIGN_CENTER_VERTICAL, 5)

        self.text_selected_sheet = wx.StaticText(
            self, wx.ID_ANY, "Selected Sheet:", wx.DefaultPosition,
            wx.DefaultSize, 0)
        fgSizer1.Add(self.text_selected_sheet, 0, wx.ALL, 5)

        self.text_sheet_name = wx.StaticText(
            self, wx.ID_ANY, "", wx.DefaultPosition, wx.DefaultSize, 0)
        fgSizer1.Add(self.text_sheet_name, 0, wx.ALL, 5)

        self.m_staticText2 = wx.StaticText(self, wx.ID_ANY, "Output file:",
                                           wx.DefaultPosition, wx.DefaultSize,
                                           0)
        self.m_staticText2.Wrap(-1)
        fgSizer1.Add(self.m_staticText2, 0, wx.ALL, 5)

        self.m_staticText3 = wx.StaticText(
            self, wx.ID_ANY, "", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText3.Wrap(-1)
        fgSizer1.Add(self.m_staticText3, 0, wx.ALL, 5)

        self.m_staticText5 = wx.StaticText(self, wx.ID_ANY, "Taxonomic group:",
                                           wx.DefaultPosition, wx.DefaultSize,
                                           0)
        self.m_staticText5.Wrap(-1)
        fgSizer1.Add(self.m_staticText5, 0, wx.ALL, 5)

        m_choice1Choices = []
        self.m_choice1 = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition,
                                   wx.DefaultSize, m_choice1Choices, 0)
        self.m_choice1.SetSelection(0)
        fgSizer1.Add(self.m_choice1, 0, wx.ALL, 5)

        self.m_staticText6 = wx.StaticText(
            self, wx.ID_ANY, "Species:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText6.Wrap(-1)
        fgSizer1.Add(self.m_staticText6, 0, wx.ALL, 5)

        m_choice2Choices = []
        self.m_choice2 = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition,
                                   wx.DefaultSize, m_choice2Choices, 0)
        self.m_choice2.SetSelection(0)
        fgSizer1.Add(self.m_choice2, 0, wx.ALL, 5)

        self.m_staticText7 = wx.StaticText(
            self, wx.ID_ANY, "Data:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText7.Wrap(-1)
        fgSizer1.Add(self.m_staticText7, 0, wx.ALL, 5)

        m_choice3Choices = []
        self.m_choice3 = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition,
                                   wx.DefaultSize, m_choice3Choices, 0)
        self.m_choice3.SetSelection(0)
        fgSizer1.Add(self.m_choice3, 0, wx.ALL, 5)

        self.m_staticText8 = wx.StaticText(
            self, wx.ID_ANY, "Taxonomic weighting:", wx.DefaultPosition,
            wx.DefaultSize, 0)
        self.m_staticText8.Wrap(-1)
        fgSizer1.Add(self.m_staticText8, 0, wx.ALL, 5)

        self.m_textCtrl2 = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString,
                                       wx.DefaultPosition, wx.DefaultSize, 0)
        fgSizer1.Add(self.m_textCtrl2, 0, wx.ALL, 5)

        m_radioBox1Choices = ["arithmetic mean", "weighted (by number of data "
                                                 "per species)", "no mean no "
                                                                 "weight (raw "
                                                                 "data)"]
        self.m_radioBox1 = wx.RadioBox(
            self, wx.ID_ANY, "Species weighting:", wx.DefaultPosition,
            wx.DefaultSize, m_radioBox1Choices, 1, wx.RA_SPECIFY_COLS)
        self.m_radioBox1.SetSelection(0)
        fgSizer1.Add(self.m_radioBox1, 0, wx.ALL, 5)

        fgSizer1.AddSpacer(5)

        self.m_staticText9 = wx.StaticText(
            self, wx.ID_ANY, "Distribution law to fit:", wx.DefaultPosition,
            wx.DefaultSize, 0)
        self.m_staticText9.Wrap(-1)
        fgSizer1.Add(self.m_staticText9, 0, wx.ALL, 5)

        m_checkList1Choices = ["normal", "empirical", "triangular"]
        self.m_checkList1 = wx.CheckListBox(
            self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
            m_checkList1Choices, wx.LB_MULTIPLE)
        fgSizer1.Add(self.m_checkList1, 0, wx.ALL, 5)

        self.m_staticText10 = wx.StaticText(
            self, wx.ID_ANY, "Number of bootstrap:", wx.DefaultPosition,
            wx.DefaultSize, 0)
        self.m_staticText10.Wrap(-1)
        fgSizer1.Add(self.m_staticText10, 0, wx.ALL, 5)

        self.m_textCtrl3 = wx.TextCtrl(self, wx.ID_ANY, "1000",
                                       wx.DefaultPosition, wx.DefaultSize, 0)
        fgSizer1.Add(self.m_textCtrl3, 0, wx.ALL, 5)

        self.m_staticText11 = wx.StaticText(
            self, wx.ID_ANY, "Hazen parameter:", wx.DefaultPosition,
            wx.DefaultSize, 0)
        self.m_staticText11.Wrap(-1)
        fgSizer1.Add(self.m_staticText11, 0, wx.ALL, 5)

        self.m_textCtrl4 = wx.TextCtrl(self, wx.ID_ANY, "0.5",
                                       wx.DefaultPosition, wx.DefaultSize, 0)
        fgSizer1.Add(self.m_textCtrl4, 0, wx.ALL, 5)

        self.m_checkBox1 = wx.CheckBox(self, wx.ID_ANY,
                                       "Save intermediate worksheets",
                                       wx.DefaultPosition, wx.DefaultSize, 0)
        fgSizer1.Add(self.m_checkBox1, 0, wx.ALL, 5)

        gSizer1 = wx.GridSizer(1, 2, 0, 0)

        self.m_button2 = wx.Button(self, wx.ID_ANY, "Launch",
                                   wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer1.Add(self.m_button2, 0, wx.ALL, 5)

        self.m_button3 = wx.Button(self, wx.ID_ANY, "Exit", wx.DefaultPosition,
                                   wx.DefaultSize, 0)
        gSizer1.Add(self.m_button3, 0, wx.ALL | wx.ALIGN_RIGHT, 5)

        fgSizer1.Add(gSizer1, 1, wx.EXPAND, 5)

        self.SetSizer(fgSizer1)
        self.Layout()
        fgSizer1.Fit(self)

        self.Centre(wx.BOTH)

        # Connect Events
        self.m_filePicker1.Bind(wx.EVT_FILEPICKER_CHANGED, self.update)
        self.m_checkList1.Bind(wx.EVT_CHECKLISTBOX, self.disabled_checkbox)
        self.m_button3.Bind(wx.EVT_BUTTON, self.exit)
        self.m_button2.Bind(wx.EVT_BUTTON, self.run)

    def __del__(self):
        """Del."""
        pass

    @staticmethod
    def exit(event):
        """Exit the IHM."""
        _ = event
        sysexit(0)

    def disabled_checkbox(self, event):
        """Disable the triangular checkbox."""
        index = event.GetSelection()
        label = self.m_checkList1.GetString(index)
        if label == "triangular":
            self.m_checkList1.Check(index, check=False)

    def run(self, event):
        """Run the program."""
        _ = event
        columns_name = [self.m_choice2.GetString(
            self.m_choice2.GetSelection()), self.m_choice1.GetString(
            self.m_choice1.GetSelection()), self.m_choice3.GetString(
            self.m_choice3.GetSelection())]
        species, taxon, concentration, test = parse_file(self.filename,
                                                         columns_name)
        pcat = list(self.m_textCtrl2.GetLineText(0))
        if not len(pcat):
            pcat = None
        checked_laws = self.m_checkList1.GetCheckedItems()
        normal = 0 in checked_laws
        emp = 1 in checked_laws
        triang = 2 in checked_laws
        bootstrap = int(self.m_textCtrl3.GetLineText(0))
        hazen = float(self.m_textCtrl4.GetLineText(0))
        nbvar = False
        save = self.m_checkBox1.IsChecked()
        lbl_list = None
        adjustq = False
        isp = self.m_radioBox1.GetSelection()
        charger_parametres(
            self.filename, 1, species, taxon, concentration, test,
            pcat, pcat is None, pcat is not None, emp, normal, triang,
            bootstrap, hazen, nbvar, save, lbl_list,
                  triang and adjustq, isp, columns_name)

    def update(self, event):
        """Update IHM on file load."""
        _ = event
        self.filename = self.m_filePicker1.GetPath()

        if splitext(self.filename)[1] in ['.xls', '.xlsx']:
            data = read_excel(self.filename, sheetname=None)
            dlg = wx.SingleChoiceDialog(self, "Sheetsname:",
                                        "Pick the sheet of your file",
                                        list(data.keys()))
            if dlg.ShowModal() == wx.ID_OK:
                self.text_sheet_name.SetLabel(dlg.GetStringSelection())
                data = data[dlg.GetStringSelection()]
            else:
                self.m_filePicker1.SetPath("")
                self.m_staticText3.SetLabel("")
                self.m_choice1.SetItems([])
                self.m_choice2.SetItems([])
                self.m_choice3.SetItems([])
                self.text_sheet_name.SetLabel("")
                wx.MessageBox('Canceled', 'Warning', wx.OK | wx.ICON_WARNING)
                dlg.Destroy()
                return ()
            dlg.Destroy()
        elif splitext(self.filename)[1] == '.csv':
            data = read_csv(self.filename, header=0)
            if len(data.columns) == 1:
                data = read_csv(self.filename, sep=";", header=0)
        else:
            raise IOError("Invalid file")
        self.m_staticText3.SetLabel(
            basename(splitext(self.filename)[0]) + "_sswd.xlsx"
            if self.filename else "")
        self.m_choice1.SetItems(data.columns)
        self.m_choice2.SetItems(data.columns)
        self.m_choice3.SetItems(data.columns)
        self.m_choice1.SetSelection(0)
        self.m_choice2.SetSelection(1)
        self.m_choice3.SetSelection(2)

    @staticmethod
    def shorten(s, n=24):
        """Shorten strings."""
        if len(s) <= n:
            return s
        n_2 = int(n / 2 - 3)
        n_1 = n - n_2 - 3
        return '{0}...{1}'.format(s[:n_1], s[-n_2:])
