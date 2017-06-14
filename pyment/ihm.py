# coding=utf-8
"""IHM file."""

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-05-29T12:32:17+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: ihm.py
# @Last modified by:   gysco
# @Last modified time: 2017-06-01T16:22:31+02:00

from os.path import basename, dirname, join, splitext
from sys import exit as sysexit

import wx
from common import parse_file
from ihm_functions import charger_parametres
from pandas import read_csv, read_excel


class mainFrame(wx.Frame):
    """IHM class."""

    def __init__(self, parent):
        """__init__ of IHM."""
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title="PyME[N]T-SSWD",
                          pos=wx.DefaultPosition, size=wx.Size(-1, -1),
                          style=(wx.CAPTION | wx.CLOSE_BOX | wx.MINIMIZE_BOX |
                                 wx.TAB_TRAVERSAL))

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.filename = None
        sizer_frame = wx.FlexGridSizer(12, 1, 0, 0)
        sizer_frame.SetFlexibleDirection(wx.BOTH)
        sizer_frame.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        box_file = wx.StaticBoxSizer(
            wx.StaticBox(self, wx.ID_ANY, "File management"), wx.VERTICAL)

        sizer_file = wx.FlexGridSizer(3, 2, 0, 0)
        sizer_file.SetFlexibleDirection(wx.BOTH)
        sizer_file.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        self.text_input = wx.StaticText(self, wx.ID_ANY, "Input file:",
                                        wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_input.Wrap(-1)
        sizer_file.Add(self.text_input, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.filepicker = wx.FilePickerCtrl(self, wx.ID_ANY, wx.EmptyString,
                                            "Select a file",
                                            "Data files (*.csv, *.xls, "
                                            "*.xlsx)|*.csv;*.xls;*.xlsx",
                                            wx.DefaultPosition, wx.DefaultSize,
                                            wx.FLP_DEFAULT_STYLE)
        sizer_file.Add(self.filepicker, 0, wx.EXPAND | wx.ALIGN_CENTER_VERTICAL,
                       5)

        self.txt_sheet = wx.StaticText(self, wx.ID_ANY, "Selected sheet:",
                                       wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_sheet.Wrap(-1)
        sizer_file.Add(self.txt_sheet, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.txt_sheet_name = wx.StaticText(self, wx.ID_ANY, wx.EmptyString,
                                            wx.DefaultPosition, wx.DefaultSize,
                                            0)
        self.txt_sheet_name.Wrap(-1)
        sizer_file.Add(self.txt_sheet_name, 0,
                       wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.txt_output = wx.StaticText(self, wx.ID_ANY, "Output file:",
                                        wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_output.Wrap(-1)
        sizer_file.Add(self.txt_output, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.txt_output_name = wx.StaticText(self, wx.ID_ANY, wx.EmptyString,
                                             wx.DefaultPosition, wx.DefaultSize,
                                             0)
        self.txt_output_name.Wrap(-1)
        sizer_file.Add(self.txt_output_name, 0,
                       wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        box_file.Add(sizer_file, 1, wx.EXPAND, 5)

        sizer_frame.Add(box_file, 1, wx.ALL | wx.EXPAND, 5)

        box_data = wx.StaticBoxSizer(
            wx.StaticBox(self, wx.ID_ANY, "Data selection"), wx.VERTICAL)

        sizer_data = wx.FlexGridSizer(2, 1, 0, 0)
        sizer_data.SetFlexibleDirection(wx.BOTH)
        sizer_data.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        sizer_columns = wx.GridSizer(2, 3, 0, 0)

        self.txt_taxo = wx.StaticText(self, wx.ID_ANY,
                                      "Taxonomic groups column:",
                                      wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_taxo.Wrap(-1)
        sizer_columns.Add(self.txt_taxo, 0,
                          wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL,
                          5)

        self.txt_species = wx.StaticText(self, wx.ID_ANY, "Species column:",
                                         wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_species.Wrap(-1)
        sizer_columns.Add(self.txt_species, 0,
                          wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL,
                          5)

        self.txt_data = wx.StaticText(self, wx.ID_ANY, "Effect Dose column:",
                                      wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_data.Wrap(-1)
        sizer_columns.Add(self.txt_data, 0,
                          wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL,
                          5)

        choice_taxoChoices = []
        self.choice_taxo = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition,
                                     wx.DefaultSize, choice_taxoChoices, 0)
        self.choice_taxo.SetSelection(0)
        sizer_columns.Add(self.choice_taxo, 0,
                          wx.ALIGN_CENTER_HORIZONTAL |
                          wx.ALIGN_CENTER_VERTICAL | wx.EXPAND | wx.RIGHT |
                          wx.LEFT,
                          5)

        choice_specieChoices = []
        self.choice_specie = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition,
                                       wx.DefaultSize, choice_specieChoices, 0)
        self.choice_specie.SetSelection(0)
        sizer_columns.Add(self.choice_specie, 0,
                          wx.ALIGN_CENTER_HORIZONTAL |
                          wx.ALIGN_CENTER_VERTICAL | wx.EXPAND | wx.RIGHT |
                          wx.LEFT,
                          5)

        choice_edChoices = []
        self.choice_ed = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition,
                                   wx.DefaultSize, choice_edChoices, 0)
        self.choice_ed.SetSelection(0)
        sizer_columns.Add(self.choice_ed, 0,
                          wx.ALIGN_CENTER_HORIZONTAL |
                          wx.ALIGN_CENTER_VERTICAL | wx.EXPAND | wx.RIGHT |
                          wx.LEFT,
                          5)

        sizer_data.Add(sizer_columns, 1, wx.EXPAND, 5)

        sizer_taxo = wx.GridSizer(1, 2, 0, 0)

        self.txt_taxo_weight = wx.StaticText(self, wx.ID_ANY,
                                             "Taxonomic group weight:",
                                             wx.DefaultPosition, wx.DefaultSize,
                                             0)
        self.txt_taxo_weight.Wrap(-1)
        sizer_taxo.Add(self.txt_taxo_weight, 0, wx.ALL, 5)

        self.txtc_taxo_weight = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString,
                                            wx.DefaultPosition, wx.DefaultSize,
                                            0)
        sizer_taxo.Add(self.txtc_taxo_weight, 0, wx.ALL | wx.EXPAND, 5)

        sizer_data.Add(sizer_taxo, 1, wx.EXPAND, 5)

        box_data.Add(sizer_data, 1, wx.EXPAND, 5)

        sizer_frame.Add(box_data, 1, wx.EXPAND | wx.ALL, 5)

        box_math = wx.StaticBoxSizer(
            wx.StaticBox(self, wx.ID_ANY, "Calculation parameters"),
            wx.VERTICAL)

        sizer_math = wx.FlexGridSizer(1, 2, 0, 0)
        sizer_math.SetFlexibleDirection(wx.BOTH)
        sizer_math.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        radiobox_weightChoices = ["arithmetic mean",
                                  "weighted (by number of data per species)",
                                  "no mean, no weight (raw data)"]
        self.radiobox_weight = wx.RadioBox(self, wx.ID_ANY,
                                           "Species weighting:",
                                           wx.DefaultPosition, wx.DefaultSize,
                                           radiobox_weightChoices, 1,
                                           wx.RA_SPECIFY_COLS)
        self.radiobox_weight.SetSelection(0)
        sizer_math.Add(self.radiobox_weight, 0, wx.RIGHT | wx.LEFT, 5)

        sizer_law = wx.StaticBoxSizer(
            wx.StaticBox(self, wx.ID_ANY, "Distribution law to fit:"),
            wx.VERTICAL)

        self.checkbox_emp = wx.CheckBox(self, wx.ID_ANY, "log-empirical",
                                        wx.DefaultPosition, wx.DefaultSize, 0)
        sizer_law.Add(self.checkbox_emp, 0, 0, 5)

        self.checkbox_norm = wx.CheckBox(self, wx.ID_ANY, "log-normal",
                                         wx.DefaultPosition, wx.DefaultSize, 0)
        sizer_law.Add(self.checkbox_norm, 0, 0, 5)

        sizer_triang = wx.GridSizer(2, 2, 0, 0)

        self.checkbox_triang = wx.CheckBox(self, wx.ID_ANY, "log-triangular",
                                           wx.DefaultPosition, wx.DefaultSize,
                                           0)
        self.checkbox_triang.Enable(False)
        sizer_triang.Add(self.checkbox_triang, 0, 0, 5)

        self.radio_quant = wx.RadioButton(self, wx.ID_ANY, "Quant. fitting",
                                          wx.DefaultPosition, wx.DefaultSize, 0)
        self.radio_quant.Enable(False)
        self.radio_quant.SetValue(True)
        sizer_triang.Add(self.radio_quant, 0, 0, 5)

        sizer_triang.AddSpacer(5)

        self.radio_prob = wx.RadioButton(self, wx.ID_ANY, "Prob. fitting",
                                         wx.DefaultPosition, wx.DefaultSize, 0)
        self.radio_prob.Enable(False)
        sizer_triang.Add(self.radio_prob, 0, 0, 5)

        sizer_law.Add(sizer_triang, 1, wx.EXPAND, 5)

        sizer_math.Add(sizer_law, 1, wx.RIGHT | wx.LEFT | wx.EXPAND, 5)

        box_math.Add(sizer_math, 1, wx.EXPAND, 5)

        sizer_frame.Add(box_math, 1, wx.EXPAND | wx.ALL, 5)

        box_option = wx.StaticBoxSizer(
            wx.StaticBox(self, wx.ID_ANY, "Advanced parameters"), wx.VERTICAL)

        sizer_option = wx.FlexGridSizer(3, 3, 0, 0)
        sizer_option.SetFlexibleDirection(wx.BOTH)
        sizer_option.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        self.txt_bootstrap = wx.StaticText(self, wx.ID_ANY,
                                           "Number of bootstrap:",
                                           wx.DefaultPosition, wx.DefaultSize,
                                           0)
        self.txt_bootstrap.Wrap(-1)
        sizer_option.Add(self.txt_bootstrap, 0,
                         wx.ALIGN_RIGHT | wx.RIGHT | wx.LEFT, 5)

        self.txtc_bootstrap = wx.TextCtrl(self, wx.ID_ANY, "1000",
                                          wx.DefaultPosition, wx.DefaultSize, 0)
        sizer_option.Add(self.txtc_bootstrap, 0, wx.RIGHT | wx.LEFT, 5)

        self.checkbox_nbvar = wx.CheckBox(self, wx.ID_ANY,
                                          "Optimize bootstrap sample size",
                                          wx.DefaultPosition, wx.DefaultSize, 0)
        sizer_option.Add(self.checkbox_nbvar, 0, wx.RIGHT | wx.LEFT, 5)

        box_option.Add(sizer_option, 1, wx.EXPAND, 5)

        radiobox_seedChoices = ["Fix (seed=42)", "Random"]
        self.radiobox_seed = wx.RadioBox(self, wx.ID_ANY,
                                         "Bootstrap seed type",
                                         wx.DefaultPosition, wx.DefaultSize,
                                         radiobox_seedChoices, 2,
                                         wx.RA_SPECIFY_COLS)
        self.radiobox_seed.SetSelection(0)
        box_option.Add(self.radiobox_seed, 0, wx.RIGHT | wx.LEFT, 5)

        fgSizer6 = wx.FlexGridSizer(1, 2, 0, 0)
        fgSizer6.SetFlexibleDirection(wx.BOTH)
        fgSizer6.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        self.txt_hazen = wx.StaticText(self, wx.ID_ANY, "Hazen parameter:",
                                       wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_hazen.Wrap(-1)
        fgSizer6.Add(self.txt_hazen, 0, wx.ALIGN_RIGHT | wx.RIGHT | wx.LEFT, 5)

        self.txtc_hazen = wx.TextCtrl(self, wx.ID_ANY, "0.5",
                                      wx.DefaultPosition, wx.DefaultSize, 0)
        fgSizer6.Add(self.txtc_hazen, 0, wx.RIGHT | wx.LEFT, 5)

        box_option.Add(fgSizer6, 1, wx.EXPAND, 5)

        sizer_frame.Add(box_option, 1, wx.EXPAND | wx.ALL, 5)

        sizer_end = wx.GridSizer(1, 2, 0, 0)

        self.checkbox_save = wx.CheckBox(self, wx.ID_ANY,
                                         "Save intermediate calculation sheets",
                                         wx.DefaultPosition, wx.DefaultSize, 0)
        sizer_end.Add(self.checkbox_save, 0, wx.ALL, 5)

        sizer_buttons = wx.GridSizer(1, 2, 0, 0)

        self.button_launch = wx.Button(self, wx.ID_ANY, "Launch",
                                       wx.DefaultPosition, wx.DefaultSize, 0)
        sizer_buttons.Add(self.button_launch, 0,
                          wx.ALL | wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL |
                          wx.ALIGN_CENTER_VERTICAL,
                          5)

        self.button_end = wx.Button(self, wx.ID_ANY, "Exit",
                                    wx.DefaultPosition, wx.DefaultSize, 0)
        sizer_buttons.Add(self.button_end, 0,
                          wx.ALL | wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL |
                          wx.ALIGN_CENTER_VERTICAL,
                          5)

        sizer_end.Add(sizer_buttons, 1, wx.EXPAND, 5)

        sizer_frame.Add(sizer_end, 1, wx.EXPAND, 5)

        self.SetSizer(sizer_frame)
        self.Layout()
        sizer_frame.Fit(self)

        self.Centre(wx.BOTH)
        # Connect Events
        self.filepicker.Bind(wx.EVT_FILEPICKER_CHANGED, self.update)
        self.checkbox_triang.Bind(wx.EVT_CHECKBOX, self.enable_radios)
        self.radio_prob.Bind(wx.EVT_RADIOBUTTON, self.change_state)
        self.radio_quant.Bind(wx.EVT_RADIOBUTTON, self.change_state)
        self.button_end.Bind(wx.EVT_BUTTON, self.exit)
        self.button_launch.Bind(wx.EVT_BUTTON, self.run)

    def __del__(self):
        """Del."""
        pass

    @staticmethod
    def exit(event):
        """Exit the IHM."""
        _ = event
        sysexit(0)

    def enable_radios(self, event):
        """Disable the triangular checkbox."""
        _ = event
        self.radio_prob.Enable(event.IsChecked())
        self.radio_quant.Enable(event.IsChecked())

    def change_state(self, event):
        """Change the state of the other radiobutton."""
        _ = event
        if self.radio_quant.GetValue():
            self.radio_prob.SetValue(self.radio_quant.GetValue())
            self.radio_quant.SetValue(not self.radio_quant.GetValue())
        else:
            self.radio_quant.SetValue(self.radio_prob.GetValue())
            self.radio_prob.SetValue(not self.radio_prob.GetValue())

    def run(self, event):
        """Run the program."""
        _ = event
        columns_name = [self.choice_specie.GetString(
            self.choice_specie.GetSelection()), self.choice_taxo.GetString(
            self.choice_taxo.GetSelection()), self.choice_ed.GetString(
            self.choice_ed.GetSelection())]
        output = join(dirname(self.filename), self.txt_output_name.GetLabel())
        species, taxon, concentration, test = parse_file(self.filename,
                                                         columns_name,
                                                         self.txt_sheet_name.GetLabel())
        pcat = list(self.txtc_taxo_weight.GetLineText(0))
        if not len(pcat):
            pcat = None
        normal = self.checkbox_norm.IsChecked()
        emp = self.checkbox_emp.IsChecked()
        triang = self.checkbox_triang.IsChecked()
        bootstrap = int(self.txtc_bootstrap.GetLineText(0))
        hazen = float(self.txtc_hazen.GetLineText(0))
        nbvar = self.checkbox_nbvar
        save = self.checkbox_save.IsChecked()
        lbl_list = None
        adjustq = self.radio_quant.GetValue()
        isp = self.radiobox_weight.GetSelection()
        seed = (self.radiobox_seed.GetSelection() == 0)
        charger_parametres(
            self.filename, output , 1, species, taxon,
            concentration, test, pcat, pcat is None, pcat is not None, emp,
            normal, triang, bootstrap, hazen, nbvar, save, lbl_list,
            triang and adjustq, isp, columns_name, seed)

    def update(self, event):
        """Update IHM on file load."""
        _ = event
        self.filename = self.filepicker.GetPath()

        if splitext(self.filename)[1] in ['.xls', '.xlsx']:
            data = read_excel(self.filename, sheetname=None)
            dlg = wx.SingleChoiceDialog(self, "Sheetsname:",
                                        "Pick the sheet of your file",
                                        list(data.keys()))
            if dlg.ShowModal() == wx.ID_OK:
                self.txt_sheet_name.SetLabel(dlg.GetStringSelection())
                data = data[dlg.GetStringSelection()]
            else:
                self.filepicker.SetPath("")
                self.txt_output_name.SetLabel("")
                self.choice_taxo.SetItems([])
                self.choice_specie.SetItems([])
                self.choice_ed.SetItems([])
                self.txt_sheet_name.SetLabel("")
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
        self.txt_output_name.SetLabel(
            basename(splitext(self.filename)[0] + (("_"
                     + self.txt_sheet_name.GetLabel()) if self.txt_sheet_name.GetLabel() else "" ) + "_sswd.xlsx")
            if self.filename else "")
        self.choice_taxo.SetItems(data.columns)
        self.choice_specie.SetItems(data.columns)
        self.choice_ed.SetItems(data.columns)
        self.choice_taxo.SetSelection(self.choice_taxo.FindString(
            "PhylumSup"))
        self.choice_specie.SetSelection(self.choice_specie.FindString(
            "Species"))
        self.choice_ed.SetSelection(self.choice_ed.FindString(
            "ED"))

    @staticmethod
    def shorten(s, n=40):
        """Shorten strings."""
        if len(s) <= n:
            return s
        n_2 = int(n / 2 - 3)
        n_1 = n - n_2 - 3
        return '{0}...{1}'.format(s[:n_1], s[-n_2:])

    def progress(self, species):
        """Progress dialog."""
        datanb = len(species.split('!')[1].split(';'))
        i = ((1 if self.checkbox_norm.IsChecked() else 0)
             + (1 if self.checkbox_emp.IsChecked() else 0)
             + (1 if self.checkbox_triang.IsChecked() else 0))
        maximum = int(
            self.txtc_bootstrap.GetLineText(0)) * datanb * i + datanb * 100
        from tqdm import tqdm
        for x in tqdm(range(0, maximum)):
            wx.Sleep(.1)
