#!/usr/bin/env python3
# coding=utf-8
"""
File to execute SSWD.

Parse command line and csv file.
"""

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-27T09:39:06+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: __main__.py
# @Last modified by:   gysco
# @Last modified time: 2017-06-01T09:56:50+02:00

import argparse
import sys

import wx
from common import parse_file
from ihm import mainFrame
from ihm_functions import charger_parametres


def isp_type(x):
    """Check isp type."""
    x = str(x)
    to_index = ["w", "u", "m"]
    if x not in to_index:
        raise argparse.ArgumentTypeError(
            "%s has to be: m(ean)/w(eighted)/u(nweighted) for ponderation" %
            (x,))
    return to_index.index(x)


def restricted_float(x):
    """Check float hazen."""
    x = float(x)
    if x < 0.0 or x > 1.0:
        raise argparse.ArgumentTypeError("%r not in range [0.0, 1.0]" % (x,))
    return x


def main():
    """Main."""
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--save", help="Save intermediate operations", action="store_true")
    parser.add_argument(
        "--emp", help="Enable empirical law", action="store_true")
    parser.add_argument(
        "--normal", help="Enable normal law", action="store_true")
    parser.add_argument(
        "--triang", help="Enable triangular law", action="store_true")
    parser.add_argument("--iproc", help="iproc value", default=1, type=int)
    parser.add_argument(
        "--hazen", help="Hazen parameter a", default=.5, type=restricted_float)
    parser.add_argument("-f", "--file", help="filename")
    parser.add_argument("--pcat", help="pcat values or colonne", type=str)
    parser.add_argument(
        "--bootstrap", help="bootstrap n times", default=1000, type=int)
    parser.add_argument("--nbvar", help="enable nbvar", action="store_true")
    parser.add_argument(
        "--adjustq", help="adjust q for triangular law", action="store_true")
    parser.add_argument(
        "--isp", help="poderation type", default=1, type=isp_type)
    parser.add_argument("--lbl_liste", help="weight of each taxonomic group")
    args = parser.parse_args()
    # columns_name = get_columns(args.file)
    columns_name = ["Species", "PhylumSup", "ED"]
    espece, taxo, concentration, test = parse_file(args.file, columns_name)
    charger_parametres(
        args.file, args.iproc, espece, taxo, concentration, test, args.pcat,
        args.pcat is None, args.pcat is not None, args.emp, args.normal,
        args.triang, args.bootstrap, args.hazen, args.nbvar, args.save,
        args.lbl_liste, args.triang and args.adjustq, args.isp, columns_name)
    return 0


if __name__ == '__main__':
    # sys.tracebacklimit = None
    if len(sys.argv) > 1:
        sys.exit(main())
    else:
        app = wx.App(False)
        frame = mainFrame(None)
        frame.Show()
        sys.exit(app.MainLoop())
