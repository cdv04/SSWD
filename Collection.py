"""
Class file.

This file is here to create the typing for Collection.
"""

# !/usr/bin/env python

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-05T09:03:15+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: Collection.py
# @Last modified by:   gysco
# @Last modified time: 2017-04-28T09:51:02+02:00


class Collection():
    """
    The types are structered as follow.

    "" -> String
    0.0 -> Double
    0 -> Integer
    """

    def __repr__(self):
        """Return str of class."""
        return (self.espece + "\t" + self.taxo + "\t" + str(self.data))

    def __init__(self, test="C", pond=1., num=1, pcum=1.):
        """Init method."""
        self.espece = ""
        self.taxo = ""
        self.test = test
        self.data = 0.0
        self.num = num
        self.pond = pond
        self.pcum = pcum
        self.std = 0.0
        self.act = 0.0
        self.pcum_a = 0.0
