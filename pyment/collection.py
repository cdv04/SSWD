# coding=utf-8
"""
Class file.

This file is here to create the typing for Collection.
"""


class Collection:
    """
    The types are structured as follow.

    "" -> String
    0.0 -> Double
    0 -> Integer
    """

    def __repr__(self):
        """Return str of class."""
        return self.espece + "\t" + self.taxo + "\t" + str(self.data)

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
