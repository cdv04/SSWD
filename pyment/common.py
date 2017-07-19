# coding=utf-8
"""
Les fonctions les plus utilisees.

A inclure dans la quasi totatilite des autres.
"""

# !/usr/bin/env python

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-06T09:22:23+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: common.py
# @Last modified by:   gysco
# @Last modified time: 2017-06-19T16:15:43+02:00

import argparse
import operator
import sys
from os.path import abspath, join as pathjoin, splitext

import initialisation
import pandas
from message_box import message_box
from worksheet import Worksheet


def sort_collection(this_collection, item, order):
    """
    Sort the data on 1 of its parameters.

    :param this_collection: collection to be sorted
    :param item: index of the item on which we sort
    :param order: order of the sort (0=down,1=up)
    """
    tmp_list = [
        "espece", "taxo", "test", "data", "num", "pond", "pcum", "std", "act",
        "pcum_a"
    ]
    if item - 1 == 1:
        this_collection.sort(
            key=operator.attrgetter(tmp_list[3]), reverse=True)
    this_collection.sort(
        key=operator.attrgetter(tmp_list[item - 1]),
        reverse=(True if order == 0 else False))


def ecrire_titre(titre, nom_feuille, lig, col):
    """
    Ecrit le titre d'un tableau.

    @param titre: titre du tableau
    @param nom_feuille: nom de la feuille de calcul
    @param lig: numero de la ligne ou ecrire le titre du tableau
    @param col: numero de la colonne ou ecrire le titre du tableau
    @param nbcol: nombre de colonnes du tableau (pour centrer le titre
                  sur toutes les colonnes)
    """
    initialisation.Worksheets[nom_feuille].Cells.set_value(lig, col, titre)


def ecrire_data_co(data_co, nom_colonne, lig, col, nom_feuille, invlog, iproc):
    """
    Ecrit la collection data_co dans une feuille de calcul.

    @param data_co: Nom de la collection
    @param lig: Numero de la premiere ligne du tableau ou est ecrit la
                collection cette premiere ligne correspond aux titres
                des colonnes du tableau
    @param col: Numero de la premiere colonne du tableau a ecrire
    @param nom_colonne: nom des colonnes de la data_co
    @param nom_feuille: nom de la feuille ou est affichee la collection

    Toutes les colonnes ne sont pas necessairement affichees
    """
    nbdata = len(data_co)
    """1. Titre des colonnes"""
    for i in range(0, len(nom_colonne)):
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig, col + i, nom_colonne[i])
    """2. Donnees"""
    for i in range(0, nbdata):
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig + i + 1, col, data_co[i].espece)
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig + i + 1, col + 1, data_co[i].taxo)
        if iproc == 2:
            if invlog is True:
                initialisation.Worksheets[nom_feuille].Cells.set_value(
                    lig + i + 1, col + 5, 10**data_co[i].act)
            else:
                initialisation.Worksheets[nom_feuille].Cells.set_value(
                    lig + i + 1, col + 5, data_co[i].act)
            initialisation.Worksheets[nom_feuille].Cells.set_value(
                lig + i + 1, col + 6, data_co[i].pcum_a)
        if invlog is True:
            initialisation.Worksheets[nom_feuille].Cells.set_value(
                lig + i + 1, col + 2, 10**data_co[i].data)
        else:
            initialisation.Worksheets[nom_feuille].Cells.set_value(
                lig + i + 1, col + 2, data_co[i].data)
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig + i + 1, col + 3, data_co[i].pond)
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig + i + 1, col + 4, data_co[i].pcum)


def verif(nom_feuille_pond, nom_feuille_stat, nom_feuille_res,
          nom_feuille_qemp, nom_feuille_qnorm, nom_feuille_sort,
          nom_feuille_Ftriang, nom_feuille_qtriang, nom_feuille_err_ve,
          nom_feuille_err_inv, nom_feuille_indice):
    """
    Verifie que le nom des feuilles resultats n'existe pas.

    supprime les feuilles intermediaires et cree les feuilles
    nom_feuille_res et nom_feuille_pond
    """
    name_list = [
        nom_feuille_pond, nom_feuille_stat, nom_feuille_res + "_emp",
        nom_feuille_res + "_norm", nom_feuille_res + "_triang",
        nom_feuille_qemp, nom_feuille_qnorm, nom_feuille_sort,
        nom_feuille_Ftriang, nom_feuille_qtriang, nom_feuille_err_ve,
        nom_feuille_err_inv, nom_feuille_indice, "details"
    ]
    for ws in initialisation.Worksheets:
        if ws.Name == nom_feuille_res:
            rep = message_box('Attention...', 'Result\'s worksheet already '
                              'exists!\
                          If you continue, this one will be destroyed.\
                          Would you like to go on?\n\
                          If you want to keep this previous results,\
                          rename the SSWD_result worksheet.', 4)
            if rep == 7 or not rep:
                sys.exit(0)
            else:
                del initialisation.Worksheets[ws.Name]
        else:
            if ws.Name in name_list:
                del initialisation.Worksheets[ws.Name]
    for name_str in name_list:
        initialisation.Worksheets[name_str] = Worksheet(name=name_str)


def minimum_tab_dif0(a):
    """Renvoie la valeur minimum <> 0 d'un tableau de reels."""
    """Recherche d'une valeur non-nulle dans le tableau"""
    i = 0
    for i in range(0, len(a)):
        if a[i] != 0:
            break
    _ret = a[i]
    for i in range(0, len(a)):
        if a[i] < _ret and a(i) != 0:
            _ret = a(i)
    return _ret


def calcul_lig_graph(lig_deb):
    """
    Calcul les indices de lignes pour les graphes dans nom_feuille_res.

    Compte tenu de la disposition du tableau de resultats HC

    @param lig_deb: ligne de debut de l'affichage du tableau de
                    resultats HC dans nom_feuille_res,
                    il s'agit d'une ligne de titre

    Attention ceci depend des choix d'affichage dans calculer_res
    """
    lig_p = lig_deb + 1
    lig_qbe = lig_deb + 2
    lig_qbi = lig_deb + 5
    lig_qbs = lig_deb + 6
    return lig_p, lig_qbe, lig_qbi, lig_qbs


def affichage_options(nom_feuille, isp, val_pcat, liste_taxo, B, lig, col,
                      lig_s, col_s, dist, nbvar, iproc, a):
    """
    Affichage options choisies par l'utilisateur dans nom_feuille_res.

    @param isp: option de traitement de l'information espece
    @param Pcat: vecteur de poids accordes a chaque categorie taxonomique
    @param b: nombre de runs du bootstrap
    @param lig: premiere ligne d'affichage options
    @param col: premiere colonne d'affichage options
    @param lig_s: premiere ligne d'affichage sigles
    @param col_s: premiere colonne d'affichage sigles
    """
    nbcol = 1
    initialisation.Worksheets[nom_feuille].Cells.set_value(lig, col, 'Options')
    """Option espece"""
    initialisation.Worksheets[nom_feuille].Cells.set_value(
        lig + 1, col, 'Species=')
    initialisation.Worksheets[nom_feuille].Cells.set_value(
        lig + 1, col + 1, sp_opt(isp))
    """Option pcat"""
    initialisation.Worksheets[nom_feuille].Cells.set_value(
        lig + 2, col, 'Taxonomy')
    if liste_taxo is not None:
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig + 2, col + 1, str(liste_taxo))
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig + 2, col + 2, str(val_pcat))
    else:
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig + 2, col + 1, 'No Weight')
    """nbruns B"""
    initialisation.Worksheets[nom_feuille].Cells.set_value(
        lig + 3, col, 'Nb bootstrap samples')
    initialisation.Worksheets[nom_feuille].Cells.set_value(
        lig + 3, col + nbcol + 1, B)
    """nbvar"""
    initialisation.Worksheets[nom_feuille].Cells.set_value(
        lig + 4, col, 'Nb data')
    initialisation.Worksheets[nom_feuille].Cells.set_value(
        lig + 4, col + nbcol + 1, nbvar)
    """parametre de Hazen : a"""
    initialisation.Worksheets[nom_feuille].Cells.set_value(
        lig + 5, col, 'Hazen parameter a')
    initialisation.Worksheets[nom_feuille].Cells.set_value(
        lig + 5, col + nbcol + 1, a)
    """Sigles=acronyms"""
    initialisation.Worksheets[nom_feuille].Cells.set_value(
        lig_s, col_s, 'SSWD=Species Sensitivity Weighted Distribution')
    if iproc == 2:
        i = 1
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig_s + i, col_s, 'ACT=Acute to Chronic Transformation')
    else:
        i = 0
    initialisation.Worksheets[nom_feuille].Cells.set_value(
        lig_s + i + 1, col_s, 'HC=Hazardous Concentration')
    initialisation.Worksheets[nom_feuille].Cells.set_value(
        lig_s + i + 2, col_s, 'Sp=Species')
    initialisation.Worksheets[nom_feuille].Cells.set_value(
        lig_s + i + 3, col_s, 'TW=Taxonomic or Trophical Weights')
    i += 4
    if dist[1] is True or dist[2] is True:
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig_s + i, col_s,
            'R_=Multiple R-square on the empirical quantiles')
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig_s + i + 1, col_s,
            'KSpvalue=pvalue of the Kolmogorov-Smirnov goodness of fit ' +
            'test (with Dallal-Wilkinson approximation)')
        i += 2
    if dist[1] is True:
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig_s + i, col_s,
            'GWM=Geometric Weighted Mean of the log-normal distribution')
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig_s + i + 1, col_s,
            'GWSD=Geometric Weighted Standard Deviation of the log-normal' +
            'distribution')
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig_s + i + 2, col_s,
            'wm.lg=Weighted Mean of the log-normal distribution of the data' +
            ' (log10)')
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig_s + i + 3, col_s,
            'wsd.lg=Weighted Standard Deviation of the log-normal' +
            'distribution of the data (log10)')
        i += 4
    if dist[2] is True:
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig_s + i, col_s,
            'GWMin=Geometric Min parameter of the Weighted log-triangular' +
            ' distribution')
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig_s + i + 1, col_s,
            'GWMax=Geometric Max parameter of the Weighted log-triangular' +
            ' distribution')
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig_s + i + 2, col_s,
            'GWMode=Geometric Mode parameter of the Weighted log-triangular' +
            ' distribution')
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig_s + i + 3, col_s,
            'wmin.lg=Min parameter of the Weighted log-triangular' +
            ' distribution (log10)')
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig_s + i + 4, col_s,
            'wmax.lg=Max parameter of the Weighted log-triangular' +
            ' distribution (log10)')
        initialisation.Worksheets[nom_feuille].Cells.set_value(
            lig_s + i + 5, col_s,
            'wmode.lg=Mode parameter of the Weighted log-triangular' +
            ' distribution (log10)')


def calcul_col_res(c_hc, nbcol_vide, pourcent, dist, ind_tax, ind_data,
                   ind_pcum, nom_colonne, ind_data_act, ind_pcum_a):
    """
    Calcul des indices de colonnes pour l'affichage des resultats.

    @param col_deb: premiere colonne affichage resultats numeriques HC
    @param col_fin: derniere colonne affichage resultats numeriques HC
    @param col_data1: premiere colonne d'affichage de data_co_feuil
                      triee selon la taxonomie
    @param col_data2: premiere colonne d'affichage de data_co_feuil
                      triee selon les concentrations croissantes
    @param col_tax: colonne d'affichage information taxonomique
    @param col_data: colonne d'affichage concentration
    @param col_pcum: colonne d'affichage probabilite cumulees
                     empiriques ponderees
    @param col_data_le: colonne d'affichage concentration pour graphe
                        distribution empirique
    @param col_pcum_le: colonne d'affichage probabilite pour graphe
                        distribution empirique
    @param c_hc: premiere colonne d'affichage resultats
    @param nbcol_vide: nombre de colonne vide entre resultats
    @param pourcent: liste des pourcentages de HCx% affiches
    @param dist: vecteur de booleens definissant les lois de
                 distribution affichees
    @param ind_tax: indice de la colonne d'information taxonomique dans
                    data_co_feuil
    @param ind_data: indice de la colonne concentration dans
                     data_co_feuil
    @param ind_pcmu: indice de la colonne probabilite cumulee ponderee
                     empirique dans data_co_feuil
    """
    # l'affichage des parametres mu et sig et/ou min, max, mode
    nbcol = 4 if dist[2] else 3
    col_data1 = c_hc + len(pourcent) + nbcol_vide + nbcol + 1
    col_data2 = col_data1 + len(nom_colonne) + nbcol_vide
    col_deb = c_hc + 1
    col_fin = c_hc + len(pourcent)
    col_tax = col_data1 + ind_tax - 1
    col_data = col_data1 + ind_data - 1
    col_pcum = col_data1 + ind_pcum - 1
    col_pcum_a = col_data1 + ind_pcum_a - 1
    col_data_act = col_data1 + ind_data_act - 1
    col_data_le = col_data2 + ind_data - 1
    col_pcum_le = col_data2 + ind_pcum - 1
    col_data_act_le = col_data2 + ind_data_act - 1
    return (col_deb, col_fin, col_data1, col_data2, col_tax, col_data,
            col_pcum, col_data_le, col_pcum_le, col_data_act, col_data_act_le,
            col_pcum_a)


def calcul_ref_pond(col_deb, l1, ind_data, ind_pond, ind_pcum, nbdata,
                    ind_data_act):
    """
    Calcul les indices lignes et colonnes dans nom_feuille_pond.

    @param col_deb: premiere colonne de donnees
    @param col_data: colonne des donnees data
    @param col_pond: colonne des ponderations
    @param col_pcum: colonne des probabilites cumulees ponderees
                     empiriques
    @param l1: premiere ligne de donnees
    @param lig_deb: premiere ligne de donnees numeriques
    @param lig_fin: derniere ligne de donnees numeriques
    @param ind_data: indice de la colonne data dans data_co_feuil
    @param ind_pond: indice de la colonne ponderation dans
                     data_co_feuil
    @param ind_pcum: indice de la colonne probabilite cumulee dans
                     data_co_feuil
    """
    lig_deb = l1 + 1
    lig_fin = lig_deb + nbdata - 1
    col_data = col_deb + ind_data - 1
    col_pond = col_deb + ind_pond - 1
    col_pcum = col_deb + ind_pcum - 1
    col_data_act = col_deb + ind_data_act - 1
    return lig_deb, lig_fin, col_data, col_pond, col_pcum, col_data_act


def efface_feuil_inter(nom_feuille_pond, nom_feuille_stat, nom_feuille_qemp,
                       nom_feuille_qnorm, nom_feuille_qtriang,
                       nom_feuille_sort, nom_feuille_Ftriang,
                       nom_feuille_err_ve, nom_feuille_err_inv,
                       nom_feuille_indice):
    """Efface les feuilles de calcul intermediaires si voulu."""
    name_list = [
        nom_feuille_pond, nom_feuille_stat, nom_feuille_qemp,
        nom_feuille_qnorm, nom_feuille_qtriang, nom_feuille_sort,
        nom_feuille_Ftriang, nom_feuille_err_ve, nom_feuille_err_inv,
        nom_feuille_indice
    ]
    for name in name_list:
        if name in initialisation.Worksheets:
            del initialisation.Worksheets[name]


def write_feuil_inter(writer, empty=False):
    """Save worksheet to excel files."""
    for x in initialisation.Worksheets:
        if (len(initialisation.Worksheets[x].Cells.columns) or empty) and x:
            df = initialisation.Worksheets[x].Cells.sort_index(
                axis=1).reindex_axis(
                    range(
                        0,
                        (initialisation.Worksheets[x].Cells.columns.max() + 1)
                        if not empty else 1),
                    axis=1)
            if (len(initialisation.Worksheets[x].Cells.index) <
                    initialisation.Worksheets[x].Cells.index.max()):
                df = df.sort_index(axis=0).reindex_axis(
                    range(0,
                          (initialisation.Worksheets[x].Cells.index.max() + 1)
                          if not empty else 1),
                    axis=0)
            df.to_excel(writer, sheet_name=x, index=False, header=False)
    writer.save()


def trier_tableau(a):
    """Trie un tableau de strings par ordre alphabetique."""
    tmp = list()
    for i in range(0, len(a)):
        tmp.append(a[i])
    num = 0
    for i in range(0, len(tmp)):
        mini = 'Z'
        for j in range(0, len(tmp)):
            if mini.upper > tmp[j].upper:
                mini = tmp[j]
                num = j
        del tmp[num]
        a[i] = mini


def rechercher_categorie(a):
    """
    Recherche les valeurs differentes dans un tableau de strings.

    @param a: tableau de strings
    @param diff: est le vecteur des chaines differentes

    !!! Le tableau A doit etre trie dans l'ordre alphabetique !!!
    """
    nb = 0
    diff = list()
    diff.append(0)
    for i in range(0, len(a)):
        if a(i) != a(i + 1):
            diff.insert(nb, a[i])
            diff.append(a[i + 1])
            nb += 1
    if nb == 0:
        diff = list(a[0])
    return diff


def isnumeric(code):
    """
    Indique si un code ascii correspond a un nombre ou non.

    (on autorise la virgule et le point)
    """
    if code < 48 or code > 57:
        if code != 44 and code != 46:
            _ret = False
        else:
            _ret = True
    else:
        _ret = True
    return _ret


def isentier(code):
    """Indique si un code ascii correspond a un entier."""
    if code < 48 or code > 57:
        _ret = False
    else:
        _ret = True
    return _ret


def cellule_gras():
    """Met le contenu d'une cellule feuille de calcul en gras."""
    # Range[Cells(l1, c1), Cells(l2, c2)].Font.Bold = True


def compt_inf(ech, ind):
    """Nombre de valeur d'un ech >= a un nombre ind."""
    _ret = 0
    for i in range(0, len(ech)):
        if ech[i] <= ind:
            _ret += 1
    return _ret


# def trier_tableau_num(a):
#     """Trie un tableau de nombres."""
#     tmp = list()
#     for i in range(0, len(a)):
#         tmp.append(a[i])
#     maxi = max(a)
#     num = 0
#     for i in range(0, len(a)):
#         mini = maxi
#         for j in range(1, len(tmp)):
#             if (mini > tmp[j]):
#                 mini = tmp[j]
#                 num = j
#         del tmp[num]
#         a[i] = mini
#     tmp = None


def rech_l1c1(_str, deb_str):
    """
    Recherche indice colonne +/- ligne dans une reference du type L1C1.

    Independamment de la langue utilisateur
    """
    return len(_str.split(";")), deb_str


def trier_tirages_feuille(nom_feuille_stat, nom_feuille_sort, nbvar):
    """
    Permet de trier les tirages aleatoires de nom_feuille_stat.

    sauvegarder tries dans une nouvelle dans nom_feuille_sort
    """
    for i in range(0, nbvar):
        initialisation.Worksheets[nom_feuille_sort].Cells.set_value(
            0, i, "RANK" + str(i + 1))
    for x in range(1, len(initialisation.Worksheets[nom_feuille_stat].Cells)):
        i = 0
        for e in sorted(
                initialisation.Worksheets[nom_feuille_stat].Cells.ix[x, :]):
            initialisation.Worksheets[nom_feuille_sort].Cells.set_value(
                x, i, e)
            i += 1


def ischainevide(texte, message, nomboite, erreur=False):
    """Test si la chain de texte est vide, error si tel est le cas."""
    if texte == '':
        message_box(nomboite, message, 0)
        erreur = True
    return erreur


def sp_opt(isp):
    """Added beacause unavailabe to get from IHM actually."""
    return "weighted" if isp == 1 else ("unweighted" if isp == 2 else "mean")


def parse_file(filename, colnames, sheetname):
    """Parse file using pandas module."""
    if not filename:
        raise argparse.ArgumentTypeError("Filename (-f) is None.")
    if splitext(filename)[1] in ['.xls', '.xlsx']:
        data = pandas.read_excel(filename, sheetname=None)
        data = data[sheetname]
    elif splitext(filename)[1] == '.csv':
        data = pandas.read_csv(filename, header=0)
        if len(data.columns) == 1:
            data = pandas.read_csv(filename, sep=";", header=0)
    else:
        raise IOError("Invalid file")
    espece = filename + "!"
    test = filename + "!"
    for n in data[colnames[0]].tolist():
        espece += str(n) + ";"
        test += str(n) + ";"
    taxo = filename + "!"
    for n in data[colnames[1]].tolist():
        taxo += str(n) + ";"
    concentration = filename + "!"
    for n in data[colnames[2]].tolist():
        concentration += str(n) + ";"
    return espece, taxo, concentration, test


def get_columns(filename):
    """Get columns for further parsing."""
    names = list()
    if not filename:
        raise argparse.ArgumentTypeError("Filename (-f) is None.")
    data = pandas.read_csv(filename, header=0)
    if len(data.columns) == 1:
        data = pandas.read_csv(filename, sep=";", header=0)
    for x in data.columns:
        names.append(str(x))
    return names


def rsrc_path(relative_path):
    """For HMI ressource essentially."""
    if hasattr(sys, '_MEIPASS'):
        return pathjoin(sys._MEIPASS, relative_path)
    return pathjoin(abspath('./rsrc/img/'), relative_path)
