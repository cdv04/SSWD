"""Controle de l'IHM."""

# !/usr/bin/env python

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-10T09:11:00+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: Fct_ihm.py
# @Last modified by:   gysco
# @Last modified time: 2017-05-09T16:12:39+02:00

import numpy as np

import Initialisation
from fct_generales import (ischainevide, rech_l1c1, rechercher_categorie,
                           trier_collection, trier_tableau)
from Lancement_sswd import lance
from MsgBox import MsgBox
from ponderation import calcul_nb_taxo
from specific_sswd import filtre_collection_act  # init_collection_act


def recherche_nom_feuille(plage):
    """Recherche nom_feuille contenu dans RefEdit et data_plage."""
    erreur = False
    separateur = '!'
    # If plage = "" Then
    #   MsgBox "Selected the range or the column containing the data!",
    #   0, "SSWD"
    #   erreur = True
    #   Exit Sub
    #   End If
    # ceci est pris en charge par une fonction ischainevide specifique
    ipos = plage.find(separateur)
    if (ipos == 0):
        MsgBox('SSWD',
               'The selected range don\'t contain any worksheet\'s name!', 0)
        erreur = True
        return
    else:
        nom_feuille = plage[1:ipos]
        data_plage = plage[ipos + 1:]
    return (nom_feuille, data_plage, erreur)


def trf_plage_cellule(nom_feuille, plage):
    """Recherche les lignes et colonnes d'une plage de cellules."""
    # erreur = False
    # if Application.ReferenceStyle == xlR1C1:  # TODO modifie en python IHM
    """
    Recherche separateur ":" specifiant la selection d'une plage
    """
    ipos_sep = plage.find(':')
    """Cas d'une selection de colonne et non de plage"""
    if ipos_sep == 0:
        # Recherche indice colonne
        l1, c1 = rech_l1c1(plage, 2)
        c2 = c1
        l2 = l1
        while Initialisation.Worksheets[nom_feuille].Cells[l2, c1]:
            l2 = l2 + 1
        l2 = l2 - 1
        """Cas d'une selection d'une plage"""
    else:
        """
        Recherche indice ligne/ colonne de la premiere cellule de
        la plage
        """
        l1, c1 = rech_l1c1(plage, 2)
        """
        Recherche indice ligne/ colonne de la premiere cellule de
        la plage
        """
        l2, c2 = rech_l1c1(plage, ipos_sep + 2)
    """else:
        # TODO pythoneries
        # plage = Worksheets[nom_feuille].Range(plage).AddressLocal(
        #     ReferenceStyle=xlA1)
        # l1 = Worksheets[nom_feuille].Range(plage).Cells(1, 1).Row
        # c1 = Worksheets[nom_feuille].Range(plage).Cells(1, 1).Column
        i = 1
        i = i + 1
        # while Worksheets[nom_feuille].Range(plage).Cells(i, 1):
        #     i = i + 1
        # l2 = Worksheets[nom_feuille].Range(plage).Cells(i - 1, 1).Row
        # c2 = Worksheets[nom_feuille].Range(plage).Cells(i - 1, 1).Column"""
    return (l1, c1, l2, c2, False)


def lire_pcat(val_pcat, pcat, dim_pcat):
    """
    Lit les valeurs de pcat entrees par l'utilisateur.

    Les ranges sous forme de vecteur.
    """
    erreur = False
    ipos = 0
    debut = 0
    for i in range(0, dim_pcat):
        ipos = val_pcat.find(';')
        if ipos != 0 and i == dim_pcat:
            erreur = True
            MsgBox('SSWD', 'The number of weight values do not corresponds' +
                   ' to the number of taxonomic groups!', 0)
            return
        if ipos == 0:
            ipos = len(val_pcat) + 1
        if ipos - debut <= 0:
            erreur = True
            MsgBox('SSWD',
                   'Please enter a weight for ' + 'every taxonomic group!', 0)
            return
        pcat[i] = float(val_pcat[debut:ipos - debut])
        debut = ipos + 1
        ipos = ipos + 1
    return (erreur)


def afficher_taxo(data_taxo, liste_taxo, erreur):
    """Bouton pcat enter weight values est actionne."""
    """Recherche du nom de la feuille"""
    nom_feuille, plage, erreur = recherche_nom_feuille(data_taxo)
    if erreur:
        return
    """Recherche plage de cellules"""
    l1, c1, l2, c2, erreur = trf_plage_cellule(nom_feuille, plage)
    if erreur:
        return
    """
    Les nom taxo sont charges dans un tableau et tries par ordre
    alphabetique
    """
    # Worksheets[nom_feuille].Activate()
    # tmp = Worksheets[nom_feuille].Range(Cells[l1, c1], Cells[l2, c2])
    tmp = np.copy(Initialisation.Worksheets[nom_feuille].Cells[l1:l2, c1:c2])
    taxo = list()
    for i in range(0, len(taxo)):
        taxo.append(tmp[i + 1, 1])
    trier_tableau(taxo)
    """Extraction des differentes categories taxo"""
    taxo_dif = rechercher_categorie(taxo)  # TODO return taxo_dif
    """Si une seule categorie taxo, pas de ponderation possible"""
    if len(taxo_dif) < 3:
        erreur = True
        MsgBox('SSWD', 'There is only one taxonomic group: \
you cannot enter weight!', 0)
        return
    """Affichage"""
    liste_taxo = ''
    for i in range(1, len(taxo_dif) - 1):
        if (i == len(taxo_dif) - 1):
            liste_taxo += taxo_dif[i]
        else:
            liste_taxo += taxo_dif[i] + ';'


def charger_parametres(iproc, r_espece, r_taxo, r_concentration, r_test, txt_p,
                       opt_bt_nul, opt_bt_val, ch_e, ch_n, ch_t, txt_val_b,
                       txt_val_a, ch_nb, ch_sauve, lbl_liste, opt_bt_q, cbx_e):
    """
    Charge les parametres receuillis par la boite de dialogue.

    Execute SSWD
    """
    data_co = list()
    nomboite = ('SSWD' if iproc == 1 else 'ACT')
    """Chargement de la collection"""
    r_x = [r_espece, r_taxo, r_concentration]
    plage_x = [None, None, None]
    # l1 = [0, 0, 0]
    # c1 = [0, 0, 0]
    # l2 = [0, 0, 0]
    # c2 = [0, 0, 0]
    str_x = [
        "the species or genus names!",
        "the trophic levels or taxonomic groups!", "concentration data!"
    ]
    for i in range(0, len(r_x)):
        data = r_x[i]
        assert (
            ischainevide(data, 'Select the range or the column of ' + str_x[i],
                         nomboite) is False)
        nom_feuille, plage_x[i], erreur = recherche_nom_feuille(data)
        assert (erreur is False)
        # l1[i], c1[i], l2[i], c2[i], erreur = trf_plage_cellule(
        #     nom_feuille, plage_x[i])
        # assert (erreur is False)
    if (iproc == 1):
        """Collection SSWD"""
        # data_co = plage_x.copy()
        delem = ';' if len(plage_x[0].split(";")) > 1 else ','
        Initialisation.init_collection(data_co, plage_x[0].split(delem),
                                       plage_x[1].split(delem),
                                       plage_x[2].split(delem))
        # init_collection(nom_feuille, l1[0] + 1, l1[1] + 1, l1[2] + 1, c1[0],
        #                 c1[1], c1[2], data_co)
    else:
        """Collection ACT"""
        data = r_test
        erreur, plage_test, nom_feuille = recherche_nom_feuille(data)
        assert (erreur is False)
        # l1_test, c1_test, l2_test, c2_test, erreur = trf_plage_cellule(
        #     nom_feuille, plage_test)
        # assert (erreur is False)
        # init_collection_act(nom_feuille, l1[0] + 1, l1[1] + 1, l1[2] + 1,
        #                     l1_test, c1[0], c1[1], c1[2], c1_test, data_co,
        #                     "", "", "", "", "")
        data_co = plage_test.copy()
        filtre_collection_act(data_co, "", "", "", "", 1, 1, erreur)
        # change_nom_taxo(data_co)
    """Titre des colonnes de data_co"""
    nom_colonne = list()
    nom_colonne.append("SpeciesComp")
    nom_colonne.append("PhylumSup")
    nom_colonne.append("ED")
    # nom_colonne.append(Worksheets[nom_feuille].Cells[l1[0], c1[0]])
    # nom_colonne.append(Worksheets[nom_feuille].Cells[l1[1], c1[1]])
    # nom_colonne.append(Worksheets[nom_feuille].Cells[l1[2], c1[2]])
    nom_colonne.append('Weight')
    nom_colonne.append('Weighted Emp. Cumul. Prob.')
    check_nom_colonne(iproc, nom_colonne)
    """Type de pondération espece"""
    isp = (cbx_e + 1) if (iproc == 1) else 1
    """Pcat : pondération taxonomie"""
    trier_collection(data_co, 2, 1)
    nb_taxo = calcul_nb_taxo(data_co)
    pcat = list()
    val_pcat = txt_p
    if opt_bt_nul is True:
        for i in range(1, nb_taxo):
            pcat.append(0)
    else:
        if val_pcat == '':
            MsgBox('SSWD', 'Please enter weight values or select No weight!',
                   0)
            erreur = True
            return
        assert (lire_pcat(val_pcat, pcat, nb_taxo) is False)
    """parametre de Hazen a"""
    if txt_val_a is None:
        MsgBox('SSWD', 'You must chose a value for the Hazen parameter \
between 0 and 1, strictly less than 1', 0)
        erreur = True
        return
    a = float(txt_val_a)
    if a >= 1:
        MsgBox('SSWD'
               'The Hazen parameter must be included between \
0 and 1, strictly less than 1', 0)
        erreur = True
        return
    """Loi statistique 1:empirique, 2:normal, 3:triangulaire"""
    dist = list()
    dist.append(ch_e)
    dist.append(ch_n)
    dist.append(ch_t)
    """B"""
    if txt_val_b == '':
        MsgBox('SSWD',
               'You must chose a value for the number of bootstrap samples', 0)
        erreur = True
        return
    B = int(txt_val_b)
    """nbvar : nombre de données tirées"""
    n_optim = ch_nb
    """Sauvegarde des feuilles de resultats intermediaires"""
    conserv_inter = ch_sauve
    """Recupération de la liste taxo"""
    if opt_bt_val is True:
        liste_taxo = lbl_liste
        pos = liste_taxo.find('\n')
        ltaxo = liste_taxo[:pos + 1]
    else:
        ltaxo = plage_x[1].split(';')
    """
    Option d'ajustement pour loi triangulaire,
    si True ajustement sur quantiles, sinon sur probabilités cumulées
    """
    triang_ajust = opt_bt_q
    lance(data_co, nom_feuille, nom_colonne, isp, pcat, dist, B, a, n_optim,
          conserv_inter, nb_taxo, val_pcat, ltaxo, triang_ajust)


def check_nom_colonne(iproc, nom_colonne):
    """Verifie le iproc et ajoute des nom de colonnes si besoin."""
    if (iproc == 2):
        nom_colonne.append('ACT data')
        nom_colonne.append('Weighted Emp. Cumul. Prob. Acute')
