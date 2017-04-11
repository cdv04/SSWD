"""
Controle de l'IHM.

Probablement inutilisable.
"""

# !/usr/bin/env python

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-10T09:11:00+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: Fct_ihm.py
# @Last modified by:   gysco
# @Last modified time: 2017-04-11T15:53:48+02:00

from fct_generales import (ischainevide, rech_l1c1, rechercher_categorie,
                           trier_collection, trier_tableau)
from Initialisation import init_collection
from MsgBox import MsgBox
from ponderation import calcul_nb_taxo
from specific_sswd import filtre_collection_act, init_collection_act


def recherche_nom_feuille(plage, erreur, nom_feuille, data_plage):
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
        nom_feuille = plage[1:ipos - 1]
        long_tot = len(plage)
        data_plage = plage[ipos + 1:long_tot]


def trf_plage_cellule(nom_feuille, plage, l1, c1, l2, c2, erreur):
    """Recherche les lignes et colonnes d'une plage de cellules."""
    # erreur = False
    global Worksheets
    if Application.ReferenceStyle == xlR1C1:  # TODO modifie en python IHM
        """
        Recherche separateur ":" specifiant la selection d'une plage
        """
        ipos_sep = plage.find(':')
        """Cas d'une selection de colonne et non de plage"""
        if ipos_sep == 0:
            # Recherche indice colonne
            rech_l1c1(plage, l1, c1, 2)
            c2 = c1
            l2 = l1
            while Worksheets[nom_feuille].Cells[l2, c1]:
                l2 = l2 + 1
            l2 = l2 - 1
            """Cas d'une selection d'une plage"""
        else:
            """
            Recherche indice ligne/ colonne de la premiere cellule de
            la plage
            """
            rech_l1c1(plage, l1, c1, 2)
            """
            Recherche indice ligne/ colonne de la premiere cellule de
            la plage
            """
            rech_l1c1(plage, l2, c2, ipos_sep + 2)
    else:
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
        # c2 = Worksheets[nom_feuille].Range(plage).Cells(i - 1, 1).Column


def lire_pcat(val_pcat, pcat, dim_pcat, erreur):
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


def afficher_taxo(data_taxo, liste_taxo, erreur):
    """Bouton pcat enter weight values est actionne."""
    """Recherche du nom de la feuille"""
    recherche_nom_feuille(data_taxo, erreur, nom_feuille,
                          plage)  # TODO return nom_feuille, plage
    if erreur is True:
        return
    """Recherche plage de cellules"""
    trf_plage_cellule(nom_feuille, plage, l1, c1, l2, c2, erreur)
    if erreur is True:
        return
    """
    Les nom taxo sont charges dans un tableau et tries par ordre
    alphabetique
    """
    # Worksheets[nom_feuille].Activate()
    tmp = Worksheets[nom_feuille].Range(Cells[l1, c1], Cells[l2, c2]).Value
    taxo = list()
    for i in range(0, len(taxo)):
        taxo.append(tmp[i + 1, 1])
    trier_tableau(taxo)
    """Extraction des differentes categories taxo"""
    rechercher_categorie(taxo, taxo_dif)  # TODO return taxo_dif
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


def charger_parametres(data_co, nom_feuille, nom_colonne, isp, pcat, dist, B,
                       a, n_optim, conserv_inter, nb_taxo, val_pcat, ltaxo,
                       liste_taxo, triang_ajust, iproc, nom_ve, nom_inv,
                       nom_testA, nom_al, nom_testC, r_espece, r_taxo,
                       r_concentration, r_test, txt_p, opt_bt_nul, opt_bt_val,
                       ch_e, ch_n, ch_t, txt_val_b, txt_val_a, ch_nb, ch_sauve,
                       lbl_liste, opt_bt_q, cbx_e, nb_vea, nb_inva, erreur):
    """
    Charge les parametres receuillis par la boite de dialogue.

    Execute SSWD
    """
    if iproc == 1:
        nomboite = 'SSWD'
    else:
        nomboite = 'ACT'
    """Chargement de la collection"""
    data = r_espece.Value
    ischainevide(data, 'Select the range or the column of the \
species or genus names!', nomboite, erreur)
    if erreur:
        return
    recherche_nom_feuille(data, erreur, nom_feuille, plage_espece)
    if erreur:
        return
    trf_plage_cellule(nom_feuille, plage_espece, l1_espece, c1_espece,
                      l2_espece, c2_espece, erreur)
    if erreur:
        return
    data = r_taxo.Value
    ischainevide(data, 'Select the range or the column of the trophic \
levels or taxonomic groups!', nomboite, erreur)
    if erreur:
        return
    recherche_nom_feuille(data, erreur, nom_feuille, plage_taxo)
    if erreur:
        return
        trf_plage_cellule(nom_feuille, plage_taxo, l1_taxo, c1_taxo, l2_taxo,
                          c2_taxo, erreur)
    if erreur:
        return
    data = r_concentration.Value
    ischainevide(data, 'Select the range or the column of concentration data!',
                 nomboite, erreur)
    if erreur:
        return
    recherche_nom_feuille(data, erreur, nom_feuille, plage_data)
    if erreur:
        return
        trf_plage_cellule(nom_feuille, plage_data, l1_data, c1_data, l2_data,
                          c2_data, erreur)
    if erreur:
        return
    if (iproc == 1):
        """Collection SSWD"""
        init_collection(nom_feuille, l1_espece + 1, l1_taxo + 1, l1_data + 1,
                        c1_espece, c1_taxo, c1_data, data_co)
    else:
        """Collection ACT"""
        data = r_test.Value
        recherche_nom_feuille(data, erreur, nom_feuille, plage_test)
        if erreur:
            return
        trf_plage_cellule(nom_feuille, plage_test, l1_test, c1_test, l2_test,
                          c2_test, erreur)
        if erreur:
            return
        init_collection_act(nom_feuille, l1_espece + 1, l1_taxo + 1,
                            l1_data + 1, l1_test, c1_espece, c1_taxo, c1_data,
                            c1_test, data_co, nom_ve, nom_inv, nom_al,
                            nom_testC, nom_testA)
        filtre_collection_act(data_co, nom_ve, nom_inv, nom_testA, nom_al,
                              nb_vea, nb_inva, erreur)
        # change_nom_taxo(data_co)
    """Titre des colonnes de data_co"""
    nom_colonne[1] = Worksheets[nom_feuille].Cells[l1_espece, c1_espece]
    nom_colonne[2] = Worksheets[nom_feuille].Cells[l1_taxo, c1_taxo]
    nom_colonne[3] = Worksheets[nom_feuille].Cells[l1_data, c1_data]
    nom_colonne[4] = 'Weight'
    nom_colonne[5] = 'Weighted Emp. Cumul. Prob.'
    if (iproc == 2):
        nom_colonne[6] = 'ACT data'
        nom_colonne[7] = 'Weighted Emp. Cumul. Prob. Acute'
    """Type de pondération espece"""
    if (iproc == 1):
        isp = cbx_e.ListIndex
        isp = isp + 1
    else:
        isp = 1
    """Pcat : pondération taxonomie"""
    trier_collection(data_co, 2, 1)
    calcul_nb_taxo(data_co, nb_taxo)
    pcat = list()
    val_pcat = txt_p.text
    if opt_bt_nul.Value is True:
        for i in range(1, nb_taxo):
            pcat.append(0)
    else:
        if val_pcat == '':
            MsgBox('SSWD', 'Please enter weight values or select No weight!',
                   0)
            erreur = True
            return
        lire_pcat(val_pcat, pcat, nb_taxo, erreur)
        if erreur:
            return
    """parametre de Hazen a"""
    if txt_val_a.text == '':
        MsgBox('SSWD', 'You must chose a value for the Hazen parameter \
between 0 and 1, strictly less than 1', 0)
        erreur = True
        return
    a = float(txt_val_a.text)
    if a >= 1:
        MsgBox('SSWD'
               'The Hazen parameter must be included between \
0 and 1, strictly less than 1', 0)
        erreur = True
        return
    """Loi statistique 1:empirique, 2:normal, 3:triangulaire"""
    dist[1] = ch_e.Value
    dist[2] = ch_n.Value
    dist[3] = ch_t.Value
    """B"""
    if txt_val_b.text == '':
        MsgBox('SSWD',
               'You must chose a value for the number of bootstrap samples', 0)
        erreur = True
        return
    B = int(txt_val_b.text)
    """nbvar : nombre de données tirées"""
    n_optim = ch_nb.Value
    """Sauvegarde des feuilles de resultats intermediaires"""
    conserv_inter = ch_sauve.Value
    """Recupération de la liste taxo"""
    if opt_bt_val.Value is True:
        liste_taxo = lbl_liste.text
        pos = liste_taxo.find('\n')
        ltaxo = liste_taxo[:pos + 1]
    """
    Option d'ajustement pour loi triangulaire,
    si True ajustement sur quantiles, sinon sur probabilités cumulées
    """
    triang_ajust = opt_bt_q.Value
