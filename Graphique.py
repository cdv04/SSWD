"""
Module permettant la mise ne place des graphiques.

A remettre completement en forme grace a numpy
"""

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-12T09:13:56+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: Graphique.py
# @Last modified by:   gysco
# @Last modified time: 2017-05-17T10:23:47+02:00

from matplotlib import pyplot as plot

import Initialisation
from fct_generales import sp_opt


def tracer_graphique(nom_feuille, lig_p, lig_qbe, lig_qbi, lig_qbs, col_deb,
                     col_fin, ligne_data, col_tax, col_data, col_pcum,
                     col_data_le, col_pcum_le, loi, titre_graf, R2, Pvalue,
                     nb_ligne_data, mup, sigmap, _min, _max, mode, titre_axe,
                     val_pcat, liste_taxo, isp, col_data_act, col_data_act_le,
                     iproc, col_pcum_a):
    """
    Trace les graphiques (Mediane, quantile 5%, quantile 95% et data).

    nom_feuille -> nom de la feuille contenant les donnees
    lig_p   -> ligne contenant les valeurs probabilites cumulees retenues
    lig_qbe   -> ligne quantiles best-estimates
    lig_qbi   -> ligne quantiles borne inferieure de l'intervalle de confiance
    lig_qbs    -> ligne qualites borne superieure de l'intervalle de confiance
    col_deb    -> colonne debut des donnees numeriques
     col_fin    -> colonne fin des donnees numeriques
     col_tax  -> colonne des indicateurs de categories taxonomiques
                 pour representation points differencies
    col_data  -> colonne des donnees de concentration
    col_pcum  -> colonne des probabilites cumulees empiriques
    ligne_data -> numero de la premiere ligne du tableau des donnees data
    col_data_le colonne des donnees data pour representation d'une ligne
    continue entre les points empiriques
    col_pcum_le colonne des donnees de probabilite cumulee ponderee empirique
    pour representation ligne continue  si loi = empirique
    loi -> type de di_stribution (1=empirique, 2=normal, 3=triang)
    titra_graf -> tableau des titres des graphiques
    R2 -> coefficient de determination
    mup, sigmap parametres moyenne et ecart type ponderes
    _min, _max, mode parametres de la loi triangulaire ponderee
    nb_ligne_data -> nombre de ligne du tableau de donnees data
    """
    nseries = 0
    """Ajout de la serie Mediane"""
    if (loi != 1):
        nseries = nseries + 1
        data_x = list()
        data_y = list()
        for x in range(col_deb, col_fin):
            data_x.append(Initialisation.Worksheets[nom_feuille]
                          .Cells.get_value(lig_qbe, x))
            data_y.append(
                float(Initialisation.Worksheets[nom_feuille]
                      .Cells.get_value(lig_p, x)[:-1]))
        data_nom = Initialisation.Worksheets[nom_feuille].Cells.get_value(
            lig_qbe, col_deb - 1)
        # TODO modifer en loi log avec range 0, 100
        plot.plot(data_x, data_y, 'k-', label=data_nom, linewidth=1)
    else:
        if iproc == 1:
            nseries = tracer_courbe_empirique(nseries, nom_feuille,
                                              ligne_data + 1, nb_ligne_data,
                                              col_data_le, col_pcum_le)
        else:
            nseries = tracer_courbe_empirique(nseries, nom_feuille,
                                              ligne_data + 1, nb_ligne_data,
                                              col_data_act_le, col_pcum_le)
    """Ajout de la serie quantile borne inf"""
    data_x = list()
    data_y = list()
    for x in range(col_deb, col_fin):
        data_x.append(Initialisation.Worksheets[nom_feuille]
                      .Cells.get_value(lig_qbi, x))
        data_y.append(
            float(Initialisation.Worksheets[nom_feuille]
                  .Cells.get_value(lig_p, x)[:-1]))
    data_nom = Initialisation.Worksheets[nom_feuille].Cells.get_value(
        lig_qbi, col_deb - 1)
    plot.plot(data_x, data_y, 'r--', label=data_nom, linewidth=0.5)
    nseries = nseries + 1
    """Ajout de la serie quantile borne superieure"""
    data_x = list()
    data_y = list()
    for x in range(col_deb, col_fin):
        data_x.append(Initialisation.Worksheets[nom_feuille]
                      .Cells.get_value(lig_qbs, x))
        data_y.append(
            float(Initialisation.Worksheets[nom_feuille]
                  .Cells.get_value(lig_p, x)[:-1]))
    data_nom = Initialisation.Worksheets[nom_feuille].Cells.get_value(
        lig_qbs, col_deb - 1)
    plot.plot(data_x, data_y, 'r--', label=data_nom, linewidth=0.5)
    nseries = nseries + 1
    """
    Ajoute les donnees data
    """
    if iproc == 1:
        ajoute_series(nom_feuille, nseries, True, ligne_data, col_tax,
                      col_data, col_pcum, col_pcum_a)
    else:
        ajoute_series(nom_feuille, nseries, True, ligne_data, col_tax,
                      col_data_act, col_pcum, col_pcum_a)
        ajoute_series(nom_feuille, nseries, False, ligne_data, col_tax,
                      col_data, col_pcum_a, col_pcum_a)
    """
    Ajoute une zone de texte avec les valeurs de R2, Pttest, GWM et GWSD
    """
    if (loi == 2):
        plot.text(
            max(data_x) - 1, 0.9, 'R_ = {:.4f}\nKSpvalue = {:.3f}'.format(
                R2, Pvalue))
        plot.text(
            max(data_x), 1, 'wm.lg = {:.2f}\nwsd.lg = {:.2f}'.format(
                mup, sigmap))
    elif (loi == 3):
        plot.text(0.1, 0.9, 'R_ = {:.4f}\nKSpvalue = {:.3f}'.format(
            R2, Pvalue))
        plot.text(
            1, 1,
            'wmin.lg = {:.2f}\nwmax.lg = {:.2f}\nwmode.lg = {:.2f}'.format(
                _min, _max, mode))
    """Rappel des options dans le titre du graphique"""
    ligne_option = 'Sp = ' + sp_opt(isp)
    if val_pcat != '':
        _str = ""
        for x in set(liste_taxo):
            if x != "":
                _str += x + " "
        ligne_option = '{}; TW: {}= {}'.format(ligne_option, _str, val_pcat)
    else:
        ligne_option = '{}; TW: none'.format(ligne_option)
    plot.title(titre_graf[loi - 1] + '\n' + ligne_option)
    plot.xlabel(titre_axe[0])
    plot.ylabel(titre_axe[1])
    plot.xscale('log')
    plot.legend()
    plot.grid()
    plot.show()
    # ActiveChart.ChartTitle.Characters.text = titre_graf(loi) + '\n' + ligne_option
    # ActiveChart.ChartTitle.Characters[Start= 1, Length= len(titre_graf(loi))].Font.Size = 10
    # ActiveChart.ChartTitle.Characters[Start= len(titre_graf(loi)) + 1, Length= nb_car].Font.Bold = False
    # ActiveChart.ChartTitle.Characters[Start= len(titre_graf(loi)) + 1, Length= nb_car].Font.Size = 8
    # On reduit la taille de police de la legende
    # ActiveChart.Legend.Select()
    # Selection.AutoScaleFont = True
    # _with13 = Selection.Font
    # _with13.Name = 'Arial'
    # _with13.FontStyle = 'Normal'
    # _with13.Size = 8
    # _with13.Strikethrough = False
    # _with13.Superscript = False
    # _with13.Subscript = False
    # _with13.OutlineFont = False
    # _with13.Shadow = False
    # _with13.Underline = xlUnderlineStyleNone
    # _with13.ColorIndex = xlAutomatic
    # _with13.Background = xlAutomatic


def ajoute_series(nom_feuille, nseries, nouveau, ligne_data, col_tax, col_data,
                  col_pcum, col_pcum_a):
    """
    Ajoute les donnees par especes.

    @param nom_feuille: nom de la feuille contenant les graphiques
    @param nseries: nombre de series de courbes tracees
    @param ligne_data: numero premiere ligne du tableau des donnees;
                       premiere ligne de donnees numeriques
    @param col_tax: colonne de la variable taxonomie
    @param col_data: colonne des donnees tox
    @param col_pcum: colonne des probabilites cumulees ponderees
                     empiriques
    """
    marker_style = [
        'o', 'v', '1', 's', 'p', '*', 'h', 'H', '+', 'x', 'D', 'd', '|', '_',
        '.', '^', '<', '>', '2', '3', '4'
    ]
    color_style = ['b', 'g', 'r', 'c', 'm', 'y', 'k', 'w']
    data_x = list(
        Initialisation.Worksheets[nom_feuille].Cells.ix[3:, col_data].dropna())
    data_y = list(
        Initialisation.Worksheets[nom_feuille].Cells.ix[3:, col_pcum].dropna())
    taxon = list(
        Initialisation.Worksheets[nom_feuille].Cells.ix[3:, col_tax].dropna())
    species = ["s"] * (len(taxon))
    # list(Initialisation.Worksheets[nom_feuille].Cells.ix[
    #     3:, col_tax - 1].dropna())
    s_species = list(set(species))
    s_taxon = list(set(taxon))
    for e in s_taxon:
        for s in s_species:
            sub_data_x = list()
            sub_data_y = list()
            for i in range(0, len(taxon)):
                if e == taxon[i] and s == species[i]:
                    sub_data_x.append(data_x[i])
                    sub_data_y.append(data_y[i] * 100)
            if sub_data_x and sub_data_y:
                plot.plot(
                    sub_data_x,
                    sub_data_y,
                    color_style[s_taxon.index(e)] +
                    marker_style[s_species.index(s)],
                    label=e)


def decaler_graph(nom_feuille):
    """
    Modifie la position des graphiques contenus dans une feuille.

    @param nom_feuille: nom de la feuille contenant les graphiques
    """
    decalage = 0
    for ch in Initialisation.Worksheets[nom_feuille].ChartObjects:
        ch.Left = ch.Left + decalage
        decalage = decalage + 200


def tracer_courbe_empirique(nseries, nom_feuille, lig_deb, nbligne, col_data,
                            col_pcum):
    """
    Relie points ponderes dans le graph correspdant a la loi empirique.

    @param nseries: numero de la serie de donnees encours a representer
    @param nom_feuille: nom de la feuille de calcul ou sont les donnees
    @param lig_deb: ligne de debut des donnees numeriques representees
    @param nbligne: nombre de lignes a traiter
    @param col_data: colonne des donnees numeriques
    @param col_pcum: colonne des donnees de probabilite cumulees ponderees
                     empiriques
    """
    nseries = nseries + 1
    data_x = list()
    data_y = list()
    for y in range(lig_deb, lig_deb + nbligne):
        data_x.append(Initialisation.Worksheets[nom_feuille]
                      .Cells.get_value(y, col_data))
        data_y.append(Initialisation.Worksheets[nom_feuille]
                      .Cells.get_value(y, col_pcum))
    data_nom = 'Weighted Empirical'
    plot.plot(data_x, data_y, 'k-', label=data_nom, linewidth=0.5)
    return (nseries)
