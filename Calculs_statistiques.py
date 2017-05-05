"""
Calculs statistiques.

Many function to refactor to python function.
"""

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-11T10:01:21+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: Calculs_statistiques.py
# @Last modified by:   gysco
# @Last modified time: 2017-05-05T11:19:37+02:00

import math

import Initialisation
from fct_generales import (cellule_gras, compt_inf, csd, ecrire_titre,
                           encadrer_colonne, trier_tirages_feuille)
from Worksheet import Worksheet


def tirage(nom_feuille_stat, nbvar, B, nom_feuille_pond, lig_deb, col_data,
           lig_fin, col_pond):
    """
    Effectue un tirage aleatoire suivant une loi multinomiale.

    @param nom_feuille_stat: nom de la feuille pour affichage
                             resultats du tirage
    @param nbvar: nombre de points tires (nombre de donnees utilisees
                  dans le calcul de la HC)
    @param B: nombre de tirages du bootstrap
    @param nom_feuille_pond: nom de la feuille contenant les donnees a
                             tirer et les probabilites associees
    @param lig_deb: ligne de debut de la plage des donnees a tirer
    remarque: il y a une ligne de titre avant cette ligne
    @param col_data: colonne des donnees a tirer
    @param lig_fin: derniere ligne de la plage des donnees a tirer
    @param col_pond: colonne des probabilites associees a chaque donnee
    """
    # Application.Run('ATPVBAEN.XLA!Random', nom_feuille_stat, nbvar, B, 7,
    #                 Initialisation.Worksheets[nom_feuille_pond].Range(
    #                     Cells(lig_deb, col_data), Cells(lig_fin, col_pond)))
    # Initialisation.Worksheets[nom_feuille_stat].Rows(1).Insert()
    for j in range(0, nbvar):
        Initialisation.Worksheets[nom_feuille_stat].Cells.set_value(
            1, j, 'POINT ' + str(j))
    # Initialisation.Worksheets[nom_feuille_stat].Cells[1, 1].Select()


def calcul_ic_empirique(l1, c1, l2, c2, c3, p, nom_feuille_stat,
                        nom_feuille_qemp, nom_feuille_sort, nbvar, a):
    """
    Calcul les centiles p% empiriques sur les echantillons du tirage aleatoire.

    @param nom_feuille_stat: nom de la feuille des tirages
                             aleatoires
    @param nom_feuille_qemp: nom de la feuille contenant les
                             quantiles empiriques calcules a
                             partir des tirages
                             dans nom_feuille_stat
    @param l1: premiere ligne de donnees numeriques tirees ;
               sera egalement la premiere ligne des qemp
    @param c1: premiere colonne de donnees tirees
    @param l2: derniere ligne de donnees tirees
    @param c2: derniere colonne de donnees tirees
               dans nom_feuille_qemp
    @param c3: premiere colonne affichage resultats quantiles empiriques
    Remarque : la premiere ligne des resultats empiriques est
               obligatoirement la meme que celle des tirages
    """
    pcum = list()
    rang = list()
    tmp = list()
    # Application.ScreenUpdating = False
    """
    On calcule la probabilite cumulee empirique de chaque point tire
    sauvegarde dans pcum
    """
    for i in range(0, nbvar):
        pcum.append((i - a) / (nbvar + 1 - 2 * a))
    """
    On calcule le rang qu'occupe la probabilite voulues (p) au sein des
    pcum, sauvegarde dans rang
    """
    for i in range(0, len(p)):
        rang.append(compt_inf(pcum, p[i]))
    """
    On trie les donnees a exploiter (issues des tirages aleatoires)
    dans l'ordre croissant
    Creation de la feuille nom_feuille_sort
    """
    data = 'RC' + c1 + ':RC' + c2
    trier_tirages_feuille(nom_feuille_stat, nom_feuille_sort, l1, c1, l2,
                          nbvar, data)
    """Creation de la feuille contenant les quantiles empiriques"""
    Initialisation.Worksheets[nom_feuille_qemp] = Worksheet()
    """Ecriture des entetes de colonnes"""
    for i in range(0, len(p)):
        Initialisation.Worksheets[nom_feuille_qemp].Cells[
            l1 - 1, c3 + i - 1] = 'QUANT ' + p[i] * 100 + ' %'
    """
    Calcul des quantiles p%
    data = "RC" & c1 & ":RC" & c2
    """
    data = nom_feuille_sort + '!RC'
    for i in range(0, len(p)):
        # Initialisation.Worksheets(nom_feuille_qemp).Cells(l1, c3 + i - 1).FormulaR1C1 =
        # "=PERCENTILE(" & nom_feuille_stat & "!" & data & "," & csd(p[i])
        # & ")"
        if (rang[i] == 0 or rang[i] == nbvar):
            tmp[i] = 0
        else:
            tmp[i] = (pcum[rang[i] + 1] - p[i]) / \
                (pcum[rang[i] + 1] - pcum[rang[i]])
        Initialisation.Worksheets[nom_feuille_qemp].Cells[
            l1, c3 + i -
            1].FormulaR1C1 = '=IF(' + rang[i] + '=0,' + data + c1 + \
            ',IF(' + rang[i] + '>=' + nbvar + ',' + data + c2 + ',' + data + \
            rang[i] + 1 + '-(' + data + rang[i] + 1 + '-' + data + rang[i] +\
            ')*' + csd(tmp[i]) + '))'
    # Range(
    #     Initialisation.Worksheets(nom_feuille_qemp).Cells(l1, c3),
    #     Initialisation.Worksheets(nom_feuille_qemp).Cells(l1, c3 + len(p) - 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_qemp).Cells(l1, c3),
    #         Initialisation.Worksheets(nom_feuille_qemp).Cells(l2, c3 + len(p) - 1)),
    #     Type=xlFillDefault)
    # Initialisation.Worksheets(nom_feuille_qemp).Cells(1, 1).Select()


def calcul_ic_normal(l1, c1, l2, c2, c3, p, nom_feuille_stat,
                     nom_feuille_qnorm, c_mu):
    """
    Calcul les centiles p% normaux pour chaque echantillon du tirage.

    @param nom_feuille_stat: nom de la feuille des tirages aleatoires
    @param nom_feuille_qnorm: nom de la feuille contenant les
                              quantiles normaux calcules a partir des
                              tirages
    dans nom_feuille_stat
        @param l1: premiere ligne de donnees tirees ;
                   correspond egalement a la premiere ligne des
                   quantiles normaux ;
                   il y a une premiere ligne de titre avant
        @param c1: premiere colonne de donnees tirees
        @param l2: derniere ligne de donnees tirees
        @param c2: derniere colonne de donnees tirees
        @param c_mu : indice de la colonne contenant la moyenne des
                      tirages ; la colonne contenant l'ecart type est
                      a c_mu+1 ;
                      remarque c_mu est definie dans cette procedure
    dans nom_feuille_qnorm
        @param c3: premiere colonne affichage resultats normaux
    """
    c_mu = c2 + 2
    # Application.ScreenUpdating = False
    """
    Calcul moyenne et ecart type correspondant a chaque tirage
    (chaque ligne de nom_feuille_stat)
    on travaille dans nom_feuille_stat
    """
    # Initialisation.Worksheets(nom_feuille_stat).Activate()
    Initialisation.Worksheets[nom_feuille_stat].Cells[l1 - 1, c_mu] = 'MEAN'
    Initialisation.Worksheets[nom_feuille_stat].Cells[l1 - 1, c_mu +
                                                      1] = 'STDEV'
    data = 'RC' + c1 + ':RC' + c2
    """1. Calcul de la moyenne des echantillons"""
    Initialisation.Worksheets[nom_feuille_stat].Cells[
        l1,
        c_mu].FormulaR1C1 = '=AVERAGE(' + nom_feuille_stat + '!' + data + ')'
    """2. Calcul de l'ecart type des echantillons"""
    Initialisation.Worksheets[nom_feuille_stat].Cells[
        l1, c_mu +
        1].FormulaR1C1 = '=STDEV(' + nom_feuille_stat + '!' + data + ')'
    # Range(
    #     Initialisation.Worksheets(nom_feuille_stat).Cells(l1, c_mu),
    #     Initialisation.Worksheets(nom_feuille_stat).Cells(l1, c_mu + 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_stat).Cells(l1, c_mu),
    #         Initialisation.Worksheets(nom_feuille_stat).Cells(l2, c_mu + 1)),
    #     Type=xlFillDefault)
    """
    3. Calcul quantiles normaux correspondant a p() et a mean et
    stdev precedemment calcules
    """
    """Affichage dans nom_feuille_qnorm"""
    Initialisation.Worksheets[nom_feuille_qnorm] = Worksheet()
    for i in range(0, len(p)):
        Initialisation.Worksheets[nom_feuille_qnorm].Cells[
            l1 - 1, c3 + i - 1] = 'QUANT ' + p[i] * 100 + ' %'
        Initialisation.Worksheets[nom_feuille_qnorm].Cells[
            l1, c3 + i - 1].FormulaR1C1 = '=NORMINV(' + csd(
                p[i]) + ',' + nom_feuille_stat + '!RC' + (
                    c_mu) + ',' + nom_feuille_stat + '!RC' + (c_mu + 1) + ')'
    # Range(
    #     Initialisation.Worksheets(nom_feuille_qnorm).Cells(l1, c3),
    #     Initialisation.Worksheets(nom_feuille_qnorm).Cells(l1, c3 + len(p) - 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_qnorm).Cells(l1, c3),
    #         Initialisation.Worksheets(nom_feuille_qnorm).Cells(l2, c3 + len(p) - 1)),
    #     Type=xlFillDefault)
    # Initialisation.Worksheets(nom_feuille_qnorm).Cells(1, 1).Select()


def calcul_ic_triang_p(l1, c1, l2, c2, c3, nbvar, a, p, nom_feuille_stat,
                       nom_feuille_sort, nom_feuille_Ftriang,
                       nom_feuille_qtriang, c_min):
    """
    Calcul les centiles p% triangulaires pour chaque echantillon.

    Avant cela, il faut estimer les parametres min, max, mode de
    la loi triangulaire correspondant a chaque echantillon aleatoire
    pour se faire : utilisation du solver (ajust proba cumulees)
    @param nom_feuille_stat: nom de la feuille des tirages aleatoires
    @param nom_feuille_qtriang: nom de la feuille contenant les
                                quantiles triangulaires calcules a
                                partir des tirages
    @param nom_feuille_sort: nom de la feuille des tirages aleatoires
                             classes dans l'ordre croissant sur chaque
                             ligne
    @param nom_feuille_Ftriang: nom de la feuille contenant les
                                probabilites cumulees triangulaires,
                                empiriques  et theoriques pour
                                ajustement et determination des
                                parametres min, max et mode
    dans nom_feuille_stat
    @param l1: premiere ligne de donnees tirees ;
               egalement premiere ligne resultats triangulaires
    @param c1: premiere colonne de donnees tirees
    @param l2: derniere ligne de donnees tirees
    @param c2: derniere colonne de donnees tirees
    @param c_min: indice de la colonne contenant le min des tirages ;
                  le max et le mode occupe les positions respectives
                  c_min+1 et c_min+2
    dans nom_feuille_sort, nom_feuille_Ftriang et nom_feuille_qtriang
    @param c3: premiere colonne affichage resultats triangulaire
    @param nbvar: nombre de points tires a chaque tirage du bootstrap
    @param a: parametre de Hazen pour calcul des probabilites
              empiriques cumulees
    """
    # Application.ScreenUpdating = False
    indic = 0
    data = 'RC' + c1 + ':RC' + c2
    """
    On trie les donnees a exploiter (issues des tirages aleatoires)
    dans l'ordre croissant
    Creation de la feuille nom_feuille_sort si pas existante
    (pour empirique)
    """
    for ws in Initialisation.Worksheets:
        if ws.Name == nom_feuille_sort:
            indic = 1
    if indic == 0:
        trier_tirages_feuille(nom_feuille_stat, nom_feuille_sort, l1, c3, l2,
                              nbvar, data)
    """
    On calcule les probabilites cumulees empiriques que l'on affiche
    dans la premiere ligne et on met en place les formules de
    probabilite triangulaire qui seront comparees aux probabilites
    empiriques ; creation de la feuille nom_feuille_Ftriang
    """
    # Initialisation.Worksheets.Add()
    # ActiveSheet.Name = nom_feuille_Ftriang
    """
    On initialise le solver en prennant le min et le max de chaque
    serie tiree et on calcule mode=(min+max)/2
    """
    c_min = c3 + nbvar + 1
    c_max = c_min + 1
    c_mode = c_max + 1
    Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1, c_min] = 'min'
    Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1, c_max] = 'max'
    Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1,
                                                         c_mode] = 'mode'
    for i in range(l1, l2):
        Initialisation.Worksheets[nom_feuille_Ftriang].Cells[
            i, c_min] = Initialisation.Worksheets[nom_feuille_sort].Cells[i,
                                                                          c3]
        Initialisation.Worksheets[nom_feuille_Ftriang].Cells[
            i, c_max] = Initialisation.Worksheets[nom_feuille_sort].Cells[
                i, c3 + nbvar - 1]
        Initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_mode] = (
            Initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_min] +
            Initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_max]) / 2
    """Calcul probabilites empiriques et theoriques pour ajustement"""
    for i in range(0, nbvar):
        Initialisation.Worksheets[nom_feuille_Ftriang].Cells[
            l1 - 1, c3 + i - 1] = (i - a) / (nbvar + 1 - 2 * a)
    # data = Cells(l1, c3).Address(
    #     False, False, xlR1C1, RelativeTo=Cells(l1, c3))
    # data1 = nom_feuille_sort + '!' + data
    ref = nom_feuille_Ftriang + '!RC'
    # Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1, c3].FormulaR1C1 = (
    #     '=IF(' + data1 + '<=' + ref + c_min + ',0, IF(' + data1 + '<=' + ref+
    #     c_mode + ', ((' + data1 + '-' + ref + c_min + ')^2)/(('+ref + c_max +
    #     '-' + ref + c_min + ')*(' + ref + c_mode + '-' +ref+ c_min + ')),' +
    #     'IF(' + data1 + '<=' + ref + c_max + ', 1-((' + data1 + '-' + ref +
    #     c_max + ')^2)/((' + ref + c_max + '-' + ref + c_min + ')*(' + ref +
    #     c_max + '-' + ref + c_mode + ')),1)))')
    # Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3),
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3 + nbvar - 1)),
    #     Type=xlFillDefault)
    # Range(
    #     Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3),
    #     Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3 + nbvar - 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3),
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l2, c3 + nbvar - 1)),
    #     Type=xlFillDefault)
    """
    On calcule la somme des carres des differences entre probabilites
    empiriques et probabilites theoriques triangulaires,
    pour ajustement
    """
    c_ssr = c_mode + 1
    # Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1, c_ssr] = 'Sum Square Res'
    # data = Cells(l1 - 1, c3).Address(True, True, xlR1C1) + ':' + Cells(
    #     l1 - 1, c3 + nbvar - 1).Address(True, True, xlR1C1)
    # data = nom_feuille_Ftriang + '!' + data
    # data1 = Cells(l1, c3).Address(
    #     False, False, xlR1C1, RelativeTo=Cells(l1, c_ssr)) + ':' + Cells(
    #         l1, c3 + nbvar - 1).Address(
    #             False, False, xlR1C1, RelativeTo=Cells(l1, c_ssr))
    # data1 = nom_feuille_Ftriang + '!' + data1
    # Initialisation.Worksheets[nom_feuille_Ftriang].Cells[
    #     l1, c_ssr].FormulaR1C1 = '=SUMXMY2(' + data + ',' + data1 + ')'
    # Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c_ssr).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c_ssr),
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l2, c_ssr)),
    #     Type=xlFillDefault)
    """
    Ajustement : on determine les valeurs de min, max, mode qui
    minimisent la sum square res et on calcule la probabilite du mode
    (necessaire pour le calcul des quantiles)
    """
    c_pmode = c_ssr + 1
    Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1,
                                                         c_pmode] = 'pmode'
    for i in range(l1, l2):
        # SolverOk(
        #     SetCell=Initialisation.Worksheets(nom_feuille_Ftriang).Cells(i, c_ssr),
        #     MaxMinVal=2,
        #     ValueOf='0',
        #     ByChange=Range(
        #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(i, c_min),
        #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(i, c_min + 2)))
        # SolverSolve(UserFinish=True)
        Initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_pmode] = (
            Initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_mode] -
            Initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_min]
        ) / (Initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_max] -
             Initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_min])
    Initialisation.Worksheets[nom_feuille_Ftriang].Cells(1, 1).Select()
    """
    Calcul des quantiles correspondant a la loi triangulaire dont les
    parametres min, max et mode viennent d'etre estimes ; creation de
    la feuille nom_feuille_qtriang
    """
    Initialisation.Worksheets[nom_feuille_qtriang] = Worksheet()
    ref = nom_feuille_Ftriang + '!RC'
    for i in range(0, len(p)):
        Initialisation.Worksheets[nom_feuille_qtriang].Cells[
            l1 - 1, c3 + i - 1] = 'QUANT ' + p[i] * 100 + ' %'
        Initialisation.Worksheets[nom_feuille_qtriang].Cells[
            l1, c3 + i - 1].FormulaR1C1 = (
                '=IF(' + csd(p[i]) + '<=' + ref + c_pmode + ',' + ref + c_min +
                '+SQRT(' + csd(p[i]) + '*(' + ref + c_max + '-' + ref + c_min +
                ')*(' + ref + c_mode + '-' + ref + c_min + ')), ' + ref + c_max
                + '-SQRT((' + csd(1 - p[i]) + ')*(' + ref + c_max + '-' + ref +
                c_min + ')*(' + ref + c_max + '-' + ref + c_mode + ')))')
    # Range(
    #     Initialisation.Worksheets(nom_feuille_qtriang).Cells(l1, c3),
    #     Initialisation.Worksheets(nom_feuille_qtriang).Cells(l1,
    #                                           c3 + len(p) - 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_qtriang).Cells(l1, c3),
    #         Initialisation.Worksheets(nom_feuille_qtriang).Cells(l2, c3 + len(p) - 1)),
    #     Type=xlFillDefault)
    # Initialisation.Worksheets(nom_feuille_qtriang).Cells(1, 1).Select()


def calcul_ic_triang_q(l1, c1, l2, c2, c3, nbvar, a, p, nom_feuille_stat,
                       nom_feuille_sort, nom_feuille_Ftriang,
                       nom_feuille_qtriang, c_min):
    """
    Calcul les centiles p% triangulaires pour chaque echantillon.

    Avant cela, il faut estimer les parametres min, max, mode de la
    loi triangulaire correspondant a chaque echantillon aleatoire
    pour se faire: utilisation du solver, ajustement sur les quantiles
    @param nom_feuille_stat: nom de la feuille des tirages aleatoires
    @param nom_feuille_qtriang: nom de la feuille contenant les
                                quantiles triangulaires calcules
                                a partir des tirages
    @param nom_feuille_sort: nom de la feuille des tirages aleatoires
                             classes dans l'ordre croissant sur chaque
                             ligne
    @param nom_feuille_Ftriang: nom de la feuille contenant les
                                quantiles triangulaires, empiriques et
                                theoriques pour ajustement et
                                determination des parametres min, max
                                et mode
    dans nom_feuille_stat
    @param l1: premiere ligne de donnees tirees ;
               egalement premiere ligne resultats triangulaires
    @param c1: premiere colonne de donnees tirees
    @param l2: derniere ligne de donnees tirees
    @param c2: derniere colonne de donnees tirees
    @param c_min: indice de la colonne contenant le min des tirages ;
                  le max et le mode occupe les positions respectives
                  c_min+1 et c_min+2
    dans nom_feuille_sort, nom_feuille_Ftriang et nom_feuille_qtriang
    @param c3: premiere colonne affichage resultats triangulaire
    @param nbvar: nombre de points tires a chaque tirage du bootstrap
    @param a: parametre de Hazen pour calcul des probabilites
              empiriques cumulees
    """
    indic = 0
    # Application.ScreenUpdating = False
    data = 'RC' + c1 + ':RC' + c2
    """
    On trie les donnees a exploiter (issues des tirages aleatoires)
    dans l'ordre croissant
    Creation de la feuille nom_feuille_sort si pas deja existant
    (pour empirique)
    """
    for ws in Initialisation.Worksheets:
        if ws.Name == nom_feuille_sort:
            indic = 1
    if indic == 0:
        trier_tirages_feuille(nom_feuille_stat, nom_feuille_sort, l1, c3, l2,
                              nbvar, data)
    """
    On calcule les probabilites cumulees empiriques que l'on affiche
    dans la premiere ligne et on met en place les formules de quantile
    triangulaire qui seront comparees aux valeurs empiriques ;
    creation de la feuille nom_feuille_Ftriang
    """
    Initialisation.Worksheets[nom_feuille_Ftriang] = Worksheet()
    """
    On initialise le solver en prennant le min et le max de chaque
    serie tiree et on calcule mode=(min+max)/2 puis pmode
    """
    c_min = c3 + nbvar + 1
    c_max = c_min + 1
    c_mode = c_max + 1
    c_ssr = c_mode + 1
    c_pmode = c_ssr + 1
    ref = nom_feuille_Ftriang + '!RC'
    Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1, c_min] = 'min'
    Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1, c_max] = 'max'
    Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1,
                                                         c_mode] = 'mode'
    Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1,
                                                         c_pmode] = 'pmode'
    for i in range(l1, l2):
        Initialisation.Worksheets[nom_feuille_Ftriang].Cells[
            i, c_min] = Initialisation.Worksheets[nom_feuille_sort].Cells[i,
                                                                          c3]
        Initialisation.Worksheets[nom_feuille_Ftriang].Cells[
            i, c_max] = Initialisation.Worksheets[nom_feuille_sort].Cells[
                i, c3 + nbvar - 1]
        Initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_mode] = (
            Initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_min] +
            Initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_max]) / 2
    # Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1, c_pmode].FormulaR1C1 = (
    #     '=(' + ref + c_mode + '-' + ref + c_min + ')/(' + ref + c_max + '-' +
    #     ref + c_min + ')')
    # Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c_pmode).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c_pmode),
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l2, c_pmode)),
    #     Type=xlFillDefault)
    """
    Calcul probabilites empiriques et quantiles triangulaires
    correspondants
    """
    for i in range(0, nbvar):
        Initialisation.Worksheets[nom_feuille_Ftriang].Cells[
            l1 - 1, c3 + i - 1] = (i - a) / (nbvar + 1 - 2 * a)
    ref2 = nom_feuille_Ftriang + '!R'
    # Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1, c3].FormulaR1C1 = (
    #     '=IF(' + ref2 + l1 - 1 + 'C<=' + ref + c_pmode + ',' + ref + c_min +
    #     '+SQRT(' + ref2 + l1 - 1 + 'C*(' + ref + c_max + '-' + ref + c_min +
    #     ')' + '*(' + ref + c_mode + '-' + ref + c_min + ')),' + ref + c_max +
    #     '-SQRT((1-' + ref2 + l1 - 1 + 'C)*(' + ref + c_max + '-' + ref +c_min
    #     + ')' + '*(' + ref + c_max + '-' + ref + c_mode + ')))')
    # Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3),
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3 + nbvar - 1)),
    #     Type=xlFillDefault)
    # Range(
    #     Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3),
    #     Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3 + nbvar - 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3),
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l2, c3 + nbvar - 1)),
    #     Type=xlFillDefault)
    """
    On calcule la somme des carres des differences entre donnees
    empiriques etquantiles theoriques triangulaires, pour ajustement
    """
    Initialisation.Worksheets[nom_feuille_Ftriang].Cells[
        l1 - 1, c_ssr] = 'Sum Square Res'
    # data = Cells(l1, c3).Address(
    #     False, False, xlR1C1, RelativeTo=Cells(l1, c_ssr)) + ':' + Cells(
    #         l1, c3 + nbvar - 1).Address(
    #             False, False, xlR1C1, RelativeTo=Cells(l1, c_ssr))
    # data1 = nom_feuille_Ftriang + '!' + data
    # data = nom_feuille_sort + '!' + data
    # Initialisation.Worksheets[nom_feuille_Ftriang].Cells[
    #     l1, c_ssr].FormulaR1C1 = '=SUMXMY2(' + data + ',' + data1 + ')'
    # Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c_ssr).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c_ssr),
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l2, c_ssr)),
    #     Type=xlFillDefault)
    """
    Ajustement : on determine les valeurs de min, max, mode qui
    minimisent la sum square res et on calcule la probabilite du mode
    (necessaire pour le calcul des quantiles)
    """
    # for i in range(l1, l2):
    #     SolverOk(
    #         SetCell=Initialisation.Worksheets(nom_feuille_Ftriang).Cells(i, c_ssr),
    #         MaxMinVal=2,
    #         ValueOf='0',
    #         ByChange=Range(
    #             Initialisation.Worksheets(nom_feuille_Ftriang).Cells(i, c_min),
    #             Initialisation.Worksheets(nom_feuille_Ftriang).Cells(i, c_min + 2)))
    #     SolverSolve(UserFinish=True)
    Initialisation.Worksheets[nom_feuille_Ftriang].Cells(1, 1).Select()
    """
    Calcul des quantiles correspondant a la loi triangulaire dont les
    parametres min, max et mode viennent d'etre estimes ; creation de
    la feuillenom_feuille_qtriang
    """
    Initialisation.Worksheets[nom_feuille_qtriang] = Worksheet()
    for i in range(0, len(p)):
        Initialisation.Worksheets[nom_feuille_qtriang].Cells[
            l1 - 1, c3 + i - 1] = 'QUANT ' + p[i] * 100 + ' %'
    #     Initialisation.Worksheets[nom_feuille_qtriang].Cells[l1, c3 + i - 1].FormulaR1C1 = (
    #         '=IF(' + csd(p[i]) + '<=' + ref + c_pmode + ',' + ref + c_min +
    #         '+SQRT(' + csd(p[i]) + '*(' + ref + c_max + '-' + ref + c_min +
    #         ')*(' + ref + c_mode + '-' + ref + c_min + ')), ' + ref+c_min + 1
    #         + '-SQRT((' + csd(1 - p[i]) + ')*(' + ref + c_max + '-' + ref +
    #         c_min + ')*(' + ref + c_max + '-' + ref + c_mode + ')))')
    # Range(
    #     Initialisation.Worksheets(nom_feuille_qtriang).Cells(l1, c3),
    #     Initialisation.Worksheets(nom_feuille_qtriang).Cells(l1,
    #                                           c3 + len(p) - 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_qtriang).Cells(l1, c3),
    #         Initialisation.Worksheets(nom_feuille_qtriang).Cells(l2, c3 + len(p) - 1)),
    #     Type=xlFillDefault)
    # Initialisation.Worksheets(nom_feuille_qtriang).Cells(1, 1).Select()


def calcul_res(l1, c1, l2, c2, ind_hc, pond_lig_deb, pond_col_deb,
               pond_col_data, pond_col_pcum, l_hc, c_hc, nbvar, ligne_tot, loi,
               titre, pcent, pourcent, data_co, nom_colonne, nom_feuille_res,
               nom_feuille_quant, nom_feuille_pond, nom_feuille, mup, sigmap,
               c_mu, min, max, mode, c_min, triang_ajust, iproc, nbdata,
               data_c):
    """
    Calcul des resultats statistiques finaux.

    @param data_co tableau des donnees exploitees pour calcul HC

    dans les feuilles contenant les quantiles
    @param l1: Numero de la premiere ligne des quantiles
    @param c1: premiere colonne contenant les quantiles issues du
               bootstrap
    @param l2: Numero de la derniere ligne des resultats quantiles
               bootstraps
    @param c2: derniere colonne contenant les quantiles

    dans nom_feuille_pond
    @param pond_lig_deb premiere ligne de donnees numeriques
    @param pond_col_deb premiere colonne de donnees numeriques
    @param pond_col_data colonne des donnees de concentrations
    @param pond_col_pcum colonne des donnees de probabilites cumulees

    dans nom_feuille_res
    @param l_hc: premiere ligne d'affichage des resultats HC
    @param c_hc: premiere colonne d'affichage des resultats HC
    @param ligne_tot: numero de la derniere ligne en cours
    @param ind_hc: indice de reperage dans pourcent du HC a encadrer
                   en gras
    @param nbvar: nombre de points par tirage bootstrap
    @param loi: distribution statistique choisie
                1: empirique, 2: normal, 3: triangulaire
    @param titre: titre des tableaux de resultats HCx%
    @param pourcent: tableau des probabilites x% definissant les HCx%
                     a estimer
    @param pcent: tableau des centiles a calculer sur chaque HCx%
                  issues des tirages bootstrap
    @param nom_feuille_quant: nom de la feuille contenant les quantiles
                             issues du bootstrap
    @param nom_feuille_res: nom de la feuille de resultats
    @param nom_feuille_pond: nom de la feuille contenant la table
                             data_co avec ponderations
    mup, sigmap moyenne et ecar type ponderes des donnees
    min, max, mode parametre de la loi triangulaire ponderee
    @param c_min: numero de colonne du parametre min de la loi
                  triangulaire les parametres max et modes se trouvent
                  respectivement a c_min+1 et c_min+2
    @param triang_ajust: option d'ajustement pour la loi triangulaire:
                         si T ajustement sur les quantiles, sinon sur
                         les probabilites cumulees
    """
    # Application.ScreenUpdating = False
    # Initialisation.Worksheets(nom_feuille_res).Activate()
    # Initialisation.Worksheets(nom_feuille_res).Cells(1, 1).Select()
    """Ecriture du titre"""
    ecrire_titre(titre(loi), nom_feuille_res, l_hc, c_hc, len(pourcent) + 1)
    """
    Affichage titre des lignes du tableau HC et Calcul ecart type
    de HC
    """
    Initialisation.Worksheets[nom_feuille_res].Cells[l_hc + 1, c_hc] = 'HC'
    for i in range(0, len(pourcent)):
        # Initialisation.Worksheets(nom_feuille_res).Cells(l_hc + 1, c_hc + i) = (
        # pourcent[i]*100 & "%")
        # ne fonctionne pas avec les ',' d'où ce qui suit
        Initialisation.Worksheets[nom_feuille_res].Cells[l_hc + 1, c_hc +
                                                         i] = pourcent[i]
        if len(str(pourcent[i])) > 4:
            Initialisation.Worksheets[nom_feuille_res].Cells[
                l_hc + 1, c_hc + i].NumberFormat = '0.0%'
        else:
            Initialisation.Worksheets[nom_feuille_res].Cells[
                l_hc + 1, c_hc + i].NumberFormat = '0%'
    Initialisation.Worksheets[nom_feuille_res].Cells[l_hc + 2,
                                                     c_hc] = 'Best-Estimate'
    Initialisation.Worksheets[nom_feuille_res].Cells[
        l_hc + 3, c_hc] = 'Geo. Stand. Deviation'
    for i in range(0, len(pcent)):
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 3 + i, c_hc] = 'Centile ' + (pcent[i] * 100) + '%'
    nbligne_res = len(pcent) + 3 + 1
    data = 'R' + l1 + 'C[' + c1 - c_hc - 1 + ']:R' + \
        l2 + 'C[' + c1 - c_hc - 1 + ']'
    data = nom_feuille_quant + '!' + data
    Initialisation.Worksheets[nom_feuille_res].Cells[
        l_hc + 3, c_hc + 1].FormulaR1C1 = '=10^(STDEV(' + data + '))'
    HC_be = list()
    """
    Calcul HC best-estimate : different suivant la loi
    Selection des donnees de concentrations utilisees suivant que
    procedure SSWD ou ACT
    """
    data_c = list()
    for i in range(0, nbdata):
        data_c.append(data_co.Item[i].data
                      if iproc == 1 else data_co.Item[i].act)
    if (loi == 1):
        calculer_be_empirique(data_co, pourcent, nom_feuille_pond,
                              pond_lig_deb, pond_col_pcum, HC_be, nbdata,
                              data_c)
    elif (loi == 2):
        calculer_be_normal(data_co, mup, sigmap, pourcent, HC_be, nbdata,
                           data_c)
    elif (loi == 3):
        if triang_ajust is True:
            calculer_be_triang_q(data_c, nom_feuille_pond, pond_lig_deb,
                                 pond_col_deb, pond_col_data, pond_col_pcum,
                                 pourcent, HC_be, min, max, mode, nom_colonne,
                                 nbdata)
        else:
            calculer_be_triang_p(data_c, nom_feuille_pond, pond_lig_deb,
                                 pond_col_deb, pond_col_data, pond_col_pcum,
                                 pourcent, HC_be, min, max, mode, nom_colonne,
                                 nbdata)
    # Initialisation.Worksheets(nom_feuille_res).Activate()
    """Affichage HC best-estimate dans la feuille de resultats"""
    for i in range(0, len(pourcent)):
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 2, c_hc + i].Value = 10**HC_be[i]
    cellule_gras(l_hc + 2, c_hc + 2, l_hc + 2, c_hc + 2)
    """
    calcul percentiles intervalles de confiance :
    empirique : pas de bias correction
    normal et triangulaire : bias correction
    """
    if loi == 1:
        for i in range(0, len(pcent)):
            Initialisation.Worksheets[nom_feuille_res].Cells[
                l_hc + 3 + i, c_hc +
                1].FormulaR1C1 = '=10^(PERCENTILE(' + data + ',' + csd(
                    pcent[i]) + '))'
    else:
        for i in range(0, len(pcent)):
            Initialisation.Worksheets[nom_feuille_res].Cells[
                l_hc + 3 + i, c_hc +
                1].FormulaR1C1 = '=10^(PERCENTILE(' + data + ',' + csd(
                    pcent[i]) + ')-MEDIAN(' + data + '))*R' + l_hc + 2 + 'C'
    # Initialisation.Worksheets(nom_feuille_res).Activate()
    # Range(
    #     Initialisation.Worksheets(nom_feuille_res).Cells(l_hc + 3, c_hc + 1),
    #     Initialisation.Worksheets(nom_feuille_res).Cells(l_hc + nbligne_res - 1,
    #                                       c_hc + 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_res).Cells(l_hc + 3, c_hc + 1),
    #         Initialisation.Worksheets(nom_feuille_res).Cells(l_hc + nbligne_res - 1,
    #                                           c_hc + len(pourcent))),
    #     Type=xlFillDefault)
        encadrer_colonne(nom_feuille_res, l_hc + 1, c_hc + ind_hc,
                         l_hc + nbligne_res - 1, c_hc + ind_hc)
    """Infos supplementaires suivant les distributions"""
    if (loi == 2):
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 2, c_hc + len(pourcent) + 1] = 'Best-Estimate'
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 3, c_hc + len(pourcent) + 1] = 'Geo. Stand. Deviation'
        for i in range(0, len(pcent)):
            Initialisation.Worksheets[nom_feuille_res].Cells[
                l_hc + 3 + i, c_hc + len(pourcent) + 1] = 'Centile ' + (
                    pcent[i] * 100) + '%'
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 1, c_hc + len(pourcent) + 2] = 'GWM'
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 1, c_hc + len(pourcent) + 3] = 'GWSD'
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 2, c_hc + len(pourcent) + 2] = 10**mup
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 2, c_hc + len(pourcent) + 3] = 10**sigmap
        data = 'R' + l1 + 'C[' + c_mu - c_hc - len(
            pourcent) - 2 + ']:R' + l2 + 'C[' + c_mu - c_hc - len(
                pourcent) - 2 + ']'
        data = nom_feuille + '!' + data
        Initialisation.Worksheets[nom_feuille_res].Cells[l_hc + 3, c_hc + len(
            pourcent) + 2].FormulaR1C1 = '=10^STDEV(' + data + ')'
        for i in range(0, len(pcent)):
            Initialisation.Worksheets[nom_feuille_res].Cells[
                l_hc + 3 + i, c_hc + len(pourcent) +
                2].FormulaR1C1 = '=10^(PERCENTILE(' + data + ',' + csd(
                    pcent[i]) + ')-MEDIAN(' + data + '))*R' + l_hc + 2 + 'C'
        # Range(
        #     Initialisation.Worksheets(nom_feuille_res).Cells(l_hc + 3,
        #                                       c_hc + len(pourcent) + 2),
        #     Initialisation.Worksheets(nom_feuille_res).Cells(
        #         l_hc + nbligne_res - 1, c_hc + len(pourcent) + 2)).Select()
        # Selection.AutoFill(
        #     Destination=Range(
        #         Initialisation.Worksheets(nom_feuille_res).Cells(l_hc + 3,
        #                                           c_hc + len(pourcent) + 2),
        #         Initialisation.Worksheets(nom_feuille_res).Cells(l_hc + nbligne_res - 1,
        #                                           c_hc + len(pourcent) + 3)),
        #     Type=xlFillDefault)
    elif (loi == 3):
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 2, c_hc + len(pourcent) + 1] = 'Best-Estimate'
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 3, c_hc + len(pourcent) + 1] = 'Geo. Stand. Deviation'
        for i in range(0, len(pcent)):
            Initialisation.Worksheets[nom_feuille_res].Cells[
                l_hc + 3 + i, c_hc + len(pourcent) + 1] = 'Centile ' + (
                    pcent[i] * 100) + '%'
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 1, c_hc + len(pourcent) + 2] = 'GWMin'
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 1, c_hc + len(pourcent) + 3] = 'GWMax'
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 1, c_hc + len(pourcent) + 4] = 'GWMode'
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 2, c_hc + len(pourcent) + 2] = 10**min
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 2, c_hc + len(pourcent) + 3] = 10**max
        Initialisation.Worksheets[nom_feuille_res].Cells[
            l_hc + 2, c_hc + len(pourcent) + 4] = 10**mode
        data = 'R' + l1 + 'C[' + c_min - c_hc - len(
            pourcent) - 2 + ']:R' + l2 + 'C[' + c_min - c_hc - len(
                pourcent) - 2 + ']'
        data = nom_feuille + '!' + data
        Initialisation.Worksheets[nom_feuille_res].Cells[l_hc + 3, c_hc + len(
            pourcent) + 2].FormulaR1C1 = '=10^STDEV(' + data + ')'
        for i in range(0, len(pcent)):
            Initialisation.Worksheets[nom_feuille_res].Cells[
                l_hc + 3 + i, c_hc + len(pourcent) +
                2].FormulaR1C1 = '=10^(PERCENTILE(' + data + ',' + csd(
                    pcent[i]) + ')-MEDIAN(' + data + '))*R' + l_hc + 2 + 'C'
        # Range(
        #     Initialisation.Worksheets(nom_feuille_res).Cells(l_hc + 3,
        #                                       c_hc + len(pourcent) + 2),
        #     Initialisation.Worksheets(nom_feuille_res).Cells(
        #         l_hc + nbligne_res - 1, c_hc + len(pourcent) + 2)).Select()
        # Selection.AutoFill(
        #     Destination=Range(
        #         Initialisation.Worksheets(nom_feuille_res).Cells(l_hc + 3,
        #                                           c_hc + len(pourcent) + 2),
        #         Initialisation.Worksheets(nom_feuille_res).Cells(l_hc + nbligne_res - 1,
        #                                           c_hc + len(pourcent) + 4)),
        #     Type=xlFillDefault)
    """3. Sauvegarde des resultats (suppression des formules)"""
    ligne_tot = l_hc + nbligne_res - 1
    # Range(Cells(1, 1), Cells(ligne_tot, c_hc + len(pourcent) + 4)).Copy()
    # Range(Cells(1, 1), Cells(ligne_tot,
    #                          c_hc + len(pourcent) + 4)).PasteSpecial(
    #                              Paste=xlValues,
    #                              Operation=xlNone,
    #                              SkipBlanks=False,
    #                              Transpose=False)
    # Initialisation.Worksheets(nom_feuille_res).Cells(1, 1).Select()


def calcul_R2(data_co, loi, mup, sigmap, min, max, mode, nbdata, data_c):
    """
    Calcul de R2 et de Pvalue paired TTest.

    Base sur quantiles ponderes empiriques versus quantiles theoriques
    ponderes: normaux ou triangulaires
    @param data_co: tableau des donnees exploitees pour calcul HCx%
    @param loi: distribution statistique retenue:
                2 pour normal et 3 pour triangulaire
    @param R2: coefficient de determination
    @param Pvalue: Proba paired TTEST comparaison quantiles
                   empiriques/quantiles theoriques
    @param mup: moyenne ponderee des donnees de concentration
    @param sigmap: ecart type ponderee des donnees de concentration
    @param min, max, mode: parametres de la loi trianguliaire ponderee
    @param nbdata: nombre de donnees exploitees pour les calculs HC
    @param data_c: donnees de concentration exploitees pour les
                   calculs HC
    """
    """Calcul de quantiles theoriques, normaux ou triangulaires"""

    resQ = [0.0] * nbdata
    Qth = [0.0] * nbdata
    Pth = [0.0] * nbdata
    dif = [0.0] * nbdata

    if loi == 2:
        for i in range(0, nbdata):
            Qth[i] = i
            # Qth[i] = Application.WorksheetFunction.NormInv(
            #     data_co.Item[i].pcum, mup, sigmap)
            resQ[i] = Qth[i] - data_c[i]
            Pth[i] = i
            # Pth[i] = Application.WorksheetFunction.NormDist(
            #     data_co.Item[i].data, mup, sigmap, True)
            dif[i] = data_co.Item[i].pcum - Pth[i]
            dif[i] = math.fabs(dif[i])
    if loi == 3:
        pmode = (mode - min) / (max - min)
        for i in range(0, nbdata):
            if data_co.Item[i].pcum <= pmode:
                Qth[i] = (min + math.sqrt(data_co.Item[i].pcum * (max - min) *
                                          (mode - min)))
            else:
                Qth[i] = (max - math.sqrt(
                    (1 - data_co.Item[i].pcum) * (max - min) * (max - mode)))
            resQ[i] = Qth[i] - data_c[i]

            if data_co.Item[i].data <= min:
                Pth[i] = 0.0
            else:
                if data_co.Item[i].data <= mode:
                    Pth[i] = ((data_co.Item[i].data - min) ^ 2) / \
                        ((max - min) * (mode - min))
                else:
                    if data_co.Item[i].data <= max:
                        Pth[i] = 1 - ((data_co.Item[i].data - max) ^ 2) / \
                            ((max - min) * (max - mode))
                    else:
                        Pth[i] = 1

            dif[i] = data_co.Item[i].pcum - Pth[i]
            dif[i] = math.fabs(dif[i])
    """
    Calcul variance et R2

    ceci conduit au calcul d'un R2 non pondere(il est max quand aucune
    ponderation n'est appliquee aux donnees, ce qui n'est pas tres
    coherent)
    var_resQ = Application.WorksheetFunction.Var(resQ)
    var_data = Application.WorksheetFunction.Var(data_c)
    calcul variance ponderee des donnees concentrations(calcul deja
    effectue dans le cas de la loi log - normale)
    """
    mu = 0
    for i in range(0, nbdata):
        mu = mu + data_co.Item[i].data * data_co.Item[i].pond
    var_data = 0
    for i in range(0, nbdata):
        var_data = (var_data + data_co.Item[i].pond *
                    (data_co.Item[i].data - mu) ^ 2)
    var_data = var_data * nbdata / (nbdata - 1)
    """calcul variance ponderee des residus"""
    mu = 0
    for i in range(0, nbdata):
        mu = mu + resQ[i] * data_co.Item[i].pond
    var_resQ = 0
    for i in range(0, nbdata):
        var_resQ = var_resQ + data_co.Item[i].pond * (resQ[i] - mu) ^ 2
    var_resQ = var_resQ * nbdata / (nbdata - 1)
    """calcul R2"""
    R2 = 1 - var_resQ / var_data
    """KS dallal wilkinson approximation pvalue"""
    n = nbdata
    KS = max(dif)
    if n < 5:
        Pvalue = 0
    else:
        if n > 100:
            KS = KS * (n / 100) ^ 0.49
            n = 100
        Pvalue = math.exp(-7.01256 * (KS ^ 2) * (n + 2.78019) +
                          2.99587 * KS * math.sqrt(n + 2.78019) - 0.122119 +
                          0.974598 / math.sqrt(n) + 1.67997 / n)
    if Pvalue > 0.1:
        Pvalue = 0.5
    return (R2, Pvalue)


def calculer_be_empirique(data_co, pourcent, nom_feuille, lig_deb, col_pcum,
                          HC_emp, nbdata, data):
    """
    Calcul des HCx% empirique meilleure estimation.

    (independant des runs bootstraps)
    a partir de la probabilite cumulee ponderee empirique
    @param data_co: tableau des donnees exploitees pour le calcul HC
    @param pourcent: table des probabilites x% correspondant aux HCx%
                     calculees
    @param nom_feuille: nom de la feuille contenant data_co_feuil
                        (il s'agit en fait de nom_feuille_pond)
    @param lig_deb: ligne debut donnees numeriques dans nom_feuille
    @param col_pcum: indice de la colonne contenant les probabilites
                     cumulees dans data_co_feuil
    @param iproc: indice de procedure: 1 pour SSWD et 2 pour ACT
    @param nbdata: nombre de donnees exploitees
    @param data: donnees de concentration exploitees pour le calcul HC

    Remarque: les donnees doivent etre classees dans l'ordre croissant
    dans data et data_co
    """
    pcum = list()
    for i in range(0, nbdata):
        pcum[i] = data_co.Item[i].pcum
    rang = list(0, len(pourcent))
    # Calcul de HC_emp
    for i in range(0, len(pourcent)):
        rang[i] = compt_inf(pcum, pourcent[i])
        if (rang[i] == 0):
            HC_emp[i] = data(1)
        elif (rang[i] >= nbdata):
            HC_emp[i] = data(nbdata)
        else:
            HC_emp[i] = (
                data(rang[i] + 1) - (data(rang[i] + 1) - data(rang[i])) *
                (data_co(rang[i] + 1).pcum - pourcent[i]) /
                (data_co.Item(rang[i] + 1).pcum - data_co.Item(rang[i]).pcum))


def calculer_be_normal(data_co, mup, sigmap, pourcent, HC_norm, nbdata, data):
    """
    Calcul des HCp% normaux meilleure estimation.

    (independant des runs bootstrap)
    pour cela, calcul prealable des moyenne et ecart type ponderes
    correspondant aux donnees

    @param data_co: collection de donnees exploitees pour le calcul
                    des HC
    @param mup: moyenne ponderee des donnees de concentration
    @param sigmap: ecart type pondere des donnees de concentration
    @param pourcent: table des probabilites x% correspondant au calcul
                     des HCx%
    @param nbdata: nombre de donnees exploitees
    @param data: donnees de concentration exploitees pour les calculs
                 de HC
    """
    mup = 0
    for i in range(0, nbdata):
        mup = mup + data[i] * data_co.Item[i].pond

    sigmap = 0
    for i in range(0, nbdata):
        sigmap = sigmap + data_co.Item[i].pond * (data[i] - mup) ^ 2
    sigmap = math.sqrt(sigmap * nbdata / (nbdata - 1))

    for i in range(0, len(pourcent)):
        HC_norm[i] = i
        # HC_norm[i] = Application.WorksheetFunction.NormInv(
        #     pourcent[i], mup, sigmap)


def calculer_be_triang_p(data_c, nom_feuille, lig_deb, col_deb, col_data,
                         col_pcum, pourcent, HC_triang, min, max, mode,
                         nom_colonne, nbdata):
    """
    Calcul des HCp% triang meilleure estimation.

    (independant des runs bootstrap)
    pour cela, calcul prealable des parametres min, max et mode,
    ponderes correspondant aux donnees ; estimation par ajustement
    sur les probabilites
    @param nom_colonne: contient les titres des colonnes de
                        data_co_feuil
    @param nom_feuille: nom_feuille_pond
    @param pourcent: table des probabilites x% correspondant au calcul
                     des HCx%
    @param min, max, mode: parametre de la loi triangulaire, ajustee
                           sur donnees ponderees
    @param lig_deb: premiere ligne de donnees numeriques dans
                    nom_feuille
    @param col_deb: premiere colonne de donnees dans nom_feuille
    @param col_data: colonne contenant les concentrations nom_feuille
    @param col_pcum: colonne contenant les probabilites cumulees dans
                     nom_feuille
    @param nbdata: nombre de donnees exploitees pour le calcul des HC
    @param data_c: donnees de concentration exploitees pour le calcul
                   des HC
    """
    # Application.ScreenUpdating = False
    col = col_deb + len(nom_colonne) + 1
    """Definition des valeurs initiale de min, max et mode"""
    min = min(data_c)
    max = max(data_c)
    mode = (min + max) / 2
    # Initialisation.Worksheets[nom_feuille].Activate()
    Initialisation.Worksheets[nom_feuille].Cells[lig_deb, col] = min
    Initialisation.Worksheets[nom_feuille].Cells[lig_deb, col].Name = 'cmin'
    Initialisation.Worksheets[nom_feuille].Cells[lig_deb + 1, col] = max
    Initialisation.Worksheets[nom_feuille].Cells[lig_deb + 1,
                                                 col].Name = 'cmax'
    Initialisation.Worksheets[nom_feuille].Cells[lig_deb + 2, col] = mode
    Initialisation.Worksheets[nom_feuille].Cells[lig_deb + 2,
                                                 col].Name = 'cmode'
    col = col - 1
    """Recherche de la colonne correspondant à data"""
    # data = Initialisation.Worksheets[nom_feuille].Cells(lig_deb, col_data).
    # Address(False, False, xlR1C1, RelativeTo= Cells(lig_deb, col))
    # data = nom_feuille + '!' + data
    """Formule correspondant à la probabilite cumulee triangulaire"""
    # Initialisation.Worksheets[nom_feuille].Cells[lig_deb, col].FormulaR1C1 = (
    #     '=IF(' + data + '<=cmin,0,IF(' + data + '<= cmode ,((' + data +
    #     '-cmin)^2)/((cmax - cmin) * (cmode - cmin))  ,IF(' + data +
    #     '<= cmax ,1-((' + data +
    #     '- cmax)^2)/ ((cmax - cmin) * (cmax - cmode)),1)))')
    # Initialisation.Worksheets[nom_feuille].Cells(lig_deb, col).Select()
    # Selection.AutoFill(Destination=Range(Initialisation.Worksheets[nom_feuille].
    # Cells(lig_deb, col), Initialisation.Worksheets[nom_feuille].Cells(lig_deb + nbdata - 1,
    # col)), Type=xlFillDefault)
    # dataF = Cells(lig_deb, col).Address(True, True, xlR1C1) + ':' +
    # Cells(lig_deb + nbdata - 1, col).Address(True, True, xlR1C1)
    """
    Recherche de la colonne correspondant à probabilite cumulee
    ponderee empirique
    """
    # dataP = Cells(lig_deb, col_pcum).Address(True, True, xlR1C1) + ':' +
    # Cells(lig_deb + nbdata - 1, col_pcum).Address(True, True, xlR1C1)
    """
    Calcul somme à minimiser pour estimer parametre min, max et mode
    puis optimisation par la procedure solver
    """
    # Initialisation.Worksheets[nom_feuille].Cells[
    #     lig_deb + 3, col +
    #     1].FormulaR1C1 = '=SUMXMY2(' + dataP + ',' + dataF + ')'
    # SolverOk(SetCell=Initialisation.Worksheets[nom_feuille].Cells(lig_deb + 3, col + 1),
    # MaxMinVal=2, ValueOf='0', ByChange=Initialisation.Worksheets[nom_feuille].
    # Range(Cells(lig_deb, col + 1), Cells(lig_deb + 2, col + 1)))
    # SolverSolve(UserFinish=True)
    """
    On rapatrie min, max, mode dans programme et on calcule HC_triang
    meilleure estimation correspondant
    """
    col = col + 1
    min = Initialisation.Worksheets[nom_feuille].Cells(lig_deb, col)
    max = Initialisation.Worksheets[nom_feuille].Cells(lig_deb + 1, col)
    mode = Initialisation.Worksheets[nom_feuille].Cells(lig_deb + 2, col)
    pmode = (mode - min) / (max - min)
    for i in range(0, len(pourcent)):
        if (pourcent[i] <= pmode):
            HC_triang[i] = min + math.sqrt(pourcent[i] * (max - min) *
                                           (mode - min))
        else:
            HC_triang[i] = max - math.sqrt(
                (1 - pourcent[i]) * (max - min) * (max - mode))
    """
    On efface la plage de cellules sur laquelle on vient de
    travailler
    """
    # Range(Initialisation.Worksheets[nom_feuille].Cells(lig_deb, col - 1),
    # Initialisation.Worksheets[nom_feuille].Cells(lig_deb + nbdata - 1, col)).Select()
    # Selection.Delete()


def calculer_be_triang_q(data_c, nom_feuille, lig_deb, col_deb, col_data,
                         col_pcum, pourcent, HC_triang, _min, _max, mode,
                         nom_colonne, nbdata):
    """
    Calcul des HCp% triang meilleure estimation.

    (independant des runs bootstrap)
    pour cela, calcul prealable des parametres _min, _max et mode,
    ponderes correspondant aux donnees ; estimation par ajustement
    sur les quantiles

    @param nom_feuille: nom_feuille_pond
    @param pourcent: table des probabilites x% correspondant au calcul
                     des HCx%
    @param _min, _max, mode: parametre de la loi triangulaire, ajustee
                           sur donnees ponderees
    @param lig_deb: premiere ligne de donnees numeriques dans
                    nom_feuille
    @param col_deb: premiere colonne de donnees dans nom_feuille
    @param col_data: colonne contenant les concentrations nom_feuille
    @param col_pcum: colonne contenant les probabilites cumulees dans
                     nom_feuille
    @param nbdata: nombre de donnees exploitees pour le calcul des HC
    @param data_c: donnees de concentration exploitees pour le calcul
                   des HC
    """
    # Application.ScreenUpdating = False
    col = col_deb + len(nom_colonne) + 1
    """Definition des valeurs initiale de _min, max, mode et pmode"""
    _min = min(data_c)
    _max = max(data_c)
    mode = (_min + max) / 2
    # Initialisation.Worksheets[nom_feuille].Activate()
    Initialisation.Worksheets[nom_feuille].Cells[lig_deb, col] = _min
    Initialisation.Worksheets[nom_feuille].Cells[lig_deb, col].Name = 'cmin'
    Initialisation.Worksheets[nom_feuille].Cells[lig_deb + 1, col] = _max
    Initialisation.Worksheets[nom_feuille].Cells[lig_deb + 1,
                                                 col].Name = 'cmax'
    Initialisation.Worksheets[nom_feuille].Cells[lig_deb + 2, col] = mode
    Initialisation.Worksheets[nom_feuille].Cells[lig_deb + 2,
                                                 col].Name = 'cmode'
    # Initialisation.Worksheets[nom_feuille].Cells[lig_deb + 4, col].FormulaR1C1 =
    # '=(cmode - cmin) / (cmax - cmin)'
    Initialisation.Worksheets[nom_feuille].Cells[lig_deb + 4,
                                                 col].Name = 'cpmode'
    col = col - 1
    """Recherche de la colonne correspondant à pcum"""
    # data = Initialisation.Worksheets[nom_feuille].Cells(lig_deb, col_pcum).
    # Address(False, False, xlR1C1, RelativeTo= Cells(lig_deb, col))
    # data = nom_feuille + '!' + data
    """Formule correspondant aux quantiles de la loi triangulaire"""
    # Initialisation.Worksheets[nom_feuille].Cells[lig_deb, col].FormulaR1C1 = (
    #     '=IF(' + data + '<=cpmode,cmin+SQRT(' + data +
    #     '*(cmax - cmin) * (cmode - cmin)),cmax -SQRT((1-' + data +
    #     ')*(cmax - cmin) * (cmax - cmode)))')
    # Initialisation.Worksheets[nom_feuille].Cells[lig_deb, col].Select()
    # Selection.AutoFill(
    #     Destination=Range(Initialisation.Worksheets[nom_feuille].Cells(lig_deb, col),
    #                       Initialisation.Worksheets[nom_feuille].Cells(
    #                           lig_deb + nbdata - 1, col)),
    #     Type=xlFillDefault)
    # dataF = Cells(lig_deb, col).Address(True, True, xlR1C1) + ':' + Cells(
    #     lig_deb + nbdata - 1, col).Address(True, True, xlR1C1)
    """Recherche de la colonne correspondant aux donnees"""
    # dataP = Cells(lig_deb, col_data).Address(True, True, xlR1C1) + ':' +
    # Cells(lig_deb + nbdata - 1, col_data).Address(True, True, xlR1C1)
    """
    Calcul somme à minimiser pour estimer parametre min, max et mode
    puis optimisation par la procedure solver
    """
    # Initialisation.Worksheets[nom_feuille].Cells[lig_deb + 3, col + 1].FormulaR1C1 = (
    #     '=SUMXMY2(' + dataP + ',' + dataF + ')')
    # SolverOk(
    #     SetCell=Initialisation.Worksheets[nom_feuille].Cells(lig_deb + 3, col + 1),
    #     MaxMinVal=2,
    #     ValueOf='0',
    #     ByChange=Initialisation.Worksheets[nom_feuille].Range(
    #         Cells(lig_deb, col + 1), Cells(lig_deb + 2, col + 1)))
    # SolverSolve(UserFinish=True)
    """
    On rapatrie min, max, mode dans le programme et on calcule
    HC_triang meilleure estimation correspondant
    """
    col = col + 1
    _min = Initialisation.Worksheets[nom_feuille].Cells(lig_deb, col)
    _max = Initialisation.Worksheets[nom_feuille].Cells(lig_deb + 1, col)
    mode = Initialisation.Worksheets[nom_feuille].Cells(lig_deb + 2, col)
    pmode = (mode - min) / (max - min)
    for i in range(1, len(pourcent)):
        if (pourcent[i] <= pmode):
            HC_triang[i] = min + math.sqrt(pourcent[i] * (max - min) *
                                           (mode - min))
        else:
            HC_triang[i] = max - math.sqrt(
                (1 - pourcent[i]) * (max - min) * (max - mode))
    """On efface la plage de cellules sur laquelle on vient de travailler"""
    # Range(Initialisation.Worksheets[nom_feuille].Cells(lig_deb, col - 1),
    #       Initialisation.Worksheets[nom_feuille].Cells(lig_deb + nbdata - 1, col)).Select()
    # Selection.Delete()
