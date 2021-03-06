# coding=utf-8
"""
Calculs statistiques.

Many function to refactor to python function.
"""

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-11T10:01:21+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: statistics.py
# @Last modified by:   gysco
# @Last modified time: 2017-06-16T11:39:51+02:00

import math

import initialisation
from common import cellule_gras, compt_inf, ecrire_titre, trier_tirages_feuille
from numpy import mean, median, percentile, std
from numpy.random import choice, seed
from scipy.stats import norm

from multiprocessing import cpu_count, Pool


def threaded_bootstrap(data, nbvar, B, pond, line_start, nom_feuille_stat):
    """Optimized bootstrap on thread."""
    i = 1
    j = 0
    bootstrap = choice(data, nbvar * B, p=pond)
    for x in bootstrap:
        if j == nbvar:
            j = 0
            i += 1
            initialisation.Worksheets[nom_feuille_stat].Cells \
                .set_value(i + line_start, j, x)
        j += 1


def tirage(nom_feuille_stat, nbvar, B, nom_feuille_pond, col_data, col_pond,
           seed_check):
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
    data = list()
    pond = list()
    if seed_check:
        seed(42)
    for x in range(1, len(initialisation.Worksheets[nom_feuille_pond].Cells)):
        data.append(initialisation.Worksheets[nom_feuille_pond]
                    .Cells.get_value(x, col_data))
        pond.append(initialisation.Worksheets[nom_feuille_pond]
                    .Cells.get_value(x, col_pond))
    for j in range(0, nbvar):
        initialisation.Worksheets[nom_feuille_stat].Cells.set_value(
            0, j, 'POINT ' + str(j + 1))
    """
    thread_number = cpu_count()
    with Pool(thread_number) as p:
        for i in range(0, thread_number):
            p.apply_async(func=threaded_bootstrap, args=(
                data, nbvar, int(B / thread_number), pond,
                int((B / thread_number) * i), nom_feuille_stat,))
    """
    i = 1
    j = 0
    for x in choice(data, nbvar * B, p=pond):
        if j == nbvar:
            j = 0
            i += 1
        initialisation.Worksheets[nom_feuille_stat].Cells.set_value(i, j, x)
        j += 1


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
    trier_tirages_feuille(nom_feuille_stat, nom_feuille_sort, nbvar)
    """Creation de la feuille contenant les quantiles empiriques"""
    """Ecriture des entetes de colonnes"""
    for i in range(0, len(p)):
        initialisation.Worksheets[nom_feuille_qemp].Cells.set_value(
            l1 - 1, c3 + i, 'QUANT ' + str(p[i] * 100) + ' %')
    """
    Calcul des quantiles p%
    data = "RC" & c1 & ":RC" & c2
    """
    for y in range(1, l2 + 1):
        tmp = list()
        for i in range(0, len(p)):
            if rang[i] == 0 or rang[i] == nbvar:
                tmp.append(0)
            else:
                tmp.append((pcum[rang[i]] - p[i]) /
                           (pcum[rang[i]] - pcum[rang[i] - 1]))
            if rang[i] == 0:
                set_data = initialisation.Worksheets[
                    nom_feuille_sort].Cells.get_value(y, c1)
            elif rang[i] >= nbvar:
                set_data = initialisation.Worksheets[
                    nom_feuille_sort].Cells.get_value(y, c2)
            else:
                set_data = (initialisation.Worksheets[nom_feuille_sort]
                            .Cells.get_value(y, rang[i]) -
                            (initialisation.Worksheets[nom_feuille_sort]
                             .Cells.get_value(y, rang[i]) -
                             initialisation.Worksheets[nom_feuille_sort]
                             .Cells.get_value(y, rang[i] - 1)) * tmp[i])
            initialisation.Worksheets[nom_feuille_qemp].Cells.set_value(
                y, c3 + i, set_data)


def calcul_ic_normal(l1, c1, l2, c2, c3, p, nom_feuille_stat,
                     nom_feuille_qnorm):
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
    dans nom_feuille_qnorm
        @param c3: premiere colonne affichage resultats normaux
    """
    c_mu = c2 + 2
    """
    Calcul moyenne et ecart type correspondant a chaque tirage
    (chaque ligne de nom_feuille_stat)
    on travaille dans nom_feuille_stat
    """
    initialisation.Worksheets[nom_feuille_stat].Cells.set_value(
        l1 - 1, c_mu, 'MEAN')
    initialisation.Worksheets[nom_feuille_stat].Cells.set_value(
        l1 - 1, c_mu + 1, 'STDEV')
    for i in range(l1, l2 + 1):
        data = list()
        for j in range(c1, c2):
            data.append(
                float(initialisation.Worksheets[nom_feuille_stat]
                      .Cells.get_value(i, j)))
        """1. Calcul de la moyenne des echantillons"""
        initialisation.Worksheets[nom_feuille_stat].Cells.set_value(
            i, c_mu, mean(data))
        """2. Calcul de l'ecart type des echantillons"""
        initialisation.Worksheets[nom_feuille_stat].Cells.set_value(
            i, c_mu + 1, std(data))
    """
    3. Calcul quantiles normaux correspondant a p() et a mean et
    stdev precedemment calcules
    """
    """Affichage dans nom_feuille_qnorm"""
    for i in range(0, len(p)):
        initialisation.Worksheets[nom_feuille_qnorm].Cells.set_value(
            l1 - 1, c3 + i, 'QUANT ' + str(p[i] * 100) + ' %')
        for x in range(l1, l2 + 1):
            initialisation.Worksheets[nom_feuille_qnorm].Cells.set_value(
                x, c3 + i,
                norm.ppf(
                    p[i],
                    loc=initialisation.Worksheets[nom_feuille_stat]
                    .Cells.get_value(x, c_mu),
                    scale=initialisation.Worksheets[nom_feuille_stat]
                    .Cells.get_value(x, c_mu + 1)))
    return c_mu


def calcul_ic_triang_p(l1, c1, l2, c2, c3, nbvar, a, p, nom_feuille_stat,
                       nom_feuille_sort, nom_feuille_Ftriang,
                       nom_feuille_qtriang):
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
    indic = 0
    """
    On trie les donnees a exploiter (issues des tirages aleatoires)
    dans l'ordre croissant
    Creation de la feuille nom_feuille_sort si pas existante
    (pour empirique)
    """
    for ws in initialisation.Worksheets:
        if ws.Name == nom_feuille_sort:
            indic = 1
    if indic == 0:
        trier_tirages_feuille(nom_feuille_stat, nom_feuille_sort, nbvar)
    """
    On calcule les probabilites cumulees empiriques que l'on affiche
    dans la premiere ligne et on met en place les formules de
    probabilite triangulaire qui seront comparees aux probabilites
    empiriques ; creation de la feuille nom_feuille_Ftriang

    On initialise le solver en prennant le min et le max de chaque
    serie tiree et on calcule mode=(min+max)/2
    """
    c_min = c3 + nbvar + 1
    c_max = c_min + 1
    c_mode = c_max + 1
    initialisation.Worksheets[nom_feuille_Ftriang].Cells.set_value(
        l1 - 1, c_min, 'min')
    initialisation.Worksheets[nom_feuille_Ftriang].Cells.set_value(
        l1 - 1, c_max, 'max')
    initialisation.Worksheets[nom_feuille_Ftriang].Cells.set_value(
        l1 - 1, c_mode, 'mode')
    for i in range(l1, l2):
        initialisation.Worksheets[nom_feuille_Ftriang].Cells.set_value(
            i, c_min,
            initialisation.Worksheets[nom_feuille_sort].Cells.get_value(i, c3))
        initialisation.Worksheets[nom_feuille_Ftriang].Cells.set_value(
            i, c_max,
            initialisation.Worksheets[nom_feuille_sort].Cells.get_value(
                i, c3 + nbvar - 1))
        initialisation.Worksheets[nom_feuille_Ftriang].Cells.set_value(
            i, c_mode,
            (initialisation.Worksheets[nom_feuille_Ftriang].Cells.get_value(
                i, c_min) + initialisation.Worksheets[nom_feuille_Ftriang]
             .Cells.get_value(i, c_max)) / 2)
    """Calcul probabilites empiriques et theoriques pour ajustement"""
    for i in range(0, nbvar):
        initialisation.Worksheets[nom_feuille_Ftriang].Cells.set_value(
            l1 - 1, c3 + i - 1, (i - a) / (nbvar + 1 - 2 * a))
    # data = Cells(l1, c3).Address(
    #     False, False, xlR1C1, RelativeTo=Cells(l1, c3))
    # data1 = nom_feuille_sort + '!' + data
    # Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1,
    # c3].FormulaR1C1 = (
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
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1,
    # c3 + nbvar - 1)),
    #     Type=xlFillDefault)
    # Range(
    #     Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3),
    #     Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3 + nbvar
    #  - 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3),
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l2,
    # c3 + nbvar - 1)),
    #     Type=xlFillDefault)
    """
    On calcule la somme des carres des differences entre probabilites
    empiriques et probabilites theoriques triangulaires,
    pour ajustement
    """
    c_ssr = c_mode + 1
    # Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1, c_ssr] =
    # 'Sum Square Res'
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
    initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1,
                                                         c_pmode] = 'pmode'
    for i in range(l1, l2):
        # SolverOk(
        #     SetCell=Initialisation.Worksheets(nom_feuille_Ftriang).Cells(i,
        #  c_ssr),
        #     MaxMinVal=2,
        #     ValueOf='0',
        #     ByChange=Range(
        #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(i,
        # c_min),
        #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(i,
        # c_min + 2)))
        # SolverSolve(UserFinish=True)
        initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_pmode] = (
            initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_mode] -
            initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_min]
        ) / (initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_max] -
             initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_min])
    initialisation.Worksheets[nom_feuille_Ftriang].Cells(1, 1).Select()
    """
    Calcul des quantiles correspondant a la loi triangulaire dont les
    parametres min, max et mode viennent d'etre estimes ; creation de
    la feuille nom_feuille_qtriang
    """
    ref = nom_feuille_Ftriang + '!RC'
    for i in range(0, len(p)):
        initialisation.Worksheets[nom_feuille_qtriang].Cells[
            l1 - 1, c3 + i - 1] = 'QUANT ' + str(p[i] * 100) + ' %'
        initialisation.Worksheets[nom_feuille_qtriang].Cells[
            l1, c3 + i - 1].FormulaR1C1 = (
                '=IF(' + p[i] + '<=' + ref + c_pmode + ',' + ref + c_min +
                '+SQRT(' + p[i] + '*(' + ref + c_max + '-' + ref + c_min +
                ')*(' + ref + c_mode + '-' + ref + c_min + ')), ' + ref + c_max
                + '-SQRT((' + 1 - p[i] + ')*(' + ref + c_max + '-' + ref +
                c_min + ')*(' + ref + c_max + '-' + ref + c_mode + ')))')
    # Range(
    #     Initialisation.Worksheets(nom_feuille_qtriang).Cells(l1, c3),
    #     Initialisation.Worksheets(nom_feuille_qtriang).Cells(l1,
    #                                           c3 + len(p) - 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_qtriang).Cells(l1, c3),
    #         Initialisation.Worksheets(nom_feuille_qtriang).Cells(l2,
    # c3 + len(p) - 1)),
    #     Type=xlFillDefault)
    # Initialisation.Worksheets(nom_feuille_qtriang).Cells(1, 1).Select()
    return c_min


def calcul_ic_triang_q(l1, c1, l2, c2, c3, nbvar, a, p, nom_feuille_stat,
                       nom_feuille_sort, nom_feuille_Ftriang,
                       nom_feuille_qtriang):
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
    """
    On trie les donnees a exploiter (issues des tirages aleatoires)
    dans l'ordre croissant
    Creation de la feuille nom_feuille_sort si pas deja existant
    (pour empirique)
    """
    for ws in initialisation.Worksheets:
        if ws.Name == nom_feuille_sort:
            indic = 1
    if indic == 0:
        trier_tirages_feuille(nom_feuille_stat, nom_feuille_sort, nbvar)
    """
    On calcule les probabilites cumulees empiriques que l'on affiche
    dans la premiere ligne et on met en place les formules de quantile
    triangulaire qui seront comparees aux valeurs empiriques ;
    creation de la feuille nom_feuille_Ftriang

    On initialise le solver en prennant le min et le max de chaque
    serie tiree et on calcule mode=(min+max)/2 puis pmode
    """
    c_min = c3 + nbvar + 1
    c_max = c_min + 1
    c_mode = c_max + 1
    c_ssr = c_mode + 1
    c_pmode = c_ssr + 1
    initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1, c_min] = 'min'
    initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1, c_max] = 'max'
    initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1,
                                                         c_mode] = 'mode'
    initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1 - 1,
                                                         c_pmode] = 'pmode'
    for i in range(l1, l2):
        initialisation.Worksheets[nom_feuille_Ftriang].Cells[
            i, c_min] = initialisation.Worksheets[nom_feuille_sort].Cells[i,
                                                                          c3]
        initialisation.Worksheets[nom_feuille_Ftriang].Cells[
            i, c_max] = initialisation.Worksheets[nom_feuille_sort].Cells[
                i, c3 + nbvar - 1]
        initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_mode] = (
            initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_min] +
            initialisation.Worksheets[nom_feuille_Ftriang].Cells[i, c_max]) / 2
    # Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1,
    # c_pmode].FormulaR1C1 = (
    #     '=(' + ref + c_mode + '-' + ref + c_min + ')/(' + ref + c_max + '-' +
    #     ref + c_min + ')')
    # Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c_pmode)
    #   .Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c_pmode)
    #         ,Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l2,
    # c_pmode)),
    #     Type=xlFillDefault)
    """
    Calcul probabilites empiriques et quantiles triangulaires
    correspondants
    """
    for i in range(0, nbvar):
        initialisation.Worksheets[nom_feuille_Ftriang].Cells[
            l1 - 1, c3 + i - 1] = (i - a) / (nbvar + 1 - 2 * a)
    # Initialisation.Worksheets[nom_feuille_Ftriang].Cells[l1,
    # c3].FormulaR1C1 = (
    #     '=IF(' + ref2 + l1 - 1 + 'C<=' + ref + c_pmode + ',' + ref + c_min +
    #     '+SQRT(' + ref2 + l1 - 1 + 'C*(' + ref + c_max + '-' + ref + c_min +
    #     ')' + '*(' + ref + c_mode + '-' + ref + c_min + ')),' + ref + c_max +
    #     '-SQRT((1-' + ref2 + l1 - 1 + 'C)*(' + ref + c_max + '-' + ref +c_min
    #     + ')' + '*(' + ref + c_max + '-' + ref + c_mode + ')))')
    # Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3),
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1,
    # c3 + nbvar - 1)),
    #     Type=xlFillDefault)
    # Range(
    #     Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3),
    #     Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3 + nbvar
    #  - 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l1, c3),
    #         Initialisation.Worksheets(nom_feuille_Ftriang).Cells(l2,
    # c3 + nbvar - 1)),
    #     Type=xlFillDefault)
    """
    On calcule la somme des carres des differences entre donnees
    empiriques etquantiles theoriques triangulaires, pour ajustement
    """
    initialisation.Worksheets[nom_feuille_Ftriang].Cells[
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
    #         SetCell=Initialisation.Worksheets(nom_feuille_Ftriang).Cells(i,
    #  c_ssr),
    #         MaxMinVal=2,
    #         ValueOf='0',
    #         ByChange=Range(
    #             Initialisation.Worksheets(nom_feuille_Ftriang).Cells(i,
    # c_min),
    #             Initialisation.Worksheets(nom_feuille_Ftriang).Cells(i,
    # c_min + 2)))
    #     SolverSolve(UserFinish=True)
    initialisation.Worksheets[nom_feuille_Ftriang].Cells(1, 1).Select()
    """
    Calcul des quantiles correspondant a la loi triangulaire dont les
    parametres min, max et mode viennent d'etre estimes ; creation de
    la feuillenom_feuille_qtriang
    """
    for i in range(0, len(p)):
        initialisation.Worksheets[nom_feuille_qtriang].Cells[
            l1 - 1, c3 + i - 1] = 'QUANT ' + p[i] * 100 + ' %'
    # Initialisation.Worksheets[nom_feuille_qtriang].Cells[l1, c3 + i -
    # 1].FormulaR1C1 = (
    #         '=IF(' + p[i] + '<=' + ref + c_pmode + ',' + ref + c_min +
    #         '+SQRT(' + p[i] + '*(' + ref + c_max + '-' + ref + c_min +
    #         ')*(' + ref + c_mode + '-' + ref + c_min + ')), ' + ref+c_min + 1
    #         + '-SQRT((' + 1 - p[i] + ')*(' + ref + c_max + '-' + ref +
    #         c_min + ')*(' + ref + c_max + '-' + ref + c_mode + ')))')
    # Range(
    #     Initialisation.Worksheets(nom_feuille_qtriang).Cells(l1, c3),
    #     Initialisation.Worksheets(nom_feuille_qtriang).Cells(l1,
    #                                           c3 + len(p) - 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Initialisation.Worksheets(nom_feuille_qtriang).Cells(l1, c3),
    #         Initialisation.Worksheets(nom_feuille_qtriang).Cells(l2,
    # c3 + len(p) - 1)),
    #     Type=xlFillDefault)
    # Initialisation.Worksheets(nom_feuille_qtriang).Cells(1, 1).Select()
    return c_min


def calcul_res(l1, l2, ind_hc, pond_lig_deb, pond_col_deb, pond_col_data,
               pond_col_pcum, l_hc, c_hc, nbvar, loi, titre, pcent, pourcent,
               data_co, nom_colonne, nom_feuille_res, nom_feuille_quant,
               nom_feuille_pond, nom_feuille, c_min, triang_ajust, iproc,
               nbdata):
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
    _min = 0
    _max = 0
    mode = 0
    mup = 0
    sigmap = 0
    """Ecriture du titre"""
    ecrire_titre(titre[loi - 1], nom_feuille_res, l_hc, c_hc)
    """
    Affichage titre des lignes du tableau HC et Calcul ecart type
    de HC
    """
    initialisation.Worksheets[nom_feuille_res].Cells.set_value(
        l_hc + 1, c_hc, 'HC')
    for i in range(0, len(pourcent)):
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 1, c_hc + i + 1, "={:.3g}".format(pourcent[i]))
    initialisation.Worksheets[nom_feuille_res].Cells.set_value(
        l_hc + 2, c_hc, 'Best-Estimate')
    initialisation.Worksheets[nom_feuille_res].Cells.set_value(
        l_hc + 3, c_hc, 'Geo. Stand. Deviation')
    for i in range(0, len(pcent)):
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 4 + i, c_hc, 'Centile ' + str(pcent[i] * 100) + '%')
    HC_be = list()
    """
    Calcul HC best-estimate : different suivant la loi
    Selection des donnees de concentrations utilisees suivant que
    procedure SSWD ou ACT
    """
    data_c = list()
    for i in range(0, nbdata):
        data_c.append(data_co[i].data if iproc == 1 else data_co[i].act)
    if loi == 1:
        calculer_be_empirique(data_co, pourcent, HC_be, nbdata, data_c)
    elif loi == 2:
        mup, sigmap = calculer_be_normal(data_co, pourcent, HC_be, nbdata,
                                         data_c)
    elif loi == 3:
        _min, _max, mode = (calculer_be_triang_q(
            data_c, nom_feuille_pond, pond_lig_deb, pond_col_deb,
            pond_col_data, pond_col_pcum, pourcent, HC_be, nom_colonne, nbdata)
                            if triang_ajust is True else calculer_be_triang_p(
                                data_c, nom_feuille_pond, pond_lig_deb,
                                pond_col_deb, pond_col_pcum, pourcent, HC_be,
                                nom_colonne, nbdata))
    """Affichage HC best-estimate dans la feuille de resultats"""
    for i in range(0, len(pourcent)):
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 2, c_hc + 1 + i, 10**HC_be[i])
    cellule_gras()
    """
    calcul percentiles intervalles de confiance :
    empirique : pas de bias correction
    normal et triangulaire : bias correction
    """
    for x in range(0, len(pourcent)):
        data = list()
        for y in range(1, l2):
            data.append(initialisation.Worksheets[nom_feuille_quant]
                        .Cells.get_value(y, x))
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 3, c_hc + 1 + x, 10**std(data))
        if loi == 1:
            for i in range(0, len(pcent)):
                initialisation.Worksheets[nom_feuille_res].Cells.set_value(
                    l_hc + 4 + i, c_hc + 1 + x, 10**percentile(
                        data, pcent[i] * 100))
        else:
            for i in range(0, len(pcent)):
                initialisation.Worksheets[nom_feuille_res].Cells.set_value(
                    l_hc + 4 + i, c_hc + 1 + x,
                    (10**(percentile(data, pcent[i] * 100) - median(data)) *
                     initialisation.Worksheets[nom_feuille_res]
                     .Cells.get_value(l_hc + 2, c_hc + 1 + x)))
    """Infos supplementaires suivant les distributions"""
    if loi == 2:
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 2, c_hc + len(pourcent) + 1, 'Best-Estimate')
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 3, c_hc + len(pourcent) + 1, 'Geo. Stand. Deviation')
        for i in range(0, len(pcent)):
            initialisation.Worksheets[nom_feuille_res].Cells.set_value(
                l_hc + 4 + i, c_hc + len(pourcent) + 1,
                'Centile ' + str(pcent[i] * 100) + '%')
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 1, c_hc + len(pourcent) + 2, 'GWM')
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 1, c_hc + len(pourcent) + 3, 'GWSD')
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 2, c_hc + len(pourcent) + 2, 10**mup)
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 2, c_hc + len(pourcent) + 3, 10**sigmap)
        for x in [nbvar + 1, nbvar + 2]:
            data = list()
            for y in range(l1, l2):
                data.append(initialisation.Worksheets[nom_feuille]
                            .Cells.get_value(y, x))
            initialisation.Worksheets[nom_feuille_res].Cells.set_value(
                l_hc + 3, c_hc + len(pourcent) + 1 + x - nbvar, 10**std(data))
            for i in range(0, len(pcent)):
                initialisation.Worksheets[nom_feuille_res].Cells.set_value(
                    l_hc + 4 + i, c_hc + len(pourcent) + 1 + x - nbvar,
                    (10**(percentile(data, pcent[i] * 100) - median(data)) *
                     initialisation.Worksheets[
                         nom_feuille_res].Cells.get_value(
                             l_hc + 2, c_hc + len(pourcent) + 1 + x - nbvar)))
    if loi == 3:
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 2, c_hc + len(pourcent) + 1, 'Best-Estimate')
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 3, c_hc + len(pourcent) + 1, 'Geo. Stand. Deviation')
        for i in range(0, len(pcent)):
            initialisation.Worksheets[nom_feuille_res].Cells.set_value(
                l_hc + 4 + i, c_hc + len(pourcent) + 1,
                'Centile ' + str(pcent[i] * 100) + '%')
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 1, c_hc + len(pourcent) + 2, 'GWMin')
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 1, c_hc + len(pourcent) + 3, 'GWMax')
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 1, c_hc + len(pourcent) + 4, 'GWMode')
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 2, c_hc + len(pourcent) + 2, 10**_min)
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 2, c_hc + len(pourcent) + 3, 10**_max)
        initialisation.Worksheets[nom_feuille_res].Cells.set_value(
            l_hc + 2, c_hc + len(pourcent) + 4, 10**mode)
        for x in [c_min, c_min + 1, c_min + 2]:
            data = list()
            for y in range(l1, l2):
                data.append(initialisation.Worksheets[nom_feuille]
                            .Cells.get_value(y, x))
            initialisation.Worksheets[nom_feuille_res].Cells.set_value(
                l_hc + 3, c_hc + len(pourcent) + 1 + x - c_min, 10**std(data))
            for i in range(0, len(pcent)):
                initialisation.Worksheets[nom_feuille_res].Cells.set_value(
                    l_hc + 4 + i, c_hc + len(pourcent) + 1 + x - c_min,
                    (10**(percentile(data, pcent[i] * 100) - median(data)) *
                     initialisation.Worksheets[
                         nom_feuille_res].Cells.get_value(
                             l_hc + 2, c_hc + len(pourcent) + 2 + x - c_min)))
    return mup, sigmap, _min, _max, mode, data_c


def calcul_R2(data_co, loi, mup, sigmap, _min, _max, mode, nbdata, data_c):
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
            Qth[i] = norm.ppf(data_co[i].pcum, mup, sigmap)
            resQ[i] = Qth[i] - data_c[i]
            Pth[i] = norm.cdf(data_co[i].data, mup, sigmap)
            dif[i] = data_co[i].pcum - Pth[i]
            dif[i] = math.fabs(dif[i])
    if loi == 3:
        pmode = (mode - _min) / (_max - _min)
        for i in range(0, nbdata):
            if data_co[i].pcum <= pmode:
                Qth[i] = (_min + math.sqrt(data_co[i].pcum * (_max - _min) *
                                           (mode - _min)))
            else:
                Qth[i] = (_max - math.sqrt(
                    (1 - data_co[i].pcum) * (_max - _min) * (_max - mode)))
            resQ[i] = Qth[i] - data_c[i]

            if data_co[i].data <= _min:
                Pth[i] = 0.0
            else:
                if data_co[i].data <= mode:
                    Pth[i] = ((data_co[i].data - _min) ** 2.) / \
                             ((_max - _min) * (mode - _min))
                else:
                    if data_co[i].data <= _max:
                        Pth[i] = 1 - ((data_co[i].data - _max) ** 2.) / \
                                     ((_max - _min) * (_max - mode))
                    else:
                        Pth[i] = 1

            dif[i] = data_co[i].pcum - Pth[i]
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
        mu = mu + data_co[i].data * data_co[i].pond
    var_data = 0
    for i in range(0, nbdata):
        var_data = (var_data + data_co[i].pond * (data_co[i].data - mu)**2.)
    var_data = var_data * nbdata / (nbdata - 1)
    """calcul variance ponderee des residus"""
    mu = 0
    for i in range(0, nbdata):
        mu += resQ[i] * data_co[i].pond
    var_resQ = 0
    for i in range(0, nbdata):
        var_resQ += data_co[i].pond * (resQ[i] - mu)**2.
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
            KS *= (n / 100)**0.49
            n = 100
        Pvalue = math.exp(-7.01256 * (KS**2) * (n + 2.78019) +
                          2.99587 * KS * math.sqrt(n + 2.78019) - 0.122119 +
                          0.974598 / math.sqrt(n) + 1.67997 / n)
    if Pvalue > 0.1:
        Pvalue = 0.5
    return R2, Pvalue


def calculer_be_empirique(data_co, pourcent, HC_emp, nbdata, data):
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
        pcum.append(data_co[i].pcum)
    rang = list()
    """Calcul de HC_emp"""
    for i in range(0, len(pourcent)):
        rang.append(compt_inf(pcum, pourcent[i]))
        if rang[i] == 0:
            HC_emp.append(data[1])
        elif rang[i] >= nbdata:
            HC_emp.append(data[nbdata - 1])
        else:
            HC_emp.append(data[rang[i]] - (data[rang[i]] - data[rang[i] - 1]) *
                          (data_co[rang[i]].pcum - pourcent[i]) /
                          (data_co[rang[i]].pcum - data_co[rang[i] - 1].pcum))


def calculer_be_normal(data_co, pourcent, HC_norm, nbdata, data):
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
        mup = mup + data[i] * data_co[i].pond
    sigmap = 0
    for i in range(0, nbdata):
        sigmap = sigmap + data_co[i].pond * (data[i] - mup)**2
    sigmap = math.sqrt(sigmap * nbdata / (nbdata - 1))
    for i in range(0, len(pourcent)):
        HC_norm.append(norm.ppf(pourcent[i], mup, sigmap))
    return mup, sigmap


def calculer_be_triang_p(data_c, nom_feuille, lig_deb, col_deb, col_pcum,
                         pourcent, HC_triang, nom_colonne, nbdata):
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
    _min = min(data_c)
    _max = max(data_c)
    mode = (min + max) / 2
    # Initialisation.Worksheets[nom_feuille].Activate()
    initialisation.Worksheets[nom_feuille].Cells[lig_deb, col] = _min
    initialisation.Worksheets[nom_feuille].Cells[lig_deb, col].Name = 'cmin'
    initialisation.Worksheets[nom_feuille].Cells[lig_deb + 1, col] = _max
    initialisation.Worksheets[nom_feuille].Cells[lig_deb + 1,
                                                 col].Name = 'cmax'
    initialisation.Worksheets[nom_feuille].Cells[lig_deb + 2, col] = mode
    initialisation.Worksheets[nom_feuille].Cells[lig_deb + 2,
                                                 col].Name = 'cmode'
    col -= 1
    """Recherche de la colonne correspondant à data"""
    # data = Initialisation.Worksheets[nom_feuille].Cells(lig_deb, col_data).
    # Address(False, False, xlR1C1, RelativeTo= Cells(lig_deb, col))
    # data = nom_feuille + '!' + data
    """Formule correspondant à la probabilite cumulee triangulaire"""
    # Initialisation.Worksheets[nom_feuille].Cells[lig_deb, col].FormulaR1C1 =
    #     ('=IF(' + data + '<=cmin,0,IF(' + data + '<= cmode ,((' + data +
    #     '-cmin)^2)/((cmax - cmin) * (cmode - cmin))  ,IF(' + data +
    #     '<= cmax ,1-((' + data +
    #     '- cmax)^2)/ ((cmax - cmin) * (cmax - cmode)),1)))')
    # Initialisation.Worksheets[nom_feuille].Cells(lig_deb, col).Select()
    # Selection.AutoFill(Destination=Range(Initialisation.Worksheets[
    # nom_feuille].
    # Cells(lig_deb, col), Initialisation.Worksheets[nom_feuille].Cells(
    # lig_deb + nbdata - 1,
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
    # SolverOk(SetCell=Initialisation.Worksheets[nom_feuille].Cells(lig_deb +
    #  3, col + 1),
    # MaxMinVal=2, ValueOf='0', ByChange=Initialisation.Worksheets[nom_feuille]
    # .Range(Cells(lig_deb, col + 1), Cells(lig_deb + 2, col + 1)))
    # SolverSolve(UserFinish=True)
    """
    On rapatrie min, max, mode dans programme et on calcule HC_triang
    meilleure estimation correspondant
    """
    col += 1
    _min = initialisation.Worksheets[nom_feuille].Cells(lig_deb, col)
    _max = initialisation.Worksheets[nom_feuille].Cells(lig_deb + 1, col)
    mode = initialisation.Worksheets[nom_feuille].Cells(lig_deb + 2, col)
    pmode = (mode - _min) / (_max - _min)
    for i in range(0, len(pourcent)):
        if pourcent[i] <= pmode:
            HC_triang[i] = _min + math.sqrt(pourcent[i] * (_max - _min) *
                                            (mode - _min))
        else:
            HC_triang[i] = _max - math.sqrt(
                (1 - pourcent[i]) * (_max - _min) * (_max - mode))
    """
    On efface la plage de cellules sur laquelle on vient de
    travailler
    """
    # Range(Initialisation.Worksheets[nom_feuille].Cells(lig_deb, col - 1),
    # Initialisation.Worksheets[nom_feuille].Cells(lig_deb + nbdata - 1,
    # col)).Select()
    # Selection.Delete()
    return _min, _max, mode


def calculer_be_triang_q(data_c, nom_feuille, lig_deb, col_deb, col_data,
                         col_pcum, pourcent, HC_triang, nom_colonne, nbdata):
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
    mode = (_min + _max) / 2
    # Initialisation.Worksheets[nom_feuille].Activate()
    initialisation.Worksheets[nom_feuille].Cells[lig_deb, col] = _min
    initialisation.Worksheets[nom_feuille].Cells[lig_deb, col].Name = 'cmin'
    initialisation.Worksheets[nom_feuille].Cells[lig_deb + 1, col] = _max
    initialisation.Worksheets[nom_feuille].Cells[lig_deb + 1,
                                                 col].Name = 'cmax'
    initialisation.Worksheets[nom_feuille].Cells[lig_deb + 2, col] = mode
    initialisation.Worksheets[nom_feuille].Cells[lig_deb + 2,
                                                 col].Name = 'cmode'
    # Initialisation.Worksheets[nom_feuille].Cells[lig_deb + 4,
    # col].FormulaR1C1 =
    # '=(cmode - cmin) / (cmax - cmin)'
    initialisation.Worksheets[nom_feuille].Cells[lig_deb + 4,
                                                 col].Name = 'cpmode'
    col -= 1
    """Recherche de la colonne correspondant à pcum"""
    # data = Initialisation.Worksheets[nom_feuille].Cells(lig_deb, col_pcum).
    # Address(False, False, xlR1C1, RelativeTo= Cells(lig_deb, col))
    # data = nom_feuille + '!' + data
    """Formule correspondant aux quantiles de la loi triangulaire"""
    # Initialisation.Worksheets[nom_feuille].Cells[lig_deb, col].FormulaR1C1 =
    #     ('=IF(' + data + '<=cpmode,cmin+SQRT(' + data +
    #     '*(cmax - cmin) * (cmode - cmin)),cmax -SQRT((1-' + data +
    #     ')*(cmax - cmin) * (cmax - cmode)))')
    # Initialisation.Worksheets[nom_feuille].Cells[lig_deb, col].Select()
    # Selection.AutoFill(
    #     Destination=Range(Initialisation.Worksheets[nom_feuille].Cells(
    # lig_deb, col),
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
    # Initialisation.Worksheets[nom_feuille].Cells[lig_deb + 3,
    # col + 1].FormulaR1C1 = (
    #     '=SUMXMY2(' + dataP + ',' + dataF + ')')
    # SolverOk(
    #     SetCell=Initialisation.Worksheets[nom_feuille].Cells(lig_deb + 3,
    # col + 1),
    #     MaxMinVal=2,
    #     ValueOf='0',
    #     ByChange=Initialisation.Worksheets[nom_feuille].Range(
    #         Cells(lig_deb, col + 1), Cells(lig_deb + 2, col + 1)))
    # SolverSolve(UserFinish=True)
    """
    On rapatrie min, max, mode dans le programme et on calcule
    HC_triang meilleure estimation correspondant
    """
    col += 1
    _min = initialisation.Worksheets[nom_feuille].Cells(lig_deb, col)
    _max = initialisation.Worksheets[nom_feuille].Cells(lig_deb + 1, col)
    mode = initialisation.Worksheets[nom_feuille].Cells(lig_deb + 2, col)
    pmode = (mode - _min) / (_max - _min)
    for i in range(1, len(pourcent)):
        if pourcent[i] <= pmode:
            HC_triang[i] = _min + math.sqrt(pourcent[i] * (_max - _min) *
                                            (mode - _min))
        else:
            HC_triang[i] = _max - math.sqrt(
                (1 - pourcent[i]) * (_max - _min) * (_max - mode))
    """On efface la plage de cellules sur laquelle on vient de travailler"""
    # Range(Initialisation.Worksheets[nom_feuille].Cells(lig_deb, col - 1),
    #       Initialisation.Worksheets[nom_feuille].Cells(lig_deb + nbdata -
    # 1, col)).Select()
    # Selection.Delete()
    return _min, _max, mode
