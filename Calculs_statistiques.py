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
# @Last modified time: 2017-04-11T11:25:39+02:00

from fct_generales import compt_inf, csd, trier_tirages_feuille
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
    global Worksheets
    # Application.Run('ATPVBAEN.XLA!Random', nom_feuille_stat, nbvar, B, 7,
    #                 Worksheets[nom_feuille_pond].Range(
    #                     Cells(lig_deb, col_data), Cells(lig_fin, col_pond)))
    # Worksheets[nom_feuille_stat].Rows(1).Insert()
    for j in range(0, nbvar):
        Worksheets[nom_feuille_stat].Cells[1, j] = 'POINT ' + j
    # Worksheets[nom_feuille_stat].Cells[1, 1].Select()


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
    Worksheets[nom_feuille_qemp] = Worksheet()
    """Ecriture des entetes de colonnes"""
    for i in range(0, len(p)):
        Worksheets[nom_feuille_qemp].Cells[l1 - 1, c3 + i -
                                           1] = 'QUANT ' + p[i] * 100 + ' %'
    """
    Calcul des quantiles p%
    data = "RC" & c1 & ":RC" & c2
    """
    data = nom_feuille_sort + '!RC'
    for i in range(0, len(p)):
        # Worksheets(nom_feuille_qemp).Cells(l1, c3 + i - 1).FormulaR1C1 =
        # "=PERCENTILE(" & nom_feuille_stat & "!" & data & "," & csd(p(i))
        # & ")"
        if (rang[i] == 0 or rang[i] == nbvar):
            tmp[i] = 0
        else:
            tmp[i] = (pcum[rang[i] + 1] - p[i]) / \
                (pcum[rang[i] + 1] - pcum[rang[i]])
        Worksheets[nom_feuille_qemp].Cells[
            l1, c3 + i -
            1].FormulaR1C1 = '=IF(' + rang[i] + '=0,' + data + c1 + \
            ',IF(' + rang[i] + '>=' + nbvar + ',' + data + c2 + ',' + data + \
            rang[i] + 1 + '-(' + data + rang[i] + 1 + '-' + data + rang[i] +\
            ')*' + csd(tmp[i]) + '))'
    # Range(
    #     Worksheets(nom_feuille_qemp).Cells(l1, c3),
    #     Worksheets(nom_feuille_qemp).Cells(l1, c3 + UBound(p) - 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Worksheets(nom_feuille_qemp).Cells(l1, c3),
    #         Worksheets(nom_feuille_qemp).Cells(l2, c3 + UBound(p) - 1)),
    #     Type=xlFillDefault)
    # Worksheets(nom_feuille_qemp).Cells(1, 1).Select()


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
    # Worksheets(nom_feuille_stat).Activate()
    Worksheets[nom_feuille_stat].Cells[l1 - 1, c_mu] = 'MEAN'
    Worksheets[nom_feuille_stat].Cells[l1 - 1, c_mu + 1] = 'STDEV'
    data = 'RC' + c1 + ':RC' + c2
    """1. Calcul de la moyenne des echantillons"""
    Worksheets[nom_feuille_stat].Cells[
        l1,
        c_mu].FormulaR1C1 = '=AVERAGE(' + nom_feuille_stat + '!' + data + ')'
    """2. Calcul de l'ecart type des echantillons"""
    Worksheets[nom_feuille_stat].Cells[
        l1, c_mu +
        1].FormulaR1C1 = '=STDEV(' + nom_feuille_stat + '!' + data + ')'
    # Range(
    #     Worksheets(nom_feuille_stat).Cells(l1, c_mu),
    #     Worksheets(nom_feuille_stat).Cells(l1, c_mu + 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Worksheets(nom_feuille_stat).Cells(l1, c_mu),
    #         Worksheets(nom_feuille_stat).Cells(l2, c_mu + 1)),
    #     Type=xlFillDefault)
    """
    3. Calcul quantiles normaux correspondant a p() et a mean et
    stdev precedemment calcules
    """
    """Affichage dans nom_feuille_qnorm"""
    Worksheets[nom_feuille_qnorm] = Worksheet()
    for i in range(0, len(p)):
        Worksheets[nom_feuille_qnorm].Cells[l1 - 1, c3 + i -
                                            1] = 'QUANT ' + p[i] * 100 + ' %'
        Worksheets[nom_feuille_qnorm].Cells[
            l1, c3 + i - 1].FormulaR1C1 = '=NORMINV(' + csd(
                p[i]) + ',' + nom_feuille_stat + '!RC' + (
                    c_mu) + ',' + nom_feuille_stat + '!RC' + (c_mu + 1) + ')'
    # Range(
    #     Worksheets(nom_feuille_qnorm).Cells(l1, c3),
    #     Worksheets(nom_feuille_qnorm).Cells(l1, c3 + UBound(p) - 1)).Select()
    # Selection.AutoFill(
    #     Destination=Range(
    #         Worksheets(nom_feuille_qnorm).Cells(l1, c3),
    #         Worksheets(nom_feuille_qnorm).Cells(l2, c3 + UBound(p) - 1)),
    #     Type=xlFillDefault)
    # Worksheets(nom_feuille_qnorm).Cells(1, 1).Select()
