"""
Les fonctions les plus utilisees.

A inclure dans la quasi totatilite des autres.
"""

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-06T09:22:23+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: fct_generales.py
# @Last modified by:   gysco
# @Last modified time: 2017-04-07T08:27:20+02:00


def somme_tableau(a):
    """
    Fonction de calcul de la somme des termes d'un vecteur.

    @param a: tableau de reels
    """
    _ret = 0
    for i in range(0, len(a)):
        _ret += a[i]
    return _ret


def trier_collection(aCollection, itri, isens):
    """
    Trie la collection data_co suivant une de ses variables donnee.

    # WARNING: A OPTIMISER DE FACON PYTHONESQUE

    @param aCollection: collection a trier
    @param itri: numero de l'item sur lequel on effectue le tri
    @param isens: sens du tri (0=decroissant,1=croissant)
    """
    tmp_col = list()
    j = 1
    while aCollection.Count > 0:
        if (itri <= 3):
            mini = 'z'
            maxi = 'A'
            """ne pas modifier sinon ça ne marche plus"""
        else:
            mini = 10**300
            maxi = -10**300
        for i in range(0, len(aCollection)):
            if (itri == 1):
                tmp = aCollection[i].espece
            elif (itri == 2):
                tmp = aCollection[i].taxo
            elif (itri == 3):
                tmp = aCollection[i].test
            elif (itri == 4):
                tmp = aCollection[i].data
            elif (itri == 5):
                tmp = aCollection[i].num
            elif (itri == 6):
                tmp = aCollection[i].pond
            elif (itri == 7):
                tmp = aCollection[i].pcum
            elif (itri == 8):
                tmp = aCollection[i].std
            elif (itri == 9):
                tmp = aCollection[i].act
            elif (itri == 10):
                tmp = aCollection[i].pcum_a
            if (isens == 1):
                if (tmp <= mini):
                    mini = tmp
                    num = i
            else:
                if (tmp >= maxi):
                    maxi = tmp
                    num = i
        tmp_col.append(aCollection[num])
        del aCollection[num]
        j += 1
    for i in range(0, len(tmp_col)):
        aCollection.append(tmp_col[i])
    tmp_col = None


def ecrire_titre(titre, nom_feuille, lig, col, nbcol):
    """
    Ecrit le titre d'un tableau.

    @param titre: titre du tableau
    @param nom_feuille: nom de la feuille de calcul
    @param lig: numero de la ligne ou ecrire le titre du tableau
    @param col: numero de la colonne ou ecrire le titre du tableau
    @param nbcol: nombre de colonnes du tableau (pour centrer le titre
                  sur toutes les colonnes)
    """
    Worksheets[nom_feuille].Cells[lig, col] = titre


def maximum(a, b):
    """Renvoie le maximum de 2 valeurs."""
    return(a if a > b else b)


def minimum(a, b):
    """Renvoie le minimum de 2 valeurs."""
    return(a if a < b else b)


def minimum_tab(a):
    """Renvoie la valeur minimum d'un tableau de reels."""
    _ret = a[0]
    for i in range(1, len(a)):
        if (a[i] < _ret):
            _ret = a[i]
    return(_ret)


def maximum_tab(a):
    """Renvoie la valeur maximum d'un tableau de reels."""
    _ret = a[0]
    for i in range(1, len(a)):
        if (a[i] > _ret):
            _ret = a[i]
    return(_ret)


def minimum_tab_dif0(a):
    """Renvoie la valeur minimum <> 0 d'un tableau de reels."""
    _ret = None
    """Recherche d'une valeur non-nulle dans le tableau"""
    for i in range(0, len(a)):
        if a[i] != 0:
            break
    _ret = a[i]
    for i in range(0, len(a)):
        if (a[i] < _ret and a(i) != 0):
            _ret = a(i)
    return _ret


def calcul_lig_graph(lig_deb, lig_p, lig_qbe, lig_qbi, lig_qbs):
    """
    Calcul les indices de lignes pour les graphes dans nom_feuille_res.

    Compte tenu de la disposition du tableau de resultats HC

    @param lig_deb: ligne de debut de l'affichage du tableau de
                    resultats HC dans nom_feuille_res,
                    il s'agit d'une ligne de titre
    @param lig_p: ligne des probabilites cumulees
    @param lig_qbe: ligne des HC best-estimates
    @param lig_qbi: ligne de la borne inferieure de l'intervalle de
                    confiance de la HC
    @param lig_qbs: ligne de la borne superieure de l'intervalle de
                    confiance de la HC

    Attention ceci depend des choix d'affichage dans calculer_res
    """
    lig_p = lig_deb + 1
    lig_qbe = lig_deb + 2
    lig_qbi = lig_deb + 5
    lig_qbs = lig_deb + 6
