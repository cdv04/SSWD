# coding=utf-8
"""Ponderation."""

# !/usr/bin/env python

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-05T10:08:45+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: weighting.py
# @Last modified by:   gysco
# @Last modified time: 2017-05-19T11:22:51+02:00

from common import sort_collection


def calcul_ponderation(data_co, pcat, isp, a, nb_taxo):
    """
    Calcul des ponderations.

    Calcul des ponderations associee a chaque resultat de test ecotox
    (concentration value) compte tenue des options retenues par
    l'utilisateur concernant les especes et les categories
    taxonomiques; Calcul des probabilites ponderees cumulees
    empiriques a partir des ponderations et par la methode de Hazen
    On travaille donc ensuite avec des points ponderees
    Cette ponderation correspond egalement a la probabilite d'occurence
    que chaque donnee doit avoir lors des tirages bootstraps

    @param data_co: tableau des donnees generees a partir des donnees
                    presentes dans une feuille de calcul
    @param pcat: vecteur de poids associes a chaque categorie
                 taxonomique
    @param isp: option concernant l'information espece ;
                prend les valeurs :
                    1 pour wted,
                    2 pour unwted,
                    3 pour mean
    @parma nom_feuille: nom de la feuille contenant les donnees
                        initiales
    @param a: parametre de Hazen pour le calcul de probabilite cumulee
              empirique
    @param nb_taxo: nombre de categories taxonomiques
    """
    if isp == 0:
        """
        Si isp=0 (mean), eliminimumation des doublons au niveau espece et
        remplacement par la moyenne des donnees
        """
        data_co = moyenne(data_co)
    p = list()
    nb_espece = compte_espece(data_co, p)
    somme_pcat = sum(pcat)
    if somme_pcat == 0:
        if isp != 1:
            for i in range(0, len(data_co)):
                p[i] = (1 / len(data_co))
        else:
            for i in range(0, len(data_co)):
                p[i] = (1 / (p[i] * nb_espece))
    else:
        if isp != 1:
            ind_debut = 0
            ind_taxo = 0
            nb = 0
            j = 0
            """
            Pour chaque categorie taxo, calcul du nombre d'espece presente puis
            des ponderations
            """
            for i in range(0, nb_taxo):
                while data_co[j].taxo == data_co[j + 1].taxo:
                    nb += 1
                    j += 1
                    if j == len(data_co):
                        break
                ind_fin = j
                for k in range(ind_debut, ind_fin):
                    p[k] = (1 / nb) * pcat[ind_taxo]
                ind_debut = ind_fin + 1
                ind_taxo += 1
                j += 1
                nb = 1
        else:
            j = 0
            nb = 0
            nespece = 1
            """
            Pour chaque categorie taxo, calcul du nombre d'espece differente
            puis des ponderations
            """
            for i in range(0, nb_taxo):
                ind_deb = j
                k = 0
                while data_co[j].taxo == data_co[j + 1].taxo:
                    tmp = data_co[j].espece
                    prem = True
                    k = j + 1
                    while k <= len(data_co):
                        if tmp == data_co[k].espece and prem:
                            nb += 1
                            prem = False
                        if data_co[k - 1].taxo != data_co[k].taxo:
                            break
                        k += 1
                    j += 1
                    nespece += 1
                    if j == len(data_co):
                        break
                ind_fin = k - 1
                nb = nespece - nb
                for l in range(ind_deb, ind_fin):
                    p[l] = (1 / (p[l] * nb)) * pcat[i]
                nespece = 1
                nb = 0
                j += 1
        """Normalisation"""
        somme_p = sum(p)
        for i in range(0, len(p)):
            p[i] /= somme_p
    for i in range(0, len(data_co)):
        data_co[i].pond = p[i]
    """Calcul des ponderations cumulees"""
    sort_collection(data_co, 4, 1)
    calcul_prob_cumul(data_co, a)


def moyenne(aCollection):
    """
    Recherche et suppression des doublons especes dans une collection.

    Calcule alors la moyenne sur les donnees en doublon et enregistre
    le nombre de doublon trouve (membre num)

    @param aCollection: la collection a traiter
    """
    i = 0
    while i in range(0, len(aCollection)):
        tmp = aCollection[i].espece
        for j in range(i + 1, len(aCollection)):
            if tmp == aCollection[j].espece:
                aCollection[i].data += aCollection[j].data
                aCollection[i].num += 1
                del (aCollection[j])
                i -= 1
                break
        i += 1
    for i in range(0, len(aCollection)):
        aCollection[i].data /= aCollection[i].num
    return aCollection


def compte_espece(aCollection, p):
    """Compte le nombre d'especes differentes presentes dans les donnees."""
    if len(aCollection) == 0:
        return
    compt = list()
    """Nombre d'espece differentes par categorie taxonomique"""
    for i in range(0, len(aCollection)):
        tmp = aCollection[i].espece
        # prem = True
        compt.append(aCollection[i].espece)
        for j in range(i + 1, len(aCollection)):
            if aCollection[j - 1].taxo != aCollection[j].taxo:
                break
            if tmp == aCollection[j].espece:
                aCollection[i].num += 1
                aCollection[j].num += 1
                """compte le nombre de doublons"""
                # if (prem):
                #     compt = compt + 1
                #     prem = False
    """Compte le nombre d'espece differentes dans la collection"""
    nb_espece = len(set(compt))  # len(aCollection) - compt
    for i in range(0, len(aCollection)):
        p.append(aCollection[i].num)
    return nb_espece


def calcul_prob_cumul(aCollection, a):
    """
    Proba ponderees cumulees empiriques de chaque donnee de toxicite.

    attention, la collection doit etre triee dans le sens croissant des data
    """
    p1 = list()
    p1.append(aCollection[0].pond * len(aCollection))
    for i in range(1, len(aCollection)):
        p1.append(p1[i - 1] + aCollection[i].pond * len(aCollection))
    """
    si les ponderations sont telles que les premiers points ont des poids
    inferieurs au paramètre de Hazen a, alors le programme prend
    automatiquement la valeur a=0
    """
    if p1[0] <= a:
        a = 0
    for i in range(0, len(aCollection)):
        """
        Anciennen approche pour prendre en compte les poids faibles
        methode abandonnee, car pas satisfaisante d'un point de vue theorique
        """
        # tmp_pond = (p1[i] - a) / (len(aCollection) + 1 - 2 * a)
        # If (tmp_pond <= 0) Then
        #   aCollection[i].pcum = 0.0001
        # Else
        aCollection[i].pcum = (p1[i] - a) / (len(aCollection) + 1 - 2 * a)


def calcul_nbvar(n_optim, data_co, pcat):
    """
    Calcul le nombre de donnees optimal a tirer lors du bootstrap.

    Compte tenu du nombre de donnees presentes dans chaque categorie
    taxonomique et des ponderations voulues ou bien renvoie le nombre
    de donnees si pas d'optimisation demandee
    """
    somme_pcat = sum(pcat)
    """Optimisation demandee et possible"""
    if n_optim is True and somme_pcat != 0:
        """Calcul du nombre de donnees par categorie taxonomique"""
        sort_collection(data_co, 2, 1)
        nb_data = list()
        nb_data.append(0)
        gr = list()
        j = 0
        for i in range(0, len(data_co)):
            nb_data[j] += 1
            if data_co[i].taxo != data_co[i + 1].taxo:
                j += 1
        if data_co[len(data_co)].taxo == data_co[len(data_co) - 1].taxo:
            nb_data[len(nb_data)] += 1
        else:
            nb_data[len(nb_data)] = 1
        """Normalisation de Pcat si pas fait"""
        for i in range(0, len(pcat)):
            pcat[i] = pcat[i] / somme_pcat
        for i in range(0, len(nb_data)):
            gr[i] = nb_data[i] / pcat[i]
        """Calcul de nbvar"""
        nbvar = min(gr)
        # x = min(gr)
        # for i in range(0, len(nb_data)):
        #   gr[i] = x * pcat[i]
        # Calcul de nbvar
        # For i in range(0, len(nb_data))
        #  nbvar += int(gr[i] / min(minimum_tab_dif0(gr), 1))
        """Pas d'optimisation : nbvar=nbdata"""
    else:
        nbvar = len(data_co)
    x = nbvar
    nbvar = min(x, 250)
    """
    250 est la limite du nombre de donnees que l'on peut tirer compte tenu de
    la limite du nombre de colonnes de excel (version 97) qui est 256
    """
    return nbvar


def calcul_nb_taxo(data_co):
    """Calcul du nombre de categories taxonomiques differentes."""
    nb_taxo = 0
    for i in range(0, len(data_co) - 1):
        if data_co[i].taxo != data_co[i + 1].taxo:
            nb_taxo += 1
    return nb_taxo
