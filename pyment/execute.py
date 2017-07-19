# coding=utf-8
"""
Program start here.

To python soon.
"""

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-10T15:43:09+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: execute.py
# @Last modified by:   gysco
# @Last modified time: 2017-06-19T16:15:30+02:00

from statistics import (calcul_R2, calcul_ic_empirique, calcul_ic_normal,
                        calcul_ic_triang_p, calcul_ic_triang_q, calcul_res,
                        tirage)

from charts import draw_chart
from common import (
    affichage_options, calcul_col_res, calcul_lig_graph, calcul_ref_pond,
    ecrire_data_co, ecrire_titre, efface_feuil_inter, verif, write_feuil_inter
)
from initialisation import initialise
from pandas import ExcelWriter
from weighting import calcul_nbvar, calcul_ponderation, sort_collection


def lance_ihm():
    """Fait apparaitre la boite de dialogue SSWD."""
    global frm_sswd
    frm_sswd.Show()


def lance_apropos():
    """Fait apparaitre la boite A propos."""
    global frm_apropos
    frm_apropos.Show()


def lance(output,
          data_co,
          nom_colonne,
          isp,
          pcat,
          dist,
          B,
          a,
          n_optim,
          conserv_inter,
          nb_taxo,
          val_pcat,
          liste_taxo,
          triang_ajust,
          seed):
    """
    Module de lancement de la procedure SSWD.

    Remarque : HC=Hazardous Concentration;
               SSWD=Species Sensitivity Weighted Distribution;
               WECP=Weighted Empirical Cumulative Probability
    Principales etapes algorithmiques :
    1. Calcul des ponderations associees a chaque resultat de test
       ecotox (concentration), compte tenu des poids et des options
       choisis par l'utilisateur et des proportions de donnees
       existantes; calcul des probabilites empiriques cumulees
       ponderees
    2. Calcul des parametres mu, sig, _min, _max, mode, suivant les cas,
       qui permettent l'estimationdes valeurs de best-estimates des
       HCx% a partir des donnees ponderees
    3. Tirages aleatoires (procedure de bootstrap) pour estimation de
       l'intervalle de confiance associee a chaque HCx%
    4. Affichage des resultats et representation graphique
    ___________________________________________________________________
    Parametres principaux
    @param data_co: tableaux des donnees exploites pour le calcul des
                    HC et genere par la procedure
                    attention : ce tableau contient des colonnes qui
                                ne sont pas affichees dans les feuilles
                                de calcul
                    ce tableau est affiche dans nom_feuille_pond pour
                    les calculs intermediaires et deux fois dans
                    nom_feuille_res pour l'affichage des graphiques
                    SSWD.
                    une fois triee en fonction des categories
                    taxonomiques et une fois dans l'ordre croissant des
                    concentrations nous appellerons data_co_feuil la
                    data_co telle qu'elle est affichees dans ces
                    feuilles
    @param nom_colonne: nom des colonnes de data_co_feuil
    @param isp: indice correspondant a la methode de traitement du
                parametre espece 1=wted, 2=unwted,3=mean
    @param pcat: poids accordes a chaque categorie taxonomique
    @param nb_taxo: nombre de categories taxonomiques ou niveaux
                    trophiques
    @param triang_ajust: option d'ajustement pour la loi triangulaire
                         si True ajustement sur les quantiles, sinon
                         sur les probabilites cumulees
    ___________________________________________________________________
    // TODO: Ajouter le reste de la docummentation a la main.
    """
    """Debut de la procedure"""
    # Application.ScreenUpdating = False
    """Valeurs specifique a la procedure SSWD"""
    iproc = 1
    ind_tax = 2
    ind_data = 3
    ind_pond = 4
    ind_pcum = 5
    tmp = 0
    """
    Initialisation : definition des valeurs par defaut pour certains
                     parametres
    modifiables par l'utilisateur averti
    """
    (nom_feuille_pond, nom_feuille_stat, nom_feuille_res, nom_feuille_qemp,
     nom_feuille_qnorm, nom_feuille_sort, nom_feuille_Ftriang,
     nom_feuille_qtriang, pourcent, ind_hc, pcent, titre_graf, titre_axe,
     titre_res, titre_data) = initialise()
    """
    Test sur l'existence de feuilles de resultats et creation des feuilles
    necessaires
    """
    verif(nom_feuille_pond, nom_feuille_stat, nom_feuille_res,
          nom_feuille_qemp, nom_feuille_qnorm, nom_feuille_sort,
          nom_feuille_Ftriang, nom_feuille_qtriang, '', '', '')
    """
    1. Calcul des ponderations et affichage resultats
    dans nom_feuille_pond
    """
    pond_lig = 0
    pond_col = 0
    calcul_ponderation(data_co, pcat, isp, a, nb_taxo)
    ecrire_data_co(data_co, nom_colonne, pond_lig, pond_col, nom_feuille_pond,
                   False, iproc)
    nbdata = len(data_co)
    """2. Calcul nbvar et Tirages aleatoires"""
    nbvar = calcul_nbvar(n_optim, data_co, pcat)
    (pond_lig_deb, pond_lig_fin, pond_col_data, pond_col_pond,
     pond_col_pcum, pond_col_data_act) = calcul_ref_pond(
         pond_col, pond_lig, ind_data, ind_pond, ind_pcum, nbdata, tmp)
    tirage(nom_feuille_stat, nbvar, B, nom_feuille_pond, pond_col_data,
           pond_col_pond, seed)
    """
    Remarque : le resultat des tirages est affiche dan@s nom_feuille_stat
    L'affichage commence a la premiere ligne et a la premiere colonne;
    la premiere ligne est une ligne de titre;
    ceci n'est pour l'instant pas parametrable
    """
    """3. Calculs valeurs best-estimates et statistiques apres tirages"""
    """Definition indice lignes et colonnes"""
    l1 = 1
    """
    l1 ne peut être modifiee : c'est en fait une constante definie par la
    procedure tirage
    """
    l2 = B + l1 - 1
    c1 = 0
    """c'est une constante definie par la sub tirage"""
    c2 = c1 + nbvar - 1
    lig_hc = 26
    """attention : il faut tenir compte de l'affichage des options"""
    col_hc = 0
    nbcol_vide = 1
    lig_data = 1
    writer = ExcelWriter(output)
    """
    Calcul des indices des colonnes d'affichage des resultats dans
    nom_feuille_res
    """
    (col_deb, col_fin, col_data1, col_data2, col_tax, col_data, col_pcum,
     col_data_le, col_pcum_le, col_data_act, col_data_act_le,
     col_pcum_a) = calcul_col_res(col_hc, nbcol_vide, pourcent, dist, ind_tax,
                                  ind_data, ind_pcum, nom_colonne, tmp, tmp)
    """Calcul des indices des lignes pour les graphes de nom_feuille_res"""
    lig_p, lig_qbe, lig_qbi, lig_qbs = calcul_lig_graph(lig_hc)
    """initialisation de ligne_tot"""
    i = 0
    feuilles_res = ['_emp', '_norm', '_triang']
    for x in dist:
        if x is True:
            """
            Ecriture de data_co_feuil triees par rapport aux categories
            taxonomiques dans nom_feuille_res
            """
            sort_collection(data_co, 2, 2)
            ecrire_titre(titre_data[0], nom_feuille_res + feuilles_res[i],
                         lig_data - 1, col_data1)
            ecrire_data_co(data_co, nom_colonne, lig_data, col_data1,
                           nom_feuille_res + feuilles_res[i], True, iproc)
            """
            Ecriture de data_co_feuil triees par ordre croissant des
            concentrations dans nom_feuille_res
            """
            sort_collection(data_co, 7, 1)
            ecrire_titre(titre_data[1], nom_feuille_res + feuilles_res[i],
                         lig_data - 1, col_data2)
            ecrire_data_co(data_co, nom_colonne, lig_data, col_data2,
                           nom_feuille_res + feuilles_res[i], True, iproc)
        i += 1
    """loi empirique"""
    if dist[0] is True:
        loi = 1
        """Calcul les valeurs correspondant a chaque tirage"""
        calcul_ic_empirique(l1, c1, l2, c2, c1, pourcent, nom_feuille_stat,
                            nom_feuille_qemp, nom_feuille_sort, nbvar, a)
        """Calcul des valeurs best-estimates et affichage des resultats"""
        mup, sigmap, _min, _max, mode, data_c = calcul_res(
            l1, l2, ind_hc, pond_lig_deb, pond_col, pond_col_data,
            pond_col_pcum, lig_hc, col_hc, nbvar, loi, titre_res, pcent,
            pourcent, data_co, nom_colonne, nom_feuille_res + "_emp",
            nom_feuille_qemp, nom_feuille_pond, '', 0, triang_ajust, iproc,
            nbdata)
        """Graphes de SSWD"""
        draw_chart(writer, nom_feuille_res + "_emp", lig_p, lig_qbe, lig_qbi,
                   lig_qbs, col_deb, col_fin, lig_data, col_tax, col_data,
                   col_pcum, col_data_le, col_pcum_le, loi, titre_graf, 0, 0,
                   nbdata, mup, sigmap, _min, _max, mode, titre_axe, val_pcat,
                   liste_taxo, isp, tmp, tmp, iproc)
        lig_p, lig_qbe, lig_qbi, lig_qbs = calcul_lig_graph(lig_hc)
    """loi normale"""
    if dist[1] is True:
        loi = 2
        calcul_ic_normal(l1, c1, l2, c2, c1, pourcent, nom_feuille_stat,
                         nom_feuille_qnorm)
        mup, sigmap, _min, _max, mode, data_c = calcul_res(
            l1, l2, ind_hc, pond_lig_deb, pond_col, pond_col_data,
            pond_col_pcum, lig_hc, col_hc, nbvar, loi, titre_res, pcent,
            pourcent, data_co, nom_colonne, nom_feuille_res + "_norm",
            nom_feuille_qnorm, nom_feuille_pond, nom_feuille_stat, 0,
            triang_ajust, iproc, nbdata)
        R2_norm, Pvalue_norm = calcul_R2(data_co, loi, mup, sigmap, _min, _max,
                                         mode, nbdata, data_c)
        draw_chart(writer, nom_feuille_res + "_norm", lig_p, lig_qbe, lig_qbi,
                   lig_qbs, col_deb, col_fin, lig_data, col_tax, col_data,
                   col_pcum, col_data_le, col_pcum_le, loi, titre_graf,
                   R2_norm, Pvalue_norm, nbdata, mup, sigmap, _min, _max, mode,
                   titre_axe, val_pcat, liste_taxo, isp, tmp, tmp, iproc)
        lig_p, lig_qbe, lig_qbi, lig_qbs = calcul_lig_graph(lig_hc)
    """loi triangulaire"""
    if dist[2] is True:
        loi = 3
        if triang_ajust is True:
            c_min = calcul_ic_triang_q(
                l1, c1, l2, c2, c1, nbvar, a, pourcent, nom_feuille_stat,
                nom_feuille_sort, nom_feuille_Ftriang, nom_feuille_qtriang)
        else:
            c_min = calcul_ic_triang_p(
                l1, c1, l2, c2, c1, nbvar, a, pourcent, nom_feuille_stat,
                nom_feuille_sort, nom_feuille_Ftriang, nom_feuille_qtriang)
        mup, sigmap, _min, _max, mode, data_c = calcul_res(
            l1, l2, ind_hc, pond_lig_deb, pond_col, pond_col_data,
            pond_col_pcum, lig_hc, col_hc, nbvar, loi, titre_res, pcent,
            pourcent, data_co, nom_colonne, nom_feuille_res + "_triang",
            nom_feuille_qtriang, nom_feuille_pond, nom_feuille_Ftriang, c_min,
            triang_ajust, iproc, nbdata)
        R2_triang, Pvalue_triang = calcul_R2(data_co, loi, mup, sigmap, _min,
                                             _max, mode, nbdata, data_c)
        draw_chart(writer, nom_feuille_res + "_triang", lig_p, lig_qbe,
                   lig_qbi, lig_qbs, col_deb, col_fin, lig_data, col_tax,
                   col_data, col_pcum, col_data_le, col_pcum_le, loi,
                   titre_graf, R2_triang, Pvalue_triang, nbdata, mup, sigmap,
                   _min, _max, mode, titre_axe, val_pcat, liste_taxo, isp, tmp,
                   tmp, iproc)
    # decaler_graph(nom_feuille_res)
    affichage_options("details", isp, val_pcat, liste_taxo, B, 0, 0, 6, 0,
                      dist, nbvar, iproc, a)
    # cellule_gras()
    if conserv_inter is False:
        efface_feuil_inter(nom_feuille_pond, nom_feuille_stat,
                           nom_feuille_qemp, nom_feuille_qnorm,
                           nom_feuille_qtriang, nom_feuille_sort,
                           nom_feuille_Ftriang, '', '', '')
    write_feuil_inter(writer)
