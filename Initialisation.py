"""
Initialisation de l'IHM et differentes variables.

Problablement inutilisable au niveau de l'IHM.
"""

# !/usr/bin/env python

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-10T09:13:55+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: Initialisation.py
# @Last modified by:   gysco
# @Last modified time: 2017-04-10T15:39:25+02:00

import math

import Collection


def init_apropos():
    """Definit le texte de la boite A propos. Specifique SSWD."""
    global frm_apropos
    frm_apropos.text = 'SSWD'
    frm_apropos.Lbl_description.text = 'Species Sensitivity Weighted ' + \
        'Distribution (SSWD) Software\nenables to estimate Hazardous' + \
        ' Concentration (HC) with confidence limits by bootstrap'
    frm_apropos.Lbl_version.text = 'Version : 1.0'
    frm_apropos.Lbl_date.text = 'December 2003'
    frm_apropos.Lbl_dev = 'Developed by Electricite de France (EDF)\n' + \
        'With Institut de l\'Environnement Industriel et des Risques ' + \
        '(INERIS)\nMethodology and design: C.Duboudin - EDF\n' + \
        'Code development: R.Aletti - Simulog and C.Duboudin\n' + \
        'Contacts: Ph.Ciffroy - EDF (philippe.ciffroy@edf.fr)\n' + \
        '          H.Magaud - INERIS (helene.magaud@ineris.fr)'
    frm_apropos.Cmd_Ok.text = 'Ok'


def init_ihm():
    """Initialisation pour les fonctions ihm boite de dialogue."""
    """
    Definition des options possibles du parametre espece ;
    correspond à isp=1,2,3
    """
    global frm_sswd
    sp_opt = list()
    sp_opt.append('weighted')
    sp_opt.append('unweighted')
    sp_opt.append('mean')
    """Initialisation des titres de la boite de dialogue"""
    frm_sswd.text = 'SSWD'
    frm_sswd.cadre_donnees.text = 'Data'
    frm_sswd.Lbl_espece = 'Species or genus information'
    frm_sswd.Lbl_taxo = 'Taxonomic or trophic information'
    frm_sswd.Lbl_concentration = 'Concentration values'
    frm_sswd.Cadre_Option_espece.text = 'Weighting procedures'
    frm_sswd.Lbl_pond.text = 'Species or genus options'
    frm_sswd.Cbx_esp.append(sp_opt[0])
    frm_sswd.Cbx_esp.append(sp_opt[1])
    frm_sswd.Cbx_esp.append(sp_opt[2])
    frm_sswd.Lbl_Pcat = 'Taxonomic weight'
    frm_sswd.Opt_Pcat_nul.text = 'No weight'
    frm_sswd.Opt_Pcat_valeur.text = 'Enter weight values'
    frm_sswd.cadre_option_stat.text = 'Statistical options'
    frm_sswd.Lbl_loi.text = 'Distribution'
    frm_sswd.Chk_emp.text = 'Log-Empirical'
    frm_sswd.Chk_normal.text = 'Log-Normal'
    frm_sswd.Chk_triangle.text = 'Log-Triangular'
    frm_sswd.Opt_ajust_q.text = 'Quant. fitting'
    frm_sswd.Opt_ajust_p.text = 'Prob. fitting'
    frm_sswd.Lbl_B.text = 'Number of bootstrap samples'
    frm_sswd.Chk_nbvar.text = 'Optimized bootstrap samples size'
    frm_sswd.Lbl_a.text = 'Hazen parameter a'
    frm_sswd.Chk_sauvegarde.text = 'Conserve the intermediate worksheets of' +\
        ' calculation'
    frm_sswd.Cmd_Ok.text = 'Ok'
    frm_sswd.Cmd_Annuler.text = 'Cancel'
    # Help string
    frm_sswd.cadre_donnees.ControlTipText = 'The data must be in columns ' +\
        'with headings; a minimum of three columns is needed'
    frm_sswd.Lbl_espece.ControlTipText = 'Select the range or column ' +\
        '(heading included) containing the name of the tested species ' +\
        '(or genus)'
    frm_sswd.Lbl_taxo.ControlTipText = 'Select the range or column ' +\
        '(heading included) containing the taxonomic groups or trophic levels'
    frm_sswd.Lbl_concentration.ControlTipText = 'Select the range or column' +\
        ' (heading included) containing the ecotoxicological test\'s results'
    frm_sswd.Lbl_pond.ControlTipText = 'Three options are proposed to ' +\
        'account for redundant data for each species or genus'
    frm_sswd.Lbl_Pcat.ControlTipText = 'Two approaches are proposed ' +\
        'regarding proportions of data of each taxonomic group or ' +\
        'trophic level'
    frm_sswd.Opt_Pcat_nul.ControlTipText = 'If you select -No Weight-,' +\
        ' the default weights will be the observed proportions of data in' +\
        ' each taxonomic group or trophic level'
    frm_sswd.Opt_Pcat_valeur.ControlTipText = 'Weights to be allocated to ' +\
        'each taxonomic group or trophic level'
    frm_sswd.Lbl_loi.ControlTipText = 'Select the distribution to be ' +\
        'applied to the weighted data'
    frm_sswd.Opt_ajust_q.ControlTipText = 'Estimation of the min, max and' +\
        ' mode parameters by fitting the quantiles'
    frm_sswd.Opt_ajust_p.ControlTipText = 'Estimation of the min, max and' +\
        ' mode parameters by fitting the cumulative probabilities'
    frm_sswd.Lbl_B.ControlTipText = 'The log-triangular distribution' +\
        ' computational time costs is much greater than the others;' +\
        ' to test first with few bootstrap runs'
    frm_sswd.Chk_nbvar.ControlTipText = 'By default bootstrap draws number' +\
        ' is the number of used data'
    frm_sswd.Lbl_a.ControlTipText = 'The Hazen parameter have an effect on' +\
        ' the log-empirical and the log-triangular distribution'
    """Options par defaut"""
    frm_sswd.Txt_B.text = 1000
    frm_sswd.Txt_a.text = 0.5
    frm_sswd.Cbx_esp.text = frm_sswd.Cbx_esp.List(0, 0)
    frm_sswd.Opt_Pcat_nul.Value = True
    frm_sswd.Chk_emp.Value = True
    frm_sswd.Chk_normal.Value = True
    frm_sswd.Chk_triangle.Value = False
    frm_sswd.Chk_sauvegarde.Value = False
    frm_sswd.Lbl_liste_taxo.visible = False
    frm_sswd.txt_pcat.visible = False
    if frm_sswd.Chk_triangle.Value is True:
        frm_sswd.Opt_ajust_q.Enabled = True
        frm_sswd.Opt_ajust_p.Enabled = True
    else:
        frm_sswd.Opt_ajust_q.Enabled = False
        frm_sswd.Opt_ajust_p.Enabled = False


def init_collection(nom_feuille, l_espece, l_taxo, l_data, c_espece, c_taxo,
                    c_data, data_co):
    """
    Charge en memoire les donnees selectionnees par l'utilisateur.

    Creation de la collection data_co
    """
    global Worksheets
    ligne = Collection()
    i = 0
    while Worksheets[nom_feuille].Cells[l_espece + i, c_espece]:
        ligne = Collection()
        ligne.espece = Worksheets[nom_feuille].Cells[l_espece + i, c_espece]
        ligne.taxo = Worksheets[nom_feuille].Cells[l_taxo + i, c_taxo]
        ligne.test = 'C'
        ligne.data = math.log(
            Worksheets[nom_feuille].Cells[l_data + i, c_data]) / math.log(10)
        ligne.pond = 1
        ligne.num = 1
        ligne.pcum = 1
        data_co.append(ligne)
        ligne = None
        i += 1


def initialise(pourcent, pcent, nom_feuille_pond, nom_feuille_stat,
               nom_feuille_res, nom_feuille_qemp, nom_feuille_qnorm,
               nom_feuille_sort, nom_feuille_Ftriang, nom_feuille_qtriang,
               titre_graf, titre_res, a, ind_hc, titre_data, titre_axe):
    """
    Initialisation de certains parametres.

    valeurs par defaut modifiables par l'utilisateur averti
    """
    """Nom des feuilles intermediaires et resultat final"""
    nom_feuille_pond = 'weight_result'
    nom_feuille_stat = 'draw_result'
    nom_feuille_res = 'SSWD_result'
    nom_feuille_qemp = 'qemp_result'
    nom_feuille_qnorm = 'qnorm_result'
    nom_feuille_sort = 'draw_sort'
    nom_feuille_Ftriang = 'ftriang_result'
    nom_feuille_qtriang = 'qtriang_result'
    """Pourcentage x des quantiles HCx% calcules à chaque run du bootstrap"""
    pourcent = [
        0.025, 0.05, 0.1, 0.15, 0.2, 0.3, 0.5, 0.7, 0.8, 0.85, 0.9, 0.95, 0.975
    ]
    ind_hc = 2
    """
    Pourcentages definissant les intervalles de confiance à 90 et 95%
    des HCx%
    """
    pcent = [0.025, 0.05, 0.95, 0.975]
    """Titres des graphiques et des tableaux de resultats HC et data"""
    titre_graf = list()
    titre_graf.append('SSWD - Log Empirical')
    titre_graf.append('SSWD - Log Normal')
    titre_graf.append('SSWD - Log Triangular')
    titre_axe = list()
    titre_axe.append('Concentration')
    titre_axe.append('Cumulative weighted probability')
    titre_res = list()
    titre_res.append('Distribution: Weighted Empirical + ' +
                     'Confidence limits by weighted bootstrap')
    titre_res.append(
        'Distribution: Weighted Normal + ' +
        'Confidence limits by weighted bootstrap and Bias Correction')
    titre_res.append(
        'Distribution: Weighted Triangular + ' +
        'Confidence limits by weighted bootstrap and Bias Correction')
    titre_data = list()
    titre_data.append('Used data sorted out by taxonomy group')
    titre_data.append('Used data sorted out by increasing concentrations')
    """Parametre de Hazen pour calcul probabilites empiriques"""
    # a = 0.5
