# coding=utf-8
"""Charts drawing module."""

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-12T09:13:56+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: charts.py
# @Last modified by:   gysco
# @Last modified time: 2017-06-19T14:17:33+02:00

import initialisation
from common import sp_opt


def draw_chart(writer, nom_feuille, lig_p, lig_qbe, lig_qbi, lig_qbs, col_deb,
               col_fin, ligne_data, col_tax, col_data, col_pcum, col_data_le,
               col_pcum_le, loi, titre_graf, r2, weight_value, nb_ligne_data,
               mup, sigmap, _min, _max, mode, titre_axe, val_pcat, liste_taxo,
               isp, col_data_act, col_data_act_le, iproc):
    """
    Draw charts (median, percentile 5%, percentile 95% & data).

    :param writer: output file opened
    :param nom_feuille: name of the worksheet containing the data
    :param lig_p: line containing values of cumulative probabilities
    :param lig_qbe: best-estimate row
    :param lig_qbi: percentile 5% row
    :param lig_qbs: percentile 95% row
    :param col_deb: start column
    :param col_fin: end column
    :param ligne_data: data line
    :param col_tax:
    :param col_data:
    :param col_pcum:
    :param col_data_le:
    :param col_pcum_le:
    :param loi: type of law
    :param titre_graf: chart title
    :param r2: parameter r2 of curve
    :param weight_value:
    :param nb_ligne_data:
    :param mup:
    :param sigmap:
    :param _min: triangular law parameter
    :param _max: triangular law parameter
    :param mode: triangular law parameter
    :param titre_axe: axis titles
    :param val_pcat:
    :param liste_taxo:
    :param isp:
    :param col_data_act:
    :param col_data_act_le:
    :param iproc:
    """
    workbook = writer.book
    initialisation.Worksheets[nom_feuille].Cells.sort_index(
        axis=1).reindex_axis(
            range(0,
                  initialisation.Worksheets[nom_feuille].Cells.columns.max() +
                  1),
            axis=1).sort_index(axis=0).reindex_axis(
                range(
                    0,
                    (initialisation.Worksheets[nom_feuille].Cells.index.max() +
                     1)),
                axis=0).to_excel(
                    writer, sheet_name=nom_feuille, index=False, header=False)
    worksheet = writer.sheets[nom_feuille]
    """Ajout de la serie Mediane"""
    if loi != 1:
        chart = workbook.add_chart({
            'type': 'scatter',
            'subtype': 'smooth_with_markers'
        })
        chart.add_series({
            'values': [nom_feuille, lig_p, col_deb, lig_p, col_fin],
            'categories': [nom_feuille, lig_qbe, col_deb, lig_qbe, col_fin],
            'name': [nom_feuille, lig_qbe, col_deb - 1],
            'line': {
                'color': 'red'
            },
            'marker': {
                'type': 'none'
            }
        })

    # TODO modifer en loi log avec range 0, 100
    else:
        chart = tracer_courbe_empirique(
            workbook, nom_feuille, ligne_data + 1, nb_ligne_data, col_data_le,
            col_pcum_le) if iproc == 1 else tracer_courbe_empirique(
                workbook, nom_feuille, ligne_data + 1, nb_ligne_data,
                col_data_act_le, col_pcum_le)
    """Ajout de la serie quantile borne inf"""
    chart.add_series({
        'values': [nom_feuille, lig_p, col_deb, lig_p, col_fin],
        'categories': [nom_feuille, lig_qbi, col_deb, lig_qbi, col_fin],
        'name': [nom_feuille, lig_qbi, col_deb - 1],
        'line': {
            'color': 'black',
            'dash_type': 'dash',
            'width': 1.25
        },
        'marker': {
            'type': 'none'
        }
    })
    """Ajout de la serie quantile borne superieure"""
    chart.add_series({
        'values': [nom_feuille, lig_p, col_deb, lig_p, col_fin],
        'categories': [nom_feuille, lig_qbs, col_deb, lig_qbs, col_fin],
        'name': [nom_feuille, lig_qbs, col_deb - 1],
        'line': {
            'color': 'black',
            'dash_type': 'dash',
            'width': 1.25
        },
        'marker': {
            'type': 'none'
        }
    })
    """
    Ajoute les donnees data
    """
    if iproc == 1:
        chart = add_species_series(chart, nom_feuille, col_tax, col_data,
                                   col_pcum)
    else:
        chart = add_species_series(chart, nom_feuille, col_tax, col_data_act,
                                   col_pcum)
    """Rappel des options dans le titre du graphique"""
    ligne_option = 'Sp = ' + sp_opt(isp)
    if val_pcat is not None:
        _str = ""
        for x in set(liste_taxo):
            if x != "":
                _str += x + "/"
        ligne_option = '{}; TW: {}= {}'.format(ligne_option, _str[:-1],
                                               val_pcat)
    else:
        ligne_option = '{}; TW: none'.format(ligne_option)
    chart.set_title({'name': titre_graf[loi - 1] + '\n' + ligne_option})
    chart.set_x_axis({
        'name': titre_axe[0],
        'log_base': 10,
        'crossing': 0,
        'major_gridlines': {
            'visible': True
        }
    })
    chart.set_y_axis({
        'min': 0,
        'max': 1,
        'major_unit': .1,
        'crossing': 0,
        'name': titre_axe[1],
        'major_gridlines': {
            'visible': True
        }
    })
    chart.set_size({'width': 896, 'height': 500})
    chart.set_legend({'position': 'bottom'})
    worksheet.insert_chart('A1', chart)
    """
    Ajoute une zone de texte avec les valeurs de r2, Pttest, GWM et GWSD
    """
    if loi > 1:
        worksheet.insert_textbox('P1', 'R² = {:.4f}\nKSpvalue = {:.3f}'.format(
            r2, weight_value), {'width': 128,
                                'height': 40})
        worksheet.insert_textbox(
            'P3', ('wm.lg = {:.2f}\nwsd.lg = {:.2f}'.format(mup, sigmap))
            if loi == 2 else
            ('wmin.lg = {:.2f}\nwmax.lg = {:.2f}\nwmode.lg = {:.2f}'.format(
                _min, _max, mode)), {'width': 128,
                                     'height': 40})


def add_species_series(chart, worksheet_name, col_tax, col_data, col_pcum):
    """
    Add data by species.

    :param chart: chart to be filled
    :param worksheet_name: name of the worksheet containing the actual chart
    :param col_tax: taxonomic column
    :param col_data: taxonomic data column
    :param col_pcum: cumulative weighted probability column
    """
    end = len(
        list(initialisation.Worksheets[worksheet_name].Cells.ix[2:, col_tax]
             .dropna()))
    x = 2
    while x < end:
        new_end = x + 1
        while (initialisation.Worksheets[worksheet_name].Cells
                .ix[new_end - 1, col_tax] ==
                initialisation.Worksheets[worksheet_name]
                .Cells.ix[new_end, col_tax]) and new_end < end:
            new_end += 1
        chart.add_series({
            'values': [worksheet_name, x, col_pcum, new_end, col_pcum],
            'categories': [worksheet_name, x, col_data, new_end, col_data],
            'name': initialisation.Worksheets[worksheet_name].Cells.ix[x - 1, col_tax],
            'marker': {
                'type': 'automatic'
                # 'type': 'square',
                # 'border': {
                #     'color': 'black'
                # },
                # 'fill': {
                #     'color': 'black'
                # }
            },
            'line': {
                'none': True
            }
        })
        x = new_end
    return chart


def decaler_graph(nom_feuille):
    """
    Modifie la position des graphiques contenus dans une feuille.

    :param nom_feuille: name of the worksheet containing the actual chart
    """
    decalage = 0
    for ch in initialisation.Worksheets[nom_feuille].ChartObjects:
        ch.Left += decalage
        decalage += 200


def tracer_courbe_empirique(workbook, nom_feuille, lig_deb, nbligne, col_data,
                            col_pcum):
    """
    Relie points ponderes dans le graph correspdant a la loi empirique.

    :param nseries: numero de la serie de donnees encours a representer
    :param nom_feuille: nom de la feuille de calcul ou sont les donnees
    :param lig_deb: ligne de debut des donnees numeriques representees
    :param nbligne: nombre de lignes a traiter
    :param col_data: colonne des donnees numeriques
    :param col_pcum: colonne des donnees de probabilite cumulees ponderees
                     empiriques
    """
    chart = workbook.add_chart({
        'type': 'scatter',
        'subtype': 'smooth_with_markers'
    })
    chart.add_series({
        'values':
        [nom_feuille, lig_deb, col_pcum, lig_deb + nbligne, col_pcum],
        'categories':
        [nom_feuille, lig_deb, col_data, lig_deb + nbligne, col_data],
        'name':
        "Weighted Empirical",
        'line': {
            'color': 'red'
        },
        'marker': {
            'type': 'none'
        }
    })
    return chart
