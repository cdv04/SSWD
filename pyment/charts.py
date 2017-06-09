# coding=utf-8
"""
Module permettant la mise ne place des graphiques.

"""

# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-04-12T09:13:56+02:00
# @Email:  zackary.beaugelin@epitech.eu
# @Project: SSWD
# @Filename: charts.py
# @Last modified by:   gysco
# @Last modified time: 2017-06-02T10:16:16+02:00

import initialisation
from common import sp_opt
from pandas import ExcelWriter


def draw_chart(writer: type(ExcelWriter),
               nom_feuille: str,
               lig_p: int,
               lig_qbe: int,
               lig_qbi: int,
               lig_qbs: int,
               col_deb: int,
               col_fin: int,
               ligne_data: int,
               col_tax: int,
               col_data: int,
               col_pcum: int,
               col_data_le: int,
               col_pcum_le: int,
               loi: int,
               titre_graf: str,
               r2: float,
               weight_value: float,
               nb_ligne_data: int,
               mup: float,
               sigmap: float,
               _min: float,
               _max: float,
               mode: float,
               titre_axe: list,
               val_pcat: list,
               liste_taxo: list,
               isp: int,
               col_data_act: int,
               col_data_act_le: object,
               iproc: int):
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
        axis=1).to_excel(
        writer, sheet_name=nom_feuille, index=False, header=False)
    worksheet = writer.sheets[nom_feuille]
    """Ajout de la serie Mediane"""
    if loi != 1:
        chart = workbook.add_chart({
            'type': 'scatter',
            'subtype': 'smooth_with_markers'
        })
        chart.add_series({
            'values': [nom_feuille, lig_p - 1, col_deb, lig_p - 1, col_fin],
            'categories':
                [nom_feuille, lig_qbe - 1, col_deb, lig_qbe - 1, col_fin],
            'name': [nom_feuille, lig_qbe - 1, col_deb - 1],
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
            workbook, nom_feuille, ligne_data + 1, nb_ligne_data,
            col_data_le,
            col_pcum_le) if iproc == 1 else tracer_courbe_empirique(
            workbook, nom_feuille, ligne_data + 1, nb_ligne_data,
            col_data_act_le, col_pcum_le)
    """Ajout de la serie quantile borne inf"""
    chart.add_series({
        'values': [nom_feuille, lig_p - 1, col_deb, lig_p - 1, col_fin],
        'categories':
            [nom_feuille, lig_qbi - 1, col_deb, lig_qbi - 1, col_fin],
        'name': [nom_feuille, lig_qbi - 1, col_deb - 1],
        'line': {
            'color': 'black',
            'dash_type': 'dash'
        },
        'marker': {
            'type': 'none'
        }
    })
    """Ajout de la serie quantile borne superieure"""
    chart.add_series({
        'values': [nom_feuille, lig_p - 1, col_deb, lig_p - 1, col_fin],
        'categories':
            [nom_feuille, lig_qbs - 1, col_deb, lig_qbs - 1, col_fin],
        'name': [nom_feuille, lig_qbs - 1, col_deb - 1],
        'line': {
            'color': 'black',
            'dash_type': 'dash'
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
        worksheet.insert_textbox('M1', 'R² = {:.4f}\nKSpvalue = {:.3f}'.format(
            r2, weight_value), {'width': 128,
                                'height': 40})
        worksheet.insert_textbox(
            'M3', ('wm.lg = {:.2f}\nwsd.lg = {:.2f}'.format(mup, sigmap))
            if loi == 2 else (
                'wmin.lg = {:.2f}\nwmax.lg = {:.2f}\nwmode.lg = {:.2f}'.format(
                    _min, _max, mode)), {'width': 128,
                                         'height': 40})


def add_species_series(chart: type(ExcelWriter.book), worksheet_name: str,
                       col_tax: int,
                       col_data: int, col_pcum: int) -> type(ExcelWriter.book):
    """
    Add data by species.

    :param chart: chart to be filled
    :param worksheet_name: name of the worksheet containing the actual chart
    :param col_tax: taxonomic column
    :param col_data: taxonomic data column
    :param col_pcum: cumulative weighted probability column
    """
    end = len(
        list(initialisation.Worksheets[worksheet_name].Cells.ix[3:, col_tax]))
    chart.add_series({
        'values': [worksheet_name, 2, col_pcum, end, col_pcum],
        'categories': [worksheet_name, 2, col_data, end, col_data],
        'name': [worksheet_name, 2, col_tax, end, col_tax],
        'marker': {
            'type': 'square',
            'border': {
                'color': 'black'
            },
            'fill': {
                'color': 'black'
            }
        },
        'line': {
            'none': True
        }
    })
    return chart


def decaler_graph(nom_feuille: str):
    """
    Modifie la position des graphiques contenus dans une feuille.

    :param nom_feuille: name of the worksheet containing the actual chart
    """
    decalage = 0
    for ch in initialisation.Worksheets[nom_feuille].ChartObjects:
        ch.Left += decalage
        decalage += 200


def tracer_courbe_empirique(workbook, nom_feuille, lig_deb, nbligne,
                            col_data, col_pcum):
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
        'line': {'color': 'red'},
        'marker': {'type': 'none'}
    })
    return chart
