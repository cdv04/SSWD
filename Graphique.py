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
# @Last modified time: 2017-05-03T15:04:50+02:00

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
    # Charts.Add()
    # ActiveChart.ChartType = xlXYScatterSmooth
    # Menage
    # if ActiveChart.SeriesCollection.Count > 0:
    # for i in vbForRange(ActiveChart.SeriesCollection.Count, 1):
    # ActiveChart.SeriesCollection(i).Delete()
    """Ajout de la serie Mediane"""
    if (loi != 1):
        nseries = nseries + 1
        data_x = 'R' + lig_qbe + 'C' + col_deb + ':R' + lig_qbe + 'C' + col_fin
        data_y = 'R' + lig_p + 'C' + col_deb + ':R' + lig_p + 'C' + col_fin
        data_nom = 'R' + lig_qbe + 'C' + col_deb - 1
        # ActiveChart.SeriesCollection.NewSeries()
        # ActiveChart.SeriesCollection[nseries].XValues = '=' + nom_feuille + '!' + data_x
        # ActiveChart.SeriesCollection[nseries].Values = '=' + nom_feuille + '!' + data_y
        # ActiveChart.SeriesCollection[nseries].Name = '=' + nom_feuille + '!' + data_nom
        # _with0 = ActiveChart.SeriesCollection(nseries).Border
        # _with0.ColorIndex = 1
        # _with0.Weight = xlThin
        # _with0.LineStyle = xlContinuous
        # ActiveChart.SeriesCollection[nseries].MarkerStyle = xlNone
    else:
        if iproc == 1:
            tracer_courbe_empirique(nseries, nom_feuille, ligne_data + 1,
                                    nb_ligne_data, col_data_le, col_pcum_le)
        else:
            tracer_courbe_empirique(nseries, nom_feuille, ligne_data + 1,
                                    nb_ligne_data, col_data_act_le,
                                    col_pcum_le)
    # Ajout de la serie quantile borne inf
    nseries = nseries + 1
    data_x = 'R' + lig_qbi + 'C' + col_deb + ':R' + lig_qbi + 'C' + col_fin
    data_y = 'R' + lig_p + 'C' + col_deb + ':R' + lig_p + 'C' + col_fin
    # data_nom = "Confidence limits 5% - 95%"
    data_nom = 'R' + lig_qbi + 'C' + col_deb - 1
    # ActiveChart.SeriesCollection.NewSeries()
    # ActiveChart.SeriesCollection[nseries].XValues = '=' + nom_feuille + '!' + data_x
    # ActiveChart.SeriesCollection[nseries].Values = '=' + nom_feuille + '!' + data_y
    # ActiveChart.SeriesCollection[nseries].Name = '=' + nom_feuille + '!' + data_nom
    # _with1 = ActiveChart.SeriesCollection(nseries).Border
    # _with1.ColorIndex = 1
    # _with1.Weight = xlThin
    # _with1.LineStyle = xlDot
    # ActiveChart.SeriesCollection[nseries].MarkerStyle = xlNone
    # Ajout de la serie quantile borne superieure
    nseries = nseries + 1
    data_x = 'R' + lig_qbs + 'C' + col_deb + ':R' + lig_qbs + 'C' + col_fin
    data_y = 'R' + lig_p + 'C' + col_deb + ':R' + lig_p + 'C' + col_fin
    data_nom = 'R' + lig_qbs + 'C' + col_deb - 1
    # ActiveChart.SeriesCollection.NewSeries()
    # ActiveChart.SeriesCollection[nseries].XValues = '=' + nom_feuille + '!' + data_x
    # ActiveChart.SeriesCollection[nseries].Values = '=' + nom_feuille + '!' + data_y
    # ActiveChart.SeriesCollection[nseries].Name = '=' + nom_feuille + '!' + data_nom
    # _with2 = ActiveChart.SeriesCollection(nseries).Border
    # _with2.ColorIndex = 1
    # _with2.Weight = xlThin
    # _with2.LineStyle = xlDot
    # ActiveChart.SeriesCollection[nseries].MarkerStyle = xlNone
    # ActiveChart.Legend.LegendEntries(nseries).Delete
    """
    Mise en forme du graphique
    """
    # ActiveChart.Location(Where=xlLocationAsObject, Name=nom_feuille)
    # position de la legende
    # ActiveChart.HasLegend = True
    # ActiveChart.Legend.Select()
    # Selection.Position = xlBottom
    # format quadrillage et axe
    # ActiveChart.Axes(xlValue).Select()
    # _with3 = ActiveChart.Axes(xlValue)
    # _with3.MinimumScaleIsAuto = True
    # _with3.MaximumScale = 1
    # _with3.MinorUnitIsAuto = True
    # _with3.MajorUnitIsAuto = True
    # _with3.Crosses = xlAutomatic
    # _with3.ReversePlotOrder = False
    # _with3.ScaleType = xlLinear
    # .DisplayUnit = xlNone
    # _with3.HasMajorGridlines = True
    # _with3.HasMinorGridlines = False
    # _with3.TickLabels.NumberFormat = '0%'
    # graphique log
    # ActiveChart.PlotArea.Select()
    # _with4 = ActiveChart.Axes(xlCategory)
    # _with4.HasMajorGridlines = True
    # _with4.HasMinorGridlines = False
    # _with4.ScaleType = xlLogarithmic
    # _with5 = Selection.Border
    # _with5.ColorIndex = 16
    # _with5.Weight = xlThin
    # _with5.LineStyle = xlContinuous
    # Selection.Interior.ColorIndex = xlNone
    # ActiveChart.Axes(xlValue).MajorGridlines.Select()
    # _with6 = Selection.Border
    # _with6.ColorIndex = 15
    # _with6.Weight = xlHairline
    # _with6.LineStyle = xlContinuous
    # ActiveChart.Axes(xlValue).MajorGridlines.Select()
    # ActiveChart.Axes(xlCategory).MajorGridlines.Select()
    # _with7 = Selection.Border
    # _with7.ColorIndex = 15
    # _with7.Weight = xlHairline
    # _with7.LineStyle = xlContinuous
    # _with8 = ActiveChart
    # _with8.HasTitle = True
    # _with8.ChartTitle.text = titre_graf(loi)
    # _with8.Axes[xlCategory, xlPrimary].HasTitle = True
    # _with8.Axes[xlCategory, xlPrimary].AxisTitle.Characters.text = titre_axe(1)
    # _with8.Axes[xlCategory, xlPrimary].AxisTitle.Font.Size = 8
    # _with8.Axes[xlCategory, xlPrimary].AxisTitle.Font.Bold = False
    # _with8.Axes[xlValue, xlPrimary].HasTitle = True
    # _with8.Axes[xlValue, xlPrimary].AxisTitle.Characters.text = titre_axe(2)
    # _with8.Axes[xlValue, xlPrimary].AxisTitle.Font.Size = 8
    # _with8.Axes[xlValue, xlPrimary].AxisTitle.Font.Bold = False
    #
    # Ajoute les donnees data
    #
    if iproc == 1:
        ajoute_series(nom_feuille, nseries, True, ligne_data, col_tax,
                      col_data, col_pcum, col_pcum_a)
    else:
        ajoute_series(nom_feuille, nseries, True, ligne_data, col_tax,
                      col_data_act, col_pcum, col_pcum_a)
        ajoute_series(nom_feuille, nseries, False, ligne_data, col_tax,
                      col_data, col_pcum_a, col_pcum_a)
    #
    # Ajoute une zone de texte avec les valeurs de R2, Pttest, GWM et GWSD
    #
    _select0 = loi
    if (_select0 == 2):
        # ActiveChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 9, 7.5, 80, 25).Select()
        #   Selection.Characters.text = "R_ = " & FormatNumber(R2, 4)
        _str = 'R_ = ' + '{}'.format(
            ':.4f',
            R2, ) + '\n' + 'KSpvalue = ' + '{}'.format(':.3f', Pvalue)
        # Selection.Characters.text = _str
        # Selection.AutoScaleFont = False
        long_chaine = len(_str)
        # _with9 = Selection.Characters(Start=1, Length=long_chaine).Font
        # _with9.Name = 'Arial'
        # _with9.FontStyle = 'Normal'
        # _with9.Size = 8
        # ActiveChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 330, 7.5,
        # 80, 25).Select()
        _str = ' wm.lg = ' + '{}'.format(
            ':.2f', mup) + '\n' + ' wsd.lg = ' + '{}'.format(':.2f', sigmap)
        # Selection.Characters.text = _str
        # Selection.AutoScaleFont = False
        long_chaine = len(_str)
        # _with10 = Selection.Characters(Start=1, Length=long_chaine).Font
        # _with10.Name = 'Arial'
        # _with10.FontStyle = 'Normal'
        # _with10.Size = 8
    elif (_select0 == 3):
        # ActiveChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 9, 7.5, 80, 25).Select()
        # _str = "R_ = " & Format(R2, "#0.0000")
        _str = 'R_ = ' + '{}'.format(
            ':.4f',
            R2, ) + '\n' + 'KSpvalue = ' + '{}'.format(':.3f', Pvalue)
        # Selection.Characters.text = _str
        # Selection.AutoScaleFont = False
        long_chaine = len(_str)
        # _with11 = Selection.Characters(Start=1, Length=long_chaine).Font
        # _with11.Name = 'Arial'
        # _with11.FontStyle = 'Normal'
        # _with11.Size = 8
        # ActiveChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 330, 7.5,
        # 80, 35).Select()
        _str = ' wmin.lg = ' + '{}'.format(
            ':.2f', _min) + '\n' + ' wmax.lg = ' + '{}'.format(
                ':.2f', _max) + '\n' + ' wmode.lg = ' + '{}'.format(
                    ':.2f', mode)
        # Selection.Characters.text = _str
        # Selection.AutoScaleFont = False
        long_chaine = len(_str)
        # _with12 = Selection.Characters(Start=1, Length=long_chaine).Font
        # _with12.Name = 'Arial'
        # _with12.FontStyle = 'Normal'
        # _with12.Size = 8
    # Rappel des options dans le titre du graphique
    ligne_option = 'Sp = ' + sp_opt(isp)
    if val_pcat != '':
        ligne_option = ligne_option + '; TW: ' + liste_taxo + ' = ' + val_pcat
    else:
        ligne_option = ligne_option + '; TW: none'
    nb_car = len(titre_graf(loi)) + len(ligne_option)
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
    nb_style = 5
    nb_col = 14
    col_possible = [1, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16]
    style_possible = [1, 2, 3, 7, 8]
    i = 1
    nseries_pts = nseries - 3
    ldeb = ligne_data + 1
    if nouveau is True:
        col1_marqueur = 0
        style_marqueur = 0
    else:
        col1_marqueur = nseries_pts % nb_col
        if col1_marqueur == 0:
            col1_marqueur = nseries_pts
        style_marqueur = nseries_pts % nb_style
        if style_marqueur == 0:
            style_marqueur = nseries_pts
    col2_marqueur = col1_marqueur
    nbligne = 0
    while Initialisation.Worksheets[nom_feuille].Cells[ligne_data + i,
                                                       col_tax]:
        if (Initialisation.Worksheets[nom_feuille].Cells[ligne_data + i,
                                                         col_pcum] != 0):
            if (Initialisation.Worksheets[nom_feuille].Cells[ligne_data + i,
                                                             col_tax] !=
                    Initialisation.Worksheets[nom_feuille]
                    .Cells[ligne_data + i + 1, col_tax]):
                nseries = nseries + 1
                if style_marqueur == nb_style:
                    style_marqueur = 1
                else:
                    style_marqueur = style_marqueur + 1
                if col1_marqueur == nb_col:
                    col1_marqueur = 1
                else:
                    col1_marqueur = col1_marqueur + 1
                col2_marqueur = col1_marqueur
                data_x = 'R' + ldeb + 'C' + col_data + ':R' + ldeb + nbligne + 'C' + col_data
                data_y = 'R' + ldeb + 'C' + col_pcum + ':R' + ldeb + nbligne + 'C' + col_pcum
                data_nom = 'R' + ldeb + 'C' + col_tax
                # ActiveChart.SeriesCollection.NewSeries()
                # ActiveChart.SeriesCollection[
                #     nseries].XValues = '=' + nom_feuille + '!' + data_x
                # ActiveChart.SeriesCollection[
                #     nseries].Values = '=' + nom_feuille + '!' + data_y
                # ActiveChart.SeriesCollection[
                #     nseries].Name = '=' + nom_feuille + '!' + data_nom
                if nouveau is True and Initialisation.Worksheets[nom_feuille].Cells(
                        ldeb, col_pcum_a) != 0:
                    print("A completer")  # TODO
                #     ActiveChart.SeriesCollection[
                #         nseries].Name = ActiveChart.SeriesCollection(
                #             nseries).Name + '_ACT'
                # _with0 = ActiveChart.SeriesCollection(nseries).Border
                # _with0.ColorIndex = 1
                # _with0.Weight = xlThin
                # _with0.LineStyle = xlNone
                # _with1 = ActiveChart.SeriesCollection(nseries)
                # _with1.MarkerBackgroundColorIndex = col_possible(col2_marqueur)
                # .MarkerBackgroundColorIndex = xlNone
                # _with1.MarkerForegroundColorIndex = col_possible(col1_marqueur)
                # _with1.MarkerStyle = style_possible(style_marqueur)
                # _with1.MarkerSize = 5
                ldeb = ligne_data + i + 1
                nbligne = 0
            else:
                nbligne = nbligne + 1
        else:
            ldeb = ligne_data + i + 1
            nbligne = 0
        i = i + 1


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
    data_x = 'R' + lig_deb + 'C' + col_data + \
        ':R' + lig_deb + nbligne - 1 + 'C' + col_data
    data_y = 'R' + lig_deb + 'C' + col_pcum + \
        ':R' + lig_deb + nbligne - 1 + 'C' + col_pcum
    data_nom = 'Weighted Empirical'
    # ActiveChart.SeriesCollection.NewSeries()
    # ActiveChart.SeriesCollection[
    #     nseries].XValues = '=' + nom_feuille + '!' + data_x
    # ActiveChart.SeriesCollection[
    #     nseries].Values = '=' + nom_feuille + '!' + data_y
    # ActiveChart.SeriesCollection[nseries].Name = data_nom
    # _with0 = ActiveChart.SeriesCollection(nseries).Border
    # _with0.ColorIndex = 1
    # _with0.Weight = xlThin
    # _with0.LineStyle = xlContinuous
    # ActiveChart.SeriesCollection[nseries].MarkerStyle = xlNone
