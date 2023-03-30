# coding: utf-8
from __future__ import unicode_literals
from typing_extensions import final

import uno
import sys
import re
from collections import namedtuple

import algorithms
import constants


Participant = namedtuple('Participant', ['row', 'name', 'club', 'country', 'rating'])


def _printDir(x, grep='.*'):
    pat = re.compile(grep)
    lines = []
    for attr in sorted(x.__dir__()):
        if not pat.match(attr):
            continue
        try:
            obj = str(getattr(x, attr, None))
        except:
            obj = str('Error representing value: {}'.format(sys.exc_info()[0]))
        lines.append('{}: {}'.format(attr, obj))
    text = '\n'.join(lines)
    print(text)


def loadParticipants(doc):
    plist = doc.Sheets[constants.PARTICIPANT_LIST]
    participants = []
    i = 0
    while True:
        row = i + 1
        name = plist.getCellByPosition(0, row).getString()
        club = plist.getCellByPosition(1, row).getString()
        country = plist.getCellByPosition(2, row).getString()
        rating = plist.getCellByPosition(3, row).getValue()
        present = plist.getCellByPosition(4, row).getString()
        
        if present == 'y':
            participants.append(Participant(row, name, club, country, rating))
        if not name:
                break
        i += 1
    return participants


def addSheet(doc, name, position=None):
    if position is None:
        position = len(doc.Sheets)
    doc.Sheets.insertNewByName(name, position)
    sheet = doc.Sheets[position]
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(True)
    address = cursor.RangeAddress
    rng = sheet.getCellRangeByPosition(0, 0, address.EndColumn, 1000)
    rng.CellStyle = 'Default'
    return sheet


def createGroups(doc, participants):
    ## prepare cell styles
    thin_border = _makeBorderLine2(0, 35 // 2)
    medium_border = _makeBorderLine2(0, 35)
    thick_border = _makeBorderLine2(0, 2 * 35)
    _makeCellStyle(doc, 'scoring_table_default', dict(
        ParaTopMargin=150,
        ParaLeftMargin=150,
        ParaBottomMargin=150,
        ParaRightMargin=150,
    ), 'Default')
    _makeCellStyle(doc, 'scoring_table_number', dict(
        VertJustify=2,
        HoriJustify=2,
        TopBorder2=medium_border,
        RightBorder2=medium_border,
        BottomBorder2=medium_border,
        LeftBorder2=medium_border
    ), 'scoring_table_default')
    _makeCellStyle(doc, 'scoring_table_name', dict(
        VertJustify=2,
        HoriJustify=1,
        TopBorder2=medium_border,
        RightBorder2=medium_border,
        BottomBorder2=medium_border,
        LeftBorder2=medium_border
    ), 'scoring_table_default')
    _makeCellStyle(doc, 'scoring_table_inner', dict(
        VertJustify=2,
        HoriJustify=2,
        TopBorder2=medium_border,
        RightBorder2=medium_border,
        BottomBorder2=medium_border,
        LeftBorder2=medium_border
    ), 'scoring_table_default')
    _makeCellStyle(doc, 'scoring_table_inner_self', dict(
        VertJustify=2,
        HoriJustify=2,
        IsCellBackgroundTransparent=False,
        CellBackColor=0x00CCCCCC,
        TopBorder2=medium_border,
        RightBorder2=medium_border,
        BottomBorder2=medium_border,
        LeftBorder2=medium_border
    ), 'scoring_table_default')
    _makeCellStyle(doc, 'scoring_sheet_header', dict(
        VertJustify=2,
        HoriJustify=2,
        CharHeight=15,
        ParaTopMargin=235,
        ParaLeftMargin=235,
        ParaBottomMargin=235,
        ParaRightMargin=235,
    ), 'scoring_table_default')
    _makeCellStyle(doc, 'group_results_eliminated', dict(
        IsCellBackgroundTransparent=False,
        CellBackColor=0x00CCCCCC,
    ), 'Default')
    
    
    max_group_size = int(doc.Sheets[constants.SETTINGS].getCellByPosition(1, 0).getValue())
    groups_per_row = int(doc.Sheets[constants.SETTINGS].getCellByPosition(1, 1).getValue())
    cut_n = doc.Sheets[constants.SETTINGS].getCellByPosition(1, 2).getValue()
    if cut_n <= 1:
        cut_n = cut_n * len(participants)
    cut_n = round(cut_n)
    rating_is_rank = doc.Sheets[constants.SETTINGS].getCellByPosition(1, 3).getValue() == 1
    if rating_is_rank:
        sort_key = lambda x: x.rating
    else:
        sort_key = lambda x: -x.rating
    
    group_sizes = algorithms.findGroupSizes(len(participants), max_group_size)
    groups = algorithms.assignGroups(group_sizes, sorted(participants, key=sort_key), [(lambda p: p.club), (lambda p: p.country)])
    max_group_size = max(group_sizes)
    
    group_list_sheet = addSheet(doc, constants.GROUP_LIST, 2)
    
    group_results_sheet = addSheet(doc, constants.GROUPS_RESULTS, 3)
    group_results_sheet.getCellByPosition(0, 0).setString('Rank')
    group_results_sheet.getCellByPosition(1, 0).setString('Name')
    group_results_sheet.getCellByPosition(2, 0).setString('Club')
    group_results_sheet.getCellByPosition(3, 0).setString('V/M (↓)')
    group_results_sheet.getCellByPosition(4, 0).setString('D-R (↓)')
    group_results_sheet.getCellByPosition(5, 0).setString('D (↓)')
    group_results_sheet.getCellByPosition(6, 0).setString('R (↑)')
    group_results_sheet.getCellByPosition(7, 0).setString('RND')
    group_results_sheet.getCellRangeByPosition(0, 0, 7, 0).HoriJustify = 3
    group_results_sheet.getCellByPosition(1, 0).HoriJustify = 0

    final_ranking_sheet = doc.Sheets[constants.FINAL_RANKING]
    list_of_fights = doc.Sheets[constants.LIST_OF_FIGHTS]

    for i, group in enumerate(groups):
        schedule = algorithms.makeGroupSchedule(list(range(len(group))))
        group_name = 'Group {}'.format(i + 1)

        # write group into summary of all groups
        group_row = (i // groups_per_row) * (1 + max_group_size)
        group_col = (i % groups_per_row) * 3
        header_range = group_list_sheet.getCellRangeByPosition(group_col, group_row, group_col + 2, group_row)
        header_range.merge(True)
        header_range.TopBorder2 = header_range.RightBorder2 = header_range.BottomBorder2 = header_range.LeftBorder2 = medium_border
        group_list_sheet.getCellByPosition(group_col, group_row).setString(group_name)
        
        # create sheet for the group
        grp_sheet = addSheet(doc, group_name, 3 + i)
        grp_sheet.getCellRangeByPosition(0, 0, 1000, 1000).CellStyle = 'scoring_table_default'
        # sheet header
        grp_sheet.getCellByPosition(0, 0).setString(group_name)
        grp_sheet.getCellRangeByPosition(0, 0, len(group) + 5, 4).merge(True)
        grp_sheet.getCellByPosition(0, 0).CellStyle = 'scoring_sheet_header'
        
        # top-left coords for the two parts of the sheet
        schedule_coords = (len(group) + 7, 0) #(0, 1)
        table_coords = (0, 5) #(3 * (len(group) // 2) + 1, 1)
        
        # table header
        grp_sheet.getCellByPosition(*_add(table_coords, 1, 0)).setString('Name')
        grp_sheet.getCellByPosition(*_add(table_coords, 2 + len(group) + 0, 0)).setString('V/M')
        grp_sheet.getCellByPosition(*_add(table_coords, 2 + len(group) + 1, 0)).setString('D')
        grp_sheet.getCellByPosition(*_add(table_coords, 2 + len(group) + 2, 0)).setString('R')
        grp_sheet.getCellByPosition(*_add(table_coords, 2 + len(group) + 3, 0)).setString('Signature')
        # inner cells style
        grp_sheet.getCellRangeByPosition(*_add(table_coords, 0, 0), *_add(table_coords, 0, 1 + len(group) - 1)).CellStyle = 'scoring_table_number'
        grp_sheet.getCellRangeByPosition(*_add(table_coords, 1, 0), *_add(table_coords, 1, 1 + len(group) - 1)).CellStyle = 'scoring_table_name'
        grp_sheet.getCellRangeByPosition(*_add(table_coords, 2, 0), *_add(table_coords, 2 + len(group) - 1 + 4, 1 + len(group) - 1)).CellStyle = 'scoring_table_inner'
        for j, p in enumerate(group):
            participant_ref = _getParticipantReference(p)
            club_ref = _getParticipantClubReference(p)

            # write into summary group list
            num_cell = group_list_sheet.getCellByPosition(group_col, group_row + 1 + j)
            num_cell.setValue(j + 1)
            num_cell.LeftBorder2 = medium_border
            name_cell = group_list_sheet.getCellByPosition(group_col + 1, group_row + 1 + j)
            name_cell.setFormula('={}'.format(participant_ref))
            club_cell = group_list_sheet.getCellByPosition(group_col + 2, group_row + 1 + j)
            club_cell.setFormula('={}'.format(club_ref))
            club_cell.RightBorder2 = medium_border
            
            # write into scoring table
            # number column
            grp_sheet.getCellByPosition(*_add(table_coords, 2 + j, 0)).setValue(j + 1)
            # number row
            grp_sheet.getCellByPosition(*_add(table_coords, 0, 1 + j)).setValue(j + 1)
            # name
            grp_sheet.getCellByPosition(*_add(table_coords, 1, 1 + j)).setFormula('={}'.format(participant_ref))
            # victories / matches
            grp_sheet.getCellByPosition(*_add(table_coords, 2 + len(group) + 0, 1 + j)).setFormula('=({1}) / {0}'.format(len(group) - 1, '+'.join(['IF({} > {}; 1; 0)'.format(_c2s(*_add(table_coords, 2 + k, 1 + j)), _c2s(*_add(table_coords, 2 + j, 1 + k))) for k in range(len(group)) if k != j])))
            # dealt
            grp_sheet.getCellByPosition(*_add(table_coords, 2 + len(group) + 1, 1 + j)).setFormula('={0}'.format('+'.join([_c2s(*_add(table_coords, 2 + k, 1 + j)) for k in range(len(group)) if k != j])))
            # received
            grp_sheet.getCellByPosition(*_add(table_coords, 2 + len(group) + 2, 1 + j)).setFormula('={0}'.format('+'.join([_c2s(*_add(table_coords, 2 + j, 1 + k)) for k in range(len(group)) if k != j])))
            # self-match cell style
            grp_sheet.getCellByPosition(*_add(table_coords, 2 + j, 1 + j)).CellStyle = 'scoring_table_inner_self'

            # write into results table
            res_row = sum(group_sizes[:i]) + j + 1
            group_results_sheet.getCellByPosition(0, res_row).setValue(res_row)
            group_results_sheet.getCellByPosition(1, res_row).setFormula("={}".format(participant_ref))
            group_results_sheet.getCellByPosition(2, res_row).setFormula("={}".format(club_ref))
            group_results_sheet.getCellByPosition(3, res_row).setFormula("=$'{}'.{}".format(group_name, _c2s(*_add(table_coords, 2 + len(group) + 0, 1 + j))))
            group_results_sheet.getCellByPosition(4, res_row).setFormula('={} - {}'.format(_c2s(5, res_row), _c2s(6, res_row)))
            group_results_sheet.getCellByPosition(5, res_row).setFormula("=$'{}'.{}".format(group_name, _c2s(*_add(table_coords, 2 + len(group) + 1, 1 + j))))
            group_results_sheet.getCellByPosition(6, res_row).setFormula("=$'{}'.{}".format(group_name, _c2s(*_add(table_coords, 2 + len(group) + 2, 1 + j))))
            if res_row > cut_n:
                rng = group_results_sheet.getCellRangeByPosition(0, res_row, 7, res_row)
                rng.CellStyle = 'group_results_eliminated'
                if res_row == cut_n + 1:
                    rng.TopBorder2 = thick_border
                final_ranking_sheet.getCellByPosition(1, res_row).setFormula("=$'{}'.{}".format(constants.GROUPS_RESULTS, _c2s(1, res_row)))
                final_ranking_sheet.getCellByPosition(2, res_row).setFormula("=$'{}'.{}".format(constants.GROUPS_RESULTS, _c2s(2, res_row)))
                final_ranking_sheet.getCellByPosition(4, res_row).setValue(res_row)
            
            # when we are done
            if j == len(group) - 1:
                # close off bottom border in summary group list
                num_cell.BottomBorder2 = name_cell.BottomBorder2 = club_cell.BottomBorder2 = medium_border
                
                # set column widths in scoring table
                for k in range(len(group) + 6):
                    grp_sheet.Columns[_add(table_coords, k, 0)[0]].OptimalWidth = True
                grp_sheet.Columns[_add(table_coords, len(group) + 6, 0)[0]].Width = 100_0
        
        # finalize styling
        tb = uno.createUnoStruct('com.sun.star.table.TableBorder2')
        tb.TopLine = thick_border
        tb.IsTopLineValid = True
        tb.LeftLine = thick_border
        tb.IsLeftLineValid = True
        tb.BottomLine = thick_border
        tb.IsBottomLineValid = True
        tb.RightLine = thick_border
        tb.IsRightLineValid = True
        grp_sheet.getCellRangeByPosition(*_add(table_coords, 2, 1), *_add(table_coords, 2 + len(group) - 1, 1 + len(group) - 1)).TableBorder2 = tb
    
        schedule_cols = 2
        for j, (a, b) in enumerate(schedule):
            row = 2 * (j // schedule_cols)
            col = 3 * (j % schedule_cols)
            # first participant header
            grp_sheet.getCellByPosition(*_add(schedule_coords, col, row)).setValue(a + 1)
            grp_sheet.getCellByPosition(*_add(schedule_coords, col + 1, row)).setFormula('={}'.format(_getParticipantReference(group[a])))
            grp_sheet.getCellByPosition(*_add(schedule_coords, col, row)).TopBorder2 = thick_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col, row)).BottomBorder2 = thin_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col, row)).LeftBorder2 = thick_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col, row)).RightBorder2 = thin_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col + 1, row)).TopBorder2 = thick_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col + 1, row)).BottomBorder2 = thin_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col + 2, row)).TopBorder2 = thick_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col + 2, row)).BottomBorder2 = thin_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col + 2, row)).LeftBorder2 = thin_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col + 2, row)).RightBorder2 = thick_border
            # second participant header
            grp_sheet.getCellByPosition(*_add(schedule_coords, col, row + 1)).setValue(b + 1)
            grp_sheet.getCellByPosition(*_add(schedule_coords, col + 1, row + 1)).setFormula('={}'.format(_getParticipantReference(group[b])))
            grp_sheet.getCellByPosition(*_add(schedule_coords, col, row + 1)).BottomBorder2 = thick_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col, row + 1)).LeftBorder2 = thick_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col, row + 1)).RightBorder2 = thin_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col + 1, row + 1)).BottomBorder2 = thick_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col + 2, row + 1)).BottomBorder2 = thick_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col + 2, row + 1)).LeftBorder2 = thin_border
            grp_sheet.getCellByPosition(*_add(schedule_coords, col + 2, row + 1)).RightBorder2 = thick_border
            # first participant binding
            grp_sheet.getCellByPosition(*_add(table_coords, 2 + b, 1 + a)).setFormula('=IF(ISBLANK({0}); ""; {0})'.format(_c2s(*_add(schedule_coords, col + 2, row))))
            # second participant binding
            grp_sheet.getCellByPosition(*_add(table_coords, 2 + a, 1 + b)).setFormula('=IF(ISBLANK({0}); ""; {0})'.format(_c2s(*_add(schedule_coords, col + 2, row + 1))))

            # write into list of fights
            k = 1
            while True:
                if list_of_fights.getCellByPosition(0, k).getString() == '':
                    break
                k += 1
            list_of_fights.getCellByPosition(0, k).setString(group_name)
            list_of_fights.getCellByPosition(1, k).setFormula('={}'.format(_getParticipantReference(group[a])))
            list_of_fights.getCellByPosition(2, k).setFormula('={}'.format(_getParticipantReference(group[b])))
            list_of_fights.getCellByPosition(3, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(group_name, _c2s(*_add(schedule_coords, col + 2, row))))
            list_of_fights.getCellByPosition(4, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(group_name, _c2s(*_add(schedule_coords, col + 2, row + 1))))
            list_of_fights.getCellByPosition(5, k).setFormula("=IF($'{0}'.{1} < $'{0}'.{2}; \"LOSS\"; \"WIN\")".format(group_name, _c2s(*_add(schedule_coords, col + 2, row)), _c2s(*_add(schedule_coords, col + 2, row + 1))))
    
        for j in range(schedule_cols):
            grp_sheet.Columns[_add(schedule_coords, 3 * j + 0, 0)[0]].OptimalWidth = True
            grp_sheet.Columns[_add(schedule_coords, 3 * j + 1, 0)[0]].OptimalWidth = True
            grp_sheet.Columns[_add(schedule_coords, 3 * j + 2, 0)[0]].Width = 200_0
        grp_sheet.Columns[table_coords[0] + 2 + len(group)].IsVisible = False
        grp_sheet.Columns[table_coords[0] + 2 + len(group) + 1].IsVisible = False
        grp_sheet.Columns[table_coords[0] + 2 + len(group) + 2].IsVisible = False
    
    for i in range(len(groups)):
        group_col = (i % groups_per_row) * 3
        group_list_sheet.Columns[group_col].OptimalWidth = True
        group_list_sheet.Columns[group_col + 1].OptimalWidth = True
        group_list_sheet.Columns[group_col + 2].OptimalWidth = True
    
    group_results_sheet.Columns[0].OptimalWidth = True
    group_results_sheet.Columns[1].OptimalWidth = True
    group_results_sheet.Columns[2].OptimalWidth = True


def createElimination(doc, participants):
    border = _makeBorderLine2(LineStyle=0, LineWidth=35)
    _makeCellStyle(doc, 'elimination_bracket_line', dict(
        LeftBorder2=border
    ), 'Default')
    _makeCellStyle(doc, 'elimination_cell', dict(
        VertJustify=2,
        TopBorder2=border,
        LeftBorder2=border,
        BottomBorder2=border,
        RightBorder2=border,
    ), 'Default')
    _makeCellStyle(doc, 'elimination_number', dict(
        HoriJustify=0
    ), 'elimination_cell')
    _makeCellStyle(doc, 'elimination_name', dict(
        HoriJustify=1
    ), 'elimination_cell')

    if constants.ELIMINATION in doc.Sheets:
        doc.Sheets.removeByName(constants.ELIMINATION)
    el = addSheet(doc, constants.ELIMINATION, len(doc.Sheets) - 2)

    final_ranking_sheet = doc.Sheets[constants.FINAL_RANKING]
    list_of_fights = doc.Sheets[constants.LIST_OF_FIGHTS]

    cut_n = doc.Sheets[constants.SETTINGS].getCellByPosition(1, 2).getValue()
    if cut_n <= 1:
        cut_n = cut_n * len(participants)
    cut_n = round(cut_n)
    
    el_participants = list(range(cut_n))

    layer, num_layers = algorithms.makeElimination(el_participants)
    ln = 0
    number_width = None
    name_width = None
    small_final = None
    processed_fights = 0
    while ln < num_layers:
        next_layer = []
        col = 4 * ln
        finish = False
        if small_final is not None:
            finish = True
        for i in range(len(layer)):
            row = (4 * 2**ln) * i

            if ln > 0:
                row += sum([2**k for k in range(1, ln + 1)])

            winner = '=IF({0} > {1}; {2}; IF({0} < {1}; {3}; ""))'
            loser = '=IF({0} < {1}; {2}; IF({0} > {1}; {3}; ""))'
            
            top_number_cell = el.getCellByPosition(col, row)
            top_name_cell = el.getCellByPosition(col + 1, row)
            top_club_cell = el.getCellByPosition(col + 2, row)
            top_score_cell = el.getCellByPosition(col + 3, row)
            bottom_number_cell = el.getCellByPosition(col, row + 1)
            bottom_name_cell = el.getCellByPosition(col + 1, row + 1)
            bottom_club_cell = el.getCellByPosition(col + 2, row + 1)
            bottom_score_cell = el.getCellByPosition(col + 3, row + 1)

            top_number_cell_addr = _c2s(top_number_cell.getCellAddress().Column, top_number_cell.getCellAddress().Row)
            top_name_cell_addr = _c2s(top_name_cell.getCellAddress().Column, top_name_cell.getCellAddress().Row)
            top_club_cell_addr = _c2s(top_club_cell.getCellAddress().Column, top_club_cell.getCellAddress().Row)
            top_score_cell_addr = _c2s(top_score_cell.getCellAddress().Column, top_score_cell.getCellAddress().Row)
            bottom_number_cell_addr = _c2s(bottom_number_cell.getCellAddress().Column, bottom_number_cell.getCellAddress().Row)
            bottom_name_cell_addr = _c2s(bottom_name_cell.getCellAddress().Column, bottom_name_cell.getCellAddress().Row)
            bottom_club_cell_addr = _c2s(bottom_club_cell.getCellAddress().Column, bottom_club_cell.getCellAddress().Row)
            bottom_score_cell_addr = _c2s(bottom_score_cell.getCellAddress().Column, bottom_score_cell.getCellAddress().Row)

            el.getCellRangeByPosition(col, row, col, row + 1).CellStyle = 'elimination_number'
            el.getCellRangeByPosition(col + 1, row, col + 1, row + 1).CellStyle = 'elimination_name'
            el.getCellRangeByPosition(col + 2, row, col + 2, row + 1).CellStyle = 'elimination_name'
            el.getCellRangeByPosition(col + 3, row, col + 3, row + 1).CellStyle = 'elimination_number'

            if ln == 0:
                if layer[i][0] is None:
                    top_score_cell.setValue(-1)
                else:
                    top_number_cell.setFormula("=$'{}'.A{}".format(constants.GROUPS_RESULTS, layer[i][0] + 2))
                    top_name_cell.setFormula("=$'{}'.B{}".format(constants.GROUPS_RESULTS, layer[i][0] + 2))
                    top_club_cell.setFormula("=$'{}'.C{}".format(constants.GROUPS_RESULTS, layer[i][0] + 2))
                    if layer[i][1] is None:
                        top_score_cell.setValue(0)
                if layer[i][1] is None:
                    bottom_score_cell.setValue(-1)
                else:
                    bottom_number_cell.setFormula("=$'{}'.A{}".format(constants.GROUPS_RESULTS, layer[i][1] + 2))
                    bottom_name_cell.setFormula("=$'{}'.B{}".format(constants.GROUPS_RESULTS, layer[i][1] + 2))
                    bottom_club_cell.setFormula("=$'{}'.C{}".format(constants.GROUPS_RESULTS, layer[i][1] + 2))
                    if layer[i][0] is None:
                        bottom_score_cell.setValue(0)
            else:
                vert_bracket_len = 2 * (2**(ln - 1) - 1)
                if vert_bracket_len > 0:
                    c = top_number_cell.getCellAddress().Column
                    r = top_number_cell.getCellAddress().Row
                    el.getCellRangeByPosition(c, r - vert_bracket_len, c, r - 1).CellStyle = 'elimination_bracket_line'
                    c = bottom_number_cell.getCellAddress().Column
                    r = bottom_number_cell.getCellAddress().Row
                    el.getCellRangeByPosition(c, r + 1, c, r + vert_bracket_len).CellStyle = 'elimination_bracket_line'
                
                top_number = layer[i][0][0]
                top_name = layer[i][0][1]
                top_club = layer[i][0][2]
                bottom_number = layer[i][1][0]
                bottom_name = layer[i][1][1]
                bottom_club = layer[i][1][2]
                top_number_cell.setFormula(top_number)
                top_name_cell.setFormula(top_name)
                top_club_cell.setFormula(top_club)
                bottom_number_cell.setFormula(bottom_number)
                bottom_name_cell.setFormula(bottom_name)
                bottom_club_cell.setFormula(bottom_club)
            
            if layer[i][0] is not None and layer[i][1] is not None and not finish:
                # write into list of fights
                k = 1
                while True:
                    if list_of_fights.getCellByPosition(0, k).getString() == '':
                        break
                    k += 1
                phase_n = 2**(num_layers - ln)
                phase_name = 'Elimination {}'.format(phase_n)
                if phase_n == 4:
                    phase_name = 'Semi-finals'
                elif phase_n == 8:
                    phase_name = 'Quarter-finals'
                else:
                    phase_name = 'Elimination 1/{}'.format(phase_n // 2)
                list_of_fights.getCellByPosition(0, k).setString(phase_name)
                list_of_fights.getCellByPosition(1, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(constants.ELIMINATION, top_name_cell_addr))
                list_of_fights.getCellByPosition(2, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(constants.ELIMINATION, bottom_name_cell_addr))
                list_of_fights.getCellByPosition(3, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(constants.ELIMINATION, top_score_cell_addr))
                list_of_fights.getCellByPosition(4, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(constants.ELIMINATION, bottom_score_cell_addr))
                list_of_fights.getCellByPosition(5, k).setFormula("=IF($'{0}'.{1} < $'{0}'.{2}; \"LOSS\"; \"WIN\")".format(constants.ELIMINATION, top_score_cell_addr, bottom_score_cell_addr))
            
            refs = (winner.format(top_score_cell_addr, bottom_score_cell_addr, top_number_cell_addr, bottom_number_cell_addr),
                    winner.format(top_score_cell_addr, bottom_score_cell_addr, top_name_cell_addr, bottom_name_cell_addr),
                    winner.format(top_score_cell_addr, bottom_score_cell_addr, top_club_cell_addr, bottom_club_cell_addr))
            if layer[i][0] is not None and layer[i][1] is not None and len(layer) > 2:
                final_ranking_sheet.getCellByPosition(1, cut_n - processed_fights).setFormula(loser.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_name_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_name_cell_addr),
                ))
                final_ranking_sheet.getCellByPosition(2, cut_n - processed_fights).setFormula(loser.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_club_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_club_cell_addr),
                ))
                final_ranking_sheet.getCellByPosition(3, cut_n - processed_fights).setValue(2**(num_layers - ln))
                final_ranking_sheet.getCellByPosition(4, cut_n - processed_fights).setFormula(loser.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_number_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_number_cell_addr),
                ))
                processed_fights += 1
            elif finish:
                final_ranking_sheet.getCellByPosition(1, cut_n - processed_fights).setFormula(loser.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_name_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_name_cell_addr),
                ))
                final_ranking_sheet.getCellByPosition(2, cut_n - processed_fights).setFormula(loser.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_club_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_club_cell_addr),
                ))
                final_ranking_sheet.getCellByPosition(3, cut_n - processed_fights).setValue(2.2)
                final_ranking_sheet.getCellByPosition(4, cut_n - processed_fights).setFormula(loser.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_number_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_number_cell_addr),
                ))
                processed_fights += 1
                final_ranking_sheet.getCellByPosition(1, cut_n - processed_fights).setFormula(winner.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_name_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_name_cell_addr),
                ))
                final_ranking_sheet.getCellByPosition(2, cut_n - processed_fights).setFormula(winner.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_club_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_club_cell_addr),
                ))
                final_ranking_sheet.getCellByPosition(3, cut_n - processed_fights).setValue(2.1)
                final_ranking_sheet.getCellByPosition(4, cut_n - processed_fights).setFormula(winner.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_number_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_number_cell_addr),
                ))
                processed_fights += 1

                # write into list of fights
                k = 1
                while True:
                    if list_of_fights.getCellByPosition(0, k).getString() == '':
                        break
                    k += 1
                list_of_fights.getCellByPosition(0, k).setString('Final')
                list_of_fights.getCellByPosition(1, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(constants.ELIMINATION, top_name_cell_addr))
                list_of_fights.getCellByPosition(2, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(constants.ELIMINATION, bottom_name_cell_addr))
                list_of_fights.getCellByPosition(3, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(constants.ELIMINATION, top_score_cell_addr))
                list_of_fights.getCellByPosition(4, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(constants.ELIMINATION, bottom_score_cell_addr))
                list_of_fights.getCellByPosition(5, k).setFormula("=IF($'{0}'.{1} < $'{0}'.{2}; \"LOSS\"; \"WIN\")".format(constants.ELIMINATION, top_score_cell_addr, bottom_score_cell_addr))
            
            if finish:
                row += 2 + vert_bracket_len + 2 + 2
                top_number_cell = el.getCellByPosition(col, row)
                top_name_cell = el.getCellByPosition(col + 1, row)
                top_club_cell = el.getCellByPosition(col + 2, row)
                top_score_cell = el.getCellByPosition(col + 3, row)
                bottom_number_cell = el.getCellByPosition(col, row + 1)
                bottom_name_cell = el.getCellByPosition(col + 1, row + 1)
                bottom_club_cell = el.getCellByPosition(col + 2, row + 1)
                bottom_score_cell = el.getCellByPosition(col + 3, row + 1)
                top_number_cell_addr = _c2s(top_number_cell.getCellAddress().Column, top_number_cell.getCellAddress().Row)
                top_name_cell_addr = _c2s(top_name_cell.getCellAddress().Column, top_name_cell.getCellAddress().Row)
                top_club_cell_addr = _c2s(top_club_cell.getCellAddress().Column, top_club_cell.getCellAddress().Row)
                top_score_cell_addr = _c2s(top_score_cell.getCellAddress().Column, top_score_cell.getCellAddress().Row)
                bottom_number_cell_addr = _c2s(bottom_number_cell.getCellAddress().Column, bottom_number_cell.getCellAddress().Row)
                bottom_name_cell_addr = _c2s(bottom_name_cell.getCellAddress().Column, bottom_name_cell.getCellAddress().Row)
                bottom_club_cell_addr = _c2s(bottom_club_cell.getCellAddress().Column, bottom_club_cell.getCellAddress().Row)
                bottom_score_cell_addr = _c2s(bottom_score_cell.getCellAddress().Column, bottom_score_cell.getCellAddress().Row)
                el.getCellRangeByPosition(col, row, col, row + 1).CellStyle = 'elimination_number'
                el.getCellRangeByPosition(col + 1, row, col + 1, row + 1).CellStyle = 'elimination_name'
                el.getCellRangeByPosition(col + 2, row, col + 2, row + 1).CellStyle = 'elimination_name'
                el.getCellRangeByPosition(col + 3, row, col + 3, row + 1).CellStyle = 'elimination_number'
                top_number = small_final[0][0]
                top_name = small_final[0][1]
                top_club = small_final[0][2]
                bottom_number = small_final[1][0]
                bottom_name = small_final[1][1]
                bottom_club = small_final[1][2]
                top_number_cell.setFormula(top_number)
                top_name_cell.setFormula(top_name)
                top_club_cell.setFormula(top_club)
                bottom_number_cell.setFormula(bottom_number)
                bottom_name_cell.setFormula(bottom_name)
                bottom_club_cell.setFormula(bottom_club)

                final_ranking_sheet.getCellByPosition(1, cut_n - processed_fights).setFormula(loser.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_name_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_name_cell_addr),
                ))
                final_ranking_sheet.getCellByPosition(2, cut_n - processed_fights).setFormula(loser.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_club_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_club_cell_addr),
                ))
                final_ranking_sheet.getCellByPosition(3, cut_n - processed_fights).setValue(2.4)
                final_ranking_sheet.getCellByPosition(4, cut_n - processed_fights).setFormula(loser.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_number_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_number_cell_addr),
                ))
                processed_fights += 1

                final_ranking_sheet.getCellByPosition(1, cut_n - processed_fights).setFormula(winner.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_name_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_name_cell_addr),
                ))
                final_ranking_sheet.getCellByPosition(2, cut_n - processed_fights).setFormula(winner.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_club_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_club_cell_addr),
                ))
                final_ranking_sheet.getCellByPosition(3, cut_n - processed_fights).setValue(2.3)
                final_ranking_sheet.getCellByPosition(4, cut_n - processed_fights).setFormula(winner.format(
                    "$'{}'.{}".format(constants.ELIMINATION, top_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_score_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, top_number_cell_addr),
                    "$'{}'.{}".format(constants.ELIMINATION, bottom_number_cell_addr),
                ))
                processed_fights += 1

                # write into list of fights
                k = 1
                while True:
                    if list_of_fights.getCellByPosition(0, k).getString() == '':
                        break
                    k += 1
                list_of_fights.getCellByPosition(0, k).setString('Small final')
                list_of_fights.getCellByPosition(1, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(constants.ELIMINATION, top_name_cell_addr))
                list_of_fights.getCellByPosition(2, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(constants.ELIMINATION, bottom_name_cell_addr))
                list_of_fights.getCellByPosition(3, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(constants.ELIMINATION, top_score_cell_addr))
                list_of_fights.getCellByPosition(4, k).setFormula("=IF(ISBLANK($'{0}'.{1}); \"\"; $'{0}'.{1})".format(constants.ELIMINATION, bottom_score_cell_addr))
                list_of_fights.getCellByPosition(5, k).setFormula("=IF($'{0}'.{1} < $'{0}'.{2}; \"LOSS\"; \"WIN\")".format(constants.ELIMINATION, top_score_cell_addr, bottom_score_cell_addr))
            
            if i % 2 == 0:
                next_layer.append((refs, None))
            else:
                next_layer[-1] = (next_layer[-1][0], refs)
            if len(layer) == 2:
                refs = (loser.format(top_score_cell_addr, bottom_score_cell_addr, top_number_cell_addr, bottom_number_cell_addr),
                        loser.format(top_score_cell_addr, bottom_score_cell_addr, top_name_cell_addr, bottom_name_cell_addr),
                        loser.format(top_score_cell_addr, bottom_score_cell_addr, top_club_cell_addr, bottom_club_cell_addr))
                if i % 2 == 0:
                    small_final = (refs, None)
                else:
                    small_final = (small_final[0], refs)

        if ln == 0:
            el.Columns[col].OptimalWidth = True
            el.Columns[col + 1].OptimalWidth = True
            el.Columns[col + 2].OptimalWidth = True
            number_width = el.Columns[col].Width
            name_width = el.Columns[col + 1].Width
            club_width = el.Columns[col + 2].Width
        else:
            el.Columns[col].Width = number_width
            el.Columns[col + 1].Width = name_width
            el.Columns[col + 2].Width = club_width
        el.Columns[col + 2].IsVisible = False
        el.Columns[col + 3].Width = 278_0
        if finish:
            break
        layer = next_layer
        ln += 1


def _getParticipantReference(participant):
    return "$'{}'.{}".format(constants.PARTICIPANT_LIST, _c2s(0, participant.row))


def _getParticipantClubReference(participant):
    return "$'{}'.{}".format(constants.PARTICIPANT_LIST, _c2s(1, participant.row))


def _makeCellStyle(doc, name, props, parent=None):
    new_style = doc.createInstance('com.sun.star.style.CellStyle')
    cell_styles = doc.getStyleFamilies()['CellStyles']
    if cell_styles.hasByName(name):
        cell_styles.removeByName(name)
    cell_styles.insertByName(name, new_style)
    new_style.setPropertyValues(tuple(props.keys()), tuple(props.values()))
    if parent is not None:
        new_style.setParentStyle(parent)


def _makeBorderLine2(LineStyle, LineWidth):
    brd = uno.createUnoStruct('com.sun.star.table.BorderLine2')
    brd.LineStyle = LineStyle
    brd.LineWidth = LineWidth
    return brd


def _c2s(col, row):
    column = ""
    rem = col % 26
    div = col // 26
    column = chr(ord('A') + rem) + column
    while div > 0:
        rem = div % 26
        div = div // 26
        column = chr(ord('A') + rem) + column
    return column + str(row + 1)


def _add(coords, col, row):
    return (coords[0] + col, coords[1] + row)


def sortGroupRanking(doc):
    participants = loadParticipants(doc)
    rng = doc.Sheets[constants.GROUPS_RESULTS].getCellRangeByPosition(1, 1, 7, len(participants))
    
    vm = uno.createUnoStruct('com.sun.star.table.TableSortField')
    vm.Field = 2
    vm.IsAscending = False
    
    dr = uno.createUnoStruct('com.sun.star.table.TableSortField')
    dr.Field = 3
    dr.IsAscending = False

    d = uno.createUnoStruct('com.sun.star.table.TableSortField')
    d.Field = 4
    d.IsAscending = False

    r = uno.createUnoStruct('com.sun.star.table.TableSortField')
    r.Field = 5
    r.IsAscending = True

    rnd = uno.createUnoStruct('com.sun.star.table.TableSortField')
    rnd.Field = 6
    rnd.IsAscending = False

    desc = rng.createSortDescriptor()
    def set_val(name, value):
        for prop in desc:
            if prop.Name == name:
                prop.Value = value
                return
    set_val('IsSortColumns', False)
    set_val('BindFormatsToContent', False)
    set_val('MaxSortFieldsCount', False)
    # do it like this because LibreOffice, for some reason, does not add more than 3 sorting criteria, it overwrites the last one with whatever is set as last
    set_val('SortFields', uno.Any('[]com.sun.star.table.TableSortField', [rnd]))
    rng.sort(desc)
    set_val('SortFields', uno.Any('[]com.sun.star.table.TableSortField', [r]))
    rng.sort(desc)
    set_val('SortFields', uno.Any('[]com.sun.star.table.TableSortField', [d]))
    rng.sort(desc)
    set_val('SortFields', uno.Any('[]com.sun.star.table.TableSortField', [dr]))
    rng.sort(desc)
    set_val('SortFields', uno.Any('[]com.sun.star.table.TableSortField', [vm]))
    rng.sort(desc)

    equals = []
    for i in range(1, len(participants)):
        prev_vm = rng.getCellByPosition(2, i - 1).getString()
        prev_dr = rng.getCellByPosition(3, i - 1).getString()
        prev_d = rng.getCellByPosition(4, i - 1).getString()
        prev_r = rng.getCellByPosition(5, i - 1).getString()
        prev_rnd = rng.getCellByPosition(6, i - 1).getString()

        vm = rng.getCellByPosition(2, i).getString()
        dr = rng.getCellByPosition(3, i).getString()
        d = rng.getCellByPosition(4, i).getString()
        r = rng.getCellByPosition(5, i).getString()
        rnd = rng.getCellByPosition(6, i).getString()

        if prev_vm == vm and prev_dr == dr and prev_d == d and prev_r == r and prev_rnd == rnd:
            if equals and equals[-1][-1] == i - 1:
                equals[-1][-1] = i
            else:
                equals.append([i - 1, i])
    rng.CharColor = -1
    for a, b in equals:
        rng.getCellRangeByPosition(0, a, 5, b).CharColor = 0x00FF0000


def sortFinalRanking(doc):
    participants = loadParticipants(doc)
    rng = doc.Sheets[constants.FINAL_RANKING].getCellRangeByPosition(1, 1, 4, len(participants))
    
    el = uno.createUnoStruct('com.sun.star.table.TableSortField')
    el.Field = 2
    el.IsAscending = True
    
    qual = uno.createUnoStruct('com.sun.star.table.TableSortField')
    qual.Field = 3
    qual.IsAscending = True

    desc = rng.createSortDescriptor()
    for prop in desc:
        if prop.Name == 'SortFields':
            prop.Value = uno.Any('[]com.sun.star.table.TableSortField', (
                el,
                qual,
            ))
        if prop.Name == 'IsSortColumns':
            prop.Value = False
        if prop.Name == 'BindFormatsToContent':
            prop.Value = False
    rng.sort(desc)
