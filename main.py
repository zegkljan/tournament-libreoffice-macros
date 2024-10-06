# coding: utf-8
from __future__ import unicode_literals

import helpers
import constants
import uno

try:
    CTX = XSCRIPTCONTEXT
except NameError:
    CTX = None

def init():
    doc = CTX.getDocument()
    
    ## prepare sheets
    # remove all but one sheet
    for _ in range(1, len(doc.Sheets)):
        sheet = doc.Sheets[-1]
        doc.Sheets.removeByName(sheet.getName())
    # rename remaining sheet to avoid conflicts
    doc.Sheets[0].setName('x')

    # do prep
    doc.getStyleFamilies()['CellStyles']['Default'].CharHeight = 12
    doc.getStyleFamilies()['CellStyles']['Default'].ParaTopMargin = 2 * 35
    doc.getStyleFamilies()['CellStyles']['Default'].ParaLeftMargin = 2 * 35
    doc.getStyleFamilies()['CellStyles']['Default'].ParaBottomMargin = 2 * 35
    doc.getStyleFamilies()['CellStyles']['Default'].ParaRightMargin = 2 * 35

    # create participant list sheet
    plist = helpers.addSheet(doc, constants.PARTICIPANT_LIST, 0)
    plist.getCellByPosition(0, 0).setString('Name')
    plist.getCellByPosition(1, 0).setString('Club/team')
    plist.getCellByPosition(2, 0).setString('Country')
    plist.getCellByPosition(3, 0).setString('Rating/rank')
    plist.getCellByPosition(4, 0).setString('Present?')

    # create settings sheet
    settings = helpers.addSheet(doc, constants.SETTINGS, 1)
    settings.getCellByPosition(0, 0).setString('Max group size')
    settings.getCellByPosition(1, 0).setValue(7)
    settings.getCellByPosition(0, 1).setString('Groups per row')
    settings.getCellByPosition(1, 1).setValue(4)
    settings.getCellByPosition(3, 1).setString('Number of groups per row in group display sheet, has no functional impact.')
    settings.getCellByPosition(0, 2).setString('To elimination')
    settings.getCellByPosition(1, 2).setValue(0.8)
    settings.getCellByPosition(3, 2).setString('Fraction of participants that will pass to the elimination.')
    settings.getCellByPosition(0, 3).setString('Rating is rank')
    settings.getCellByPosition(1, 3).setValue(1)
    settings.getCellByPosition(3, 3).setString('Indicates whether the number in the rating/rank column is rating (bigger is better), or rank (smaller is better). 1 => rank.')
    settings.getCellByPosition(0, 4).setString('Large groups first')
    settings.getCellByPosition(1, 4).setValue(1)
    settings.getCellByPosition(3, 4).setString('If the groups have different sizes, where should the people in the last layer be put? 1 = to the first groups, 0 = to the last groups.')
    settings.getCellByPosition(0, 5).setString('Team ranking N')
    settings.getCellByPosition(1, 5).setValue(0)
    settings.getCellByPosition(3, 5).setString('Teams are ranked by summing rank of best N members of each team. If set to <=0, team processing does not happen. If set to >0, teams proceed to elimination instead of participants, with the cut being applied to the teams.')
    settings.getCellByPosition(0, 6).setString('Fill groups random')
    settings.getCellByPosition(1, 6).setValue(0)
    settings.getCellByPosition(3, 6).setString('If >0, results of group bouts will be filled by random integers in the range [0, 5]')
    settings.getCellByPosition(0, 7).setString('Fill elimination random')
    settings.getCellByPosition(1, 7).setValue(0)
    settings.getCellByPosition(3, 7).setString('If >0, results of elimination bouts will be filled by random integers in the range [0, 5]')
    settings.Columns[0].OptimalWidth = True

    # remove the last sheet
    doc.Sheets.removeByName(doc.Sheets[-1].getName())
    
    ## set focus to participant list
    doc.getCurrentController().setActiveSheet(plist)


def schedule():
    doc = CTX.getDocument()
    for s in list(doc.Sheets):
        if s.getName() not in [constants.PARTICIPANT_LIST, constants.SETTINGS]:
            doc.Sheets.removeByName(s.getName())

    participants = helpers.loadParticipants(doc)
    if not participants:
        toolkit = CTX.getComponentContext().getServiceManager().createInstance('com.sun.star.awt.Toolkit')
        parent = toolkit.getDesktopWindow()
        from com.sun.star.awt import MessageBoxButtons
        mb = toolkit.createMessageBox(parent, 'errorbox', MessageBoxButtons.BUTTONS_OK, 'No participants', 'No participants were loaded. Are present participants marked as such?')
        mb.execute()
        return
    
    team = doc.Sheets[constants.SETTINGS].getCellByPosition(1, 5).getValue() > 0

    # create final ranking sheet
    final_ranking = helpers.addSheet(doc, constants.FINAL_RANKING, 2)
    final_ranking.getCellByPosition(0, 0).setString('Final rank')
    if team:
        final_ranking.getCellByPosition(1, 0).setString('Team')
        final_ranking.getCellByPosition(2, 0).setString('Elim. round')
        final_ranking.getCellByPosition(3, 0).setString('Quali')
        for i, _ in enumerate(set([p.club for p in participants])):
            final_ranking.getCellByPosition(0, i + 1).setValue(i + 1)
        helpers.defineDatabaseRange(doc, 'finalRanking', final_ranking, 0, 0, 3, i + 1)
    else:
        final_ranking.getCellByPosition(1, 0).setString('Name')
        final_ranking.getCellByPosition(2, 0).setString('Club')
        final_ranking.getCellByPosition(3, 0).setString('Elim. round')
        final_ranking.getCellByPosition(4, 0).setString('Quali')
        for i, _ in enumerate(participants):
            final_ranking.getCellByPosition(0, i + 1).setValue(i + 1)
        helpers.defineDatabaseRange(doc, 'finalRanking', final_ranking, 0, 0, 4, i + 1)
    
    # create list of fights sheet
    list_of_fights = helpers.addSheet(doc, constants.LIST_OF_FIGHTS, 3)
    list_of_fights.getCellByPosition(0, 0).setString('Phase')
    list_of_fights.getCellByPosition(1, 0).setString('Fighter 1')
    list_of_fights.getCellByPosition(2, 0).setString('Fighter 2')
    list_of_fights.getCellByPosition(3, 0).setString('Fighter 1 score')
    list_of_fights.getCellByPosition(4, 0).setString('Fighter 2 score')
    list_of_fights.getCellByPosition(5, 0).setString('Result')

    helpers.createGroups(doc, participants)
    helpers.createElimination(doc, participants)

def evalGroups():
    doc = CTX.getDocument()

    helpers.sortGroupRanking(doc)

def evalFinal():
    doc = CTX.getDocument()

    helpers.sortFinalRanking(doc)
