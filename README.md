# LibreOffice macros for tournament organization
This repository contains a set of several python macros that can be used in LibreOffice Calc to organize a match-based tournament (i.e. pairs of competitors have matches which determine one winner) consisting of a group phase and a single-elimination phase.

Developed for the purposes of organizing small [HEMA](https://en.wikipedia.org/wiki/Historical_European_martial_arts) tournaments.

## Installation
Just clone (or downlad) this repository into the python scripts folder for your LibreOffice, which is

**Linux**, **MacOS**  
`$HOME/.config/libreoffice/4/user/Scripts/python`

**Windows**  
`%APPDATA%\LibreOffice\4\user\Scripts\python`

and restart LibreOffice.
You may have to install python support for your LibreOffice, depending on how it is packaged on your system (tested only on Arch Linux).

For further details about macro locations see [Python Scripts Organization and Location](https://help.libreoffice.org/6.3/en-US/text/sbasic/python/python_locations.html).

### APSO
It is totally optional, but recommended, to install the [APSO - Alternative Python Script Organizer](https://extensions.libreoffice.org/en/extensions/show/apso-alternative-script-organizer-for-python) extension into your LibreOffice.
APSO enables better access (and debugging, if needed) to the Python macros.
For example (at least on Linux), the keyboard shortcut `Alt+Shift+F11` opens up a window with all the macros which can be executed from there, instead of the need to go through the menu Tools → Macros → Run macro...

## Usage
There are four macros (functions within the `main.py` file) which do all the work.
Now follows their description, listed in the typical calling order.

### `init`
**Deletes all sheets from the document** and initializes:
* *Participant list* sheet which holds the list of all participants.
  There are four columns:
  * *Name* - the name of the participant
  * *Club* - the participant's club
  * *Rating/rank* - the participant's rating (higher number means better) or rank (lower number is better)
  * *Present* - indicator whether the participant checked in at the start of the tournament.
    Fill in `y` for those who check in.
* *Settings* sheet which holds the settings of the tournament:
  * *Max group size* - maximum size of a group in the group phase.
    The group generating algorithm will aim to create a set of groups such that the biggest and smallest groups differ at most by 1 with the biggest group having no more than *Max group size* participants.
    It can have less (e.g. if there are just 10 participants and the size is 7, there will be two groups of 5).
  * *Groups per row* - number of groups per row on Group list (see the macro `schedule`).
  * *To elimination* - fraction (i.e. number between 0 and 1) of participants that will be admitted to the elimination phase.
  * *Rating is rank* - if set to `1`, the value in the *Rating/rank* column in the *Participant list* sheet will be used as rank (i.e. lower is better), otherwise as rating (i.e. higher is better).

### `schedule`
Schedules the whole tournament according to the settings and the list of participants.
Namely:
* Determines the number and size of the groups.
* Assigns participants into the groups according to their rating/rank.
* Schedules the (order of) fights in each group.
* Prepares evaluation of the group phase.
* Schedules the elimination bracket.
* Prepares final evaluation.

Document-wise, a number of sheets is created:
* *Group list* - only lists all the groups that were created.
* *Group N* - for each group, a sheet with its number is created.
  This sheet contains the scoring table and a list of fights in the group.
  The fights are to be executed in the order left to right, top to bottom.
  The results (poins for each fighter) should be written into the cells in the list of fights, and they will be automatically written into the scoring table.
  The sheet can, of course, be printed.
* *Groups - results* - ranking of participants after the group phase.
  See the next macro `evalGroups` for details.
* *Elimination* - the elimination bracket.
  The participants are automatically filled into the proper positions, and it is live as *Groups - results* is being changed.
  When the elimination is ongoing, simply fill the elimination fight score into the cells following the names, and the correct participant will appear in the next layer.
* *Final ranking* - total final ranking of all participants when the tournament is over.
  See the macro `evalFinal` for details.
* *List of fights* - a list of all fights that will have taken place during the tournament.
  The column *Result* is from the point of view of *Fighter 1*.
  This sheet is also 'live', meaning that the contents are updated as the corresponding results are filled in (the group sheets, the elimination sheet).

**IMPORTANT** - running `schedule` will delete and re-create all sheets except for *Participant list* and *Settings*.
That means that any possible tournament progress **will be lost**, if this macro is called again.

### `evalGroups`
Evaluates the ranking of the participants when the group phase is over.
It sorts the participants in the sheet *Group - results*.
The sorting is done based on the columns *V/M* (victories divided by \# of fights) descending, *D-R* (points dealt minus points received) descending, *D* (points dealt) descending, *R* (points received) ascending, *RND* (random number) descending.
The column *RND* is used for breaking ties, if one should arise.
Normally, the *RND* column should be left empty.
If there is any tie, the rows where there is a tie will be highlighted in red after `evalGroups` is called.
In that case, you should put some values in the *RND* column for these rows and call `evalGroups` again.

### `evalFinal`
Performs the final ranking of the tournament in the sheet *Final ranking*.
It sorts the participants by their highest elimination bracket layer, and then by their group phase rankings (i.e. mutual ranking of participants who dropped out in the same elimination layer will be the same as their mutual group phase ranking).

## Limiations
These macros **do not** take care of the following:
* a participant dropping out of the tournament - you need to encode this information into the score (e.g. put 1:0 for all their fights)
* draws - each fight has to have one winner and one loser
* other tournament formats - the groups+elimination combo is fixed and cannot be changed
* group sizes <= 4 - if there should be a group of size 4 or less (example 1: you set *Max group size* to 4; example 2: there are 7 participants and *Max group size* is smaller than 7), an error is thrown, because the groups cannot be scheduled such that each member of the group has a pause between their fights at least 1 other fight long