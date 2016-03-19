Ulli's VB Companion AddIn
=========================

(formerly IDE Mousewheel Support and AutoComplete)

This Add-In adds configurable mousewheel support and a few other goodies to the VB IDE. Check it out; download is only 85kB (the best 85 kB you ever downloaded *g*).

It will also autocomplete words in the IDE while you type. There are three sources from which autocompletion is drawn: Private, Friend and Public variable- and procedure-names, a rather long list of VB keywords and last but not least the Win32API definitions. When Autocomplete pops up you can either continue typing if the word is not the one you expected, or you can press any of the dead keys (eg Cursor left/right). Also, there is an option to open a selection list and copy names into the IDE from there. 

Compile the DLL into your VB folder, use the AddIns manager to add it to the Addins menu, and then restart VB.

======
How to     »»  Hold down secondary (right) mouse button and rotate wheel to open
======         Options Dialog Box.

           »»  Hold down Shift or Cntl or both while scrolling to decrease the
               scroll distance temporarily.

           »»  Press Pause key to open member list.
               The member names shown are those only which are within scope of the current
               code module.
               Click on any member name to insert it into your code at the black arrow 
               location.
               Click on column header to sort by that column.
               Click on Exit or press Escape to abandon.

   NEW     »»  Press Shift+Pause keys to open multiline literal box. The literal box
               will help you to design long multiline text literals and convert them to
               the proper VB syntax, including newlines, quotes and line continuation
               marks.
               Enter the text as you would like to see it during program execution.
               Press Pause key or click on Preview to see an actual example of the
               VB-interpretation of the converted "syntaxed" text in a message box.
               Press Shift+Return to insert the converted text into your code at the
               black arrow location.
               Press Escape to abandon.
               You can also copy a literal from the codepane and paste that into
               the box, the VB syntax will be un-converted during the paste process.

   NEW     »»  Click on Compare (main VB menu bar) to see all alterations of the current
               module (code AND visual elements) which were made since you last saved 
               the current module. The current and saved modules are considered equal 
               if each non-blank line in one of them has a matching non-blank line in 
               the other. Lines match when they are in the same ordinal position and 
               have identical contents (disregarding upper/lower case(option) and 
               leading/trailing/embedded spaces). The algorithm does not depend on line 
               numbers or such, the synchronization and matching process is by contents only.

   NEW     »»  Click on OpenAll (main VB menu bar) to open all available codepanes.

   NEW     »»  Click on Reset (main VB menu bar) to reset all changes (code AND visual
               elements) which were made since you last saved the current module.

   NEW     »»  Click on Copy (main VB menu bar) to open the Copy Facility. In the CF then open 
               any VB file, select the text you wish to copy and add it to the clipboard; this
               process can be repeated until you have collected all the code you wish to paste.
               Then finally click Paste to insert the code at the arrow position.

           »»  Press any dead key (Cursor left/right/up/down, Page up/down, Pos1, End)
               to get you out of the selected text and confirm the autocomplete.

           »»  Type a questionmark after an API name which fails to trigger autocompletion. 
               If it still fails it isn't there (or faulty; yes, there are some faulty entries
               in Win32Api.mdb).

           »»  Click middle button or mousewheel to return to caret position.

           »»  Click on the horiz scrollbar thumb just once if the raster fails to adjust
               to the horizontal position or doesn't show at all.

           »»  In the Options Popup, when you turn on the Raster option a color dialog
               will let you select the raster color.

           »»  In the Options Popup, click 'Refresh/Reset' if you have altered anything in 
               the IDE that would affect this Add-In (Font, FontSize, Indicatorbar, Color
               or TabWidth) or if you wish to go back to the last saved registry settings.


The following keywords had to be replaced to trigger autocomplete from Win32Api

      instead of                type
      ------------------------------------
      Declare Sub/Function      ApiDeclare
      Const                     ApiConst
      Type                      ApiType


...and then type a space and then the API member name you wish to autocomplete.

The API database is usally located there --> 
C:\Programs\Microsoft Visual Studio\Common\Tools\Winapi\WIN32API.MDB

When you don't have that you can simply run APILOAD.EXE using WIN32API.TXT and 
convert that to a .MDB by selecting the appropriate menu-items in menu 'Files'

Registry item at "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\VBCompanion\Options\ApiLocation" can be changed to [unknown] to inhibit 
searching for the Api database.

The previous registry entry named "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Mousewheel..." may be deleted manually.

...............................................................................

Known quirk:
------------

When a member (variable or whatever) is deleted from the code then this addin has no chance to notice that change until the user also changes some other codeline and then leaves that line. Therefore a member-name may still pop up in autocomplete or the selection box when in fact that member does not exist any more.

.................................................................................

It is suggested that you also read the notes and comments at the top of module dCompanion and of some others.


Good luck and happy hacking

Ulli
