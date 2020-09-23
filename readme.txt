Q-Clips - Windows Clipboard Extender ©
iQ proPlus Software & Design
Phoenix, Arizona
Author:  Tom Moran

Thanks to all who have sent me requests for new features, bug reports, coding suggestions, etc. Some of the requests, like Scrolling Window Capture (Web Page), additional image formats (gif, png, tif, etc.), and multiple clip pasting are beyond the scope of the PSC project but will be included in the commercial version now in production.

========================================================================

Q-Clips© is a Windows Clipboard Extender featuring a Screen Capture with Magnification Window and Color Capture utilities.  This readme file contains an update history and a brief overview of the project.  For full details on using Q-Clips and it's features view the included Q-Clips Help Manual available from the Help Menu in the Q-Clips program.

IMPORTANT:  When running in the VB IDE do not use the Break or End options and always run with the Full Compile option.  The Q-Clips code subclasses both the Clipboard Chain and the system wide Hot Keys.  Terminate the program by choosing the Shutdown option of Q-Clips otherwise you will surely crash and freeze.

Requirements:  This program has been tested in Windows XP, Vista and 2000 only.  The program may not work in the older Windows 9x operating systems.
____________________________________________________________________________

================
 Update History
================

Version 4.01 - December 27th, 2007
----------------------------------
• Added View Color Tool: In Ver. 4.0 we added the color capture tool. Feedback said if you were capturing a number of different color's (color numbers) it was near impossible to remember what color was represented by the captured numbers.  The View Color Tool will now allow you to view the actual color represented by the captured color number.  Here's the way it works:  Highlight the captured color number in the Q-Clips List. Now, press the Ctrl key and click the Left Mouse Button on the number.  The background of the thumbnail will turn to the color represented by the captured number. This works for all types of color numbers... Long, Hex, RGB or the combo of all three.

• Color Capture Bug Fix: There were some instances where the long and hex options were dropping the last digit of the color number. This has been corrected.

• Updated Q-Clips Help Manual to cover the new View Color Tool.  The help manual is in PDF format and thus requires Adobe Reader or other PDF file reader to view.  This manual is accessible from the Help Menu of Q-Clips


Version 4.0 - December 15th, 2007
----------------------------------
This may be the final release of Q-Clips code save for any bug fixes. A commercial version of Q-Clips is now in production.  I've had many requests for the changes in this release, especially for the option of setting color themes.

• Added Set Themes option: The only theme color available in earlier versions of Q-Clips was blue. This looked kind of funky for those running other XP/Vista themes like Silver or Olive.  Since XP and Vista has a themes feature we've had many requests to add this feature. We've included standard Blue (default), Silver and Olive. In addition there is Black and also a Windows Default color theme.

• Added Color Capture Tool: This tool will allow you to capture the RGB Colors, Hex Color, and Long Color number of any pixel appearing on your screen.  This tool is activated either from a system wide Hot Key (F8 is default) or from the Color Pick button on the Q-Clips window. The color numbers are saved to the clipboard as text and can then be pasted into programs, like Visual Basic, Paint and Imaging programs, that work with color numbers. Please see the included Q-Clips Help Manual for details on using the Color Capture Tool.

• Replaced Set Hot Keys button on View Pane with Color Pick (color capture) button. The Set Hot Keys option is accessible from the Options Menu of the Q-Clips Window.

• Modified the Set Hot Keys window to include setting the system wide hot key for the color capture tool.

• Updated Q-Clips Help Manual to cover new features.  This is in PDF format and thus requires Adobe Reader or other PDF file reader to view.  This manual is accessible from the Help Menu of Q-Clips.


Version 3.52 - November 19th, 2007
----------------------------------
Bug Fix:  Sorry... SetHotKey menu from the Options menu actually started a region screen capture in 3.51 although the Set Hot Key button worked ok.  It's now fixed.


Version 3.51 - November 19th, 2007
----------------------------------
• Added Screen Capture Types/Profiles. You can choose Region, Desktop or Active Window. The default is Region. Please see the Q-Clips Help Manual for full details on using the Screen Capture and the different capture types.

• Added saving to JPEG option to "Save Clip As..." feature.

• Added Image Type to Options. Clip images can now be saved in bmp or jpg. Default is bmp. For speed choose bmp; for size choose jpg.

• Changed Turn Clipboard On/Off caption to Turn Q-Clips Capture On/Off. This more accurately describes the function since turning off capturing in Q-Clips does not turn off the Windows clipboard capture.

• Q-Clips Capture On/Off state now saved to ini file. Last state of this button will be the default state upon next start of Q-Clips. This provides the option of starting Q-Clips with Q-Clips capture turned off.

• Added Code to play 25clips.wav which was inadvertently left out of last release. 

• Minor bug fixes  

 
Version 3.11 - October 18th, 2007
---------------------------------
  • Changed default clip from first to most recent when activating
    from Hotkey.
  • When turning clipboard back on after toggled off, clip in Windows
    clipboard will now be copied to Q-Clip collection
  • Set max clip warning to show only once per clip collection when
    Show Warnings is true
  • Added 25clips.wav to play 1 time when 25th clip is added and
    sound option is on. Also will restore window if minimized to tray.
  • Optimized code
  • Corrected Ini file code and Declare assignment
  • Minor bug fixes

Version 3.0 - Oct. 2nd, 2007
-----------------------------
  • Added Run on Start-Up option to the Options Menu
  • Added Save Clip As... to the Edit Menu
  • Changed Files Clips to actually copy file items when pasted
    to programs like Windows Explorer / My Computer.  (See below)
  • Minor bug fixes.

Version 2.12 - Sept. 22nd, 2007
--------------------------------
  • Bug fixes
  • Added VBHotKey ctl
  • Code Release
  
Version 2.0 - Sept 1st, 2007
-----------------------------
  • Vista and XP Beta Test

____________________________________________________________________________

Q-Clips Overview:   Q-Clips enhances the native Windows clipboard functionality by remembering all items (text, graphic, color numbers and files) that are copied to the clipboard and storing them in a collection for later pasting... even after you shut off your computer.  With Q-Clips custom load and save option you have an unlimited amount of clips available to you with just a couple mouse clicks. 

You can work in virtually any Windows program while Q-Clips sits in your windows system tray... capturing all data that is copied to the Windows Clipboard.  When you need to access those clips a simple press of a HotKey (default is  Ctrl+Q keys) and a click on the desired clip will automatically paste your selected clip directly into that program.  The clip will be pasted in at the current cursor location of that program. Q-Clips knows what program you're working in and will automatically paste the clip into your program.  

Be sure the focus is set on the edit or image area of the program in which you are working when pressing the Q-Clips Hot Key.  Once you've hit the Hot Key you can continue to paste multiple clips to your program without having to press the HotKey again as Q-Clips remembers that program until you change it by pressing the HotKey in some other Windows program.

You can immediately view and edit all text and graphic clips captured by Q-Clips just by left clicking the thumbnail clip or, by right clicking on the list item and choose Open from the pop-up menu.  And if you've captured an Internet address (URL) you can click on the Thumbnail of that clip in the View Pane and your Internet Browser will open and take you directly to that page.

Pasting and Viewing File Clips:  A Files Clip is a list of files to be copied that was sent to the clipboard by a program such as Windows Explorer (My Computer).  This type of clip can be identified by the file folder icon that appears to the left of the clip in the Q-Clips list.  This clip is designed to copy the files in the list to a specific folder.  To paste a list of files to a folder start the files program (Windows Explorer or My Computer for example).  Go to the folder where you wish to copy the files.  Press the Q-Clips Hotkey, click on the Files Clip and those files will be copied to the selected folder.  Alternatively you can paste the Files Clip to the Windows clipboard and manually copy them by choosing Paste from the Windows Explorer edit menu

Editing and Viewing File Clips:   To view or edit a File Clip move your mouse to the Thumbnail area of the View Pane and press the left mouse button.  The list of files will open in your default text editor.   Once loaded you can edit or save that list as a text file.  It does not change the content of the actual Q-Clip file clip.  If you copy the displayed list in your text editor it will appear as a new text clip and not a file clip.

When the program begins, Q-Clips hooks into the Windows Clipboard Chain.  A system wide HotKey is assigned and the program is assigned to the system tray.  The Q-Clips window will appear on your screen.  You do have an option of starting Q-Clips minimized.  Once started, all data, text, files, graphics, color numbers and formatted text, copied to the Windows clipboard will be stored in Q-Clips.  The clips are actually stored in control arrays and later saved for retrieval when loading Q-Clips.  There is a maximum of 25 clips per collection.  On the 26th capture the oldest clip (first one) is discarded and the new one is added.

A screen capture tool is also available and includes a Magnify Window for more precision captures.  To capture images from the Internet, right click on any image that appears in your browser and click copy.  The image will be saved to Q-Clips.

========================================================================

There are multiple features and options in Q-Clips.  Please read the included manual for details on all the features available and their use.  The manual is available from the Help Menu while running Q-Clips.  It is in PDF format and requires Adobe Acrobat or other PDF Reader.

If you would like to offer suggestions, bug reports, comments and so forth you can email me directly at tmoran4511@hotmail.com

Acknowledgements:  Candy Buttons Control source from Mario Villanueva. Q-Clips List display is the control ucCoolList from Carles P.V.  VBHotKey control from Merrion Computing. JPEG code from John Korejwa’s JPEG Encoder Class Module. 

Note: The help manual was created in iQ WordPad and then converted to PDF format. iQ WordPad source code is also available on Planet-Source-Code at:

http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=69067&lngWId=1

Copyrights and Use:  Q-Clips is a copyrighted name and the code provided is likewise copyrighted by iQ proPlus Software and Design.  This code is intended for personal use.  You can use the program, modify the code, compile the program and use on multiple computers for personal use.  You may not, however, distribute this code publicly in any form without the express permission of iQ proPlus.  That means even if you distribute as freeware, shareware or donations requested.

The purpose of open source code projects is to help each other learn and for personal growth and use.

We've recently encountered a nasty piracy of our code for iQ WordPad by a Mr. Michael Hardy. He simply took the iQ WordPad code and design we uploaded to PSC, renamed it as UltraPad, claimed all credit for authoring (even dedicating the program to his daughter Zoe), and uploaded the compiled version to several web sites requesting donations.  Our lawyer has begun action against Mr. Hardy who has been down this road before.  If you have any questions in regard to use of our software please contact me at tmoran4511@hotmail.com.
















