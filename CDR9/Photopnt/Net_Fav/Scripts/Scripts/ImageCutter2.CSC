REM	Image Cutter - cut up file into numbered pieces
REM	For use with Photo-Paint 9 

REM
REM	To use with earlier versions of Paint, uncomment the 
REM	lines with PAINT8, comment out the lines with PAINT9,
REM	and change the Automation version of CorelPhotoPaint
REM	from 9 to 8
REM
REM	To Do List:
REM		- add more filters
REM
REM	Wish List
REM		- add a 'Reset' button for the Horizontal/Vertical 
REM		  values
REM		- add support for allowing different Horizontal and 
REM	  	  Vertical units
REM		- add a Browse button(!)
REM
REM	Created on April 6, 1998 by Rick Mutzke
REM	********************************************************

REM	-----------------------------------------------------------------
REM	Function declarations.
REM

Declare Function ConvertToPixels(Value as double, Res as double, Units as long) as long
Declare Function ConvertToUnit(Value as double, Res as double, OldUnits as long, NewUnits as long) as double
Declare Function ConvertBaseToUnit(Value as double, Res as double, Units as long) as double
Declare Function PointsToString(dValue as double, unNumDecPlaces as long) as string
Declare Function StringToPoints(Value as string) as double
Declare Function GetPath(FullPath as string) as string
Declare Function GetExtension(FullPath as string) as string
Declare Function GetVersionString(FileNum as long, NumDigits as long) as string
Declare Function GetFilterID(Extension as string) as long
Declare Function UnitsToIndex(Idx as long) as long
Declare Function RemoveBackslash(NewSavePath as string) as string
Declare Function GetFileCount(nFiles as long) as string
Declare Function Round(Value as double) as long
Declare Function Floor(Value as double) as double
Declare Function Power(Base as double, Expon as long) as double
Declare Function RemoveExt(FileName as string)as string
Declare Sub GenerateHTML (ArrayIndex as integer)
Declare Sub CreatePage (DesDirectory as String, BaseName as string)
Declare Sub SubDimension()
Declare Sub SubSubDimension()


REM	-----------------------------------------------------------------
REM	Variable delarations.
REM
Global    nNumFiles 		as long
Global	nNumXTiles		as long		
Global	nNumYTiles		as long
Global 	URLArray(100) 		as string
Global 	CommentArray(100)	as string
Global 	FileNameArray(100)	as string
Global	BMPFiles(100)		as string
Global 	ArraySize			as long
'Global	ImageSize 		as long
'Global 	PercentComp		as long
Dim		DocName$			as string		'full name of current document
Dim		ExtName$			as string		'file extension of current document
Dim		NewExtName$		as string		'user-specified file extension
Dim		NewSavePath$		as string		'user-specified pathname
Dim		VersionName$		as string		'zero-filled (current) file number
Dim		SaveName$			as string		'final name of currently-saved file
Dim		dWidth			as double		'temporary width value
Dim		dHeight			as double		'temporary height value
Dim		nWidth			as long		'temporary width value (integer form)
Dim		nHeight			as long		'temporary height value (integer form)
Dim		nPixelWidth 		as long		'width, in pixels, of single file 'piece'
Dim		nPixelHeight 		as long		'height, in pixels, of single file 'piece'
Dim		nStartX			as long		'current X start point for crop
Dim		nStartY			as long		'current Y start point for crop
'Dim		nNumXTiles		as long		'number of X tiles, based on current input
'Dim		nNumYTiles		as long		'number of Y tiles, based on current inputs
Dim		nOrgX			as long		'current X center-of-origin for crop
Dim		nOrgY			as long		'current Y center-of-origin for crop
Dim		nCountX			as long		'current X-piece pointer
Dim		nCountY			as long		'current Y-piece pointer
'Dim		nNumFiles			as long		'total number of files generated, based on current inputs
Dim		nNumDigits		as long		'number of digits in the total file count
Dim		nFileNum			as long		'currently-selected file 'piece'
Dim		nFilterID			as long		'Filter Manager-type filter ID
Dim		nXGridFreq		as long		'Horizontal grid freequency of current document
Dim		nYGridFreq		as long		'Vertical grid freequency of current document

Const	RESOLUTION		as double		= 5905512.0

GLOBAL	gnDocWidth		as long		'width of current document - treat as constant
GLOBAL	gnDocHeight		as long		'height of current document -   "    "    "
GLOBAL	gnUnitsXIdx		as long		'numbox X Units index
GLOBAL	gnUnitsYIdx		as long		'numbox X Units index (not used)
GLOBAL	gbDivide			as boolean	'TRUE if dividing by number of pieces, FALSE if dividing by size
GLOBAL	gdWidth			as double		'current numbox Width value
GLOBAL	gdHeight			as double		'current numbox Height value
GLOBAL	gdResX			as double		'Horizontal resolution of current document
GLOBAL	gdResY			as double		'Vertical resolution of current document

GLOBAL	gdOldWidth		as double		'previous numbox Width value
GLOBAL	gdOldHeight		as double		'previous numbox Height value
GLOBAL	gnOldXSize		as long		'previous numbox X Size value
GLOBAL	gnOldYSize		as long		'previous numbox Y Size value
GLOBAL	gdOrigWidth		as double		'needed if we ever have a 'Reset' button
GLOBAL	gdOrigHeight		as double		'
GLOBAL	gnOrigUnitsXIdx	as long		'needed if we ever have a 'Reset' button
GLOBAL	gnOrigUnitsYIdx	as long		'

GLOBAL	gdMaxXRes			as double		'current numbox Width value - treat as constant
GLOBAL	gdMaxYRes			as double		'current numbox Height value  "    "    "

GLOBAL	FileName$			as string		'output file name (with no path specification)
GLOBAL	SavePath$			as string		'path for output file

GLOBAL gUnitsTxt$(8)
gUnitsTxt(1) = "inches"
gUnitsTxt(2) = "millimeters"
gUnitsTxt(3) = "picas, points"
gUnitsTxt(4) = "points"
gUnitsTxt(5) = "centimeters"
gUnitsTxt(6) = "pixels"
gUnitsTxt(7) = "ciceros, didots"
gUnitsTxt(8) = "didots"


REM	-----------------------------------------------------------------
REM	Code for the dynamic dialog box
REM

Note1 = "The files will be created in the width and height you specify, " \\
		+ "and numbered in sequential order, based on the specified filename. " \\
		+ "The default sizes are based on the current document's grid settings." \\
		+ CHR(13)+CHR(10) \\ 
		+ CHR(13)+CHR(10) + "Assumptions/Comments:" \\
		+ CHR(13)+CHR(10) \\
		+ CHR(13)+CHR(10) + "  1) The currently-opened file is in CPT format." \\
		+ CHR(13)+CHR(10) + "  2) Error checking for units 'picas, points' and 'ciceros, didots' is limited."

BEGIN DIALOG OBJECT Dimensions 268, 166, "Image Cutter", SUB SPINS
	GROUPBOX  4, 2, 162, 67, .Group1, "Size Options"	 'Controlid =1
	TEXT  12, 14, 95, 8, .WidthText, "Width of each piece:"	 'Controlid =2
	SPINCONTROL  123, 11, 35, 14, .gdWidth	 'Controlid =3
	TEXTBOX  124, 11, 35, 14, .WidthTextBox	 'Controlid =4
	TEXT  12, 29, 95, 8, .HeightText, "Height of each piece:"	 'Controlid =5
	SPINCONTROL  123, 26, 35, 14, .gdHeight	 'Controlid =6
	TEXTBOX  122, 26, 35, 14, .HeightTextBox	 'Controlid =7
	TEXT  12, 44, 70, 8, .UnitsText, "Units:"	 'Controlid =8
	DDLISTBOX  98, 41, 60, 85, .ListBox	 'Controlid =9
	OPTIONGROUP .OptionGroup
		OPTIONBUTTON  14, 57, 50, 10, .OptionButton1, "Specify size"	 'Controlid =11
		OPTIONBUTTON  72, 57, 95, 10, .OptionButton2, "Specify no. of pieces"	 'Controlid =12
	GROUPBOX  4, 71, 162, 74, .Group2, "File Naming Options"	 'Controlid =13
	TEXT  12, 82, 40, 8, .SaveNameText, "Save Name:"	 'Controlid =14
	TEXTBOX  56, 79, 100, 14, .FileName	 'Controlid =15
	TEXT  12, 96, 40, 8, .SaveInText, "Save In:"	 'Controlid =16
	TEXTBOX  56, 94, 100, 14, .SavePath	 'Controlid =17
	TEXT  12, 112, 120, 8, .FileCount, "File Count"	 'Controlid =18
	GROUPBOX  169, 2, 95, 143, .Group3, "Note:"	 'Controlid =19
	TEXT  176, 12, 81, 128, .NoteText, Note1	 'Controlid =20
	OKBUTTON  181, 149, 40, 14, .OK	 'Controlid =21
	CANCELBUTTON  224, 149, 40, 14, .Cancel	 'Controlid =22
	CHECKBOX  12, 125, 68, 13, .GenerateHTML, "Generate HTML"	 'Controlid =24
END DIALOG


'Dialog Event Handler subroutine
SUB SPINS(BYVAL ControlID%, BYVAL Event%)

	DIM dWidth		as double		'temporary width value
	DIM dHeight		as double		'temporary height value
	DIM dMaxXRes		as double		'temporary maximum width value
	DIM dMaxYRes		as double		'temporary maximum height value
	DIM nOldUnits		as long		'previously-selected Unit numbox index
	DIM nXPixels		as long		'current width in pixels
	DIM nYPixels		as long		'current height in pixels
	DIM nFiles		as long		'number of files, based on current inputs
	DIM bUpdate		as boolean	'update the height/width numboxes if TRUE
	DIM StrXValue		as string		'current X value in ASCII format (for c,d and p,p units)
	DIM StrYValue		as string		'current Y value in ASCII format (for c,d and p,p units)
	Dim FileCount$		as string		'status string indicating number of files, based on current inputs

	bUpdate = FALSE
	
	'Set initial Main spin control options
	IF (Event = 0) THEN
	
		gnOldXSize = 1		
		gnOldYSize = 1
		gdOldWidth = gdWidth
		gdOldHeight = gdHeight
		gbDivide = FALSE
		nFiles = 1
		ArraySize=1
		
		Dimensions.ListBox.SETARRAY gUnitsTxt$
		Dimensions.ListBox.SETSELECT gnUnitsXIdx
		Dimensions.ListBox.ENABLE TRUE
		Dimensions.FileName.SETTEXT FileName$
		Dimensions.SavePath.SETTEXT SavePath$
		Dimensions.FileCount.SETTEXT (GetFileCount(nFiles))
		Dimensions.GenerateHTML.SETTHREESTATE FALSE
		Dimensions.GenerateHTML.SETVALUE 0
		Dimensions.GenerateHTML.ENABLE TRUE
		
		Dimensions.gdWidth.SETMAXRANGE 10
		Dimensions.gdHeight.SETMAXRANGE 10
		
		bUpdate = TRUE
		GOTO UpdateNumbox
		
	ENDIF
	
	IF ((Event = 1) AND (ControlID = 3) AND (gbDivide = TRUE)) THEN	'horiz. pieces
	
		dWidth = Dimensions.gdWidth.GETVALUE()
		dHeight = Dimensions.gdHeight.GETVALUE()
		nFiles = dWidth*dHeight
		Dimensions.FileCount.SETTEXT (GetFileCount(nFiles))
		
	ENDIF
	
	IF ((Event = 1) AND (ControlID = 6) AND (gbDivide = TRUE)) THEN	'vert. pieces

		dWidth = Dimensions.gdWidth.GETVALUE()
		dHeight = Dimensions.gdHeight.GETVALUE()
		nFiles = dWidth*dHeight
		Dimensions.FileCount.SETTEXT (GetFileCount(nFiles))
		
	ENDIF
	
	IF ((Event = 1) AND (ControlID = 3) AND (gbDivide = FALSE)) THEN	'horiz. size
	
		dWidth = Dimensions.gdWidth.GETVALUE()
		dHeight = Dimensions.gdHeight.GETVALUE()
		nXPixels = ConvertToPixels( dWidth, gdResX, gnUnitsXIdx )
		nYPixels = ConvertToPixels( dHeight, gdResY, gnUnitsXIdx )
		if ((nXPixels <> 0) AND (nYPixels <> 0)) then
			nFiles = Round(gnDocWidth/nXPixels) * Round(gnDocHeight/nYPixels)
			Dimensions.FileCount.SETTEXT (GetFileCount(nFiles))
		endif		
		
	ENDIF
	
	IF ((Event = 1) AND (ControlID = 6) AND (gbDivide = FALSE)) THEN	'vert. size

		dWidth = Dimensions.gdWidth.GETVALUE()
		dHeight = Dimensions.gdHeight.GETVALUE()
		nXPixels = ConvertToPixels( dWidth, gdResX, gnUnitsXIdx )
		nYPixels = ConvertToPixels( dHeight, gdResY, gnUnitsXIdx )
		if ((nXPixels <> 0) AND (nYPixels <> 0)) then
			nFiles = Round(gnDocWidth/nXPixels) * Round(gnDocHeight/nYPixels)
			Dimensions.FileCount.SETTEXT (GetFileCount(nFiles))
		endif	
		
	ENDIF
	
	IF ((Event = 1) AND (ControlID = 4) AND (gbDivide = FALSE)) THEN	'horiz. size  ('picas,points' or 'ciceros,didots')
	
		StrXValue	= Dimensions.WidthTextBox.GETTEXT()
		dWidth = StringToPoints( StrXValue )
		if ((dWidth > 0.0) AND (dWidth <= gdMaxXRes)) then
			gdWidth = dWidth
		else 
			if (dWidth > gdMaxXRes) then
				gdWidth = gdMaxXRes
			endif
		endif

		nXPixels = ConvertToPixels( gdWidth, gdResX, gnUnitsXIdx )
		nYPixels = ConvertToPixels( gdHeight, gdResY, gnUnitsXIdx )
		if ((nXPixels <> 0) AND (nYPixels <> 0)) then
			nFiles = Round(gnDocWidth/nXPixels) * Round(gnDocHeight/nYPixels)
			Dimensions.FileCount.SETTEXT (GetFileCount(nFiles))
		endif		
		
	ENDIF
	
	IF ((Event = 1) AND (ControlID = 7) AND (gbDivide = FALSE)) THEN	'vert. size ('picas,points' or 'ciceros,didots')

		StrYValue = Dimensions.HeightTextBox.GETTEXT()
		dHeight = StringToPoints( StrYValue )
		if ((dHeight > 0.0) AND (dHeight <= gdMaxYRes)) then
			gdHeight = dHeight
		else 
			if (dHeight > gdMaxYRes) then
				gdHeight = gdMaxYRes
			endif
		endif

		nXPixels = ConvertToPixels( gdWidth, gdResX, gnUnitsXIdx )
		nYPixels = ConvertToPixels( gdHeight, gdResY, gnUnitsXIdx )
		if ((nXPixels <> 0) AND (nYPixels <> 0)) then
			nFiles = Round(gnDocWidth/nXPixels) * Round(gnDocHeight/nYPixels)
			Dimensions.FileCount.SETTEXT (GetFileCount(nFiles))
		endif
		
	ENDIF
	
	
	IF ((Event = 2) AND (ControlID = 9)) THEN	'units
	
		nOldUnits = gnUnitsXIdx
		gnUnitsXIdx = Dimensions.ListBox.GETSELECT()
		
		' Convert the numbox values
		if ((nOldUnits <> 3) AND (nOldUnits <> 7)) then
			dWidth = Dimensions.gdWidth.GETVALUE()
			dHeight = Dimensions.gdHeight.GETVALUE()
		else
			dWidth = gdWidth
			dHeight = gdHeight
		endif
		gdWidth = ConvertToUnit( dWidth, gdResX, gnUnitsXIdx, nOldUnits )
		gdHeight = ConvertToUnit( dHeight, gdResY, gnUnitsXIdx, nOldUnits )
		
		' Convert the max values
		gdMaxXRes = ConvertToUnit( gdMaxXRes, gdResX, gnUnitsXIdx, nOldUnits )
		gdMaxYRes = ConvertToUnit( gdMaxYRes, gdResY, gnUnitsXIdx, nOldUnits )
		
		' Recalculate the number of files
		nXPixels = ConvertToPixels( gdWidth, gdResX, gnUnitsXIdx )
		nYPixels = ConvertToPixels( gdHeight, gdResY, gnUnitsXIdx )
		if ((nXPixels <> 0) AND (nYPixels <> 0)) then
			nFiles = Round(gnDocWidth/nXPixels) * Round(gnDocHeight/nYPixels)
			Dimensions.FileCount.SETTEXT (GetFileCount(nFiles))
		endif
		
		bUpdate = TRUE
		GOTO UpdateNumbox
		
	ENDIF	
	
	IF ((Event = 2) AND (ControlID = 11)) THEN	'by size
	
		gbDivide = FALSE
		gnOldXSize = INT(Dimensions.gdWidth.GETVALUE())
		gnOldYSize = INT(Dimensions.gdHeight.GETVALUE())
		gdWidth = gdOldWidth
		gdHeight = gdOldHeight		
		nFiles = gdWidth*gdHeight
		
		Dimensions.ListBox.ENABLE TRUE
		Dimensions.HeightText.SETTEXT "Height of image:"
		Dimensions.WidthText.SETTEXT "Width of image:"
		Dimensions.gdWidth.SETVALUE gdWidth
		Dimensions.gdHeight.SETVALUE gdHeight
		Dimensions.FileCount.SETTEXT (GetFileCount(nFiles))
		
		GOTO UpdateNumbox
		
	ENDIF
	
	IF ((Event = 2) AND (ControlID = 24))THEN
		checked% = Dimensions.GenerateHTML.GETVALUE ( )
		
		
	END IF	
	
	
	IF ((Event = 2) AND (ControlID = 12)) THEN	'by number of pieces
	
		gbDivide = TRUE
		gdOldWidth = Dimensions.gdWidth.GETVALUE()
		gdOldHeight = Dimensions.gdHeight.GETVALUE()
		gdWidth = gnOldXSize
		gdHeight = gnOldYSize
		nFiles = gdWidth*gdHeight
		
		Dimensions.gdWidth.SETVALUE gdWidth
		Dimensions.gdHeight.SETVALUE gdHeight
		Dimensions.ListBox.ENABLE FALSE		
		Dimensions.HeightText.SETTEXT "Number of vertical pieces:"
		Dimensions.WidthText.SETTEXT "Number of horizontal pieces:"
		Dimensions.FileCount.SETTEXT (GetFileCount(nFiles))
		
		GOTO SetToInt
		
	ENDIF
	
	
	GOTO EndDim
	
UpdateNumbox:
	select case gnUnitsXIdx
		case 1	'inches
			GOTO SetToDec
		case 2	'millimeters
			GOTO SetToDec
		case 3	'picas, points
			GOTO SetToTxt
		case 4	'points
			GOTO SetToDec
		case 5	'centimeters2
			GOTO SetToDec
		case 6	'pixels
			GOTO SetToInt
		case 7	'ciceros, didots
			GOTO SetToTxt
		case 8	'didots	
			GOTO SetToDec
	end select

SetToDec:
	Dimensions.gdWidth.SETDOUBLEMODE TRUE
	Dimensions.gdWidth.SETPRECISION 2
	Dimensions.gdWidth.SETINCREMENT 0.01
	Dimensions.gdWidth.SETMINRANGE 0.01
	Dimensions.gdWidth.SETMAXRANGE gdMaxXRes
	
	Dimensions.gdHeight.SETDOUBLEMODE TRUE
	Dimensions.gdHeight.SETPRECISION 2
	Dimensions.gdHeight.SETINCREMENT 0.01
	Dimensions.gdHeight.SETMINRANGE 0.01
	Dimensions.gdHeight.SETMAXRANGE gdMaxYRes	
	
	Dimensions.gdWidth.SETVISIBLE TRUE
	Dimensions.gdHeight.SETVISIBLE TRUE
	Dimensions.WidthTextBox.SETVISIBLE FALSE
	Dimensions.HeightTextBox.SETVISIBLE FALSE
	
	GOTO EndDim

SetToInt:
	Dimensions.gdWidth.SETDOUBLEMODE FALSE
	Dimensions.gdWidth.SETINCREMENT 1
	Dimensions.gdWidth.SETMINRANGE 1
	Dimensions.gdWidth.SETMAXRANGE gnDocWidth
	
	Dimensions.gdHeight.SETDOUBLEMODE FALSE		
	Dimensions.gdHeight.SETINCREMENT 1
	Dimensions.gdHeight.SETMINRANGE 1
	Dimensions.gdHeight.SETMAXRANGE gnDocHeight
	
	Dimensions.gdWidth.SETVISIBLE TRUE
	Dimensions.gdHeight.SETVISIBLE TRUE
	Dimensions.WidthTextBox.SETVISIBLE FALSE
	Dimensions.HeightTextBox.SETVISIBLE FALSE
	
	GOTO EndDim
	
SetToTxt:
	StrXValue = PointsToString( gdWidth, 0 )
	StrYValue = PointsToString( gdHeight, 0 )
	
	Dimensions.gdWidth.SETVISIBLE FALSE
	Dimensions.gdHeight.SETVISIBLE FALSE
	Dimensions.WidthTextBox.SETVISIBLE TRUE
	Dimensions.HeightTextBox.SETVISIBLE TRUE
	
	Dimensions.WidthTextBox.SETTEXT StrXValue
	Dimensions.HeightTextBox.SETTEXT StrYValue
	
	GOTO EndDim
		
EndDim:	

	if (bUpdate) then
		Dimensions.gdWidth.SETVALUE gdWidth
		Dimensions.gdHeight.SETVALUE gdHeight
	endif
	
END SUB

REM	-----------------------------------------------------------------
REM	Code for the static dialog box
REM

label1 = "This script divides the current document into sequentially numbered files."
BEGIN DIALOG IntroDialog 196, 74, "Image Cutter"
	GROUPBOX  2, 1, 191, 53
	TEXT  8, 12, 178, 37, label1
	OKBUTTON  109, 57, 40, 14
	CANCELBUTTON  153, 57, 40, 14
END DIALOG

REM	-----------------------------------------------------------------

REM	-----------------------------------------------------------------
REM	Code for the dynamic Saving dialog box
REM
if ret1=2 then stop
BEGIN DIALOG OBJECT Saving 281, 18, "Generating HTML Page", SUB SubSaving
	PROGRESS 2, 5, 244, 10, .Progress1
	TEXT  247, 4, 32, 10, .Text4, "Text4"
END DIALOG


Sub SubSaving (BYVAL ControlID%, BYVAL Event%)
	if Event = 0 Then
		Saving.Progress1.SETMINRANGE 0
		Saving.Progress1.SETMAXRANGE 100
		call SubSubDimension
	End if	
'	if event=2 and ControlId= 1 then
	
End Sub	
REM	-----------------------------------------------------------------


REM	-----------------------------------------------------------------
REM	Code for the dynamic Saving dialog box
REM


BEGIN DIALOG OBJECT HTML 243, 105, "Generate HTML", SUB SubHTML
	TEXT  3, 4, 38, 14, .Page, "Page Title:"	 'ControlId = 1
	TEXTBOX  43, 2, 134, 13, .PageTitle	 'ControlId = 2
	TEXT  3, 20, 38, 14, .FileName, "File Name:"	 'ControlId = 3
	DDCOMBOBOX  43, 17, 134, 51, .FilePieces	 'ControlId = 4
	TEXT  3, 34, 38, 14, .URL, "URL:"	 'ControlId = 5
	TEXTBOX  43, 33, 134, 13, .URLText	 'ControlId = 6			
	TEXT  3, 48, 38, 14, .Comment, "Comment:"	 'ControlId = 7
	TEXTBOX  43, 47, 134, 13, .CommentText	 'ControlId = 8
	IMAGE  182, 3, 58, 58, .Image1	 'ControlId = 9
	'PUSHBUTTON  155, 88, 40, 14, .PUSHBUTTON1, "OK"
	OKBUTTON  155, 88, 40, 14, .OK1	 'ControlId = 10
	CANCELBUTTON  200, 88, 40, 14, .Cancel1	 'ControlId = 11 
	PROGRESS 3, 73, 203, 11, .Progress1	 'Control 12
	TEXT  211, 72, 32, 10, .Text4, "Text4"	 'Control 13
END DIALOG

 
Sub SubHTML(BYVAL ControlID%, BYVAL Event%)
'Global SavePath as string
Global BaseName as string
Global WantHTML as integer
Global nPixelWidth as long
Global nPixelHeight	as long
'Global nNumYTiles as integer
'Global nNumXTiles as integer
	if Event= 0 THEN
		if ret1 = 2 then stop
		HTML.Progress1.SETMINRANGE 0
		HTML.Progress1.SETMAXRANGE 100
		HTML.Image1.SETVISIBLE False
		HTML.URLText.Enable False
		HTML.CommentText.Enable False
		HTML.PageTitle.Enable False
		HTML.filePieces.Enable False
		HTML.OK1.Enable False
		HTML.Cancel1.Enable False
		call SubDimension 
		HTML.Image1.SETSTYLE 16  'Auto-resize
		HTML.Image1.SETIMAGE bmpfiles$(1) 
		HTML.CommentText.SetText ""
		HTML.URLText.SetText ""
		HTML.CommentText.SetText ""
		HTML.PageTitle.SetText ""
		Savingto$=Dimensions.FileName.gettext()
		NoExt = RemoveExt(Savingto$)
		For j%=1 to ArraySize STEP 1
			FileNameArray(j%) = NoExt & STR (j)&".BMP"
		NEXT j%
		HTML.filePieces.SETARRAY FileNameArray	' At this point the FileName array is charged properly.
		HTML.filePieces.SETSELECT FileNameArray(1) 
		HTML.filePieces.SETARRAY FileNameArray	' At this point the FileName array is charged properly.
		HTML.filePieces.SETSELECT FileNameArray(1)
		HTML.Image1.SETVISIBLE true
		HTML.URLText.Enable true
		HTML.CommentText.Enable true
		HTML.PageTitle.Enable true
		HTML.filePieces.Enable true 
		HTML.OK1.Enable True
		HTML.Cancel1.Enable True	
		HTML.Progress1.SETMINRANGE 0
		HTML.Progress1.SETMAXRANGE 100			
		HTML.text4.SetText ""
		HTML.text4.Enable False
		HTML.Progress1.SetValue 0
		HTML.Progress1.SetVisible false				
	End if
	IF Event = 1 THEN
		Select Case ControlID
			Case 6
				FileName = HTML.filePieces.getselect ()
				index=1
				While FileName <> FilenameArray(index)
					index=index+1		
				WEND	
				if URLArray(index) = "" Then
					URLWanted$= HTML.URLtext.GETTEXT ()
					URLArray(index)= URLWanted$
				Else
					URLWanted$= HTML.URLtext.GETTEXT ()
					URLArray(index)= URLWanted$	
				END IF
				
				HTML.CommentText.SETTEXT CommentArray(index)
							
			Case 8
				FileName = HTML.filePieces.getselect ()
				index=1
				While FileName <> FilenameArray(index)
					index=index+1		
				WEND	
				if CommentArray(index) = "" Then
					CommentWanted$= HTML.CommentText.GETTEXT ()
					CommentArray(index)= CommentWanted$
				Else
					CommentWanted$= HTML.CommentText.GETTEXT ()
					CommentArray(index)= CommentWanted$	
				END IF
				
				HTML.URLtext.SETTEXT URLArray(index)		
		End Select
	END IF
	
	if Event = 2 THEN
		Select Case ControlID 
			Case 4
				FileName = HTML.filePieces.getselect ()
				index=1
				While FileName <> FilenameArray(index)
					index=index+1		
				WEND	
				HTML.CommentText.SETTEXT CommentArray(index)
				HTML.URLtext.SETTEXT URLArray(index)
				HTML.Image1.SETIMAGE BMPFiles (index)
			'Case 10 
			'	HTML.Image1.SETVISIBLE False
			'	HTML.URLText.Enable False
			'	HTML.CommentText.Enable False
			'	HTML.PageTitle.Enable False
			'	HTML.filePieces.Enable False
			'	HTML.PUSHBUTTON1.Enable False
			'	HTML.Cancel1.Enable False
			'	HTML.text4.Enable True
			'	HTML.Progress1.SetVisible True
						
		End Select

	End if
End Sub	
REM	---------------------------------------------Sub SubDimension()--------------------



REM
REM	----------------
REM	 MAIN PROCEDURE
REM	----------------
REM

Rem	Variable initialization

nOrgX = 0
nOrgY = 0
nCountX = 0
nCountY = 0
nNumFiles = 0

ret = DIALOG(IntroDialog)	'displays the introduction
IF ret = 2 THEN STOP		' If Cancel is selected, stop the script
	WITHOBJECT "CorelPhotoPaint.Automation.9"
	gnDocWidth = .GetDocumentWidth()
	gnDocHeight = .GetDocumentHeight()
	DocName$ = .GetDocumentName()
	
	' File Must be saved on Disk Already.
	retval = FILEATTR(DocName$)
	if retval = 0 then 
		messagebox "File must be saved on disk. Script can not continue", "Warning", 48
		stop
	End if 
	
	' Extract the path and filename
	SavePath$ = GetPath( DocName$ )
	if (SavePath$ <> "") then
		FileName$ = RIGHT( DocName$, LEN(DocName$)-LEN(SavePath$)-1 )
	else
		FileName$ = DocName$
	endif
	ExtName$ = GetExtension( FileName$ )

	' Get the default units, and initialize numboxes accordingly
	gdResX = .GetDocumentHRes()				'PAINT9
	gdResY = .GetDocumentVRes()				'PAINT9
	gnUnitsXIdx = .GetDocumentXRulerUnits()		'PAINT9
	gnUnitsYIdx = .GetDocumentYRulerUnits()		'PAINT9
	nXGridFreq = .GetDocumentXGridFrequency()	'PAINT9
	nYGridFreq = .GetDocumentYGridFrequency()	'PAINT9
REM	gdResX = RESOLUTION						'PAINT8
REM	gdResY = RESOLUTION						'PAINT8
REM	gnUnitsXIdx = 6						'PAINT8
REM	gnUnitsYIdx = 6						'PAINT8
REM	gnXGridFreq = 100						'PAINT8
REM	gnYGridFreq = 100						'PAINT8

	dWidth = gnDocWidth
	dHeight = gnDocHeight
	gdMaxXRes = ConvertToUnit( dWidth, gdResX, gnUnitsXIdx, 6 )	
	gdMaxYRes = ConvertToUnit( dHeight, gdResY, gnUnitsYIdx, 6 )
	
	dWidth = nXGridFreq
	dHeight = nYGridFreq
	gdWidth = ConvertBaseToUnit( dWidth, gdResX, gnUnitsXIdx )	
	gdHeight = ConvertBaseToUnit( dHeight, gdResY, gnUnitsYIdx )
	
	gdOrigWidth = gdWidth		'needed if we ever have a 'Reset' button
	gdOrigHeight = gdHeight		'

		
	' Prompt the user for the required data
	ret = DIALOG(Dimensions) 'displays Dimensions Dialog box
	if ret= 2 then Stop
	ret1 = dialog(HTML)
	if ret1 = 1 Then
	ret2 = dialog(Saving)
	End if
end withobject

Sub SubDimension()	
	WITHOBJECT "CorelPhotoPaint.Automation.9"
	Dim dWidth as double
	Dim dHeight as double
	Dim nPixelWidth as Long 
	Dim nPixelHeight as Long
	Dim nFileNum as long
	Dim nNumDigits as long
	'Dim nNumXTiles as long
'	Dim nNumYTiles as long
	
	nXGridFreq = .GetDocumentXGridFrequency()	'PAINT9
	nYGridFreq = .GetDocumentYGridFrequency()	'PAINT9
	dWidth = nXGridFreq
	dHeight = nYGridFreq
	
	' Calculate the width/height of the tiles	
	DocName$ = .GetDocumentName()
	gnUnitsXIdx = Dimensions.ListBox.GETSELECT()
	if (gbDivide = FALSE) then
		dWidth = Dimensions.gdWidth.GETVALUE()
		dHeight = Dimensions.gdHeight.GETVALUE()
		IF dWidth = 0.0 THEN STOP
		IF dHeight = 0.0 THEN STOP
	
		nPixelWidth = ConvertToPixels( dWidth, gdResX, gnUnitsXIdx )
		nPixelHeight = ConvertToPixels( dHeight, gdResY, gnUnitsXIdx )
	else
		nWidth = INT(Dimensions.gdWidth.GETVALUE())
		nHeight = INT(Dimensions.gdHeight.GETVALUE())
		IF nWidth = 0 THEN STOP
		IF nHeight = 0 THEN STOP
	
		nPixelWidth = gnDocWidth/nWidth
		nPixelHeight = gnDocHeight/nHeight
	endif
	
	' Convert the Units
	if (gbDivide = FALSE) then
		nNumXTiles = Round(gnDocWidth/nPixelWidth)
		nNumYTiles = Round(gnDocHeight/nPixelHeight)
	else
		nNumXTiles = nWidth
		nNumYTiles = nHeight
	endif
	nNumFiles = nNumXTiles * nNumYTiles
	nNumDigits = LEN( CSTR( nNumFiles ) )
	nStartX = 0
	nStartY = 0
	nFileNum = 1
	' Generate the output filenames	
	FileName$ = Dimensions.FileName.GETTEXT()
	NewSavePath$ = Dimensions.SavePath.GETTEXT()
	SavePath$ = RemoveBackslash( NewSavePath$ )
	
	NewExtName$ = GetExtension( FileName$ )
	if (NewExtName$ <> "") then
		BaseName$ = LEFT( FileName$, LEN(FileName$)-LEN(NewExtName$)-1 )
		ExtName$ = NewExtName$
	else
		BaseName$ = FileName$
	endif
	nFilterID = GetFilterID( ExtName$ )
	
		
	
	WantHTML% = Dimensions.GenerateHTML.GETVALUE ( )
	
	
	' Go and divvy up the file
	' ------------------------	
	i=1
	completedPart%=0
	TotalSize$ = Filesize (DocName$)
	for nCountY = 0 to (nNumYTiles-1)

		for nCountX = 0 to (nNumXTiles-1)
			
			' Crop the file at the current location
			nOrgX = nStartX + nPixelWidth/2
			nOrgY = nStartY + nPixelHeight/2
			.ImageDeskewCrop 0, nPixelWidth, nPixelHeight, nOrgX, nOrgY
	
			' Save the file (= 'Save As' if filename specified)
			VersionName$ = GetVersionString( nFileNum, nNumDigits )
			SaveName$ = SavePath$ + "\" + BaseName$ + VersionName$ + "." + "BMP"
	  		if WantHTML = 1 then
				BMPFiles(i) = SaveName$
				i=i + 1
			END IF
			.FileSave SaveName$,  769, TRUE
			'ImageSize = Filesize (SaveName$)
			'IncPercent%= Int((ImageSize /TotalSize)* 100)
			'PercentComp =  PercentComp + IncPercent%
			'completedPart%=completedPart%+ImageSize
			IncPercent%= int(100/nNumFiles)
			HTML.Progress1.setIncrement IncPercent%
			HTML.Progress1.step 
			ImageComp%=ImageComp%+IncPercent%
			'ImageComp% = PercentComp	
			HTML.Text4.settext (STR (ImageComp%)+ " %")
			.FileClose

			' Restore the file
			.FileOpen DocName$, 0, 0, 0, 0, 0, 1, 1

			' Update the current location
			nStartX = nStartX + nPixelWidth
			nFileNum = nFileNum + 1
					
		next nCountX
		
		nStartX = 0
		nStartY = nStartY + nPixelHeight
	
	next nCountY
if percentComp <> 100 then
	HTML.Progress1.setvalue 100
	HTML.Text4.settext ("100 %")
End if
if wantHTML = 0 then stop
End Withobject
End Sub
Sub SubSubDimension()
WITHOBJECT "CorelPhotoPaint.Automation.9"
Dim dWidth as double
	Dim dHeight as double
	
	Dim nFileNum as long
	Dim nNumDigits as long
nNumFiles = nNumXTiles * nNumYTiles
nNumDigits = LEN( CSTR( nNumFiles ) )
nFileNum = 1
nStartX = 0
nStartY = 0

if (gbDivide = FALSE) then
	dWidth = Dimensions.gdWidth.GETVALUE()
	dHeight = Dimensions.gdHeight.GETVALUE()
	IF dWidth = 0.0 THEN STOP
	IF dHeight = 0.0 THEN STOP

	nPixelWidth = ConvertToPixels( dWidth, gdResX, gnUnitsXIdx )
	nPixelHeight = ConvertToPixels( dHeight, gdResY, gnUnitsXIdx )
else
	nWidth = INT(Dimensions.gdWidth.GETVALUE())
	nHeight = INT(Dimensions.gdHeight.GETVALUE())
	IF nWidth = 0 THEN STOP
	IF nHeight = 0 THEN STOP

	nPixelWidth = gnDocWidth/nWidth
	nPixelHeight = gnDocHeight/nHeight
endif

if WantHTML = 1 Then
	DocName$ = .GetDocumentName()
	TotalSize$ = Filesize (DocName$)
	'
	'If ret1 = 2 Then stop
	'ret2 = Dialog (saving)
	call CreatePage SavePath$ ,BaseName
	PRINT #1, "<table border=0 CELLSPACING=0 CELLPADDING=0 Width= " + Str (nPixelWidth) + " >"
	ArrayIndex=1
	for nCountY = 0 to (nNumYTiles-1)
		PRINT #1, "<tr>"
		for nCountX = 0 to (nNumXTiles-1)
			nOrgX = nStartX + nPixelWidth/2
			nOrgY = nStartY + nPixelHeight/2
			.ImageDeskewCrop 0, nPixelWidth, nPixelHeight, nOrgX, nOrgY
			VersionName$ = GetVersionString( nFileNum, nNumDigits )
			JPEGFILESave$ = SavePath$ + "\" + BaseName$ + VersionName$ + "." + "JPG"
			.FileSave JPEGFILESave$ ,774, TRUE
			ImageSize = Filesize (JPEGFILESave$)
			IncPercent%= Int((ImageSize /TotalSize)* 100)
			PercentComp =  PercentComp + IncPercent%
			Saving.Progress1.setIncrement IncPercent%
			Saving.Progress1.step 
			ImageComp% = PercentComp	
			Saving.Text4.settext (STR (ImageComp%)+ " %")
			.FileClose
			
			JPEGFILE$=BaseName$ + VersionName$ + "." + "JPG"
			.FileOpen DocName$, 0, 0, 0, 0, 0, 1, 1
			PRINT #1, "<td>"
			URL= URLArray(ArrayIndex)
			if URL<>"" Then
				DesURL = LEFT(URL, 7)
				DesURLLC = LCASE(DesURL)
				if (DesURLLC = "http://") then
					PRINT #1,  "<A HREF=" + chr(34)+ URL + chr(34)+ "</A>"
				Else
					PRINT #1,  "<A HREF=" + chr(34)+ "http://" + URL + chr(34)+ "</A>"
				End if
			End if
			Comment$ = CommentArray(ArrayIndex)
			PRINT #1, "<IMG SRC=" + chr(34)+ JPEGFILE$+ chr(34)  	
			if Comment$<>"" Then
				PRINT #1, " alt="+  chr(34)+ Comment$ + chr(34)+ " lowsrc=" + chr(34)+ JPEGFILE$ + chr(34)
			Else
				PRINT #1,  " lowsrc=" + chr(34)+ JPEGFILE$ + chr(34)
			End if
			PRINT #1, " border=0 " + "width="+ chr(34) + Str (nPixelWidth) + chr(34)+  " height=" + chr(34) + Str (nPixelHeight) + chr(34) + ">" + "</td>"
			'PRINT #1, "</td>"
			ArrayIndex=ArrayIndex+1
			nFileNum = nFileNum + 1
			nStartX = nStartX + nPixelWidth
		next nCountX
		nStartX = 0
		nStartY = nStartY + nPixelHeight
		PRINT #1, "</tr>"
	next nCountY		
End if
PRINT #1, "</table><I>This page was generated by <A HREF=" + chr(34)+ "http://www.Photopaint.com" + chr(34)+ ">Corel Photo Paint 9</A></I>"
PRINT #1, "</body>"
PRINT #1, "</html>"
CLOSE #1
if percentComp <> 100 then
	Saving.Progress1.setvalue 100
	Saving.Text4.settext ("100 %")
End if	
stop		
END WITHOBJECT	
End Sub
'End of program


REM Function definitions.
Sub CreatePage (DesDirectory as String, BaseName as string)
	
	Dim URL as String
	Dim DesURL as string
	Dim DesURLLC as string
	Dim NameString as string
	L=1
	npages = 1
	NameString = HTML.PageTitle.GETTEXT ()
	pageName$=DesDirectory & "\" & BaseName & ".html" 
	OPEN  pageName$ FOR OUTPUT AS 1
		PRINT #1, "<html>"
		PRINT #1, "<head>"
		if NameString <>"" then
			PRINT #1, "<title>" + NameString + "</title>"
		End if		
		PRINT #1, "<meta name=" + chr(34) + "GENERATOR" + chr(34) + "content=" + chr(34) + "Image Cutter - Corel Corporation." + chr(34) + ">"
		PRINT #1, "<meta name=" + chr(34) + "keywords"  + chr(34) + "content=" + chr(34) + "Image Cutter WebPageWizard"+ chr(34) + ">"
		PRINT #1, "<meta http-equiv=" + chr(34) + "Content-Type" + chr(34) + "content=" + chr(34) + "text/html; charset=iso-8859-1" + chr(34) +" >"
		URL = HTML.URLText.GETTEXT()
		if NameString <>"" then
			PRINT #1, "<h1 align="+ chr(34) + " center" + chr(34) + ">" +NameString+"</h1>"	
		End if
		PRINT #1, "</head>"
		PRINT #1, "<body>"
	End Sub

REM *SDOC*************************************************************
REM 
REM 	Name:	ConvertToPixels
REM 
REM 	Action:	Convert the user's input into a pixel value
REM 
REM 	Params: 	dValue	- the input value
REM			dRes		- the document's X or Y display resolution
REM			nUnits	- the display units
REM 
REM 	Returns: 	LONG		- pixels
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/
Function ConvertToPixels(dValue as double, dRes as double, nUnits as long) as long

	DIM   dBaseValue		as double

	Const INCHTOBASE		as double		= 254000.0			
	Const MMTOBASE      	as double		= 10000.0
	Const PICATOBASE		as double		= (INCHTOBASE / 6.0)
	Const FRACTOBASE		as double		= (INCHTOBASE / 72.0)
	Const CICEROTOBASE		as double		= 45118.7
	Const DIDOTTOBASE		as double		= 3759.2
	Const UNITS_EPSILON		as double		= 0.000001
	Const DEFAULTRESOLUTION	as double		= 11811024.0
	
	if (nUnits = 6) then
		ConvertToPixels = INT(dValue)
		goto GetOut
	endif
	
	select case nUnits
		case 1		'inches
			dBaseValue = dValue * INCHTOBASE
		case 2		'millimeters
			dBaseValue = dValue * MMTOBASE
		case 3		'picas, points
			dBaseValue = dValue * PICATOBASE
		case 4		'points
			dBaseValue = dValue * FRACTOBASE
		case 5		'centimeters
			dBaseValue = dValue * MMTOBASE * 10
		case 6		'pixels
		case 7		'ciceros, didots
			dBaseValue = dValue * CICEROTOBASE
		case 8		'didots
			dBaseValue = dValue * DIDOTTOBASE
	end select
	
	if (dRes <= UNITS_EPSILON) then
		dRes = DEFAULTRESOLUTION
	endif
	
	if (dBaseValue <> 0.0) then
		ConvertToPixels = INT( ( (dRes / (MMTOBASE * 1000000.0)) * dBaseValue ) )
	else
		ConvertToPixels = 100
	endif

GetOut:

End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	ConvertToUnit
REM 
REM 	Action:	Convert the user's pixel input into the specified unit
REM 
REM 	Params: 	nValue		- the input value
REM			dRes			- the document's X or Y display resolution
REM			nNewUnits		- the new display units
REM			nPrevUnits	- the previous display units
REM 
REM 	Returns: 	LONG			- pixels
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/
Function ConvertToUnit(dValue as double, dRes as double, nNewUnits as long, nOldUnits as long) as long

	DIM	dBaseValue		as double
	DIM	dResult			as double 

	Const MMTOBASE      	as double		= 10000.0
	Const UNITS_EPSILON		as double		= 0.000001
	Const DEFAULTRESOLUTION	as double		= 11811024.0

	nValue = ConvertToPixels( dValue, dRes, nOldUnits )
	
	if (nNewUnits = 6) then
		dResult = nValue
		goto GetOut
	endif
	
	if (dRes <= UNITS_EPSILON) then
		dRes = DEFAULTRESOLUTION
	endif
	
		
	if (nValue <> 0) then
		dBaseValue = nValue / (dRes / (MMTOBASE * 1000000.0)) 
	else
		dBaseValue = 100
	endif
	
	dResult = ConvertBaseToUnit( dBaseValue, dRes, nNewUnits )
	
GetOut:

	ConvertToUnit = dResult	


End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	ConvertBaseToUnit
REM 
REM 	Action:	Convert the user's pixel input into the specified unit
REM 
REM 	Params: 	nValue		- the input value
REM			dRes			- the document's X or Y display resolution
REM			nNewUnits		- the new display units
REM			nPrevUnits	- the previous display units
REM 
REM 	Returns: 	LONG			- pixels
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/
Function ConvertBaseToUnit(dBaseValue as double, dRes as double, nUnits as long) as double

	Const INCHTOBASE		as double		= 254000.0			
	Const MMTOBASE      	as double		= 10000.0
	Const PICATOBASE		as double		= (INCHTOBASE / 6.0)
	Const FRACTOBASE		as double		= (INCHTOBASE / 72.0)
	Const CICEROTOBASE		as double		= 45118.7
	Const DIDOTTOBASE		as double		= 3759.2
	Const DEFAULTRESOLUTION	as double		= 11811024.0

	select case nUnits
		case 1		'inches
			ConvertBaseToUnit = dBaseValue / INCHTOBASE
		case 2		'millimeters
			ConvertBaseToUnit = dBaseValue / MMTOBASE
		case 3		'picas, points
			ConvertBaseToUnit = dBaseValue / PICATOBASE
		case 4		'points
			ConvertBaseToUnit = dBaseValue / FRACTOBASE
		case 5		'centimeters
			ConvertBaseToUnit = dBaseValue / MMTOBASE / 10
		case 6		'pixels
			if (dblResolution <= UNITS_EPSILON) then
				dblResolution = DEFAULTRESOLUTION
			endif
			ConvertBaseToUnit = (dRes / (MMTOBASE * 1000000.0)) * dBaseValue 
		case 7		'ciceros, didots
			ConvertBaseToUnit = dBaseValue / CICEROTOBASE
		case 8		'didots
			ConvertBaseToUnit = dBaseValue / DIDOTTOBASE
	end select

End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	PointsToString
REM 
REM 	Action:	Convert the value to string format
REM 
REM 	Params:	dValue		- Value to be converted
REM			unNumDecPlaces	- number of decimal places required
REM 
REM 	Returns:	STRING		- Output text
REM 
REM 	Comments:	'unNumDecPlaces' is actually assumed to be 0 for most
REM			of the string formatting
REM 
REM *************************************************************EDOC*/
Function PointsToString(dValue as double, unNumDecPlaces as long) as string

	Dim		dAbs			as double
	Dim		epsilon		as double
	Dim		dPicas		as double
	Dim		dPoints		as double
	Dim		TmpRound		as long
	Dim		bNegSign		as boolean
	Dim		szFormat		as string
	Const	kdPointMax	as double		= 12.0

	' determine # of points in fractional part of picas.
	dAbs = abs(dValue)
	dPicas = Floor(dAbs)
	dPoints = (dAbs - dPicas) * kdPointMax

	' Do floating point equivalence test
	epsilon = 1/Power(10, unNumDecPlaces)
		
	' Add half the precision for rounding of Points
	dPoints = ( Floor(dPoints / epsilon + 0.5) ) * epsilon
	
	if (dPoints >= 12.0) then
		dPoints = dPoints - 12.0
		dPicas = dPicas + 1
	endif

	' Check if a negative sign is going to be necessary
	bNegSign = (dValue < 0.0 AND NOT(( dPicas = 0.0 ) AND ( dPoints = 0.0 )))

	' Build the format string
	PointsToString = ""
  
	' Output the string
	if (dValue < 0.0 AND NOT(( dPicas = 0.0 ) AND ( dPoints = 0.0 ))) then
		PointsToString = "-" + CSTR(INT(FIX(dPicas))) + "," + CSTR(INT(FIX(dPoints)))
	else
		PointsToString = CSTR(INT(FIX(dPicas))) + "," + CSTR(INT(FIX(dPoints)))
	endif

End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	StringToPoints
REM 
REM 	Action:	Convert an input string (assume in 'picas, points'
REM			or 'ciceros, didots' format)
REM 
REM 	Params:	Value	- Input string
REM 
REM 	Returns:	DOUBLE	- Output value
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/
Function StringToPoints(Value as string) as double

	DIM		dVal1		as double
	DIM		dVal2		as double
	DIM		nLen			as long
	DIM		nCnt			as long
	DIM		nComma		as long
	DIM		nPos			as long
	DIM		TmpStr		as string
	Const	kdPointMax	as double		= 12.0
	
	dVal1 = 0.0
	dVal2 = 0.0
	
	' Basic error-checking
	nPos = INSTR(Value, ",")
	if (nPos > 1) then
	
		nLen = LEN(Value)
		nCnt = 0
		
		' Get the first number
		TmpStr = LEFT(Value, nPos-1)
		dVal1 = VAL(TmpStr)
				
		' Get the second number
		if (nLen > nPos) then
			TmpStr = RIGHT(Value, nLen - nPos)
			dVal2 = VAL(TmpStr)
		endif
		
	endif
	
	StringToPoints = dVal1 + dVal2 / kdPointMax
	
End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	GetPath
REM 
REM 	Action:	Returns the folder, given a full pathname
REM 
REM 	Params:	FullPath$	- full pathname
REM 
REM 	Returns:	STRING	- folder in which the specified file resides
REM 
REM 	Comments:	removes the trailing backslash '\'
REM 
REM *************************************************************EDOC*/
Function GetPath(FullPath as string) as string

	DIM bFound	as boolean
	DIM TmpPath	as string
	DIM nLen		as long
	DIM nPos		as long
	
	bFound = TRUE
	nLen = 0
	TmpPath = FullPath
	
	DO until (bFound = FALSE)
	
		nPos = INSTR(TmpPath, "\")
		if (nPos <> 0) then
			TmpPath = RIGHT(TmpPath, LEN(TmpPath) - nPos)
		else
			bFound = FALSE
		endif
		
	LOOP
	
	GetPath = LEFT(FullPath, LEN(FullPath)-LEN(TmpPath)-1)

End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	GetExtension
REM 
REM 	Action:	Returns the file extension, given a full pathname
REM 
REM 	Params:	FullPath$	- full pathname
REM 
REM 	Returns:	STRING	- file extension
REM 
REM 	Comments:	does not include the dot '.'
REM 
REM *************************************************************EDOC*/
Function GetExtension(FullPath as string) as string

	DIM bFound	as boolean
	DIM TmpPath	as string
	DIM TmpExt	as string
	DIM nLen		as long
	DIM nPos		as long
	
	bFound = TRUE
	nLen = 0
	TmpExt = ""
	TmpPath = FullPath
	
	DO until (bFound = FALSE)
	
		nPos = INSTR(TmpPath, ".")
		if (nPos <> 0) then
			TmpPath = RIGHT(TmpPath, LEN(TmpPath) - nPos)
			TmpExt = TmpPath
		else
			bFound = FALSE
		endif
		
	LOOP
	
	GetExtension = TmpExt

End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	GetVersionString
REM 
REM 	Action:	Returns the ASCII version of the current file number,
REM			padded by the number of zeros corresponding to the
REM			total number of files
REM 
REM 	Params:	FileNum	- current file number
REM			NumDigits	- number of digits required
REM 
REM 	Returns:	STRING	- current file number
REM 
REM 	Comments:	Example:  GetVersionString( 72, 4 ) returns "0072"
REM 
REM *************************************************************EDOC*/
Function GetVersionString(FileNum as long, NumDigits as long) as string

	DIM Version	as string
	DIM Length	as long
	Version = ""
	
	' error checking
	if ((NumDigits > 0) AND (NumDigits < 10)) then
	
		Version = CSTR( FileNum )
		Length = LEN( CSTR( FileNum ) )
		
		if( Length < NumDigits ) then
		
			DO until (Length = NumDigits)
			
				Version = "0" + Version
				Length = Length + 1
				
			LOOP
		
		endif
			
	endif
	
	GetVersionString = Version

End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	RemoveBackslash
REM 
REM 	Action:	Removes the trailing backslash (if any)
REM 
REM 	Params:	NewSavePath$	- pathname
REM 
REM 	Returns:	STRING		- pathname without the backslash "\"
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Function RemoveBackslash(NewSavePath as string) as string

	DIM char		as string
	DIM TmpPath	as string
	DIM bFound	as boolean

	bFound = TRUE	
	TmpPath = NewSavePath
	
	DO until (bFound = FALSE)
	
		char = RIGHT(NewSavePath, 1)
		if ((char = "\") OR (char = "\")) then
			TmpPath = LEFT(NewSavePath, LEN(NewSavePath)-1)
			NewSavePath = TmpPath
		else
			bFound = FALSE
		endif
		
	LOOP
	
	RemoveBackslash = NewSavePath

End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	GetFilterID
REM 
REM 	Action:	Returns the Filter ID, given the file extension
REM 
REM 	Params:	Extension$	- file extension
REM 
REM 	Returns:	LONG			- filter ID
REM 
REM 	Comments:	a REALLY crude Filter Manager
REM 
REM *************************************************************EDOC*/
Function GetFilterID(Extension as string) as long

	DIM EXT	as string
	EXT = UCASE( Extension )
	
	select case EXT

		case "CPT"
			GetFilterID = 1799	'DEC("710")	Huh?

		case "JPG"
			GetFilterID = DEC("306")
						
		case "BMP"
			GetFilterID = DEC("301")
						
		case "GIF"
			GetFilterID = DEC("305")
						
		case "TIF"
			GetFilterID = DEC("304")
						
		case "PSD"
			GetFilterID = DEC("314")
						
		case "WI"
			GetFilterID = DEC("319")
						
		case "WVL"
			GetFilterID = DEC("319")
						
		case else
			GetFilterID = DEC("710")
			
	end select
	
End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	UnitsToIndex
REM 
REM 	Action:	Converts the units to dropdown list index
REM 
REM 	Params:	Units	- Units to be converted
REM 
REM 	Returns:	LONG		- Index
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/
Function UnitsToIndex(Units as long) as long

	select case Units
	
		case 1	'WUN_INCH
			UnitsToIndex = 1

		case 2	'WUN_MM
			UnitsToIndex = 2
		 			
		case 3	'WUN_PICA_PT
			UnitsToIndex = 3

		case 4	'WUN_FRAC_PT
			UnitsToIndex = 4
		 			
		case 5	'WUN_CM
			UnitsToIndex = 5

		case 6	'WUN_PIXEL
			UnitsToIndex = 6
		 			
		case 12	'WUN_CICEROS_DIDOTS
			UnitsToIndex = 7

		case 13	'WUN_DIDOTS
			UnitsToIndex = 8
			
		case else
			UnitsToIndex = 6	'just use good 'ol pixel units

	end select
		
End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	GetFileCount
REM 
REM 	Action:	Formats the file count string, according to the
REM			specified file count
REM 
REM 	Params:	nFiles	- Number of files
REM 
REM 	Returns:	STRING	- Output text
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/
Function GetFileCount(nFiles as long) as string

	if (nFiles = 1) THEN
		GetFileCount = "A total of 1 file will be generated."
	else
		GetFileCount = "A total of " + CSTR(nFiles) + " files will be generated."
	endif
	ArraySize = nFiles
End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	Round
REM 
REM 	Action:	Rounds a DOUBLE value to to the nearest integer
REM 
REM 	Params:	Value	- Value to be rounded up
REM 
REM 	Returns:	LONG		- Rounded value
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/
Function Round(Value as double) as long

	Dim Frac		as double
	Dim TmpValue	as double
	Dim TmpRound	as long

	TmpRound = INT(Value)
	Frac = Value - TmpRound
	
	if (Frac <> 0.0) then
		Round = TmpRound + 1
	else
		Round = TmpRound
	endif
	
End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	Floor
REM 
REM 	Action:	Returns a floating-point value representing the largest 
REM			integer that is less than or equal to x
REM 
REM 	Params:	Value	- Value to be converted
REM 
REM 	Returns:	DOUBLE	- Converted value
REM 
REM 	Comments:	INT function returns the greatest integer less than or 
REM			equal to the number
REM 
REM *************************************************************EDOC*/
Function Floor(Value as double) as double

	DIM nTemp	as long
	
	nTemp = INT( Value )
	Floor = nTemp

End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	Power
REM 
REM 	Action:	Calculates X raised to the power of Y
REM 
REM 	Params:	Base		- base value (X)
REM			Expon	- exponent (Y)
REM 
REM 	Returns:	DOUBLE	- Converted value
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Function Power(Base as double, Expon as long) as double

	DIM	nExp	as long
	nExp = Expon
	
	Power = 1
	if (nExp >= 1) then
	
		Do While (nExp > 0)
			Power = Power * Base
			nExp = nExp - 1
		Loop
		
	endif
	
End Function
REM *SDOC*************************************************************
REM 
REM 	Name:	RemoveExt
REM 
REM 	Action:	Removes the extension from a file name
REM 
REM 	Params: 	FileName 
REM			
REM			
REM 
REM 	Returns: 	String		- Returns string
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/
Function RemoveExt(FileName as string) as string
	Dim IndexR as integer
	Dim tempChar as string
	Dim TempString as string
	
	indexR = 1
	Size = len (FileName)
	tempChar = RIGHT(FileName, indexR)
	While (tempChar <>".") 
		indexR = indexR + 1
		tempRStr = RIGHT(FileName, indexR)
		tempChar = left (tempRStr,1)
	WEND	
	tempSize = Size-indexR	
	TempString = left(FileName, tempSize)
	RemoveExt= TempString
			
End Function
