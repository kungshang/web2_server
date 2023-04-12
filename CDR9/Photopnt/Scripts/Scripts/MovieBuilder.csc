REM Script for Building individual Numbered Files into a movie
REM Choose Load Image Type and Save to which Format

Declare Function GetDirectory (DirStr as string,textSize as integer ) as string

Global gDir as string
Global gDirDes as String

Dim MovieName as string

' Feel Free to add more Filter Here
Global LoadTypes$(7)
LoadTypes(1) = "CPT"
LoadTypes(2) = "BMP"
LoadTypes(3) = "TGA"
LoadTypes(4) = "PCX"
LoadTypes(5) = "TIF"
LoadTypes(6) = "JPG"
LoadTypes(7) = "GIF"

' All the movies Formats supported
Global SaveTypes$(4)
SaveTypes(1) = "AVI"
SaveTypes(2) = "MPEG"
SaveTypes(3) = "MOV"
SaveTypes(4) = "GIF"

Note1 = "All files must be in the same format and use the following naming convention:" \\
				 + CHR(13)+CHR(10)+ CHR(13)+CHR(10) + "FileName##.EXT" \\
				 + CHR(13)+CHR(10)+ CHR(13)+CHR(10) + "Example: "   \\
				 + CHR(13)+CHR(10)+ "C:\MyFile0.CPT" \\
				 + CHR(13)+CHR(10)+ "C:\MyFile1.CPT"

BEGIN DIALOG OBJECT Movie 268, 167, "Movie Builder", SUB SubMovie
	GROUPBOX  4, 2, 162, 90, .Frame, "Frame Options"	 'id 1
	PUSHBUTTON  13, 13, 40, 14, .Open, "Open"	 'id 2
	TEXT  13, 30, 40, 8, .BaseName, "File Name:"	 'id 3
	TEXTBOX  58, 26, 100, 14, .BaseNameStr	 'id 4
	TEXT  13, 46, 40, 8, .type, "File Type:"	 'id 5
	DDLISTBOX  58, 43, 100, 50, .LoadType	 'id 6
	TEXT  13, 62, 74, 8, .Start, "Start at File Number:"	 'id 7
	SPINCONTROL  103, 58, 35, 14, .StartFrame	 'id 8
	TEXT  13, 77, 86, 8, .NOfFrame, "Number of Files to Load:"	 'id 9
	SPINCONTROL  104, 75, 35, 14, .NumFrame	 'id 10
	GROUPBOX  4, 93, 162, 54, .SaveOption, "Save Options"	 'id 11
	PUSHBUTTON  13, 102, 40, 14, .Open2, "Open"	 'id 12
	TEXT  13, 119, 40, 8, .SaveName, "Save Name:"	 'id 13
	TEXTBOX  58, 115, 100, 14, .TextBox2	 ', MovieName$ 'id 14
	TEXT  13, 133, 40, 8, .SaveAs, "Save As:"	 'id 15
	DDLISTBOX  58, 131, 100, 50, .SaveType	 'id 16
	GROUPBOX  170, 2, 95, 145, .Note, "Note:"	 'id 17
	TEXT  177, 13, 81, 119, .Note1, Note1	 'id 18
	OKBUTTON  179, 150, 40, 14, .Ok	 'id 19
	CANCELBUTTON  224, 150, 40, 14, .Cancel	 'id 20
	TEXT  59, 13, 103, 11, .DirOpen, " "
	TEXT  58, 102, 103, 11, .DirDes, " "
END DIALOG


 SUB SubMovie (BYVAL ControlID%, BYVAL Event%)
	Dim retDir as string
	Dim tempDir as string
	
	If Event= 0 Then
		WITHOBJECT "CorelPhotoPaint.Automation.9" 
			gDir=.getPhotopaintdir()
			gDirDes = .getPhotopaintdir()
		END WITHOBJECT
		retDir = GetDirectory(gDir, 15)
		retDirDes = GetDirectory(gDir, 15)
		
		Movie.DirOpen.SETTEXT retDir
		Movie.DirDes.SETTEXT retDirDes
		
		Movie.StartFrame.SETMINRANGE 0
		Movie.StartFrame.SETDOUBLEMODE False
		Movie.StartFrame.SETPRECISION 0
		
		Movie.NumFrame.SETMINRANGE 1
		Movie.NumFrame.SETDOUBLEMODE False
		Movie.NumFrame.SETPRECISION 0
		
		Movie.SaveType.SETARRAY SaveTypes$
		Movie.SaveType.SETSELECT 1 'Pixels is chosen at the begining
		Movie.SaveType.ENABLE TRUE
	
		Movie.LoadType.SETARRAY LoadTypes$
		Movie.LoadType.SETSELECT 2 'Pixels is chosen at the begining
		Movie.LoadType.ENABLE TRUE
		
		
	End if
	if Event = 1  Then 
		Select case ControlId
			Case 4
				Basename = Movie.BaseNameStr.GETTEXT ()
			
			Case 14
				SaveName = Movie.TextBox2.GETTEXT ()
						
		End Select
	End if
	
	if Event = 2 Then
	Select Case ControlId
		Case 2
			gDir = GETFOLDER(gDir)
			retDir = GetDirectory(gDir, 15)
			Movie.DirOpen.SETTEXT retDir
			tempDir = Right ( gDir,1)
			if tempDir<> "\" Then
				gDir = gDir & "\"
				Movie.DirOpen.SETTEXT gDir
		End if	
		Case 12
			gDirDes = GETFOLDER(gDirDes)
			retDir = GetDirectory(gDirDes, 15)
			Movie.DirDes.SETTEXT retDir
			tempDir = Right ( gDirDes,1)
			if tempDir<> "\" Then
				retDir = retDir & "\"
				Movie.DirDes.SETTEXT retDir
		End if	
	End Select		
	
	End if	

End Sub
label1 = "This script builds a movie using sequentially numbered files."
BEGIN DIALOG Dialog1 196, 74, "Movie Builder"
	GROUPBOX  2, 1, 191, 53
	TEXT  8, 12, 178, 37, label1
	OKBUTTON  109, 57, 40, 14
	CANCELBUTTON  153, 57, 40, 14
END DIALOG

ret = DIALOG(Dialog1)
' If Cancel is selected, stop the script
IF ret = 2 THEN STOP

ret = DIALOG(Movie) 'displays Movie Dialog box
IF ret = 2 THEN STOP ' the Cancel Button
	
' Setup the Extensions
LoadType = Movie.LoadType.GETSELECT()
StartFrame = Movie.StartFrame.GETVALUE()
NumFrame = Movie.NumFrame.GETVALUE()
MovieName = Movie.textBox2.GetText()
SaveType = Movie.SaveType.GETSELECT()
Basename = Movie.BaseNameStr.GETTEXT ()
tempDir = Right ( gDir,1)
	if tempDir<> "\" Then
		FileToOpen = gDir + "\" + Basename
	Else 
		FileToOpen = gDir + Basename
	end if	

tempDes = Right ( gDirDes,1)
	if tempDes<> "\" Then
		MovieName = gDirDes + "\" +  MovieName 
	Else 
		MovieName = gDirDes+  MovieName 
	end if	
		
select case LoadType
	case 1
		NameExt$ = ".cpt"
	case 2
		NameExt$ = ".bmp"
	case 3
		NameExt$ = ".tga"
	case 4
		NameExt$ = ".pcx"
	case 5
		NameExt$ = ".tif"
	case 6
		NameExt$ = ".jpg"
	case 7
		NameExt$ = ".gif"
end select

WITHOBJECT "CorelPhotoPaint.Automation.9"

	'ltrim is used to remove leading spaces if any	 
	FileName$ = FileToOpen  + ltrim(str(StartFrame)) + NameExt$  
	' Check if FileName Exists	
	ret = FileAttr(FileName$)
	if ret = 0 then
		 messagebox FileName$ + " does not exist. Script cancelled.", "Warning", 48
		stop
	end if
	.FileOpen FileName$, 0, 0, 0, 0, 0, 1, 1
  .MovieCreate
	
	Count% = 0
	for i = StartFrame+1 to (StartFrame + NumFrame -1 )
		Count = Count + 1
		FileName$ = FileToOpen  + ltrim(str(i)) + NameExt$  
  	.MovieInsertFile FileName$, 0, 0, 0, 0, 0, Count, FALSE
	next i
  
	' Get the Proper Filter ID and Extension
	select case SaveType
		case 1
			FilterID% = 1536
			NameExt$ = ".avi"
		case 2
			FilterID% = 1551
			NameExt$ = ".mpg"
		case 3
			FilterID% = 1542
			NameExt$ = ".mov"
		case 4
			FilterID% = 1558  ' 773 for paint 8
			NameExt$ = ".gif"
	end select

	' Build the Save Name
	MovieName$ = MovieName + NameExt$
	.FileSave MovieName$, FilterID, TRUE
END WITHOBJECT


REM *SDOC*************************************************************
REM 
REM 	Name:	GetDirectory
REM 
REM 	Action:	if the Directory Size >Text Size put Directory as exp: c:\...\Dorectpru
REM 
REM 	Params: 	DirStr		- Photo paint directory 
REM			textSize		- The size of the text in the dialog
REM			
REM 
REM 	Returns: 	String		- Returns string
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/

Function GetDirectory (DirStr as string,textSize as integer ) as string
	Dim Size as integer
	Dim tempRStr as string
	Dim tempLStr as string
	Dim indexL as integer
	Dim indexR as integer
	Dim tempChar as string
	Dim tempLChar as string
	Dim retStr as string
	
	indexL = 1
	indexR = 1
	Size = len (DirStr)
	if Size > textSize Then
		tempChar = RIGHT(DirStr, indexR)
		if tempChar = "\" then
			indexR = indexR + 1
			tempChar = RIGHT(DirStr, indexR)
		END IF	
			
		While (tempChar <>"\") and (indexR <> Size)
			indexR = indexR + 1
			tempRStr = RIGHT(DirStr, indexR)
			tempChar = left (tempRStr,1)
		WEND	
		
		tempLChar = Left(DirStr, indexL)	 
		While (tempLChar <> "\") and (indexR <> Size)
			indexL = indexL +1
			tempLStr = Left (DirStr, indexL)
			tempLChar = Right(tempLStr,1) 		
		WEND
			
		retStr = tempLStr + "..."+ tempRStr
		GetDirectory = retStr
		
	Else
		GetDirectory = DirStr
	End if
End function
