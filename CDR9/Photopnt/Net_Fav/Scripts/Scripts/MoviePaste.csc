REM Paste an Object to all Movie Frames for PhotoPaint 9

REM Array of Names - Room for 10, increase if you need more
Global nFrames as integer

Global ObjNames(10) as STRING

Dim ReturnValue1 as integer
Dim ReturnValue2 as integer
Dim ReturnValue3 as integer
Dim ReturnValue4 as integer

NameVal% = 1
StartOpacity% = 100
EndOPacity% = 100

REM Set to FALSE if you don't want the dialog box
bWarning% = TRUE

BEGIN DIALOG OBJECT PasteObject 193, 123, "Paste Object To Frames", SUB SubPast
	TEXT  14, 20, 38, 8, .StartFrame, "Start Frame"	 'Id 1
	TEXT  14, 37, 36, 8, .EndFrame, "End Frame"	 'Id 2
	SPINCONTROL  58, 18, 30, 13, .nStart	 'Id 3
	SPINCONTROL  58, 36, 30, 14, .nEnd	 'Id 4 
	OKBUTTON  106, 106, 40, 14, .OK1
	CANCELBUTTON  149, 106, 40, 14, .Cancel1
	LISTBOX  105, 14, 84, 90, .ListBox1	 ', NameVal% 'Id 7
	GROUPBOX  5, 4, 95, 55, .FrameRange, "Frame Range"	 'Id 8
	SPINCONTROL  58, 78, 30, 14, .SpinControl3	 'Id 9
	SPINCONTROL  58, 98, 30, 14, .SpinControl4	 'Id 10 
	TEXT  12, 82, 41, 8, .StartOpacity, "Start Opacity"	 'Id 11
	TEXT  12, 100, 39, 8, .EndOpacity, "End Opacity"	 'Id 12 
	TEXT  106, 4, 50, 8, .SelectObject, "Select Object"	 'Id 13
	GROUPBOX  5, 65, 96, 55, .GroupBox2	 'Id 14
END DIALOG

label1 = "This script combines all objects with the background."  \\
					+ CHR(13)+CHR(10)+ CHR(13)+CHR(10)+ \\
					"NOTE: A 24-bit or grayscale movie must be open in Corel PHOTO-PAINT. Ensure all objects are properly positioned before running the script."

Sub SubPast(BYVAL ControlID%, BYVAL Event%)
	IF Event=0 THEN
		WITHOBJECT "CorelPhotoPaint.Automation.9"
			nFrames% = .GetFrameCount()
		END WITHOBJECT
		PasteObject.nStart.SETMINRANGE 1
		PasteObject.nStart.SETMAXRANGE nFrames
		
		PasteObject.nEnd.SETMINRANGE 1
		PasteObject.nEnd.SETMAXRANGE nFrames
		PasteObject.nEnd.setvalue nFrames
		
		PasteObject.SpinControl3.SETMINRANGE 0
		PasteObject.SpinControl3.SETMAXRANGE 100
		
		PasteObject.SpinControl4.SETMINRANGE 0
		PasteObject.SpinControl4.SETMAXRANGE 100
		
		PasteObject.ListBox1.SETARRAY ObjNames 
		PasteObject.ListBox1.SETSELECT 1 ' Default Selection
		
		PasteObject.SpinControl3.setvalue 100
		PasteObject.SpinControl4.setvalue 100
		
	End if
	if Event=1 and ControlID = 3 Then
		RetStart% = PasteObject.nStart.GETVALUE()
		RetEnd% = PasteObject.nEnd.GETVALUE()
		if RetStart%> RetEnd% Then
			PasteObject.nEnd.SETVALUE (RetStart)
		End if
	End if	
	if Event=1 and ControlID = 4 Then
		RetStart% = PasteObject.nStart.GETVALUE()
		RetEnd% = PasteObject.nEnd.GETVALUE()
		if RetStart%> RetEnd% Then
			PasteObject.nEnd.SETVALUE (RetStart)
		End if
	End if	
			
End Sub	


 
BEGIN DIALOG Dialog1 196, 80, "Movie Paste"
	GROUPBOX  2, 1, 191, 59
	TEXT  8, 9, 178, 47, label1
	OKBUTTON  109, 64, 40, 14
	CANCELBUTTON  153, 64, 40, 14
END DIALOG

ret = DIALOG(Dialog1)
' If Cancel is selected, stop the script
IF ret = 2 THEN STOP


WITHOBJECT "CorelPhotoPaint.Automation.9"
	REM Get the Number of Frames in the Movie	
	nFrames% = .GetFrameCount()
	nStart% = 1
	nEnd% = nFrames

	if (nFrames = 0) THEN
		BEEP
		message "This script runs on Movie Files"
		STOP
	ENDIF
	
	nObjects% = .GetObjectCount()
	if (nObjects = 0) then
		BEEP
		message "This script needs objects"
		STOP
	endif

	'Make the Pick Tool Active
	.SetActiveTool 34501
	
	FOR i% = 1 TO nObjects 
		ObjNames(i%) = .GetObjectName(i%)
	NEXT i%

	REM Do the Dialog. No error checking so enter proper values
	ret2 = DIALOG(PasteObject)
	IF ret2 = 2 THEN STOP

	' get Opacity values
	StartOpacity = PasteObject.SpinControl3.Getvalue()
	EndOpacity = PasteObject.SpinControl4.Getvalue()
	nStart = PasteObject.nStart.GETVALUE()
	nEnd = PasteObject.nEnd.GETVALUE()
	NameVal = PasteObject.ListBox1.GETSELECT() 
	REM Select the Chosen Object, and make a shadow for it
	.ObjectSelectNone
	.ObjectSelect NameVal, -1
	.GetObjectRectangle NameVal, oLeft&, oTop&, oRight&, oBottom&
	.EditCut
	
	REM Calculate Fade Factor
	increment& = (EndOpacity - StartOpacity) / (nEnd - nStart)
	
	For i% = nStart to nEnd	
		.MovieGotoFrame i
		.EditPasteObject oLeft, oBottom, "" 
		
		.ObjectOpacity StartOpacity + (i - nStart) * increment
			.ObjectSelectNone 
			.ObjectSelect nObjects, -1
			.EndObject
		
		.ObjectMerge 0
			.EndObject
	Next i
END WITHOBJECT
