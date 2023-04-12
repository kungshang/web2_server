REM Melting/Dripping Effect, start with Single Frame
REM PhotoPaint 9, January 1999 by Rob Wineck


REM *** Set the Number of Frames to Add
NUMBER_FRAMES = 10
nFrame& = 1

label1 = "The script produces animated melting text. Enter a text string to create a new movie."
BEGIN DIALOG Dialog1 196, 74, "Wet Paint"
	GROUPBOX  2, 1, 191, 53
	TEXT  8, 12, 178, 37, label1
	OKBUTTON  109, 57, 40, 14
	CANCELBUTTON  153, 57, 40, 14
END DIALOG

ret = DIALOG(Dialog1)
' If Cancel is selected, stop the script
IF ret = 2 THEN STOP

TxtStr$ = inputbox("Melting Text")
if TxtStr$ = "" then
	message "You have entered nothing"
	stop
end if

WITHOBJECT "CorelPhotoPaint.Automation.9"
	.FileNew 400, 200, 1, 96, 96, FALSE, FALSE, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0
	
	' The Text String
	.TextTool 10, 100, FALSE, TRUE ' add Mask Mode 0
		.SetPaintColor 5, 255, 0, 0, 0
		.TextSetting "Fill", "255, 0, 0"
		.TextSetting "Font", "Arial Black"
		.TextSetting "TypeSize", "64.0"
		.TextAppend TxtStr$
		.TextRender 


	'Center object and Merge with Background
	.ObjectAlign 3,3,FALSE,FALSE,FALSE,TRUE,TRUE,TRUE,TRUE
	.EndObject
	.ObjectMerge TRUE
	.EndObject

	' Make a Movie
	.MovieCreate
	.MovieGotoFrame 1
	FOR i% = 1 to NUMBER_FRAMES
	  .MovieForward
	  ' Insert to After the Last Frame
		.MovieInsertFrame i, 1, 0, 1
		REM Smooth Dripping Paint Effect
		.EffectWetPaint 40, 100
		.EffectMedian 1, 100

	NEXT i%

END WITHOBJECT
