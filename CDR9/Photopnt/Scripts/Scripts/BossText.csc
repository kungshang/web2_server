REM The Boss Text Effect for PhotoPaint 9

label1 = "This script applies The Boss Effect to text. Enter a text string to create a new 24-bit image."
BEGIN DIALOG Dialog1 196, 74, "The Boss Text Effect"
	GROUPBOX  2, 1, 191, 53
	TEXT  8, 12, 178, 37, label1
	OKBUTTON  109, 57, 40, 14
	CANCELBUTTON  153, 57, 40, 14
END DIALOG

ret = DIALOG(Dialog1)
' If Cancel is selected, stop the script
IF ret = 2 THEN STOP
TxtStr$ = inputbox("The Boss Effect Text")
if TxtStr$ = "" then
	message "You have entered nothing"
	stop
end if

WITHOBJECT "CorelPhotoPaint.Automation.9"
	.FileNew 640, 480, 1, 72, 72, FALSE, FALSE, 1, 0, 0, 0, 0, 255, 255, 255, 0, FALSE
	
	' Example of Gradient Fill Tool - Rainbow Gradient - 
	.GradientTool 4, 0, 0, 2
		.SetPaintColor 5, 0, 255, 255, 0
		.SetPaperColor 5, 0, 255, 0, 0
		.GradientPoint 0, 321, 242, 5, 0, 255, 255, 0, 255
		.GradientPoint 1, 641, 242, 5, 0, 255, 0, 0, 255
		.Gradient 4, 50

	.TextTool 10, 100, TRUE, TRUE, 0
		.TextSetting "Font", "Futura XBlk BT"
		.TextSetting "TypeSize", "100.0"
		.TextAppend TxtStr$
		.TextRender 

	' Center the Mask
	.MaskAlign 3,3,FALSE,FALSE,FALSE,FALSE,TRUE

	' The Boss Effect - Outside 
	.EffectTheBoss 45, 20, 30, 60, 135, 2, 30, 70, FALSE
	' The Boss Effect - Inside 
	.EffectTheBoss 45, 20, 30, 60, 135, 2, 30, 70, TRUE

END WITHOBJECT
