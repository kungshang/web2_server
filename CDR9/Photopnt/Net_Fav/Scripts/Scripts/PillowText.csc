REM Pillow Text Effect for PhotoPaint 9
REM Created by Rob Wineck

label1 = "This script produces pillow embossed text. Enter a text string to create a new 24-bit image."
BEGIN DIALOG Dialog1 196, 74, "Pillow Text"
	GROUPBOX  2, 1, 191, 53
	TEXT  8, 12, 178, 37, label1
	OKBUTTON  109, 57, 40, 14
	CANCELBUTTON  153, 57, 40, 14
END DIALOG

ret = DIALOG(Dialog1)
' If Cancel is selected, stop the script
IF ret = 2 THEN STOP
TxtStr$ = inputbox("Pillow Effect Text")
if TxtStr$ = "" then
	message "You have entered nothing"
	stop
end if

WITHOBJECT "CorelPhotoPaint.Automation.9"
	.FileNew 640, 200, 1, 96, 96, FALSE, FALSE, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0
	
	' The Text String
	.TextTool 10, 100, FALSE, TRUE ' add Mask Mode 0
		.SetPaintColor 5, 255, 0, 0, 0
		.TextSetting "Font", "Arial Black"
		.TextSetting "TypeSize", "96.0"
		.TextAppend TxtStr$
		.TextRender 

	' Center the Object in the Doucment
	.ObjectAlign 3, 3, FALSE, FALSE, FALSE, TRUE, FALSE, TRUE, TRUE
		.EndObject
	' Create a Mask from the object
	.MaskCreate TRUE, 0
		.EndMaskCreate
	'Delete the Object
	.ObjectDelete 
		.EndObject		
	' Emboss effect in the Mask
	.EffectEmboss 2, 500, 63, 1, 5, 0, 0, 0, 0
	' Save the Mask for Later
	.MaskChannelAdd "Alpha 1"
	.MaskRemove 
	
	.EffectGaussianBlur 3.00
	.EffectEmboss 1, 500, 63, 1, 5, 0, 0, 0, 0

	.MaskChannelToMask 0, 0
	.ImageInvert 
		.EndColorEffect 
	.EditFill 0, 60, 100, 1, 0, 12, 58, 77, 34, 0, 0
		.FillSolid 5, 255, 0, 0, 0
		.EndEditFill 
	' Recreate Text Object
	.ObjectCreate FALSE
END WITHOBJECT
