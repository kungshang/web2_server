REM Chrome Text Effect for PHOTO-PAINT 9
REM Prompt User for text then create a New File and Text Object

label1 = "This script produces metallic gold text. Enter a Text string to create a new 24-bit image."
BEGIN DIALOG Dialog1 196, 74, "Metal Text"
	GROUPBOX  2, 1, 191, 53
	TEXT  8, 12, 178, 37, label1
	OKBUTTON  109, 57, 40, 14
	CANCELBUTTON  153, 57, 40, 14
END DIALOG

ret = DIALOG(Dialog1)
' If Cancel is selected, stop the script
IF ret = 2 THEN STOP

TxtStr$ = inputbox("Metal Text Effect")
if TxtStr$ = "" then
	message "You have entered nothing"
	stop
end if

WITHOBJECT "CorelPhotoPaint.Automation.9"
	.FileNew 640, 480, 1, 72, 72, FALSE, FALSE, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0
	' The Text String
	.TextTool 10, 100, FALSE, TRUE ' add Mask Mode 0
		.TextSetting "Fill", "255, 204, 102"
		.TextSetting "Font", "Futura XBlk BT"
		.TextSetting "TypeSize", "120.0"
		.TextAppend TxtStr$
		.TextRender 
	
	' Center the Object
	.ObjectAlign 3,3,FALSE,FALSE,FALSE,TRUE,TRUE,TRUE,TRUE
		.EndObject
	'Lock Object Transparency	
	.ObjectEditTransparency 0
	
	' Add Noise
	.EffectAddNoise 50, 50, FALSE, 0, 2, 5, 0, 0, 0, 0
		.EndColorEffect
	.EffectMotionBlur 10, 225, 0
	.EffectEmboss 10, 25, 45, 0, 5, 0, 0, 0, 0
	.EffectSharpen 27, 53, FALSE
		.EndColorEffect
		
END WITHOBJECT
