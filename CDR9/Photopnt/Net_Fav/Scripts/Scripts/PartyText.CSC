REM Created in CorelPhotoPaint Version 9.334
REM Created On Monday, April, 12, 1999 by stevet

BaseSize = 72

label1 = "This script produces Party text. Enter a text string, to create a new 24-bit image."
BEGIN DIALOG Dialog1 196, 74, "Party Text"
	GROUPBOX  2, 1, 191, 53
	TEXT  8, 12, 178, 37, label1
	OKBUTTON  109, 57, 40, 14
	CANCELBUTTON  153, 57, 40, 14
END DIALOG

ret = DIALOG(Dialog1)
' If Cancel is selected, stop the script
IF ret = 2 THEN STOP

' Get the Text
TxtStr$ = inputbox("Party Text Effect")
if TxtStr$ = "" then
	messagebox "You have entered nothing", "Warning", 48
	stop
end if


WITHOBJECT "CorelPhotoPaint.Automation.9"
	.FileNew 800, 200, 1, 72, 72, FALSE, FALSE, 1, 0, 0, 0, 0, 255, 255, 255, 0, FALSE
	.TextTool 400, 120, FALSE, TRUE, 0
		.TextSetting "Fill", "0,0,0"
		.TextSetting "Font", "Arial Black"
		.TextSetting "TypeSize", STR(BaseSize)
		.TextSetting "Justify", "2"

		TxtStr$ = " " + TxtStr$
		
		RandomFill = 1
		RandomAngle = 1
		RandomSize = 1
		
		Randomize

		For i = 1 to Len(TxtStr$)
			if RandomAngle = 1 Then
				AngleStr$ = STR(INT((RND(35)-17))*1000000)
				.TextSetting "TypeAngle", AngleStr$
				Endif

			if RandomSize = 1 Then
				NewSize = BaseSize + (RND(BaseSize/2) - BaseSize/4)
				.TextSetting "TypeSize", STR(NewSize)
				Endif

			if RandomFill = 1 Then
				FillStr$ = Str(INT(RND(256))) + "," + STR(INT(RND(256))) + "," + STR(INT(RND(256)))
				.TextSetting "Fill", FillStr
				Endif

			.TextAppend Mid(TxtStr$, i, 1)
		Next i
		
		.TextRender

END WITHOBJECT
