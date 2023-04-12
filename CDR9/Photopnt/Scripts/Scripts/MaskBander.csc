REM Interlaced Mask Creator for PhotoPaint 9
REM January 1999 by Rob Wineck

Direction% = 1
StartBand% = 0
BandSize% = 2
NonBandSize% = 2
Feather% = 0
nStart% = 0

BEGIN DIALOG Mask 143, 99, "Mask Interlacer"
	TEXT  8, 19, 52, 8, "Mask band size:"
	SPINCONTROL  90, 15, 40, 14, BandSize%
	OPTIONGROUP Direction%
		OPTIONBUTTON  8, 50, 50, 10, "Vertical"
		OPTIONBUTTON  8, 62, 50, 10, "Horizontal"
	CANCELBUTTON  100, 81, 40, 14
	OKBUTTON  56, 81, 40, 14
	GROUPBOX  2, 5, 138, 72, "Mask Options"
	OPTIONGROUP StartBand%
		OPTIONBUTTON  90, 50, 35, 10, "Even"
		OPTIONBUTTON  90, 62, 35, 10, "Odd"
	TEXT  8, 34, 70, 8, "Non-Mask band size:"
	SPINCONTROL  90, 31, 40, 14, NonBandSize%
END DIALOG

label1 = "This script produces an interlaced mask."
BEGIN DIALOG Dialog1 196, 74, "Interlaced Mask Creation"
	GROUPBOX  2, 1, 191, 53
	TEXT  8, 12, 178, 37, label1
	OKBUTTON  109, 57, 40, 14
	CANCELBUTTON  153, 57, 40, 14
END DIALOG

ret = DIALOG(Dialog1)
' If Cancel is selected, stop the script
IF ret = 2 THEN STOP
nRet = Dialog(Mask)
if nRet = 2 then stop

WITHOBJECT "CorelPhotoPaint.Automation.9"

	REM **** Use Document's Width and Height
	Width% = .GetDocumentWidth()
	Height% = .GetDocumentHeight()
	
	
	if StartBand then StartBand = NonBandSize
	if Direction then
		WHILE StartBand < Height
			.MaskRectangle 0, (StartBand) , Width, (StartBand + BandSize)-1, 1, 0 		
			StartBand = StartBand + BandSize + NonBandSize
		WEND
	else
		WHILE StartBand < Width
			.MaskRectangle StartBand, 0 , StartBand + BandSize - 1, Height, 1, 0 		
			StartBand = StartBand + BandSize + NonBandSize
		WEND
	endif 
END WITHOBJECT


