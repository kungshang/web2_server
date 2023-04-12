REM Crop To Objects For use with Photo-Paint 9.

REM ********************************************************
REM	This script will crop the active image to the bounding
REM	rectangle of all objects, selected objects, or visible
REM	objects in the image.
REM
REM Created On October 26, 1995 by Rennie Houtman
REM Updated for Version 9, January 1999
REM ********************************************************

REM Function declarations.

Declare Function MaxLong(x as long, y as long) as Long
Declare Function MinLong(x as long, y as long) as Long

REM Constants.

Const CANCELLED	= 2
Const	strDialogTitle = "Crop To Objects"
Const CROP_ALL			= 0
Const CROP_SELECTED	= 1
Const CROP_VISIBLE	= 2

REM Variable delarations.

Dim nObjectCount			as long
Dim nObjectIndex			as long
Dim nDialogResult			as integer
Dim nL								as long
Dim nT								as long
Dim nR								as long
Dim nB								as long
Dim nLCrop						as long
Dim nTCrop						as long
Dim nRCrop						as long
Dim nBCrop						as long
Dim nImageRight				as long
Dim nImageBottom			as long
Dim nCropObjectCount	as long
Dim bIncludeObject		as boolean

REM The Crop Options dialog definition.

BEGIN DIALOG CropOptionsDialog 135, 94, strDialogTitle
	GROUPBOX  4, 4, 126, 61, "Crop To"
	OPTIONGROUP CropOption%
		OPTIONBUTTON  14, 15, 94, 14, "All Objects"
		OPTIONBUTTON  14, 30, 94, 14, "Selected Objects"
		OPTIONBUTTON  14, 45, 94, 14, "Visible Objects"
	OKBUTTON  10, 75, 54, 15
	CANCELBUTTON  70, 75, 55, 15
END DIALOG

REM Start of the code.
label1 = "This script crops to all objects in the image."
BEGIN DIALOG Dialog1 196, 74, "Object Crop"
	GROUPBOX  2, 1, 191, 53
	TEXT  8, 12, 178, 37, label1
	OKBUTTON  109, 57, 40, 14
	CANCELBUTTON  153, 57, 40, 14
END DIALOG

ret = DIALOG(Dialog1)
' If Cancel is selected, stop the script
IF ret = 2 THEN STOP

WITHOBJECT "CorelPhotoPaint.Automation.9"

	REM Get the number of objects in the image.
	
	nObjectCount = .GetObjectCount()
	
	if nObjectCount <= 0 then
		MessageBox "The image has no objects.", strDialogTitle, 0
		STOP
	endif

	REM Display the Crop Options dialog.

	CropOption% = CROP_ALL
	nDialogResult = Dialog(CropOptionsDialog)

	if nDialogResult = CANCELLED then
		STOP
	endif	

	REM Now get the bounding rectangle of all the
	REM objects that are to be cropped to.

	nImageRight = .GetDocumentWidth()
	nImageBottom = .GetDocumentHeight()
	nImageRight = nImageRight - 1
	nImageBottom = nImageBottom - 1

	nLCrop = nImageRight
	nTCrop = nImageBottom
	nRCrop = 0
	nBCrop = 0
	nCropObjectCount = 0

	for nObjectIndex = 1 to nObjectCount
		Select Case CropOption%
			Case CROP_ALL
				bIncludeObject = TRUE
			Case CROP_SELECTED
				bIncludeObject = .GetObjectIsSelected(nObjectIndex)
			Case CROP_VISIBLE
				bIncludeObject = .GetObjectIsVisible(nObjectIndex)
			Case Else
				MessageBox "Programming error!", strDialogTitle, 0
				STOP
		End Select

		if bIncludeObject then
			.GetObjectRectangle nObjectIndex, nL, nT, nR, nB
			nLCrop = MinLong(nLCrop, nL)
			nTCrop = MinLong(nTCrop, nT)
			nRCrop = MaxLong(nRCrop, nR)
			nBCrop = MaxLong(nBCrop, nB)
			nCropObjectCount = nCropObjectCount + 1
		endif
	next nObjectIndex

	if nCropObjectCount <= 0 then
		Select Case CropOption%
			Case CROP_SELECTED
				MessageBox "No objects are selected.", strDialogTitle, 0
			Case CROP_VISIBLE
				MessageBox "There are no visible objects.", strDialogTitle, 0
			Case Else
				MessageBox "Programming error!", strDialogTitle, 0
		End Select
		STOP
	endif

	REM Make sure the crop rectangle doesn't extend outside the image.
	REM This is possible since objects can extend beyond the edges
	REM of the image.

	nLCrop = MaxLong(nLCrop, 0)
	nTCrop = MaxLong(nTCrop, 0)
	nRCrop = MinLong(nRCrop, nImageRight)
	nBCrop = MinLong(nBCrop, nImageBottom)

	REM If the crop rectangle is the same size as the image or empty
	REM do nothing.

	if (nLCrop = 0) and (nTCrop = 0) and (nRCrop = nImageRight) and (nBCrop = nImageBottom) then
		STOP
	endif

	if (nLCrop > nRCrop) or (nTCrop > nBCrop) then
		STOP
	endif

	REM Do the crop.

	.ImageCrop nLCrop, nTCrop, nRCrop, nBCrop

	REM All done.

END WITHOBJECT



REM Function definitions.

Function MaxLong(x as long, y as long) as Long
	if x >= y then
		MaxLong = x
	else
		MaxLong = y
	endif
End Function

Function MinLong(x as long, y as long) as Long
	if x <= y then
		MinLong = x
	else
		MinLong = y
	endif
End Function

