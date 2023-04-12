REM Split Objects With Mask For use with Photo-Paint 9.

REM *******************************************************
REM	This script splits selected objects using the current
REM	mask. It does this by duplicating each object and using
REM	the ClipObject (to Mask) command on each copy, using
REM the mask with one and the inverse of the mask with
REM the other.
REM
REM	The objects can be split with or without object order
REM	being maintained; the latter is faster. In both cases
REM	the part of the object inside the mask will be ordered
REM	in front of the part outside the mask.
REM
REM Created On October 26, 1995 by Rennie Houtman
REM Updated for Version 9, January 1999
REM ********************************************************

REM Function declarations.

Declare Function RectanglesIntersect(nL1&, nT1&, nR1&, nB1&, nL2&, nT2&, nR2&, nB2&) as Boolean

REM Constants.

Const MAX_OBJECTS = 100
Const CANCELLED = 2
Const TO_FRONT = 0
Const TO_BACK = 1
Const FORWARD = 2
Const BACK = 3
Const strDialogTitle = "Split Objects with Mask"

REM Variable delarations.

Dim nObjectCount			as long
Dim nObjectIndex			as long
Dim nCopyIndex				as long
Dim nIndex						as long
Dim nObjectsToSplit		as long
Dim nL								as long
Dim nT								as long
Dim nR								as long
Dim nB								as long
Dim nML								as long
Dim nMT								as long
Dim nMR								as long
Dim nMB								as long
Dim nLeft							as long
Dim nTop							as long
Dim bFoundSelectedObjects as Boolean
Dim anObjectIndices(1 to MAX_OBJECTS) as long

REM Options dialog definition.

BEGIN DIALOG SplitOptionsDialog 142, 40, strDialogTitle
	CANCELBUTTON  75, 22, 50, 13
	OKBUTTON  17, 22, 50, 13
	CHECKBOX  30, 4, 100, 12, "Preserve object order", MaintainOrder%
END DIALOG

REM Start of the code.
label1 = "This script split selected objects using the current mask."
BEGIN DIALOG Dialog1 196, 74, "Object Splitter"
	GROUPBOX  2, 1, 191, 53
	TEXT  8, 12, 178, 37, label1
	OKBUTTON  109, 57, 40, 14
	CANCELBUTTON  153, 57, 40, 14
END DIALOG

ret = DIALOG(Dialog1)
' If Cancel is selected, stop the script
IF ret = 2 THEN STOP

WITHOBJECT "CorelPhotoPaint.Automation.9"

	REM Check for a mask.

	if .GetMaskPresent() = False then
		MessageBox "A mask is needed for this script.", strDialogTitle, 0
		STOP
	endif

	REM Check for objects.

	nObjectCount = .GetObjectCount()
	
	if nObjectCount <= 0 then
		MessageBox "The image has no objects.", strDialogTitle, 0
		STOP
	endif

	REM Count and save indices of selected objects that intersect
	REM the mask bounding rectangle.

	nObjectsToSplit = 0
	bFoundSelectedObjects = FALSE
	.GetMaskRectangle nML, nMT, nMR, nMB

	for nObjectIndex = 1 to nObjectCount
		if .GetObjectisSelected(nObjectIndex) then
			bFoundSelectedObjects = TRUE
			.GetObjectRectangle nObjectIndex, nL, nT, nR, nB
			if RectanglesIntersect(nL, nT, nR, nB, nML, nMT, nMR, nMB) then
				nObjectsToSplit = nObjectsToSplit + 1
				if nObjectsToSplit > MAX_OBJECTS then
					MessageBox "This script can split a maximum of " & Str(MAX_OBJECTS) & " objects at a time.", strDialogTitle, 0
					STOP
				endif
				anObjectIndices(nObjectsToSplit) = nObjectIndex
			endif
		endif
	next nObjectIndex

	if Not bFoundSelectedObjects then
		MessageBox "No objects are selected.", strDialogTitle, 0
		STOP
	endif

	if nObjectsToSplit <= 0 then
		MessageBox "No selected objects are within the mask bounding rectangle.", strDialogTitle, 0
		STOP
	endif

	REM	Display the Options dialog.

	MaintainOrder% = 1
	HideDoc% = 1
	nDialogResult = Dialog(SplitOptionsDialog)

	if nDialogResult = CANCELLED then
		STOP
	endif

	REM Here's the (faster) version that doesn't maintain object order.

	if MaintainOrder% <> 1 then
		REM Select the objects and duplicate them en masse. The
		REM duplicates will be placed in front of the other objects.

		.ObjectDuplicate
		.ObjectSelectNone
		for nIndex = 1 to nObjectsToSplit
			.ObjectSelect anObjectIndices(nIndex), TRUE
		next nIndex
		nObjectCount = nObjectCount + nObjectsToSplit
		.EndObject

		REM Correct the duplicate's positions.

		.ObjectSelectNone
		for nIndex = 1 to nObjectsToSplit
			.GetObjectRectangle anObjectIndices(nIndex), nLeft, nTop, nR, nB
			nCopyIndex = nObjectCount - nObjectsToSplit + nIndex
			.GetObjectRectangle nCopyIndex, nL, nT, nR, nB
			.ObjectTranslate nLeft - nL, nTop - nT
			.ObjectSelect nCopyIndex, TRUE
			.EndObject
			.ObjectSelect nCopyIndex, FALSE
		next nIndex

		REM Select and clip the duplicates en masse.

		.ObjectClip
		for nIndex = (nObjectCount - nObjectsToSplit + 1) to nObjectCount
			.ObjectSelect nIndex, TRUE
		next nIndex
		.EndObject

		REM Select and clip the original objects en masse.

		.MaskInvert
		.ObjectClip
		.ObjectSelectNone
		for nIndex = 1 to nObjectsToSplit
			.ObjectSelect anObjectIndices(nIndex), TRUE
		next nIndex
		.EndObject

		.MaskInvert

		if HideDoc% = 1 then 
			.SetDocVisible TRUE
		endif

		REM All done.

		STOP
	endif

	REM Here's the version that maintains object ordering.

	REM Duplicate each selected object and use the object order commands
	REM to put the duplicate just ahead of the original.
	REM Then clip the original, taking care to account for objects that
	REM	are deleted by the clip.

	nIndexCorrection = 0
	.ObjectSelectNone
	.MaskInvert

	for nIndex = 1 to nObjectsToSplit
		anObjectIndices(nIndex) = anObjectIndices(nIndex) + nIndexCorrection
		.ObjectDuplicate
		.ObjectSelectNone
		.ObjectSelect anObjectIndices(nIndex), TRUE
		.EndObject
		.GetObjectRectangle anObjectIndices(nIndex), nLeft, nTop, nR, nB
		nObjectCount = nObjectCount + 1
		nIndexCorrection = nIndexCorrection + 1
		.GetObjectRectangle nObjectCount, nL, nT, nR, nB
		.ObjectTranslate nLeft - nL, nTop - nT
		.EndObject

		nDestIndex = anObjectIndices(nIndex) + 1

		if (nDestIndex - 1) < (nObjectCount - nDestIndex) then
			nNumberOfMoves = nDestIndex - 1
			nOrderChange = FORWARD
		else
			nNumberOfMoves = nObjectCount - nDestIndex
			nOrderChange = BACK
		endif

		if nNumberOfMoves > 0 then
			if nOrderChange = FORWARD then
				.ObjectOrder TO_BACK
				.EndObject
			endif

			while nNumberOfMoves > 0
				.ObjectOrder nOrderChange
				.EndObject
				nNumberOfMoves = nNumberOfMoves - 1
			wend
		endif

		REM Clip the original object.

		.ObjectClip
		.ObjectSelect nDestIndex, FALSE
		.ObjectSelect anObjectIndices(nIndex), TRUE
		.EndObject

		REM Check whether the object was deleted by the clip.
		REM If it was, it won't be necessary to clip the copy
		REM to the inverse mask, so mark it with a negative index.

		if .GetObjectCount() < nObjectCount then
			nIndexCorrection = nIndexCorrection - 1
			nObjectCount = nObjectCount - 1
			anObjectIndices(nIndex) =  -1
		else
			.ObjectSelect anObjectIndices(nIndex), FALSE
		endif
	next nIndex

	REM Now clip all the duplicate objects in one step.

	.MaskInvert

	.ObjectClip
	for nIndex = 1 to nObjectsToSplit
		if anObjectIndices(nIndex) >= 0 then
			.ObjectSelect anObjectIndices(nIndex) + 1, TRUE
		endif
	next nIndex
	.EndObject


	.ObjectSelectNone

	REM All done.

END WITHOBJECT



REM Function definitions.

Function RectanglesIntersect(nL1&, nT1&, nR1&, nB1&, nL2&, nT2&, nR2&, nB2&) as Boolean
	RectanglesIntersect = (nL1<=nR2) and (nR1>=nL2) and (nT1<=nB2) and (nB1>=nT2)
End Function

