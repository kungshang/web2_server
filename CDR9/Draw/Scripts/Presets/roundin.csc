REM Grey contour to center.

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyOutline 762, 0, 0, 0, 100, 0, 0, -1, -1, FALSE, 1, 0, FALSE
	.StoreColor 2, 0, 0, 0, 50
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 0, 0, 50
	.ApplyUniformFillColor
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 0, 0, 0
	.StoreColor 2, 0, 0, 0, 0
	.StoreColor 2, 0, 0, 0, 0
	.ApplyContour 0, 1269, 13, 0
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

