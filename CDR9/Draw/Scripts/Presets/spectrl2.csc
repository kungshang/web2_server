REM Off center spectral fill with reverse spectral fill on contour.

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 100, 100, 0, 0, 0, 100,  0
	.StoreColor 2, 100, 0, 0, 0, 0, 0, 100,  65
	.StoreColor 2, 0, 100, 100, 0, 0, 0, 100,  75
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  85
	.StoreColor 2, 0, 100, 100, 0, 0, 0, 100,  100
	.ApplyFountainFill 2, 0, -50, 0, 256, 0, 1, 50
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 0, 0, 0
	.StoreColor 2, 100, 100, 0, 0
	.StoreColor 2, 0, 0, 0, 0
	.ApplyContour 2, 1523, 5, 0
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

