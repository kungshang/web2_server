REM Spectral conical fill.

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 100, 100, 0, 0, 0, 100,  0
	.StoreColor 2, 0, 100, 100, 0, 0, 0, 100,  100
	.ApplyFountainFill 1, 0, 0, 900, 256, 0, 1, 50
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 100, 100, 0, 0, 0, 100,  0
	.StoreColor 2, 0, 100, 100, 0, 0, 0, 100,  100
	.ApplyFountainFill 2, 0, 0, 900, 256, 0, 1, 50
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

