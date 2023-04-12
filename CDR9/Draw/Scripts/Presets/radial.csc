REM Yellow to red off center fill, extruded.

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 100, 100, 0, 0, 0, 100,  0
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  80
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  100
	.ApplyFountainFill 1, 0, -50, 900, 256, 0, 3, 50
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyExtrude 4, 0, 0, 20, -29463, -47243, TRUE, 0, 0, 0, 0, 0, 0, 0, 0
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

