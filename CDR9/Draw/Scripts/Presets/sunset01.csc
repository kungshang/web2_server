REM Off center radial fill, extruded.

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 60, 60, 40, 0, 0, 100,  0
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  80
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  100
	.ApplyFountainFill 1, 0, -50, 900, 256, 0, 3, 50
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyExtrude 4, 0, 0, 20, 17017, -14477, TRUE, 0, 0, 0, 0, 0, 0, 0, 0
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

