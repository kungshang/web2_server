REM Conical fill with extrude.

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
rem	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  0
	.StoreColor 2, 0, 60, 60, 40, 0, 0, 100,  10
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  20
	.StoreColor 2, 0, 60, 60, 40, 0, 0, 100,  30
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  40
	.StoreColor 2, 0, 60, 60, 40, 0, 0, 100,  50
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  60
	.StoreColor 2, 0, 60, 60, 40, 0, 0, 100,  70
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  80
	.StoreColor 2, 0, 60, 60, 40, 0, 0, 100,  90
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  100
	.ApplyFountainFill 2, 0, -50, 900, 256, 0, 3, 50
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyExtrude 4, 0, 0, 20, -32765, -33527, TRUE, 0, 0, 0, 0, 0, 0, 0, 0
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

