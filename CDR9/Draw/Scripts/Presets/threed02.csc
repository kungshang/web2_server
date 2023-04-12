REM Gold fountain fill face, VP locked to object, 3 light sources    Font
REM REM used: BahamasHeavy

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 20, 100, 0, 0, 0, 100,  0
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  28
	.StoreColor 2, 0, 0, 0, 0, 0, 0, 100,  50
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  72
	.StoreColor 2, 0, 20, 100, 0, 0, 0, 100,  100
	.ApplyFountainFill 0, 0, 0, 900, 256, 0, 3, 50
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyOutline 3556, 0, 0, 0, 100, 0, 0, -1, -1, FALSE, 2, 0, FALSE
	.StoreColor 2, 0, 0, 0, 100
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 20, 100, 0
	.StoreColor 2, 0, 20, 100, 0
	.StoreColor 2, 0, 0, 0, 100
	.ApplyExtrude 0, 0, 0, 10, -9397, 292861, FALSE, 3, 63, 5, 25, 0, 0, 2
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

