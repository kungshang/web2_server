REM Blue fountain fill face, VP locked to object, 3 light sources    Font
REM REM used: AvantGarde Bk BT

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 100, 100, 0, 0, 0, 0, 100,  0
	.StoreColor 2, 100, 0, 0, 0, 0, 0, 100,  100
	.ApplyFountainFill 0, 0, 0, 900, 256, 0, 0, 50
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyOutline 3556, 0, 0, 0, 100, 0, 0, -1, -1, FALSE, 2, 0, FALSE
	.StoreColor 2, 0, 0, 0, 100
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 100, 100, 0, 0
	.StoreColor 2, 100, 100, 0, 0
	.StoreColor 2, 0, 0, 0, 100
	.ApplyExtrude 0, 0, 0, 20, -49275, -39369, TRUE, 3, 100, 10, 32, 7, 35, 2
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

