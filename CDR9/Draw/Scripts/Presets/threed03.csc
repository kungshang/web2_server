REM Purple / blue fountain fill face, VP locked to object, 2 light sources
REM REM    Font used: BankGothic Md BT

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 100, 100, 0, 0, 0, 0, 100,  0
	.StoreColor 2, 0, 100, 0, 0, 0, 0, 100,  46
	.StoreColor 2, 100, 100, 0, 0, 0, 0, 100,  100
	.ApplyFountainFill 0, 0, 0, 900, 256, 0, 3, 50
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyOutline 3556, 0, 0, 0, 100, 0, 0, -1, -1, FALSE, 2, 0, FALSE
	.StoreColor 2, 0, 0, 0, 100
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 100, 100, 0, 0
	.StoreColor 2, 0, 0, 0, 0
	.StoreColor 2, 0, 0, 0, 100
	.ApplyExtrude 4, 0, 0, 10, 24637, -22605, TRUE, 12, 79, 5, 56, 0, 0, 2
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

