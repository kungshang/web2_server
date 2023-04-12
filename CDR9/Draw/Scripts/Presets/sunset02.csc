REM Off center radial fill with extrude and  perspective.

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.RecorderApplyPerspective 1, 0, -237743, -67817, -109727, 67817, 109727, 67817, 237743, -67817, 0, 0, 0, 0, 0, 0
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 60, 60, 40, 0, 0, 100,  0
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  80
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  100
	.ApplyFountainFill 1, 0, -50, 900, 256, 0, 3, 50
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyExtrude 4, 0, 0, 20, 1523, -43687, TRUE, 0, 0, 0, 0, 0, 0, 0, 0
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

