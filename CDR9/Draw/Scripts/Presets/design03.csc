REM Purple fountain fill with 75% transparent  drop shadow.    Font used:
REM REM BANK GOTHIC Md BT

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyOutline 0, 0, 0, 0, 0, 0, 0, -1, -1, FALSE, 0, 0, FALSE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 0, 0, 50
	.ApplyLensEffect 10, FALSE, FALSE, FALSE, 0, 0, 250
	.RecorderSelectObjectByIndex TRUE, 1
	.DuplicateObject 0, 0
	.RecorderSelectObjectByIndex TRUE, 2
	.MoveObject -5841, 6349
	.RecorderSelectObjectByIndex TRUE, 2
	.StoreColor 2, 0, 100, 0, 0, 0, 0, 100,  0
	.StoreColor 2, 20, 80, 0, 20, 0, 0, 100,  100
	.ApplyFountainFill 0, 0, 0, 900, 256, 0, 3, 50
	.RecorderSelectPreselectedObjects TRUE
	.RecorderSelectObjectsByIndex FALSE, 2, -1, -1, -1, -1
	.Group 
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

