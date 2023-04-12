REM Blue fountain fill with black to purple  blended edge    Font used:
REM REM Bassoon

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyOutline 5080, 1, 0, 0, 100, 0, 0, -1, -1, FALSE, 2, 0, TRUE
	.StoreColor 2, 0, 0, 0, 100
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 1
	.DuplicateObject 0, 0
	.RecorderSelectObjectByIndex TRUE, 2
	.ApplyOutline 508, 1, 0, 0, 100, 0, 0, -1, -1, FALSE, 2, 0, TRUE
	.StoreColor 2, 0, 100, 0, 0
	.SetOutlineColor
	.RecorderSelectPreselectedObjects TRUE
	.RecorderSelectObjectsByIndex FALSE, 2, -1, -1, -1, -1
	.ApplyBlend TRUE, 10, 0, FALSE, 0, FALSE, FALSE, 0, 0, 0, TRUE, TRUE, FALSE, FALSE, 0, 0, 0, 0
	.RecorderSelectObjectByIndex TRUE, 2
	.DuplicateObject 0, 0
	.RecorderSelectObjectByIndex TRUE, 4
	.ApplyOutline 508, 1, 0, 0, 100, 0, 0, -1, -1, FALSE, 2, 0, TRUE
	.StoreColor 2, 0, 0, 0, 0
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 4
	.StoreColor 2, 0, 0, 0, 100, 0, 0, 100,  0
	.StoreColor 2, 100, 100, 0, 0, 0, 0, 100,  25
	.StoreColor 2, 40, 0, 0, 0, 0, 0, 100,  48
	.StoreColor 2, 100, 100, 0, 0, 0, 0, 100,  69
	.StoreColor 2, 0, 0, 0, 100, 0, 0, 100,  100
	.ApplyFountainFill 0, 0, 0, 900, 256, 0, 3, 50
	.RecorderSelectPreselectedObjects TRUE
	.RecorderSelectObjectsByIndex FALSE, 4, 3, 2, -1, -1
	.Group 
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

