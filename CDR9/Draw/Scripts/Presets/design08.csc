REM Blue fountain fill with feathered drop  shadow    Font used: AvantGarde
REM REM Bk BT

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 0, 0, 100
	.ApplyUniformFillColor
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyOutline 3556, 1, 1, 1, 100, 0, 0, -1, -1, FALSE, 2, 0, TRUE
	.StoreColor 2, 0, 0, 0, 0
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 1
	.DuplicateObject 0, 0
	.RecorderSelectObjectByIndex TRUE, 2
	.StoreColor 5, 153, 153, 153, 0
	.ApplyUniformFillColor
	.RecorderSelectObjectByIndex TRUE, 2
	.ApplyOutline 254, 1, 1, 1, 100, 0, 0, -1, -1, FALSE, 2, 0, TRUE
	.StoreColor 2, 0, 0, 0, 50
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 2
	.DuplicateObject 0, 0
	.RecorderSelectObjectByIndex TRUE, 3
	.MoveObject -7619, 6095
	.RecorderSelectObjectByIndex TRUE, 3
	.StoreColor 2, 100, 20, 0, 0, 0, 0, 100,  0
	.StoreColor 2, 100, 100, 0, 0, 0, 0, 100,  100
	.ApplyFountainFill 0, 0, 0, 900, 256, 0, 0, 50
	.RecorderSelectPreselectedObjects TRUE
	.RecorderSelectObjectsByIndex FALSE, 2, -1, -1, -1, -1
	.ApplyBlend TRUE, 10, 0, FALSE, 0, FALSE, FALSE, 0, 0, 0, TRUE, TRUE, FALSE, FALSE, 0, 0, 0, 0
	.RecorderSelectPreselectedObjects TRUE
	.RecorderSelectObjectsByIndex FALSE, 4, 3, 2, -1, -1
	.Group 
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

