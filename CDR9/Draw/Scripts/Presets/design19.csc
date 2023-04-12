REM 90 Deg. Swirl with 10 step blend      Font used: AvantGarde Bk BT

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.DuplicateObject 0, 0
	.RecorderSelectObjectByIndex TRUE, 2
	.RotateObject 90000000, FALSE, 0, 0
	.RecorderSelectObjectByIndex TRUE, 2
	.OrderToBack 
	.RecorderSelectObjectByIndex TRUE, 2
	.StoreColor 2, 0, 0, 0, 0
	.ApplyUniformFillColor
	.RecorderSelectPreselectedObjects TRUE
	.RecorderSelectObjectsByIndex FALSE, 2, -1, -1, -1, -1
	.ApplyBlend TRUE, 10, 0, FALSE, 0, FALSE, FALSE, 0, 0, 0, TRUE, TRUE, FALSE, FALSE, 0, 0, 0, 0
	.RecorderSelectObjectByIndex TRUE, 1
	.DuplicateObject 0, 0
	.RecorderSelectObjectByIndex TRUE, 4
	.StoreColor 2, 0, 100, 100, 0
	.ApplyUniformFillColor
	.RecorderSelectPreselectedObjects TRUE
	.RecorderSelectObjectsByIndex FALSE, 4, 3, 2, -1, -1
	.Group 
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

