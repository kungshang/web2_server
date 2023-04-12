REM Rainbow fill, extruded

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.DuplicateObject 0, 0
	.RecorderSelectObjectByIndex TRUE, 2
	.OrderToBack 
	.RecorderSelectObjectByIndex TRUE, 2
	.ApplyExtrude 0, 0, 0, 5, 253, 253, FALSE, 0, 0, 0, 0, 0, 0, 0, 1
	.RecorderSelectObjectsByIndex TRUE, 3, 2, -1, -1, -1
	.StoreColor 2, 40, 40, 0, 20, 0, 0, 100,  0
	.StoreColor 2, 0, 0, 0, 0, 0, 0, 100,  100
	.ApplyFountainFill 0, 0, 0, 900, 256, 0, 0, 50
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 100, 20, 0, 0, 0, 0, 100,  0
	.StoreColor 2, 100, 20, 0, 0, 0, 0, 100,  14
	.StoreColor 2, 20, 80, 0, 20, 0, 0, 100,  20
	.StoreColor 2, 40, 60, 0, 0, 0, 0, 100,  30
	.StoreColor 2, 100, 100, 0, 0, 0, 0, 100,  40
	.StoreColor 2, 100, 0, 100, 0, 0, 0, 100,  50
	.StoreColor 2, 0, 0, 100, 0, 0, 0, 100,  60
	.StoreColor 2, 0, 60, 100, 0, 0, 0, 100,  70
	.StoreColor 2, 0, 100, 100, 0, 0, 0, 100,  80
	.StoreColor 2, 100, 20, 0, 0, 0, 0, 100,  85
	.StoreColor 2, 100, 20, 0, 0, 0, 0, 100,  100
	.ApplyFountainFill 1, 0, -50, 900, 256, 0, 3, 50
	.RecorderSelectPreselectedObjects TRUE
	.RecorderSelectObjectsByIndex FALSE, 3, 2, -1, -1, -1
	.Group 
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

