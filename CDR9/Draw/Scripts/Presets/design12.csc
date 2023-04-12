REM 3D Drop shadow to front    Font used: USA Black (Italic)

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyOutline 0, 0, 0, 0, 0, 0, 0, -1, -1, FALSE, 0, 0, FALSE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 20, 80, 0, 20, 0, 0, 100,  0
	.StoreColor 2, 20, 20, 0, 0, 0, 0, 100,  48
	.StoreColor 2, 20, 80, 0, 20, 0, 0, 100,  100
	.ApplyFountainFill 0, 0, 0, 900, 256, 0, 3, 50
	.RecorderSelectObjectByIndex TRUE, 1
	.DuplicateObject 0, 0
	.RecorderSelectObjectByIndex TRUE, 2
	.StretchObject 1, 1, 0.253968, 1, FALSE, TRUE, 2
	.RecorderSelectObjectByIndex TRUE, 2
	.SkewObject 31800000, 0, 2
	.RecorderSelectObjectByIndex TRUE, 2
	.MoveObject -2539, 6857
	.RecorderSelectObjectByIndex TRUE, 2
	.OrderToBack 
	.RecorderSelectObjectByIndex TRUE, 2
	.StoreColor 2, 0, 0, 0, 50, 0, 0, 100,  0
	.StoreColor 2, 20, 20, 0, 0, 0, 0, 100,  48
	.StoreColor 2, 0, 0, 0, 0, 0, 0, 100,  100
	.ApplyFountainFill 0, 0, 0, -900, 256, 0, 0, 50
	.RecorderSelectPreselectedObjects TRUE
	.RecorderSelectObjectsByIndex FALSE, 2, -1, -1, -1, -1
	.Group 
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

