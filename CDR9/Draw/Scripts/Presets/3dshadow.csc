REM 3D SHADOW - a combination of blend and extrude effects give a nice
REM REM 3D look to text and other objects.

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyOutline 127000, 1, 1, 1, 100, 0, 0, -1, -1, FALSE, 2, 0, FALSE
	.StoreColor 2, 0, 0, 0, 0
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 0, 0, 0
	.ApplyUniformFillColor
	.RecorderSelectObjectByIndex TRUE, 1
	.DuplicateObject 0, 0
	.RecorderSelectObjectByIndex TRUE, 2
	.StoreColor 2, 0, 0, 0, 40
	.ApplyUniformFillColor
	.RecorderSelectObjectByIndex TRUE, 2
	.ApplyOutline 127000, 1, 1, 1, 100, 0, 0, -1, -1, FALSE, 1, 0, FALSE
	.StoreColor 2, 0, 0, 0, 40
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 2
	.ApplyOutline 0, 0, 0, 0, 0, 0, 0, -1, -1, FALSE, 0, 0, FALSE
	.RecorderSelectObjectByIndex TRUE, 2
	.DuplicateObject 0, 0
	.RecorderSelectObjectByIndex TRUE, 3
	.ApplyExtrude 5, 0, 0, 20, 0, 65277, FALSE, 0, 0, 0, 0, 0, 0, 0, 1
	.RecorderSelectObjectsByIndex TRUE, 4, 3, -1, -1, -1
	.StoreColor 2, 0, 0, 0, 40
	.StoreColor 2, 0, 0, 0, 40
	.StoreColor 2, 0, 100, 100, 0
	.ApplyExtrude 5, 0, 0, 20, 0, 12953, FALSE, 0, 0, 0, 0, 0, 0, 2
	.RecorderSelectPreselectedObjects TRUE
	.RecorderSelectObjectsByIndex FALSE, 2, -1, -1, -1, -1
	.ApplyBlend TRUE, 20, 0, FALSE, 0, FALSE, FALSE, 0, 0, 0, TRUE, TRUE, FALSE, FALSE, 0, 0, 0, 0
	.RecorderSelectPreselectedObjects TRUE
	.RecorderSelectObjectsByIndex FALSE, 5, 4, 3, 2, -1
	.Group 
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

