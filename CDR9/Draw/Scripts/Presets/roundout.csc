REM Rounded outline look.

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyNoFill 
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyOutline 25400, 0, 0, 0, 100, -900, 0, -1, -1, FALSE, 1, 0, TRUE
	.StoreColor 2, 0, 0, 0, 50
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyOutline 12700, 1, 0, 0, 100, 0, 0, -1, -1, FALSE, 2, 0, TRUE
	.StoreColor 2, 0, 0, 0, 50
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 1
	.DuplicateObject 0, 0
	.RecorderSelectObjectByIndex TRUE, 2
	.ApplyOutline 12700, 1, 0, 0, 100, -900, 0, -1, -1, FALSE, 1, 0, TRUE
	.StoreColor 2, 0, 0, 0, 0
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 2
	.ApplyOutline 3302, 1, 0, 0, 100, 0, 0, -1, -1, FALSE, 2, 0, TRUE
	.StoreColor 2, 0, 0, 0, 0
	.SetOutlineColor
	.RecorderSelectPreselectedObjects TRUE
	.RecorderSelectObjectsByIndex FALSE, 2, -1, -1, -1, -1
	.ApplyBlend TRUE, 5, 0, FALSE, 0, FALSE, FALSE, 0, 0, 0, TRUE, TRUE, FALSE, FALSE, 0, 0, 0, 0
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

