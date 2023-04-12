REM Orange face, VP locked to object, 2 light sources      Font used:
REM REM BravoEngraved

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 60, 100, 0
	.ApplyUniformFillColor
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyOutline 3556, 0, 0, 0, 100, 0, 0, -1, -1, FALSE, 2, 0, FALSE
	.StoreColor 2, 0, 0, 0, 100
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 60, 100, 0
	.StoreColor 2, 0, 60, 100, 0
	.StoreColor 2, 0, 0, 0, 100
	.ApplyExtrude 4, 0, 0, 5, -177545, 62229, FALSE, 12, 100, 5, 58, 0, 0, 2
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

