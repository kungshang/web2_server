REM Black / white fountain fill face, VP locked to object, 2 light sources
REM REM    Font used: Bedrock

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects TRUE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 0, 0, 0, 100, 0, 0, 100,  0
	.StoreColor 2, 0, 0, 0, 0, 0, 0, 100,  100
	.ApplyFountainFill 0, 0, 0, 900, 256, 0, 0, 50
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyOutline 3556, 0, 0, 0, 100, 0, 0, -1, -1, FALSE, 2, 0, FALSE
	.StoreColor 2, 0, 0, 0, 100
	.SetOutlineColor
	.RecorderSelectObjectByIndex TRUE, 1
	.ApplyExtrude 4, 0, 0, 20, 0, -31495, FALSE, 12, 100, 5, 58, 0, 0, 0, 0
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT

