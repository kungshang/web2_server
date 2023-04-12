REM Favorite : Fill

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects FALSE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 100, 0, 100, 0
	.StoreColor 2, 20, 80, 0, 20
	.ApplyTwoColorFill "dots_grf.bmp", 127000, 127000, 0, 0, TRUE, 0, TRUE, FALSE
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT


