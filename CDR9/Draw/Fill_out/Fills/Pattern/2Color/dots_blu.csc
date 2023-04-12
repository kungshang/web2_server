REM Favorite : Fill

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects FALSE
	.RecorderSelectObjectByIndex TRUE, 1
	.StoreColor 2, 100, 100, 0, 0
	.StoreColor 2, 0, 100, 100, 0
	.ApplyTwoColorFill "dots_buf.bmp", 127000, 127000, 0, 0, TRUE, 0, TRUE, FALSE
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT


