REM Favorite : Outline,Outline Style,Outline Width,Outline Arrow Head,Outline Arrow Tail

WITHOBJECT "CorelDraw.Automation.9"
	.StartOfRecording 
	.SuppressPainting 
	.RecorderStorePreselectedObjects FALSE
	.SetOutlineWidth 7056
	.SetOutlineMiscProperties 2, 2, 0, 0, 100, 0, FALSE, FALSE
	.StoreColor 2, 0, 100, 0, 0
	.SetOutlineColor
	.BeginDrawArrow TRUE, 0, 5
	.AddArrowPoint -254000, 254000, FALSE, FALSE, TRUE, TRUE, 0, 0
	.AddArrowPoint 254000, 254000, FALSE, FALSE, TRUE, FALSE, 0, 1
	.AddArrowPoint 254000, -254000, FALSE, FALSE, TRUE, FALSE, 0, 1
	.AddArrowPoint -254000, -254000, FALSE, FALSE, TRUE, FALSE, 0, 1
	.AddArrowPoint -254000, 254000, FALSE, FALSE, FALSE, TRUE, 0, 1
	.EndDrawArrow 
	.SetOutlineArrow 0
	.BeginDrawArrow FALSE, 0, 5
	.AddArrowPoint -254000, 254000, FALSE, FALSE, TRUE, TRUE, 0, 0
	.AddArrowPoint 254000, 254000, FALSE, FALSE, TRUE, FALSE, 0, 1
	.AddArrowPoint 254000, -254000, FALSE, FALSE, TRUE, FALSE, 0, 1
	.AddArrowPoint -254000, -254000, FALSE, FALSE, TRUE, FALSE, 0, 1
	.AddArrowPoint -254000, 254000, FALSE, FALSE, FALSE, TRUE, 0, 1
	.EndDrawArrow 
	.SetOutlineArrow 1
	.ResumePainting 
	.EndOfRecording 
END WITHOBJECT


