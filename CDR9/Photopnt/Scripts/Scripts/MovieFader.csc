REM Movie Fader - Just like on some Camcorders
REM Created On November 18, 1999 by Rob Wineck
Global nFrames as long
Global gRed as integer
Global gGreen as integer
Global gBlue as integer

gRed = 0
gGreen = 0
gBlue = 0


REM Default Fade Color is Black
'***********************************Dialog declaration
BEGIN DIALOG OBJECT Fader 177, 111, "Screen Fade Options", SUB SubFade
	TEXT  13, 10, 45, 11, .StartFrametxt, "Start Frame: "	 'id 1
	SPINCONTROL  13, 21, 40, 15, .startFrame	 'id 2
	TEXT  70, 10, 41, 11, .EndFrametxt, "End Frame: "	 'id 3
	SPINCONTROL  70, 21, 40, 15, .endFrame	 'id 4
	OPTIONGROUP .Fade
		OPTIONBUTTON  122, 15, 41, 12, .Fadein, "Fade In"	 'id 6
		OPTIONBUTTON  122, 26, 47, 14, .Fadeout, "Fade Out"	 'id 7
	GROUPBOX  3, 0, 169, 90, .GroupBox1	 'id 8
	TEXT  13, 45, 68, 10, .FadeColor, "Fade Color:"	 'id 9
	PUSHBUTTON  13, 61, 50, 15, .ColorOpt, "Color"	 'id 10
	TEXT  70, 63, 98, 11, .ColorGreen, "Green: "	 'id 11
	OKBUTTON  69, 94, 50, 15, .OK1	 'id 12
	CANCELBUTTON  123, 94, 50, 15, .Cancel1	 'id 13
	TEXT  70, 50, 98, 11, .ColorRed, "Red: "	 'id 11
	TEXT  70, 76, 98, 11, .ColorBlue, "Blue: "	 'id 11
END DIALOG

SUB SubFade (BYVAL ControlID%, BYVAL Event%)
	IF Event=0 THEN
		WITHOBJECT "CorelPhotoPaint.Automation.9"
			nFrames = .GetFrameCount()
		END WITHOBJECT
		
		Fader.StartFrame.SETMINRANGE 1
		Fader.StartFrame.SETMAXRANGE nFrames
		
		Fader.endFrame.SETMINRANGE 1
		Fader.endFrame.SETMAXRANGE nFrames
		Fader.endFrame.SETVALUE nFrames
	
		Fader.Fadeout.SETVALUE 1
				
		ColorOptRed$=" Red: "& STR (gRed)  
		ColorOptGreen$=" Green: " & STR (gGreen)
		ColorOptBlue$=" Blue: " & STR (gBlue) 
		Fader.ColorGreen.SETTEXT ColorOptGreen
		Fader.ColorRed.SETTEXT ColorOptRed
		Fader.ColorBlue.SETTEXT ColorOptBlue
			
	End if 
	
	IF EVENT = 1  THEN 
		SELECT CASE ControlID
			CASE 2
				RetStart% = Fader.StartFrame.GETVALUE()
				RetEnd% = Fader.endFrame.GETVALUE()
				if RetStart%> RetEnd% Then
					Fader.endFrame.SETVALUE (RetStart)
				End if
			
			CASE 4
				RetStart% = Fader.StartFrame.GETVALUE()
				RetEnd% = Fader.endFrame.GETVALUE()
				if RetStart%> RetEnd% Then
					Fader.endFrame.SETVALUE (RetStart)
				End if
			
		END SELECT
	END IF	
	IF EVENT = 2 AND ControlID =10 THEN
		GETCOLOR gRed, gGreen, gBlue
		ColorOptRed$=" Red: "& STR (gRed)  
		ColorOptGreen$=" Green: " & STR (gGreen)
		ColorOptBlue$=" Blue: " & STR (gBlue) 
		Fader.ColorGreen.SETTEXT ColorOptGreen
		Fader.ColorRed.SETTEXT ColorOptRed
		Fader.ColorBlue.SETTEXT ColorOptBlue
	
	End if	 
End Sub	

'***************************Set defaults

label1 = "This script produces a fade out effect over a series of movie frames." \\
					+ CHR(13)+CHR(10)+ CHR(13)+CHR(10)+ \\
					"NOTE: A 24-bit or grayscale movie must be open in Corel PHOTO-PAINT."
BEGIN DIALOG Dialog1 196, 77, "Movie Fader"
	GROUPBOX  2, 1, 191, 55
	TEXT  8, 9, 178, 41, label1
	OKBUTTON  109, 60, 40, 14
	CANCELBUTTON  153, 60, 40, 14
END DIALOG

ret = DIALOG(Dialog1)
' If Cancel is selected, stop the script
IF ret = 2 THEN STOP


WITHOBJECT "CorelPhotoPaint.Automation.9"

	Dim CheckFadeIn as integer
	Dim CheckFadeOut as integer
	
	nFrames = .GetFrameCount()
	if (nFrames < 1) then
		message "This script only runs on a mutliple frame movie"
		STOP
	end if

	REM *** CHECK to see if valid movie type
	docType& = .GetDocumentType()
	if (docType <> 1) AND (docType <> 2) then
		message "Script only works on 24bit and Grayscale Movies"
		STOP
	endif	

	num%  = 0

	ret = DIALOG(Fader)
	IF ret = 2 THEN STOP
	startFrame% = Fader.StartFrame.GETVALUE()
	endFrame% = Fader.endFrame.GETVALUE()
	IF Fader.Fade.GetValue() = 0 THEN
		CheckFadeIn= 1
	ELSE
		CheckFadeIn= 0
	ENDIF
	
	nFrames = (endFrame% - startFrame%) + 1
	REM Calculate the Fade Increment
	FadeValue% = 100 / nFrames
	REM the first frame we are working on
	.MovieGotoFrame startFrame%
	
	FOR i% = 1 to nFrames
		if CheckFadeIn = 0 then
			num% = 100 - ( FadeValue * (i-1) )
			if num% < 0 then num% = 0
	  else
			num% = FadeValue% * i%
			if num% > 100 then num% = 100
		end if 
		.EditFill 0, num, 100, 1, 0, 0, 0, 0, 0, 0, 0
		.FillSolid 5, gRed, gGreen, gBlue, 0
		.EndEditFill 
		.MovieForwardOne
	NEXT i%

END WITHOBJECT
