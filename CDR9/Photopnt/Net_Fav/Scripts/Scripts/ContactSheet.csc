Declare Function ConvertToPixels(Value as double, Res as double, Units as long) as long
Declare Function ConvertToUnit(Value as double, Res as double, OldUnits as long, NewUnits as long) as double
Declare Function ConvertBaseToUnit(Value as double, Res as double, Units as long) as double

Declare Function MeasurmentOption(OptionArr as Long) as String
Declare Function GetThumbW(Byref dWidth as double )as Boolean
Declare Function GetThumbH( Byref dHeight as double)as Boolean
Declare Function GetPageBorder(Byref dBorder as double )as Boolean
Declare Function  PutThumbWidthString()as string
Declare Function PutThumbHeightString()as string
Declare Function FourDecPoint(nValue as double) as double
Declare Function SetDecimalPoints(nValue as Long, nSize as double)as double
Declare Function CreateHTML()as Boolean
Declare Function PutText() as Boolean
Declare Function PutName()as Boolean
Declare Function PutDate()as Boolean
Declare Function PutSize()as Boolean
Declare Function PutDim()as Boolean
Declare Function PutLink()as Boolean
Declare Function GetDirectory (DirStr as string,textSize as integer) as string
Declare Sub CreateHTMLText( index as integer,npages as integer, DesDirectory as string)
Declare Function RemoveExt(FileName as string)as string
Declare Sub SetRange(nValue as long)
Declare Sub SubBuild()
Declare Sub CovertToPaletted()

REM Create an array called arr
GLOBAL FileTypeArr$(10)
FileTypeArr(1) = "All Files"
FileTypeArr(2) = "BMP"
FileTypeArr(3) = "CPT"
FileTypeArr(4) = "GIF"
FileTypeArr(5) = "JPG"
FileTypeArr(6) = "PCD"
FileTypeArr(7) = "PCX"
FileTypeArr(8) = "RIFF"
FileTypeArr(9) = "TGA"
FileTypeArr(10) = "TIF"

GLOBAL WidthOptionsArr$(8)
WidthOptionsArr(1) = "Inches"
WidthOptionsArr(2) = "millimeters" 
WidthOptionsArr(3) = "picas, points"  
WidthOptionsArr(4) = "points"
WidthOptionsArr(5) = "centimeters"
WidthOptionsArr(6) = "Pixels"
WidthOptionsArr(7) = "ciceros, didots"
WidthOptionsArr(8) = "didots"


GLOBAL ColorArr$(6)
ColorArr(1) = "1-bit Black and White"
ColorArr(2) = "8-bit Grayscale"
ColorArr(3) = "8-bit Paletted"
ColorArr(4) = "24-bit RGB"
ColorArr(5) = "24-bit Lab"
ColorArr(6) = "32-bit CMYK"

Global	GWidth 		as double
Global 	GHeight 		as double
Global 	Gborder 		as double
Global 	GResolution 	as double
Global 	gMeasurment	as long
Global 	gDirectory	as string
Global 	gDirectoryDes 	as string
Global	gtextdest		as string
Global	gtextDir		as string
Global 	ret2			as integer



BEGIN DIALOG OBJECT Sheet 373, 223, "Contact Sheet", SUB CSHEET
	GROUPBOX  2, 2, 184, 70, .Directory, "Directory"	 'ControlID 1 
	PUSHBUTTON  8, 15, 50, 14, .Open1, "Open"	 'ControlID 2  
	TEXT  67, 17, 115, 14, .openDir, " "	 'ControlID 3
	TEXT  8, 33, 37, 14, .FileType, "File type:"	 'ControlID 4
	DDCOMBOBOX  50, 32, 130, 70, .FileTypeArr	 'ControlID 5 
	TEXT  8, 50, 37, 14, .FileName, "File name:"	 'ControlID 6
	TEXTBOX  49, 50, 132, 14, .FileNameTxt	 'ControlID 7
	GROUPBOX  2, 76, 184, 126, .PageLayout, "Page Layout"	 'ControlID 8 
	DDLISTBOX  103, 87, 77, 70, .WidthOptionsArr	 'ControlID 9 
	TEXT  8, 86, 37, 14, .Width, "Width:"	 'ControlID 10
	SPINCONTROL  46, 86, 48, 14, .SpinControlw	 'ControlID 11 
	TEXT  8, 103, 37, 14, .Height, "Height:"	 'ControlID 12 
	SPINCONTROL  46, 103, 48, 14, .SpinControlh	 'ControlID 13
	TEXT  8, 120, 37, 14, .Columns, "Columns:"	 'ControlID 14
	SPINCONTROL  46, 120, 48, 14, .SpinControlc	 'ControlID 15 
	TEXT  102, 120, 23, 14, .Rows, "Rows:"	 'ControlID 16
	SPINCONTROL  132, 120, 48, 14, .SpinControlo	 'ControlID 17 
	TEXT  8, 137, 37, 14, .Border, "Border:"	 'ControlID 18 
	SPINCONTROL  46, 137, 48, 14, .SpinControlb	 'ControlID 19
	TEXT  102, 137, 80, 14, .inches, ""	 'ControlID 20 
	TEXT  8, 154, 37, 14, .Resolution, "Resolution:"	 'ControlID 21
	SPINCONTROL  46, 154, 48, 14, .SpinControlr	 'ControlID 22 
	TEXT  102, 154, 37, 14, .dpi, "dpi"	 'ControlID 23 
	TEXT  8, 170, 37, 14, .Color, "Color:"	 'ControlID 24
	DDCOMBOBOX  46, 170, 86, 70, .ColorArr	 'ControlID 25 
	CHECKBOX  5, 188, 57, 10, .Flatten, "Flatten Image"	 'ControlID 26
	GROUPBOX  189, 2, 180, 129, .Group	 'ControlID27
	CHECKBOX  198, 1, 64, 10, .GenerateHTML, "Generate HTML"	 'ControlID 28 
	TEXT  194, 17, 43, 10, .Destination, "Destination:"	 'ControlID 29
	PUSHBUTTON  236, 15, 50, 14, .Open2, "Open"	 'ControlID 30
	TEXT  292, 17, 74, 14, .DesFile, " "	 'ControlID 31
	OPTIONGROUP .OptionGroup1
		OPTIONBUTTON  194, 37, 31, 9, .GIF, "GIF"	 'ControlID 33
		OPTIONBUTTON  240, 36, 31, 10, .JPEG, "JPEG"	 'ControlID 34 
		OPTIONBUTTON  292, 36, 31, 10, .PNG, "PNG"	 'ControlID 35
	CHECKBOX  194, 55, 105, 11, .LinkOriginalToTumbnails, "Link Originals to Thumbnails"	 'ControlID 36
	TEXT  194, 72, 30, 10, .URL, "URL"	 'ControlID 37	
	TEXTBOX  213, 69, 151, 15, .URLTxt	 'ControlID 38
	TEXT  194, 90, 80, 11, .InfoUnderThumbnail, "Info Under Thumbnail: "	 'ControlID 39'
	CHECKBOX  194, 103, 50, 10, .NameFile, "File Name"	 'ControlID 40 
	CHECKBOX  281, 103, 50, 10, .FileDate, "File Date"	 'ControlID 41 
	CHECKBOX  194, 116, 50, 10, .FileSize, "File Size"	 'ControlID 42
	CHECKBOX  281, 116, 77, 10, .Dimensions, "Dimensions (W * H)"	 'ControlID 43 
	GROUPBOX  189, 135, 180, 67, .Thumbnails, "Thumbnails"	 'ControlID 44
	TEXT  194, 147, 21, 11, .Order, "Order:"	 'ControlID 45 
	OPTIONGROUP .OptionGroup2
		OPTIONBUTTON  229, 147, 37, 10, .Across, "Across"	 'ControlID 47 
		OPTIONBUTTON  285, 147, 37, 10, .Down, "Down"	 'ControlID 48
	TEXT  194, 160, 19, 11, .Size, "Size:"	 'ControlID 49
	TEXT  194, 174, 170, 11, .SizeW, "Width:"	 'ControlID 50
	TEXT  194, 188, 170, 11, .SizeH, "Height:"	 'ControlID 51
	OKBUTTON  209, 206, 50, 14, .OK	 'ControlID 52
	CANCELBUTTON  265, 206, 50, 14, .Cancel	 'ControlID 53 
	PUSHBUTTON  319, 206, 50, 14, .Help, "Help"	 'ControlID 54 
END DIALOG

ret1 = dialog(sheet) 
SUB CSheet(BYVAL ControlID%, BYVAL Event%)
Dim OptionStr as string
Dim ThumbWidth as double
Dim ThumbHeight as double
Dim pageBorder as double
Dim tempb1 as Boolean
Dim tempb2 as Boolean
Dim tempb3 as Boolean
Dim tempw1 as Boolean
Dim temph1 as Boolean
dim width as double
dim height as double
dim Border as double
dim returnStr as string
dim returnStrDes as string
IF Event=0 THEN
	' this part is for test only
		
	WITHOBJECT "CorelPhotoPaint.Automation.9" 
		gDirectoryDes=.getPhotopaintdir()
		gDirectory=.getPhotopaintdir()
	END WITHOBJECT
	returnStr = GetDirectory (gDirectory ,30)
	returnStrDes = GetDirectory (gDirectoryDes, 20)
	Sheet.openDir.SETTEXT  returnStr
	Sheet.DesFile.SETTEXT  returnStrDes
	Sheet.FileTypeArr.SETARRAY FileTypeArr$
	Sheet.FileTypeArr.SETSELECT FileTypeArr(1)
	Sheet.inches.SETTEXT "Pixels" 
	Sheet.WidthOptionsArr.SETARRAY WidthOptionsArr$
	Sheet.WidthOptionsArr.SETSELECT 6 'Pixels is chosen at the begining
	Sheet.WidthOptionsArr.ENABLE TRUE
	Sheet.ColorArr.SETARRAY ColorArr$
	Sheet.ColorArr.SETSELECT ColorArr(4) 
	gMeasurment = Sheet.WidthOptionsArr.GETSELECT()
	Sheet.SpinControlw.SETVALUE 640
	Sheet.SpinControlh.SETVALUE 480
	Sheet.SpinControlb.SETVALUE 30
	Sheet.SpinControlr.SETVALUE 96
	Sheet.SpinControlc.SETVALUE 5
	Sheet.SpinControlo.SETVALUE 5
	SetRange(gMeasurment)
	Sheet.SpinControlc.SETDOUBLEMODE False
	Sheet.SpinControlc.SETPRECISION 0
	Sheet.SpinControlo.SETDOUBLEMODE False
	Sheet.SpinControlo.SETPRECISION 0
	Sheet.SpinControlw.SETMINRANGE 1
	Sheet.SpinControlb.SETMINRANGE 1
	Sheet.SpinControlh.SETMINRANGE 1
	Sheet.SpinControlc.SETMINRANGE 1
	Sheet.SpinControlc.SETMAXRANGE 20
	Sheet.SpinControlo.SETMINRANGE 1
	Sheet.SpinControlo.SETMAXRANGE 20
	
	
	Sheet.Flatten.SETTHREESTATE FALSE
	Sheet.Flatten.SETVALUE 0
	Sheet.LinkOriginalToTumbnails.SETTHREESTATE FALSE
	Sheet.LinkOriginalToTumbnails.SETVALUE 0
	Sheet.GenerateHTML.SETTHREESTATE FALSE
	Sheet.GenerateHTML.SETVALUE 0
	Sheet.NameFile.SETTHREESTATE FALSE
	Sheet.NameFile.SETVALUE 0
	Sheet.FileDate.SETTHREESTATE FALSE
	Sheet.FileDate.SETVALUE 0
	Sheet.Dimensions.SETTHREESTATE FALSE
	Sheet.Dimensions.SETVALUE 0
	Sheet.FileSize.SETTHREESTATE FALSE
	Sheet.FileSize.SETVALUE 0
	Sheet.gif.SETVALUE 1
	Sheet.Across.SETVALUE 1
	GWidth=Sheet.SpinControlw.GETVALUE ()
	Gborder=Sheet.SpinControlb.GETVALUE ()
	GHeight= Sheet.SpinControlh.GETVALUE () 		
	GResolution= Sheet.SpinControlr.GETVALUE () 
	Sheet.Sizew.SetText "Width: 92 Pixels"
	Sheet.SizeH.SetText "Height: 60 Pixels"	
	Sheet.GenerateHTML.ENABLE TRUE
	Sheet.GIF.ENABLE FALSE
	Sheet.JPEG.ENABLE FALSE
	Sheet.PNG.ENABLE FALSE
	Sheet.LinkOriginalToTumbnails.ENABLE FALSE
	Sheet.Open2.ENABLE False
	Sheet.NameFile.ENABLE FALSE
	Sheet.URLTxt.ENABLE FALSE
	Sheet.FileDate.ENABLE FALSE
	Sheet.FileSize.ENABLE FALSE
	Sheet.Dimensions.ENABLE FALSE
	
ENDIF

	IF Event = 1 THEN
		Select Case ControlID
			Case 11, 15
				OptionStr= MeasurmentOption(gMeasurment)
				Width = Sheet.SpinControlw.GETVALUE ()
				dRes# = (GResolution * (100000/2.54))+ 0.5
				GWidth = ConvertToPixels( Width, dRes#, gMeasurment )				
				tempw1 =GetThumbW(ThumbWidth)
				if  tempw1 = false then
					retval% = MESSAGEBOX("Thumbnail width is too small either increase page width or decrease number of colums!", "WARNING", 0 OR 48)
					exit
				endif
				Result1 = PutThumbWidthString()
				Sheet.Sizew.SETTEXT Result1					
			
			Case 13 , 17
				OptionStr= MeasurmentOption(gMeasurment)
				Height = Sheet.SpinControlh.GETVALUE ()
				dRes# = (GResolution * (100000/2.54))+ 0.5
				GHeight= ConvertToPixels( Height, dRes#, gMeasurment )				
				temph1 =GetThumbH(ThumbHeight)
				if  temph1 = false then
					retval% = MESSAGEBOX("Thumbnail height is too small either increase page height or decrease number of rows!", "WARNING", 0 OR 48)
					exit
				endif
				Result2 = PutThumbHeightString()
				Sheet.Sizeh.SETTEXT Result2	
			Case 19
				OptionStr= MeasurmentOption(gMeasurment)
				Border= Sheet.SpinControlb.GETVALUE () 	
				dRes# = (GResolution * (100000/2.54))+ 0.5
				GBorder= ConvertToPixels( Border, dRes#, gMeasurment )	
				tempb1 =GetThumbW(ThumbWidth)
				tempb2 = GetThumbH( ThumbHeight ) 
				tempb3= GetPageBorder(pageBorder)
				if  tempb1 = false then
				retval% = MESSAGEBOX("The page border is too large", "WARNING", 0 OR 48)
				exit
				endif
				if  tempb2 = false then
				retval% = MESSAGEBOX("The page border is too large", "WARNING", 0 OR 48)
				exit
				endif
				Result1 = PutThumbWidthString()
				Result2 = PutThumbHeightString()
				Sheet.Sizew.SETTEXT Result1
				Sheet.Sizeh.SETTEXT Result2
			Case 22
				GResolution= Sheet.SpinControlr.GETVALUE () 
				temp1 =GetThumbW(ThumbWidth)
				temp2 = GetThumbH( ThumbHeight ) 
				temp3= GetPageBorder(pageBorder)
				Result1 = PutThumbWidthString()
				Result2 = PutThumbHeightString()
				Sheet.Sizew.SETTEXT Result1
				Sheet.Sizeh.SETTEXT Result2								
				
		End Select
	End if

	IF Event = 2 THEN
		Select Case ControlID
			
			Case 2
				gDirectory = GETFOLDER(gDirectory)
				gtextDir	 = GetDirectory(gDirectory, 30)
				Sheet.openDir.SETTEXT  gtextDir
				tempDir$ = Right ( gtextDir,1)
				if tempDir$<> "\" Then
					gtextDir =gtextDir & "\"
				End if	
			
			case 5
				ext$=Sheet.FileTypeArr.GETSELECT()	
			
			Case 7
				Sheet.FileNameTxt.SETTEXT " "
						
			Case 9
				dRes# = (GResolution * (100000/2.54))+ 0.5
				gMeasurment = Sheet.WidthOptionsArr.GETSELECT()
				SetRange(gMeasurment)
				Width = ConvertToUnit( GWidth, dRes#, gMeasurment, 6 )
				convW =SetDecimalPoints(gMeasurment,Width)
				Height = ConvertToUnit( GHeight, dRes#, gMeasurment,6)
				convH =SetDecimalPoints(gMeasurment,Height)
				Sheet.SpinControlw.SETVALUE convW
				Sheet.SpinControlh.SETVALUE convH
				Border = ConvertToUnit( GBorder, dRes#, gMeasurment, 6 )
				convB =SetDecimalPoints(gMeasurment,Border)
				Sheet.SpinControlb.SETVALUE convB
				OptionStr= MeasurmentOption(gMeasurment)
				GetThumbW(ThumbWidth )
				GetThumbH(ThumbHeight )
				Result1 = PutThumbWidthString()
				Result2 = PutThumbHeightString()
				Sheet.Sizew.SETTEXT Result1
				Sheet.Sizeh.SETTEXT Result2
				Sheet.inches.SETTEXT OptionStr
				 	
			case 25
				newColor$= Sheet.ColorArr.GETSELECT()	
			
			case 26
				CheckVal1 =  Sheet.Flatten.GETVALUE()
				if CheckVal1 = 0 THEN
					Sheet.Flatten.SETVALUE 0
				EndIF
				if CheckVal1 = 1 THEN
					 Sheet.Flatten.SETVALUE 1	
				Endif
				
			case 28
				WantHTML = CreateHTML()
				if WantHTML Then
					Sheet.Down.ENABLE FALSE
					Sheet.LinkOriginalToTumbnails.SETVALUE 0
					Sheet.GenerateHTML.ENABLE True
					Sheet.GIF.ENABLE TRUE
					Sheet.JPEG.ENABLE TRUE
					Sheet.PNG.ENABLE TRUE
					Sheet.LinkOriginalToTumbnails.ENABLE TRUE
					Sheet.NameFile.ENABLE TRUE
					Sheet.URLTxt.ENABLE TRUE
					Sheet.FileDate.ENABLE TRUE
					Sheet.FileSize.ENABLE TRUE
					Sheet.Dimensions.ENABLE TRUE
					Sheet.Open2.ENABLE True
				Else
					Sheet.Down.ENABLE True
					Sheet.GIF.ENABLE FALSE
					Sheet.JPEG.ENABLE FALSE
					Sheet.PNG.ENABLE FALSE
					Sheet.LinkOriginalToTumbnails.ENABLE FALSE
					Sheet.NameFile.ENABLE FALSE
					Sheet.URLTxt.ENABLE FALSE
					Sheet.FileDate.ENABLE FALSE
					Sheet.FileSize.ENABLE FALSE
					Sheet.Dimensions.ENABLE FALSE
					Sheet.Open2.ENABLE False
				End if
					
			
			Case 30
				gDirectoryDes = GETFOLDER(gDirectoryDes)
				gtextdest = GetDirectory(gDirectoryDes, 20)
				Sheet.DesFile.SETTEXT  gtextdest
				if tempDest$<> "\" Then
					gtextdest =gtextdest & "\"
				End if	
				
			case 33 
				CheckValGif% =  Sheet.gif.GETVALUE()
				if CheckValGif% = 0 THEN
					Sheet.gif.SETVALUE 0
					Sheet.png.SETVALUE 0
					Sheet.jpeg.SETVALUE 1
				EndIF
				if CheckValGif% = 1 THEN
				 	Sheet.gif.SETVALUE 1
					Sheet.png.SETVALUE 0
					Sheet.jpeg.SETVALUE 0
				EndIF
			
			case 34 
				CheckValjpeg% =  Sheet.jpeg.GETVALUE()
				if CheckValjpeg% = 0 THEN
					Sheet.gif.SETVALUE 1
					Sheet.png.SETVALUE 0
					Sheet.jpeg.SETVALUE 0
				EndIF
				if CheckValjpeg% = 1 THEN
				 	Sheet.gif.SETVALUE 0
					Sheet.png.SETVALUE 0
					Sheet.jpeg.SETVALUE 1
				EndIF
		
							
			case 36
				CheckVal1 =  Sheet.LinkOriginalToTumbnails.GETVALUE()
				if CheckVal1 = 0 THEN
					Sheet.LinkOriginalToTumbnails.SETVALUE 0
				ENDIF
				if CheckVal1 = 1 THEN
					
				EndIF
	
			Case 38
				Sheet.URLTxt.SETTEXT " "
				
	
			case 40
				CheckValname%=  Sheet.NameFile.GETVALUE()
				if CheckValname% = 0 THEN
					Sheet.NameFile.SETVALUE 0
				ENDIF
				if CheckValname% = 1 THEN
					Sheet.NameFile.SETVALUE 1
				EndIF
		
			case 41
				CheckValdate% =  Sheet.FileDate.GETVALUE()
				if CheckValdate% = 0 THEN
					Sheet.FileDate.SETVALUE 0
				ENDIF
				if CheckValdate% = 1 THEN
					 Sheet.FileDate.SETVALUE 1
				EndIF
				
			case 42
				CheckValSize% =  Sheet.FileSize.GETVALUE()
				if CheckValSize% = 0 THEN
					Sheet.FileSize.SETVALUE 0
				EndIF
				if CheckValSize% = 1 THEN
				 	Sheet.FileSize.SETVALUE 1
				EndIF
			
			case 43
				CheckValDim% =  Sheet.Dimensions.GETVALUE()
				if CheckValDim% = 0 THEN
					Sheet.Dimensions.SETVALUE 0
				ENDIF
				if CheckValDim% = 1 THEN
					Sheet.Dimensions.SETVALUE 1
				EndIF
			case 47
				CheckValAcross% =  Sheet.Across.GETVALUE()
				if CheckValAcross% = 0 THEN
					Sheet.Across.SETVALUE 0
					Sheet.Down.SETVALUE 1
					EndIF
				if CheckValAcross% = 1 THEN
				 	Sheet.Across.SETVALUE 1
					Sheet.Down.SETVALUE 0
				EndIF
			
			case 48	
			
				WantHTML = CreateHTML()
				if WantHTML Then
					Sheet.Down.ENABLE FALSE
				Else
					Sheet.Down.ENABLE True
				End if
				CheckVal1 =  Sheet.Down.GETVALUE()
				if CheckVal1 = 0 THEN
					Sheet.Across.SETVALUE 1
					Sheet.Down.SETVALUE 0
					EndIF
				if CheckVal1 = 1 THEN
				 	Sheet.Across.SETVALUE 0
					Sheet.Down.SETVALUE 1
				EndIF
		End Select
	End if				
End Sub 

if ret1=2 then	stop
BEGIN DIALOG OBJECT Build 303, 34, "Build Contact Sheet", SUB CBuild
	TEXT  3, 5, 296, 12, .Copying, "Copying..."	 'control ID 1
	PROGRESS 2, 18, 267, 11, .Progress1	 'control ID 2
	TEXT  272, 17, 27, 12, .Text2, " "	 'control ID 3
END DIALOG

if ret1 = 1 Then
	ret2 = dialog(Build)
'if ret2 = 2 then stop

SUB CBuild(BYVAL ControlID%, BYVAL Event%)

	IF Event=0 THEN
		if ret2 = 2 then stop
		Build.Progress1.SETMINRANGE 0
		Build.Progress1.SETMAXRANGE 100
		CALL SubBuild
		End if	' closing of if ret1=1
	END IF
	
End Sub

Sub SubBuild()
	WITHOBJECT "CorelPhotoPaint.Automation.9" 
		
		DIM FCOUNT%, NFILES% 
		DIM FILESARR(150) AS STRING
		DIM nItems%
		DIM CheckVal1%,CheckVal2%,CheckVal3%   'temporary varialble for the destination ext
		DIM CheckOrder%                        'temporary varialbe for the order of input image
		DIM nPages%,comp1#,comp#
		DIM extD as string
		DIM ext as string
		DIM COMPNAME as String
		DIM retname as DATE	
		DIM PicThumbH#, PicThumbW# 
		DIM PixFortxt#,Border#,newBorder#
		DIM info as String
		DIM txtXCoor#,	txtYCoor#,MakeFlat%
		DIM Flat As BOOLEAN
		DIM WantedText AS BOOLEAN
		Dim Result1$, Result2$
		Dim TotalSize as Long
		Dim DocSize as long
		Dim ImageSize as long
		Dim PercentComp as integer
		DIM WantHTML as boolean
		Dim WantName as Boolean
		Dim WantDate as Boolean
		Dim WantSize as Boolean
		Dim WantDim as Boolean
		Dim WantLink as Boolean	
		DIM OptionStr as string
		DIM Width as double
		Dim Height as double
		DIM ThumbWidth as double
		DIM ThumbHeight as double
		DIM temp# 
		DIM NoExt as string
		
		ON ERROR RESUME NEXT 'GOTO MessageHandler
				
		extChosen$=Sheet.FileTypeArr.GETSELECT()
		select case extChosen$
			case "All Files"
				ext = "*.*"
			case "BMP"
				ext = "*.bmp"
			case "CPT"
				ext = "*.cpt"
			case  "GIF"
				ext = "*.gif"
			case "JPG"
				ext = "*.jpg"
			case	"PCD"
				ext = "*.pcd"
			case "PCX"
				ext = "*.pcx"
			case  "RIFF"
				ext = "*.riff"
			case "TGA"
				ext = "*.tga"
			case "TIF"
				ext = "*.tif"
			case else
				ext = extChosen$
		End Select
		
		OptionStr = MeasurmentOption(gMeasurment)
		dRes# = (GResolution * (100000/2.54))+ 0.5
		gMeasurment = Sheet.WidthOptionsArr.GETSELECT()
		Width = ConvertToUnit( GWidth, dRes#, gMeasurment, 6 )
		convW =SetDecimalPoints(gMeasurment,Width)
		Border = ConvertToUnit( Gborder, dRes#, gMeasurment, 6 )
		convB =SetDecimalPoints(gMeasurment,Border)
		Height = ConvertToUnit( GHeight, dRes#, gMeasurment,6)
		convH =SetDecimalPoints(gMeasurment,Height)
		Sheet.SpinControlw.SETVALUE convW
		Sheet.SpinControlh.SETVALUE convH
		Sheet.SpinControlh.SETVALUE convB
		OptionStr= MeasurmentOption(gMeasurment)
		GetThumbW(ThumbWidth )
		GetThumbH(ThumbHeight )
		Result1 = PutThumbWidthString()
		Result2 = PutThumbHeightString()
		Sheet.Sizew.SETTEXT Result1
		Sheet.Sizeh.SETTEXT Result2
		Sheet.inches.SETTEXT OptionStr
		nColumn=Sheet.SpinControlc.GETVALUE ()	
		nRow=Sheet.SpinControlo.GETVALUE ()
		ThumbWidth = ConvertToUnit( ThumbWidth, dRes#, 6,gMeasurment)
		ThumbHeight = ConvertToUnit( ThumbHeight#, dRes#,6, gMeasurment)
		Width = ConvertToUnit( gWidth, dRes#, 6,gMeasurment)
		Height = ConvertToUnit( gHeight, dRes#,6, gMeasurment)
		Color$= Sheet.ColorArr.GETSELECT()
		select case Color$
			case ColorArr(1)
				colorId=3
			case ColorArr(2)
				colorId=2
			case ColorArr(3)
				colorId=4
			case ColorArr(4)
				colorId=1
			case ColorArr(5)
				colorId=7
			case ColorArr(6)
				colorId=5
		End Select
	
		CheckVal1% = Sheet.optionGroup1.GETVALUE()
		if CheckVal1%=0 THEN
			extD$= "GIF"
			id%= 773 
		ENDIF	
		if CheckVal1%=2 THEN
			extD$= "PNG"
			id%= 802 
		ENDIF	
		if CheckVal1%=1 THEN
			extD$= "JPEG"
			id%= 774 
		ENDIF	
		
		CheckOrder% =  Sheet.OptionGroup2.GETVALUE()
		IF CheckOrder%= 0 THEN
			Filling$= "Across"
		ELSE
			Filling$= "Down"
		END IF
		
		WantedText= PutTexT()
		MakeFlat=sheet.Flatten.GetValue()
		if MakeFlat=1 then
			Flat=true
		Endif 
		OpenDir$ = gDirectory
		OpenDes$ = gDirectoryDes
		DirLen%= LEN (OpenDir$)
		DesLen%= LEN (OpenDes$)
		if DirLen% > 4 then
			OpenDir$ = OpenDir$ &"\"
		End if	
		if DesLen% > 4 then
			openDes$ = openDes$ &"\"
		End if	
		FCOUNT = 1
		TotalSize = 0
		
		FILESARR(FCOUNT) = FINDFIRSTFOLDER(OpenDir$ + "\" + ext$, 1 OR 2 OR 4 OR 32 OR 128)
		DocName$ = OpenDir$+ FILESARR(FCOUNT)
		TotalSize = Filesize (DocName$)
	
		WHILE (FILESARR(FCOUNT) <> "" )
			FCOUNT = FCOUNT + 1
			FILESARR(FCOUNT) = FINDNEXTFOLDER()
			DocName$ = OpenDir$+ FILESARR(FCOUNT)
			DocSize = Filesize (DocName$)
			TotalSize = TotalSize + DocSize
		WEND
		nItems = nColumn * nRow
		temp=(FCOUNT/nItems)
		IF INT(temp#)< temp# THEN
			nPages% = INT (temp#)+1
		ELSE 
			nPages% = INT (temp#)
		EndIF
		IF ret = 2 THEN STOP 	' If Cancel is selected, stop the script	
		NFILES = FCOUNT-1
		I=1
		PercentComp = 0
		WantHTML = CreateHTML()
		WantName = putName()
		WantDate = PutDate()
		WantSize = PutSize() 
		WantLink = PutLink()
		WantDim = PutDim()
		if WantHTML Then
			Sheet.Down.ENABLE FALSE
		End if
	
		For L%=1 to nPages	
				if WantHTML Then
				call CreateHTMLText L ,npages ,openDes$	
			Else
				.FileNew gWidth ,gHeight, colorId, GResolution, GResolution, FALSE, FALSE, 0, 0, 0, 0, 0, 255, 255, 255, 0, FALSE
			End if
			xCoor#=0
			yCoor#=0
			txtXCoor=0
			txtYCoor=0
			PixFortxt=((gResolution/72)*12)*3
			IF Filling$="Across" THEN
				yCoor#=gBorder+ThumbHeight
				txtYCoor#=yCoor#+(gBorder-PixFortxt)/2
				pageName$=DesDirectory & "page" & Str (L)& ".html" 
				if WantHTML Then
					PRINT #L, "<tr>"
				End if
				For Row%=1 to nRow
					if  I > NFiles THEN EXIT FOR
					xCoor#=gBorder
					for Column%=1 to nColumn
						if  I > NFiles THEN EXIT FOR
							NoExt = RemoveExt(FILESARR(I))
							COMPNAME = OpenDir$ & FILESARR(I)					
							savedto$ = OpenDes$ & FILESARR(I)
							Copyto$ = gtextdest & NoExt & "."& extD$
							CopyFrom$ = gtextDir & FILESARR(I)
							if WantHTML Then
								Build.Copying.SetText "Copyin From:  " + CopyFrom$ + "  To  " + Copyto$
							Else
								Directoryname$= Sheet.FileNametxt.GETTEXT()
								newName$ = OpenDir$& Directoryname$ &"Page" & str(L)& ".cpt" 
								Copyto$ = gtextDir & Directoryname$ &"Page" & str(L)& ".cpt" 
								CopyFrom$ = gtextDir & FILESARR(I)
								Build.Copying.SetText "Copyin From:  " + CopyFrom$ + "  To  " + Copyto$ 
							End if
							.FileOpen  COMPNAME, 0,0,0,0,0,1,1
							txtXCoor = xCoor#
							name$ = .GetDocumentName()
							if WantHTML Then
								gnDocWidth = .GetDocumentWidth()
								gnDocHeight = .GetDocumentHeight()
								ReturnValueResX& = .GetDocumentXdpi()
								ReturnValueResY& = .GetDocumentYdpi()
								.ImageResample gnDocWidth,gnDocHeight , ReturnValueResX&, ReturnValueResY&, TRUE
								
								Lr_nameSaved$=OpenDes$& "Lr_" & NoExt & "." & extD$
								Lr_image$ = "Lr_" & NoExt & "." & extD$
								Lr_Name$ = FILESARR(I)
								if id% = 773 Then 
									' **** This part of the code is to covert GIF to Paletted
									Call CovertToPaletted
								End if	
								.filesave Lr_nameSaved$,id%,0
							END IF	
							ImageName$= FILESARR(I)
							'**** this is the case when want information under each thumbnail at .cpt file
								'if WantHTML = False then
								'	info=""
								'	if WantName Then
								'		info= info & ImageName$
								'	endif
								'	if WantDim Then
								'		info = info & "WxH" & STR (gnDocWidth)& STR(gnDocHeight)
								'	end if
								'	if WantSize then
								'		info = info & CHR(13)& retsize$
								'	end if
								'	if WantDate Then
								'		info = info & retdate$
								'	end if
								'End if	
							ImageSize = Filesize (name$)
							retsize$ = STR (Filesize (name$))
							retdate$ = STR (FileDate (name$))
							gnDocWidth = .GetDocumentWidth()
							gnDocHeight = .GetDocumentHeight()
							if gnDocWidth > gnDocHeight then
								PicThumbH = gnDocHeight* (ThumbWidth / gnDocWidth)
								PicThumbW = ThumbWidth
							else
								PicThumbW = gnDocWidth* (ThumbHeight / gnDocHeight)
								PicThumbH = ThumbHeight
							End if
							.ImageResample PicThumbW,PicThumbH , GResolution, GResolution, TRUE
							if WantHTML then
								Tn_nameSaved$=OpenDes$& "tn_" & NoExt &"."& extD$
								Tn_image$= "tn_" & NoExt &"."& extD$
								if id% = 773 Then
									' **** This code part of the code is to covert GIF to Paletted
									Call CovertToPaletted
								End if	
								.filesave Tn_nameSaved$,id%,0
							End if
							.EditCopy
							Directoryname$= Sheet.FileNametxt.GETTEXT()
							'**** this is the case when want information under each thumbnail at .cpt file
								'if WantHTML= False then
								'	if (WantedText) then
								'		.TextTool txtXCoor,txtYCoor, FALSE, TRUE, 0 
								'		.TextSetting "Fill", "0,0,0"
								'		.TextSetting "Font", "Arial"
								'		.TextSetting "TypeSize", "12.0"
								'		.TextAppend info
								'		.TextRender 
								'	EndIf	
								'End if
							.FileClose
							IncPercent%= Int((ImageSize /TotalSize)* 100)
							PercentComp =  PercentComp + IncPercent%
							Build.Progress1.setIncrement IncPercent%
							Build.Progress1.step 
							ImageComp% = PercentComp	
							if ret2 = 2 then stop
							Build.Text2.settext (STR (ImageComp%)+ " %")
							if WantHTML then
								PRINT #L, "<td align=" + chr(34) + "center" + chr(34) + ">"
								if WantLink then
									Print #L, "<A HREF= " + chr(34)+ Lr_image$ +chr(34) + "</A>"
								End if
								PRINT #L, "<IMG SRC=" + chr(34)+ Tn_image$ + chr(34) + "border="+ chr(34) + "0" + chr(34) 
								PRINT #L, " alt="+  chr(34)+ Tn_image$ +chr(34)+ " lowsrc=" + chr(34)+ Tn_image$ + chr(34) 
								PRINT #L, " width="+ chr(34) + Str (PicThumbW) + chr(34)+  "height=" + chr(34) + Str (PicThumbH) + chr(34) + ">" 
								if WantName and WantLink Then
									Print #L, "<BR>"+"<A HREF=" + chr(34)+ Lr_image$ + chr(34)+ ">"+ Lr_Name$ +"</A>"
								End if
								if WantName and WantLink=False then
									Print #L,"<BR>" & Lr_Name$ 
									
								End if
								
								if WantSize Then
									Print #L, "</A>"+"<BR>" + retsize$ + " bytes"
								end if
								if WantDate Then
									 Print #L, "</A>"+"<BR>" + retdate$
								End if
								
								if WantDim then
									Print #L,"</A>"+"<BR>" + Str(gnDocWidth) + " x" + STR (gnDocHeight)
								end if
								PRINT #L, "</td>"
							Else
								.BindToActiveDocument
								.EditPasteObject xCoor#, yCoor#, ""
								.objectName I,FILESARR(I)
							End if	
							xCoor#=xCoor#+gBorder+ThumbWidth
							txtXCoor=txtXCoor+gBorder+ThumbWidth
							I=I+1
					NEXT Column%
					yCoor#= yCoor#+gBorder+ ThumbHeight
					txtYCoor#=txtYCoor#+gBorder+ ThumbHeight
					if WantHTML then
						PRINT #L, "</tr>"
					End if
				NEXT Row%
			END IF
			IF Filling$="Down" THEN
				xCoor#= newBorder	
				txtXCoor#=xCoor#+(Gorder-PixFortxt)/2
				for Column%=1 to nColumn
					if  I > NFiles THEN EXIT FOR
					yCoor#= GBorder + ThumbHeight
					For Row%=1 to nRow
						if  I > NFiles THEN EXIT FOR
							COMPNAME = OpenDir$ & FILESARR(I)					
							savedto$ = OpenDes$ & FILESARR(I)
							Directoryname$= Sheet.FileNametxt.GETTEXT()
							newName$ = OpenDir$& Directoryname$ &"Page" & str(L)& ".cpt" 
							Copyto$ = gtextDir & Directoryname$ &"Page" & str(L)& ".cpt" 
							CopyFrom$ = gtextDir & FILESARR(I)
							Build.Copying.SetText "Copyin From:  " + CopyFrom$  + "  To  " +Copyto
							.FileOpen  COMPNAME, 0,0,0,0,0,1,1
							txtYCoor = yCoor#
							name$ = .GetDocumentName()
							'**** this is the case when want information under each thumbnail at .cpt file
								'info=""
								'if wantName Then
								'	info= info & name$
								'endif
								'if wantDim Then
								'	info = info & "WxH" & STR (gnDocWidth)& STR(gnDocHeight)
								'end if
								'if wantSize then
								'	info = info & CHR(13)& retsize$
								'end if
								'if wantDate Then
								'	info = info & retdate$
								'end if	
						ImageSize = Filesize (name$)
						retsize$ = STR (Filesize (name$))
						retdate$ = STR (FileDate (name$))
						gnDocWidth = .GetDocumentWidth()
						gnDocHeight = .GetDocumentHeight()
						if gnDocWidth > gnDocHeight then
							PicThumbH = gnDocHeight* (ThumbWidth / gnDocWidth)
							PicThumbW = ThumbWidth
						else
							PicThumbW = gnDocWidth* (ThumbHeight / gnDocHeight)
							PicThumbH = ThumbHeight
						End if
						.ImageResample PicThumbW,PicThumbH , GResolution, GResolution, TRUE
						.EditCopy
									
						Directoryname$= Sheet.FileNametxt.GETTEXT()
						.FileClose
						'**** this is the case when want information under each thumbnail at .cpt file
							'if (WantedText) then
							'		.TextTool txtXCoor,txtYCoor, FALSE, TRUE, 0 
							'		.TextSetting "Fill", "0,0,0"
							'		.TextSetting "Font", "Arial"
							'		.TextSetting "TypeSize", "12.0"
							'		.TextAppend info	
							'		.TextRender 
							'EndIf	
						PercentComp =  PercentComp + Int((ImageSize /TotalSize)* 100)
						Build.Progress1.step 
						ImageComp% = PercentComp	
						Build.Text2.settext (STR (ImageComp%)+ " %")
						if ret2 = 2 then stop
						.BindToActiveDocument
						.EditPasteObject xCoor#, yCoor#, ""
						.objectName I,FILESARR(I)
						yCoor#=yCoor#+GBorder+ThumbHeight
						txtYCoor#=txtYCoor#+GBorder+ ThumbHeight
						I=I+1					
					NEXT Row%
					xCoor#= xCoor#+GBorder+ ThumbWidth
					txtXCoor=txtXCoor+GBorder+ThumbWidth
				NEXT Column%
			END IF
			if flat then
			.ObjectMerge -1
				.EndObject
			Endif
			if WantHTML then
				PRINT #L, "</table><I>This thumbnail page was generated by <A HREF=" + chr(34)+ "http://www.Photopaint.com" + chr(34)+ ">Corel Photo Paint 9</A></I>"
				PRINT #L, "<p><i>Page" & Str(L) & "of" & Str (nPages) &"</i></p>"
				PRINT #L, "</body>"
				PRINT #L, "</html>"
				CLOSE #L
			Else
				newName$ = OpenDir$& Directoryname$ &"Page" & str(L)& ".cpt" 
				.fileSave newName$ , 1808, 0
			End if	
			if WantHTML=false then
			.FileClose
			Endif
			
		NEXT L%
		if percentComp <> 100 then
			Build.Progress1.setvalue 100
			Build.Text2.settext ("100 %")
		End if 
		ret2= 2
		Stop	
	END WITHOBJECT
	
	'MessageHandler:
	'retVal% = MESSAGEBOX("Script cancelled", "ERROR", 16)
	'Stop	
	
End Sub

REM Function definitions.
				
REM *SDOC*************************************************************
REM 
REM 	Name:	MeasurmentOption
REM 
REM 	Action:	
REM 
REM 	Params:	
REM			
REM 
REM 	Returns: String
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/			
			
Function MeasurmentOption(OptionArr as Long) as String
	Dim OptionChosen as String
	
	select case OptionArr
		case 1
			MeasurmentOption = "Inches"
		case 2
			MeasurmentOption = "millimeters" 
		case 3
			MeasurmentOption = "picas, points"  
		case 4
			MeasurmentOption = "points" 
		case 5
			MeasurmentOption = "centimeters"
		case 6
			MeasurmentOption = "Pixels"
		case 7
			MeasurmentOption = "ciceros, didots"
		case 8
			MeasurmentOption = "didots"
		END SELECT
'	MeasurmentOption = OptionChosen
End Function

REM *SDOC*************************************************************
REM 
REM 	Name:	ThumbSize
REM 
REM 	Action:	finds Thumbsize
REM 
REM 	Params:	two double	
REM			
REM 
REM 	Returns: double
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/

Function GetThumbW(Byref dWidth as double)as Boolean

	Dim newWidth 		as double
	Dim nColumn 		as integer
	Dim ThumbWidth1 	as double
	Dim ThumbWidth2 	as double
	Dim ThumbWidth 	as double
	Dim BorderW		as double
	Dim Width 		as double
	
	dRes# = (GResolution * (100000/2.54))+ 0.5
	dWidth = GWidth
	nColumn=Sheet.SpinControlc.GETVALUE ()	
	BorderW=(nColumn+1)*GBorder
	if BorderW >= dWidth Then
		GetThumbW = False
		exit
	Endif
	ThumbWidth1=((dWidth- BorderW)/nColumn)*10000
	ThumbWidth2=Int (ThumbWidth1)
	ThumbWidth=(ThumbWidth2)/10000
	Width = ConvertToUnit(ThumbWidth, dRes#, gMeasurment, 6 )
	Width =  SetDecimalPoints(gMeasurment,Width)
	dWidth = Width
	GetThumbW = true
End Function

REM *SDOC*************************************************************
REM 
REM 	Name:	ThumbSize
REM 
REM 	Action:	finds Thumbsize
REM 
REM 	Params:	two double	
REM			
REM 
REM 	Returns: double
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Function PutThumbWidthString()as string
	
	Dim Stringarr 		as string
	Dim Stringw 		as string
	Dim Result$ 		as string
	Dim ThumbWidth 	as double
		
	Stringarr= MeasurmentOption(gMeasurment)
	Stringw= "Width:"
	GetThumbW(ThumbWidth )
	Result$ =Stringw & " " & STR (ThumbWidth) & " " & Stringarr
	PutThumbWidthString = Result
End Function

REM *SDOC*************************************************************
REM 
REM 	Name:	GetThumbH
REM 
REM 	Action:	Gets Thumbnaild height
REM 
REM 	Params:	two double	
REM			
REM 
REM 	Returns: double
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Function GetThumbH( Byref dHeight as double)as Boolean
	Dim nRow	 		as integer
	Dim ThumbHeight1 	as double
	Dim ThumbHeight2 	as Long
	Dim ThumbHeight 	as double
	Dim BorderH		as double
	Dim Height		as double
	
	dRes# = (GResolution * (100000/2.54))+ 0.5
	dHeight = GHeight
	nRow=Sheet.SpinControlo.GETVALUE ()	
	BorderH=(nRow+1)*GBorder	
	if BorderH >= dHeight Then
		GetThumbH = False
		exit
	Endif
	ThumbHeight1=((dHeight-BorderH)/nRow)*10000
	ThumbHeight2=Int (ThumbHeight1)
	ThumbHeight=(ThumbHeight2)/10000
	Height = ConvertToUnit( ThumbHeight, dRes#, gMeasurment,6)
	Height =  SetDecimalPoints(gMeasurment ,Height)
	dHeight = Height
	GetThumbH = true
	
End Function

REM *SDOC*************************************************************
REM 
REM 	Name:	GetPageBorder
REM 
REM 	Action:	Gets Thumbnaild height
REM 
REM 	Params:	two double	
REM			
REM 
REM 	Returns: double
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Function GetPageBorder(Byref Border as double )as Boolean
	Dim nRow	 		as integer
	Dim ThumbHeight1 	as double
	Dim ThumbHeight2 	as Long
	Dim ThumbHeight 	as double
	Dim dBorder		as double
	Dim Height		as double
	
	dRes# = (GResolution * (100000/2.54))+ 0.5
	dBorder = ConvertToUnit( GBorder, dRes#, gMeasurment, 6)
	dBorder = SetDecimalPoints(gMeasurment ,dBorder)	
	Border = dBorder
	GetPageBorder =True
	End Function
	
REM *SDOC*************************************************************
REM 
REM 	Name:	PutThumbHeightString
REM 
REM 	Action:	puts thumbnails string
REM 
REM 	Params:	
REM			
REM 
REM 	Returns: String
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/

Function PutThumbHeightString()as string

	
	Dim StringArr 		as string
	Dim Stringh 		as string
	Dim Result$ 		as string
	Dim ThumbHeight	as double

	Stringarr= MeasurmentOption(gMeasurment)
	Stringh= "Height:"
	GetThumbH(ThumbHeight)
	Result$ =Stringh & " " & STR(ThumbHeight) & " " & StringArr
	PutThumbHeightString = Result$
	
End Function
REM *SDOC*************************************************************
REM 
REM 	Name:	SetRange
REM 
REM 	Action:	finds Thumbsize
REM 
REM 	Params:	two double	
REM			
REM 
REM 	Returns: double
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/

Sub SetRange(nValue as long)
	Select case nValue 
		case 1 '"Inches"
			Sheet.SpinControlh.SETMINRANGE 1
			Sheet.SpinControlh.SETMAXRANGE 100
			Sheet.SpinControlw.SETMINRANGE 1
			Sheet.SpinControlw.SETMAXRANGE 100
			Sheet.SpinControlb.SETMINRANGE 0
			Sheet.SpinControlb.SETMAXRANGE 1
			
			Sheet.SpinControlw.SETDOUBLEMODE TRUE
			Sheet.SpinControlw.SETPRECISION 4
			Sheet.SpinControlh.SETDOUBLEMODE  TRUE
			Sheet.SpinControlh.SETPRECISION 4
			Sheet.SpinControlb.SETDOUBLEMODE  TRUE
			Sheet.SpinControlb.SETPRECISION 4
			
			
									
		case 2 '"millimeters" 
			Sheet.SpinControlh.SETMINRANGE 100
			Sheet.SpinControlh.SETMAXRANGE 100000
			Sheet.SpinControlw.SETMINRANGE 100
			Sheet.SpinControlw.SETMAXRANGE 100000
			Sheet.SpinControlb.SETMINRANGE 0
			Sheet.SpinControlb.SETMAXRANGE 500
			
			Sheet.SpinControlw.SETDOUBLEMODE TRUE
			Sheet.SpinControlw.SETPRECISION 4
			Sheet.SpinControlh.SETDOUBLEMODE  TRUE
			Sheet.SpinControlh.SETPRECISION 4
			Sheet.SpinControlb.SETDOUBLEMODE  TRUE
			Sheet.SpinControlb.SETPRECISION 4
			
						
		case 3 '"picas, points" 
			Sheet.SpinControlh.SETMINRANGE 1000
			Sheet.SpinControlh.SETMAXRANGE 1000
			Sheet.SpinControlw.SETMINRANGE 1000
			Sheet.SpinControlw.SETMAXRANGE 1000
			Sheet.SpinControlb.SETMINRANGE 0
			Sheet.SpinControlb.SETMAXRANGE 50
			
			Sheet.SpinControlw.SETDOUBLEMODE TRUE
			Sheet.SpinControlw.SETPRECISION 4
			Sheet.SpinControlh.SETDOUBLEMODE  TRUE
			Sheet.SpinControlh.SETPRECISION 4
			Sheet.SpinControlb.SETDOUBLEMODE  TRUE
			Sheet.SpinControlb.SETPRECISION 4
						
		case 4 ' "points"
			Sheet.SpinControlh.SETMINRANGE 1000
			Sheet.SpinControlh.SETMAXRANGE 1000
			Sheet.SpinControlw.SETMAXRANGE 1000
			Sheet.SpinControlw.SETMINRANGE 1000
			Sheet.SpinControlb.SETMINRANGE 0
			Sheet.SpinControlb.SETMAXRANGE 50
			
			Sheet.SpinControlw.SETDOUBLEMODE TRUE
			Sheet.SpinControlw.SETPRECISION 4
			Sheet.SpinControlh.SETDOUBLEMODE  TRUE
			Sheet.SpinControlh.SETPRECISION 4
			Sheet.SpinControlb.SETDOUBLEMODE  TRUE
			Sheet.SpinControlb.SETPRECISION 4
			
		case 5 '"centimeters"
			Sheet.SpinControlh.SETMINRANGE 1
			Sheet.SpinControlh.SETMAXRANGE 1000
			Sheet.SpinControlw.SETMAXRANGE 1000
			Sheet.SpinControlw.SETMINRANGE 1
			Sheet.SpinControlb.SETMINRANGE 0
			Sheet.SpinControlb.SETMAXRANGE 50
			
			Sheet.SpinControlw.SETDOUBLEMODE TRUE
			Sheet.SpinControlw.SETPRECISION 4
			Sheet.SpinControlh.SETDOUBLEMODE  TRUE
			Sheet.SpinControlh.SETPRECISION 4
			Sheet.SpinControlb.SETDOUBLEMODE  TRUE
			Sheet.SpinControlb.SETPRECISION 4
			
		case 6 '"Pixels"
			Sheet.SpinControlh.SETMINRANGE 100
			Sheet.SpinControlh.SETMAXRANGE 10000
			Sheet.SpinControlw.SETMAXRANGE 10000
			Sheet.SpinControlw.SETMINRANGE 100
			Sheet.SpinControlb.SETMINRANGE 0
			Sheet.SpinControlb.SETMAXRANGE 100
			
			Sheet.SpinControlw.SETDOUBLEMODE False
			Sheet.SpinControlw.SETPRECISION 0
			Sheet.SpinControlh.SETDOUBLEMODE  False
			Sheet.SpinControlh.SETPRECISION 0
			Sheet.SpinControlb.SETDOUBLEMODE  False
			Sheet.SpinControlb.SETPRECISION 0
			
		case 7 '"ciceros, didots"
			Sheet.SpinControlh.SETMINRANGE 1
			Sheet.SpinControlh.SETMAXRANGE 1000
			Sheet.SpinControlw.SETMAXRANGE 1000
			Sheet.SpinControlw.SETMINRANGE 1
			Sheet.SpinControlb.SETMINRANGE 0
			Sheet.SpinControlb.SETMAXRANGE 50
			
			Sheet.SpinControlw.SETDOUBLEMODE TRUE
			Sheet.SpinControlw.SETPRECISION 4
			Sheet.SpinControlh.SETDOUBLEMODE  TRUE
			Sheet.SpinControlh.SETPRECISION 4
			Sheet.SpinControlb.SETDOUBLEMODE  TRUE
			Sheet.SpinControlb.SETPRECISION 4
			
		case 8  '"didots"
			Sheet.SpinControlh.SETMINRANGE 1
			Sheet.SpinControlh.SETMAXRANGE 1000
			Sheet.SpinControlw.SETMAXRANGE 1000
			Sheet.SpinControlw.SETMINRANGE 1
			Sheet.SpinControlb.SETMINRANGE 0
			Sheet.SpinControlb.SETMAXRANGE 50
			
			Sheet.SpinControlw.SETDOUBLEMODE TRUE
			Sheet.SpinControlw.SETPRECISION 4
			Sheet.SpinControlh.SETDOUBLEMODE  TRUE
			Sheet.SpinControlh.SETPRECISION 4
			Sheet.SpinControlb.SETDOUBLEMODE  TRUE
			Sheet.SpinControlb.SETPRECISION 4
	End Select
End sub 
REM *SDOC*************************************************************
REM 
REM 	Name:	SetRange
REM 
REM 	Action:	finds Thumbsize
REM 
REM 	Params:	two double	
REM			
REM 
REM 	Returns: double
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Function FourDecPoint(nValue as double) as double

	Dim temp as double
	
	temp = nValue * 10000
	temp = (int(temp))/10000
	
	FourDecPoint=temp
End Function	

REM *SDOC*************************************************************
REM 
REM 	Name:	SetRange
REM 
REM 	Action:	finds Thumbsize
REM 
REM 	Params:	two double	
REM			
REM 
REM 	Returns: double
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Function SetDecimalPoints(nValue as Long,  nSize as double)as double
	Dim temp as double
	if nValue <> 6 then
		temp = FourDecPoint(nSize)
		SetDecimalPoints = temp
		
	Else 
		temp = int (nSize)
		SetDecimalPoints = temp
		
	End if	
End Function

REM *SDOC*************************************************************
REM 
REM 	Name:	ConvertToPixels
REM 
REM 	Action:	Convert the user's input into a pixel value
REM 
REM 	Params: 	dValue	- the input value
REM			dRes		- the document's X or Y display resolution
REM			nUnits	- the display units
REM 
REM 	Returns: 	LONG		- pixels
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/
Function ConvertToPixels(dValue as double, dRes as double, nUnits as long) as long

	DIM   dBaseValue		as double

	Const INCHTOBASE		as double		= 254000.0			
	Const MMTOBASE      	as double		= 10000.0
	Const PICATOBASE		as double		= (INCHTOBASE / 6.0)
	Const FRACTOBASE		as double		= (INCHTOBASE / 72.0)
	Const CICEROTOBASE		as double		= 45118.7
	Const DIDOTTOBASE		as double		= 3759.2
	Const UNITS_EPSILON		as double		= 0.000001
	Const DEFAULTRESOLUTION	as double		= 11811024.0
	
	if (nUnits = 6) then
		ConvertToPixels = INT(dValue)
		goto GetOut
	endif
	
	select case nUnits
		case 1		'inches
			dBaseValue = dValue * INCHTOBASE
		case 2		'millimeters
			dBaseValue = dValue * MMTOBASE
		case 3		'picas, points
			dBaseValue = dValue * PICATOBASE
		case 4		'points
			dBaseValue = dValue * FRACTOBASE
		case 5		'centimeters
			dBaseValue = dValue * MMTOBASE * 10
		case 6		'pixels
		case 7		'ciceros, didots
			dBaseValue = dValue * CICEROTOBASE
		case 8		'didots
			dBaseValue = dValue * DIDOTTOBASE
	end select
	
	if (dRes <= UNITS_EPSILON) then
		dRes = DEFAULTRESOLUTION
	endif
	
	if (dBaseValue <> 0.0) then
		ConvertToPixels = INT( ( (dRes / (MMTOBASE * 1000000.0)) * dBaseValue ) )
	else
		ConvertToPixels = 0
	endif

GetOut:

End Function


REM *SDOC*************************************************************
REM 
REM 	Name:	ConvertToUnit
REM 
REM 	Action:	Convert the user's pixel input into the specified unit
REM 
REM 	Params: 	nValue		- the input value
REM			dRes			- the document's X or Y display resolution
REM			nNewUnits		- the new display units
REM			nPrevUnits	- the previous display units
REM 
REM 	Returns: 	LONG			- pixels
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/
Function ConvertToUnit(dValue as double, dRes as double, nNewUnits as long, nOldUnits as long) as double

	DIM	dBaseValue		as double
	DIM	dResult			as double 

	Const MMTOBASE      	as double		= 10000.0
	Const UNITS_EPSILON		as double		= 0.000001
	Const DEFAULTRESOLUTION	as double		= 11811024.0

	nValue = ConvertToPixels( dValue, dRes, nOldUnits )
	
	if (nNewUnits = 6) then
		dResult = nValue
		goto GetOut
	endif
	
	if (dRes <= UNITS_EPSILON) then
		dRes = DEFAULTRESOLUTION
	endif
	
		
	if (nValue <> 0) then
		dBaseValue = nValue / (dRes / (MMTOBASE * 1000000.0)) 
	else
		dBaseValue = 100
	endif
	
	dResult = ConvertBaseToUnit( dBaseValue, dRes, nNewUnits )
	
GetOut:

	ConvertToUnit = dResult	


End Function

REM *SDOC*************************************************************
REM 
REM 	Name:	ConvertBaseToUnit
REM 
REM 	Action:	Convert the user's pixel input into the specified unit
REM 
REM 	Params: 	nValue		- the input value
REM			dRes			- the document's X or Y display resolution
REM			nNewUnits		- the new display units
REM			nPrevUnits	- the previous display units
REM 
REM 	Returns: 	LONG			- pixels
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/
Function ConvertBaseToUnit(dBaseValue as double, dRes as double, nUnits as long) as double

	Const INCHTOBASE		as double		= 254000.0			
	Const MMTOBASE      	as double		= 10000.0
	Const PICATOBASE		as double		= (INCHTOBASE / 6.0)
	Const FRACTOBASE		as double		= (INCHTOBASE / 72.0)
	Const CICEROTOBASE		as double		= 45118.7
	Const DIDOTTOBASE		as double		= 3759.2
	Const DEFAULTRESOLUTION	as double		= 11811024.0

	select case nUnits
		case 1		'inches
			ConvertBaseToUnit = dBaseValue / INCHTOBASE
		case 2		'millimeters
			ConvertBaseToUnit = dBaseValue / MMTOBASE
		case 3		'picas, points
			ConvertBaseToUnit = dBaseValue / PICATOBASE
		case 4		'points
			ConvertBaseToUnit = dBaseValue / FRACTOBASE
		case 5		'centimeters
			ConvertBaseToUnit = dBaseValue / MMTOBASE / 10
		case 6		'pixels
			if (dblResolution <= UNITS_EPSILON) then
				dblResolution = DEFAULTRESOLUTION
			endif
			ConvertBaseToUnit = (dRes / (MMTOBASE * 1000000.0)) * dBaseValue 
		case 7		'ciceros, didots
			ConvertBaseToUnit = dBaseValue / CICEROTOBASE
		case 8		'didots
			ConvertBaseToUnit = dBaseValue / DIDOTTOBASE
	end select

End Function

REM *SDOC*************************************************************
REM 
REM 	Name:	CreateHtml
REM 
REM 	Action:	
REM 
REM 	Params:	
REM			
REM 
REM 	Returns: Boolean
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Function CreateHTML()as Boolean
	Dim WantHTML as integer 
	WantHTML= Sheet.GenerateHTML.GETVALUE()
	if WantHTML=1 Then
		CreateHTML = True
	Else
		CreateHTML = False
	End if

End Function
REM *SDOC*************************************************************
REM 
REM 	Name:	PutText
REM 
REM 	Action:	
REM 
REM 	Params:	
REM			
REM 
REM 	Returns: Boolean	
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Function PutText as Boolean
	
	DIM PutName as Boolean
	DIM PutDate as Boolean
	DIM PutSize as Boolean
	DIM PutDim as Boolean
	DIM WantName as Integer
	DIM WantDate as integer
	DIM WantDim as integer
	DIM WantSize as integer

	PutName = False
	PutDate = False
	PutSize = False
	PutDim = False
	WantName=Sheet.NameFile.GETVALUE()
	WantDate=Sheet.FileDate.GETVALUE()
	WantDim=Sheet.Dimensions.GETVALUE()
	WantSize=Sheet.FileSize.GETVALUE()
	if WantName=1 then 
		PutName = True
	End if 	
	if WantDate=1 then 
		PutDate= True
	End if 		
	if WantDim=1 then 
		PutDim= True
	End if 		
	if PutSize=1 then 
		PutSize= True
	End if 			
	
	if PutName or PutDate or PutDim or PutDim Then
		PutText = True
	Else 
		PutText = False
	End if
End Function	
REM *SDOC*************************************************************
REM 
REM 	Name:	PutName
REM 
REM 	Action:	
REM 
REM 	Params:	
REM			
REM 
REM 	Returns: Boolean	
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Function PutName()as Boolean
	DIM WantName as Integer
	
	WantName=Sheet.NameFile.GETVALUE()
	if WantName=1 then 
		PutName = True
	Else
		PutName = False
	End if
End Function

REM *SDOC*************************************************************
REM 
REM 	Name:	PutDate
REM 
REM 	Action:	
REM 
REM 	Params:	
REM			
REM 
REM 	Returns: Boolean	
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Function PutDate()as Boolean
	DIM WantDate as Integer
	
	WantDate=Sheet.FileDate.GETVALUE()
	if WantDate=1 then 
		PutDate = True
	Else
		PutDate = False
	End if
End Function
REM *SDOC*************************************************************
REM 
REM 	Name:	PutSize
REM 
REM 	Action:	
REM 
REM 	Params:	
REM			
REM 
REM 	Returns: Boolean	
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Function PutSize()as Boolean
DIM WantSize as Integer
	
	WantSize=Sheet.FileSize.GETVALUE()
	if WantSize=1 then 
		PutSize = True
	Else
		PutSize = False
	End if
End Function
REM *SDOC*************************************************************
REM 
REM 	Name:	PutDim
REM 
REM 	Action:	
REM 
REM 	Params:	
REM			
REM 
REM 	Returns: Boolean	
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/

Function PutDim()as Boolean
DIM WantDim as Integer
	
	WantDim=Sheet.Dimensions.GETVALUE()
	if WantDim=1 then 
		PutDim = True
	Else
		PutDim = False
	End if
End Function
REM *SDOC*************************************************************
REM 
REM 	Name:	putLink
REM 
REM 	Action:	
REM 
REM 	Params:	
REM			
REM 
REM 	Returns: String
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Function PutLink()as Boolean
DIM WantLink as Integer
	
	WantLink =Sheet.LinkOriginalToTumbnails.GETVALUE()
	if WantLink=1 then 
		PutLink = True
	Else
		PutLink = False
	End if
End Function

REM *SDOC*************************************************************
REM 
REM 	Name:	CreateHtml
REM 
REM 	Action:	
REM 
REM 	Params:	
REM			
REM 
REM 	Returns: Boolean
REM 
REM 	Comments:	
REM 
REM *************************************************************EDOC*/
Sub CreateHTMLText ( L as integer,npages as integer, DesDirectory as String)
	
	Dim URL as String
	Dim DesURL as string
	Dim DesURLLC as string
	Dim NameString as string
	Dim DestinationName as string
	NameString = Sheet.FileNameTxt.GETTEXT ()
	pageName$=DesDirectory & NameString & "page" & Str (L)& ".html" 
	OPEN  pageName$ FOR OUTPUT AS L
		PRINT #L, "<!DOCTYPE HTML PUBLIC " + Chr(34) + "-//IETF//DTD HTML//EN/" + chr(34) + ">"
		PRINT #L, "<html>"
		PRINT #L, "<head><meta name=" + chr(34) + "GENERATOR" + chr(34) + "content=" + chr(34) + "Thumb Nails - Corel Corporation." + chr(34) + ">"
		PRINT #L, "<meta name=" + chr(34) + "keywords"  + chr(34) + "content=" + chr(34) + "Thumb Nails WebPageWizard"+ chr(34) + ">"
		PRINT #L, "<meta http-equiv=" + chr(34) + "Content-Type" + chr(34) + "content=" + chr(34) + "text/html; charset=iso-8859-1" + chr(34) +" >"
		URL = Sheet.URLTxt.GETTEXT()
		PRINT #L, "<title>Contact Sheet</title>"
		PRINT #L, "</head>"
		PRINT #L, "<body>"
		if URL<>"" Then
			DesURL = LEFT(URL, 7)
			DesURLLC = LCASE(DesURL)
			if (DesURLLC = "http://") then
				PRINT #L,  "<h1 align="+ chr(34) + " center" + chr(34) + ">" +"<A HREF=" + chr(34)+ URL + chr(34)+ ">Contact Sheet</A></h1>"
			Else
				PRINT #L,  "<h1 align="+ chr(34) + " center" + chr(34) + ">" +"<A HREF=" + chr(34)+ "http://" + URL + chr(34)+ ">Contact Sheet</A></h1>"
			End if
		Else
			PRINT #L, "<h1 align="+ chr(34) + " center" + chr(34) + ">Contact Sheet</h1>"				
		End if
		PRINT #L, "<p>"
		DestinationName = Sheet.FileNameTxt.GETTEXT()
		if L<>1 then
			Previouspage$= DestinationName +"page" & Str (L-1)& ".html"  
			PRINT #L, "<A HREF=" + chr(34) + PreviousPage$+ chr(34) + "> Previous</A>"  '<! ***Put the correct file name in here>
		End if
		if L <> nPages Then
			NextPage$ = DestinationName + "page" & Str (L+1)& ".html"   
			PRINT #L, "<A HREF="+  chr(34) + NextPage$ +  chr(34) +"> Next</A>"
		End if
		PRINT #L, "</p>"		
		PRINT #L, "<table border=" + chr(34) + "2" + chr(34) + ">"
	'	CLOSE #L
		
End Sub
REM *SDOC*************************************************************
REM 
REM 	Name:	RemoveExt
REM 
REM 	Action:	Removes the extension from a file name
REM 
REM 	Params: 	FileName 
REM			
REM			
REM 
REM 	Returns: 	String		- Returns string
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/
Function RemoveExt(FileName as string) as string
	Dim IndexR as integer
	Dim tempChar as string
	Dim TempString as string
	
	indexR = 1
	Size = len (FileName)
	tempChar = RIGHT(FileName, indexR)
	While (tempChar <>".") 
		indexR = indexR + 1
		tempRStr = RIGHT(FileName, indexR)
		tempChar = left (tempRStr,1)
	WEND	
	tempSize = Size-indexR	
	TempString = left(FileName, tempSize)
	RemoveExt= TempString
			
End Function
REM *SDOC*************************************************************
REM 
REM 	Name:	GetDirectory
REM 
REM 	Action:	if the Directory Size >Text Size put Directory as exp: c:\...\Dorectpru
REM 
REM 	Params: 	DirStr		- Photo paint directory 
REM			textSize		- The size of the text in the dialog
REM			
REM 
REM 	Returns: 	String		- Returns string
REM 
REM 	Comments:
REM 
REM *************************************************************EDOC*/

Function GetDirectory (DirStr as string,textSize as integer ) as string
	Dim Size as integer
	Dim tempRStr as string
	Dim tempLStr as string
	Dim indexL as integer
	Dim indexR as integer
	Dim tempChar as string
	Dim tempLChar as string
	Dim retStr as string
	
	indexL = 1
	indexR = 1
	Size = len (DirStr)
	if Size > textSize Then
		tempChar = RIGHT(DirStr, indexR)
		if tempChar = "\" then
			indexR = indexR + 1
			tempChar = RIGHT(DirStr, indexR)
		END IF	
			
		While (tempChar <>"\") and (indexR <> Size)
			indexR = indexR + 1
			tempRStr = RIGHT(DirStr, indexR)
			tempChar = left (tempRStr,1)
		WEND	
			tempLChar = Left(DirStr, indexL)	 
			While (tempLChar <> "\") and (indexR <> Size)
				indexL = indexL +1
				tempLStr = Left (DirStr, indexL)
				tempLChar = Right(tempLStr,1) 		
			WEND
			
			retStr = tempLStr + "..."+ tempRStr
			GetDirectory = retStr
		'End if	
	Else
		GetDirectory = DirStr
	End if
End function

REM *SDOC*************************************************************
REM 
REM 	Name:	CovertToPaletted
REM 
REM 	Action:	Change to paletted 
REM 
REM  
REM
REM			
REM 
REM 
REM 
REM Comments: This sub is used for HTML option if GIF extension is used
REM 
REM *************************************************************EDOC*/
 Sub CovertToPaletted ()
	WITHOBJECT "CorelPhotoPaint.Automation.9" 
		.ImageConvertPaletted 1, 216, 0, FALSE
		.PaletteColor 5, 255, 255, 255, 0, 0
		.PaletteColor 5, 204, 255, 255, 0, 1
		.PaletteColor 5, 153, 255, 255, 0, 2
		.PaletteColor 5, 102, 255, 255, 0, 3
		.PaletteColor 5, 51, 255, 255, 0, 4
		.PaletteColor 5, 0, 255, 255, 0, 5
		.PaletteColor 5, 255, 204, 255, 0, 6
		.PaletteColor 5, 204, 204, 255, 0, 7
		.PaletteColor 5, 153, 204, 255, 0, 8
		.PaletteColor 5, 102, 204, 255, 0, 9
		.PaletteColor 5, 51, 204, 255, 0, 10
		.PaletteColor 5, 0, 204, 255, 0, 11
		.PaletteColor 5, 255, 153, 255, 0, 12
		.PaletteColor 5, 204, 153, 255, 0, 13
		.PaletteColor 5, 153, 153, 255, 0, 14
		.PaletteColor 5, 102, 153, 255, 0, 15
		.PaletteColor 5, 51, 153, 255, 0, 16
		.PaletteColor 5, 0, 153, 255, 0, 17
		.PaletteColor 5, 255, 102, 255, 0, 18
		.PaletteColor 5, 204, 102, 255, 0, 19
		.PaletteColor 5, 153, 102, 255, 0, 20
		.PaletteColor 5, 102, 102, 255, 0, 21
		.PaletteColor 5, 51, 102, 255, 0, 22
		.PaletteColor 5, 0, 102, 255, 0, 23
		.PaletteColor 5, 255, 51, 255, 0, 24
		.PaletteColor 5, 204, 51, 255, 0, 25
		.PaletteColor 5, 153, 51, 255, 0, 26
		.PaletteColor 5, 102, 51, 255, 0, 27
		.PaletteColor 5, 51, 51, 255, 0, 28
		.PaletteColor 5, 0, 51, 255, 0, 29
		.PaletteColor 5, 255, 0, 255, 0, 30
		.PaletteColor 5, 204, 0, 255, 0, 31
		.PaletteColor 5, 153, 0, 255, 0, 32
		.PaletteColor 5, 102, 0, 255, 0, 33
		.PaletteColor 5, 51, 0, 255, 0, 34
		.PaletteColor 5, 0, 0, 255, 0, 35
		.PaletteColor 5, 255, 255, 204, 0, 36
		.PaletteColor 5, 204, 255, 204, 0, 37
		.PaletteColor 5, 153, 255, 204, 0, 38
		.PaletteColor 5, 102, 255, 204, 0, 39
		.PaletteColor 5, 51, 255, 204, 0, 40
		.PaletteColor 5, 0, 255, 204, 0, 41
		.PaletteColor 5, 255, 204, 204, 0, 42
		.PaletteColor 5, 204, 204, 204, 0, 43
		.PaletteColor 5, 153, 204, 204, 0, 44
		.PaletteColor 5, 102, 204, 204, 0, 45
		.PaletteColor 5, 51, 204, 204, 0, 46
		.PaletteColor 5, 0, 204, 204, 0, 47
		.PaletteColor 5, 255, 153, 204, 0, 48
		.PaletteColor 5, 204, 153, 204, 0, 49
		.PaletteColor 5, 153, 153, 204, 0, 50
		.PaletteColor 5, 102, 153, 204, 0, 51
		.PaletteColor 5, 51, 153, 204, 0, 52
		.PaletteColor 5, 0, 153, 204, 0, 53
		.PaletteColor 5, 255, 102, 204, 0, 54
		.PaletteColor 5, 204, 102, 204, 0, 55
		.PaletteColor 5, 153, 102, 204, 0, 56
		.PaletteColor 5, 102, 102, 204, 0, 57
		.PaletteColor 5, 51, 102, 204, 0, 58
		.PaletteColor 5, 0, 102, 204, 0, 59
		.PaletteColor 5, 255, 51, 204, 0, 60
		.PaletteColor 5, 204, 51, 204, 0, 61
		.PaletteColor 5, 153, 51, 204, 0, 62
		.PaletteColor 5, 102, 51, 204, 0, 63
		.PaletteColor 5, 51, 51, 204, 0, 64
		.PaletteColor 5, 0, 51, 204, 0, 65
		.PaletteColor 5, 255, 0, 204, 0, 66
		.PaletteColor 5, 204, 0, 204, 0, 67
		.PaletteColor 5, 153, 0, 204, 0, 68
		.PaletteColor 5, 102, 0, 204, 0, 69
		.PaletteColor 5, 51, 0, 204, 0, 70
		.PaletteColor 5, 0, 0, 204, 0, 71
		.PaletteColor 5, 255, 255, 153, 0, 72
		.PaletteColor 5, 204, 255, 153, 0, 73
		.PaletteColor 5, 153, 255, 153, 0, 74
		.PaletteColor 5, 102, 255, 153, 0, 75
		.PaletteColor 5, 51, 255, 153, 0, 76
		.PaletteColor 5, 0, 255, 153, 0, 77
		.PaletteColor 5, 255, 204, 153, 0, 78
		.PaletteColor 5, 204, 204, 153, 0, 79
		.PaletteColor 5, 153, 204, 153, 0, 80
		.PaletteColor 5, 102, 204, 153, 0, 81
		.PaletteColor 5, 51, 204, 153, 0, 82
		.PaletteColor 5, 0, 204, 153, 0, 83
		.PaletteColor 5, 255, 153, 153, 0, 84
		.PaletteColor 5, 204, 153, 153, 0, 85
		.PaletteColor 5, 153, 153, 153, 0, 86
		.PaletteColor 5, 102, 153, 153, 0, 87
		.PaletteColor 5, 51, 153, 153, 0, 88
		.PaletteColor 5, 0, 153, 153, 0, 89
		.PaletteColor 5, 255, 102, 153, 0, 90
		.PaletteColor 5, 204, 102, 153, 0, 91
		.PaletteColor 5, 153, 102, 153, 0, 92
		.PaletteColor 5, 102, 102, 153, 0, 93
		.PaletteColor 5, 51, 102, 153, 0, 94
		.PaletteColor 5, 0, 102, 153, 0, 95
		.PaletteColor 5, 255, 51, 153, 0, 96
		.PaletteColor 5, 204, 51, 153, 0, 97
		.PaletteColor 5, 153, 51, 153, 0, 98
		.PaletteColor 5, 102, 51, 153, 0, 99
		.PaletteColor 5, 51, 51, 153, 0, 100
		.PaletteColor 5, 0, 51, 153, 0, 101
		.PaletteColor 5, 255, 0, 153, 0, 102
		.PaletteColor 5, 204, 0, 153, 0, 103
		.PaletteColor 5, 153, 0, 153, 0, 104
		.PaletteColor 5, 102, 0, 153, 0, 105
		.PaletteColor 5, 51, 0, 153, 0, 106
		.PaletteColor 5, 0, 0, 153, 0, 107
		.PaletteColor 5, 255, 255, 102, 0, 108
		.PaletteColor 5, 204, 255, 102, 0, 109
		.PaletteColor 5, 153, 255, 102, 0, 110
		.PaletteColor 5, 102, 255, 102, 0, 111
		.PaletteColor 5, 51, 255, 102, 0, 112
		.PaletteColor 5, 0, 255, 102, 0, 113
		.PaletteColor 5, 255, 204, 102, 0, 114
		.PaletteColor 5, 204, 204, 102, 0, 115
		.PaletteColor 5, 153, 204, 102, 0, 116
		.PaletteColor 5, 102, 204, 102, 0, 117
		.PaletteColor 5, 51, 204, 102, 0, 118
		.PaletteColor 5, 0, 204, 102, 0, 119
		.PaletteColor 5, 255, 153, 102, 0, 120
		.PaletteColor 5, 204, 153, 102, 0, 121
		.PaletteColor 5, 153, 153, 102, 0, 122
		.PaletteColor 5, 102, 153, 102, 0, 123
		.PaletteColor 5, 51, 153, 102, 0, 124
		.PaletteColor 5, 0, 153, 102, 0, 125
		.PaletteColor 5, 255, 102, 102, 0, 126
		.PaletteColor 5, 204, 102, 102, 0, 127
		.PaletteColor 5, 153, 102, 102, 0, 128
		.PaletteColor 5, 102, 102, 102, 0, 129
		.PaletteColor 5, 51, 102, 102, 0, 130
		.PaletteColor 5, 0, 102, 102, 0, 131
		.PaletteColor 5, 255, 51, 102, 0, 132
		.PaletteColor 5, 204, 51, 102, 0, 133
		.PaletteColor 5, 153, 51, 102, 0, 134
		.PaletteColor 5, 102, 51, 102, 0, 135
		.PaletteColor 5, 51, 51, 102, 0, 136
		.PaletteColor 5, 0, 51, 102, 0, 137
		.PaletteColor 5, 255, 0, 102, 0, 138
		.PaletteColor 5, 204, 0, 102, 0, 139
		.PaletteColor 5, 153, 0, 102, 0, 140
		.PaletteColor 5, 102, 0, 102, 0, 141
		.PaletteColor 5, 51, 0, 102, 0, 142
		.PaletteColor 5, 0, 0, 102, 0, 143
		.PaletteColor 5, 255, 255, 51, 0, 144
		.PaletteColor 5, 204, 255, 51, 0, 145
		.PaletteColor 5, 153, 255, 51, 0, 146
		.PaletteColor 5, 102, 255, 51, 0, 147
		.PaletteColor 5, 51, 255, 51, 0, 148
		.PaletteColor 5, 0, 255, 51, 0, 149
		.PaletteColor 5, 255, 204, 51, 0, 150
		.PaletteColor 5, 204, 204, 51, 0, 151
		.PaletteColor 5, 153, 204, 51, 0, 152
		.PaletteColor 5, 102, 204, 51, 0, 153
		.PaletteColor 5, 51, 204, 51, 0, 154
		.PaletteColor 5, 0, 204, 51, 0, 155
		.PaletteColor 5, 255, 153, 51, 0, 156
		.PaletteColor 5, 204, 153, 51, 0, 157
		.PaletteColor 5, 153, 153, 51, 0, 158
		.PaletteColor 5, 102, 153, 51, 0, 159
		.PaletteColor 5, 51, 153, 51, 0, 160
		.PaletteColor 5, 0, 153, 51, 0, 161
		.PaletteColor 5, 255, 102, 51, 0, 162
		.PaletteColor 5, 204, 102, 51, 0, 163
		.PaletteColor 5, 153, 102, 51, 0, 164
		.PaletteColor 5, 102, 102, 51, 0, 165
		.PaletteColor 5, 51, 102, 51, 0, 166
		.PaletteColor 5, 0, 102, 51, 0, 167
		.PaletteColor 5, 255, 51, 51, 0, 168
		.PaletteColor 5, 204, 51, 51, 0, 169
		.PaletteColor 5, 153, 51, 51, 0, 170
		.PaletteColor 5, 102, 51, 51, 0, 171
		.PaletteColor 5, 51, 51, 51, 0, 172
		.PaletteColor 5, 0, 51, 51, 0, 173
		.PaletteColor 5, 255, 0, 51, 0, 174
		.PaletteColor 5, 204, 0, 51, 0, 175
		.PaletteColor 5, 153, 0, 51, 0, 176
		.PaletteColor 5, 102, 0, 51, 0, 177
		.PaletteColor 5, 51, 0, 51, 0, 178
		.PaletteColor 5, 0, 0, 51, 0, 179
		.PaletteColor 5, 255, 255, 0, 0, 180
		.PaletteColor 5, 204, 255, 0, 0, 181
		.PaletteColor 5, 153, 255, 0, 0, 182
		.PaletteColor 5, 102, 255, 0, 0, 183
		.PaletteColor 5, 51, 255, 0, 0, 184
		.PaletteColor 5, 0, 255, 0, 0, 185
		.PaletteColor 5, 255, 204, 0, 0, 186
		.PaletteColor 5, 204, 204, 0, 0, 187
		.PaletteColor 5, 153, 204, 0, 0, 188
		.PaletteColor 5, 102, 204, 0, 0, 189
		.PaletteColor 5, 51, 204, 0, 0, 190
		.PaletteColor 5, 0, 204, 0, 0, 191
		.PaletteColor 5, 255, 153, 0, 0, 192
		.PaletteColor 5, 204, 153, 0, 0, 193
		.PaletteColor 5, 153, 153, 0, 0, 194
		.PaletteColor 5, 102, 153, 0, 0, 195
		.PaletteColor 5, 51, 153, 0, 0, 196
		.PaletteColor 5, 0, 153, 0, 0, 197
		.PaletteColor 5, 255, 102, 0, 0, 198
		.PaletteColor 5, 204, 102, 0, 0, 199
		.PaletteColor 5, 153, 102, 0, 0, 200
		.PaletteColor 5, 102, 102, 0, 0, 201
		.PaletteColor 5, 51, 102, 0, 0, 202
		.PaletteColor 5, 0, 102, 0, 0, 203
		.PaletteColor 5, 255, 51, 0, 0, 204
		.PaletteColor 5, 204, 51, 0, 0, 205
		.PaletteColor 5, 153, 51, 0, 0, 206
		.PaletteColor 5, 102, 51, 0, 0, 207
		.PaletteColor 5, 51, 51, 0, 0, 208
		.PaletteColor 5, 0, 51, 0, 0, 209
		.PaletteColor 5, 255, 0, 0, 0, 210
		.PaletteColor 5, 204, 0, 0, 0, 211
		.PaletteColor 5, 153, 0, 0, 0, 212
		.PaletteColor 5, 102, 0, 0, 0, 213
		.PaletteColor 5, 51, 0, 0, 0, 214
		.PaletteColor 5, 0, 0, 0, 0, 215
		.EndConvertPaletted 
	END WITHOBJECT
End Sub
