REM Creates Header in current script
REM Copyright 1996 Corel Corporation. All rights reserved.

REM *****************************************************************
REM This script creates a header in the currently open script.
REM To run this script, assign it to a menu or toolbar button. 
REM Then open the script in which you would like to insert a  
REM header and run the header script.
REM ******************************************************************

REM Perform a sanity check
REM to ensure that this script isn't 
REM being run from the script editor...
WITHOBJECT "CorelScript.Automation.9"
	.GoToLine 11
	Match$ = .GetLineText ()
	IF Match$ = "REM Perform a sanity check" THEN
		MESSAGE "This script should not be run from the Corel SCRIPT Editor."
		END
	ENDIF
END WITHOBJECT

' Defines
#define IDOK  1
#define IDCANCEL  2

'  Adds Specified header info
sub WriteHeader(Name as string, Desc as string)
	DIM DateStamp AS DATE
	DateStamp=CDAT(Fix(GetCurrDate()))
	withobject "CorelScript.Automation.9"
		.GotoLine 1
		.AddLineBefore "REM " & Name & " for v9.0, created on " & str(DateStamp) 
		.AddLineBefore "REM " & Desc
	end withobject
end sub

'####################################################################
' Main section
DIM Name AS STRING
DIM Desc AS STRING
BEGIN DIALOG Header 199, 92, "Header Maker" ' Ask for header name and description
	TEXT  5, 5, 105, 10, "Script Name"
	TEXTBOX  5, 15, 190, 15, Name$
	TEXT  5, 35, 75, 10, "Description"
	TEXTBOX  5, 45, 190, 15, Desc$
	OKBUTTON  50, 70, 40, 14
	CANCELBUTTON  110, 70, 40, 14
END DIALOG
if (dialog(Header)=IDOK) then WriteHeader Name$, Desc$




