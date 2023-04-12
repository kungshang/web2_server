REM Chrome Text Effect for PHOTO-PAINT 9
REM Prompt User for text then create a New File and Text Object

label1 = "This script produces chrome text. Enter a text string, to create a new 24-bit image."
BEGIN DIALOG Dialog1 196, 74, "Chrome Text"
	GROUPBOX  2, 1, 191, 53
	TEXT  8, 12, 178, 37, label1
	OKBUTTON  109, 57, 40, 14
	CANCELBUTTON  153, 57, 40, 14
END DIALOG

ret = DIALOG(Dialog1)
' If Cancel is selected, stop the script
IF ret = 2 THEN STOP

' Get the Text
TxtStr$ = inputbox("Chrome Text Effect")
if TxtStr$ = "" then
	messagebox "You have entered nothing", "Warning", 48
	stop
end if

WITHOBJECT "CorelPhotoPaint.Automation.9"
	.FileNew 640, 480, 1, 72, 72, FALSE, FALSE, 1, 0, 0, 0, 0, 0, 0, 0, 0, FALSE

	' The Text String
	.TextTool 10, 100, FALSE, TRUE ' add Mask Mode 0
		.TextSetting "Fill", "255, 255, 255"
		.TextSetting "Font", "Arial Black"
		.TextSetting "TypeSize", "96.0"
		.TextAppend TxtStr$
		.TextRender 
	
	.ObjectAlign 3, 3, FALSE, FALSE, FALSE, TRUE, TRUE, TRUE, TRUE
		.EndObject
	
	.MaskCreate TRUE, 0
		.EndMaskCreate
	.MaskChannelAdd "original text"

	.ObjectMerge FALSE
		.EndObject
	
	.MaskRemove 
	.EffectGaussianBlur 4.0
	.MaskSelectAll 
	.ObjectCreate TRUE

	.ObjectTranslate -3, -3, FALSE
		.EndObject
	.ObjectMergeMode 3
		.EndObject
	.ObjectMerge FALSE
		.EndObject

	.ImageInvert 
		.EndColorEffect 

	.ImageAutoEqualize 5, 5

	.ImageToneCurve 
		.ImageToneTable 0, 121, 121, 121, 121
		.ImageToneTable 1, 122, 122, 122, 122
		.ImageToneTable 2, 123, 123, 123, 123
		.ImageToneTable 3, 124, 124, 124, 124
		.ImageToneTable 4, 125, 125, 125, 125
		.ImageToneTable 5, 126, 126, 126, 126
		.ImageToneTable 6, 128, 128, 128, 128
		.ImageToneTable 7, 130, 130, 130, 130
		.ImageToneTable 8, 132, 132, 132, 132
		.ImageToneTable 9, 134, 134, 134, 134
		.ImageToneTable 10, 136, 136, 136, 136
		.ImageToneTable 11, 138, 138, 138, 138
		.ImageToneTable 12, 141, 141, 141, 141
		.ImageToneTable 13, 143, 143, 143, 143
		.ImageToneTable 14, 146, 146, 146, 146
		.ImageToneTable 15, 148, 148, 148, 148
		.ImageToneTable 16, 151, 151, 151, 151
		.ImageToneTable 17, 154, 154, 154, 154
		.ImageToneTable 18, 157, 157, 157, 157
		.ImageToneTable 19, 159, 159, 159, 159
		.ImageToneTable 20, 162, 162, 162, 162
		.ImageToneTable 21, 165, 165, 165, 165
		.ImageToneTable 22, 168, 168, 168, 168
		.ImageToneTable 23, 171, 171, 171, 171
		.ImageToneTable 24, 174, 174, 174, 174
		.ImageToneTable 25, 177, 177, 177, 177
		.ImageToneTable 26, 180, 180, 180, 180
		.ImageToneTable 27, 183, 183, 183, 183
		.ImageToneTable 28, 186, 186, 186, 186
		.ImageToneTable 29, 189, 189, 189, 189
		.ImageToneTable 30, 192, 192, 192, 192
		.ImageToneTable 31, 195, 195, 195, 195
		.ImageToneTable 32, 197, 197, 197, 197
		.ImageToneTable 33, 200, 200, 200, 200
		.ImageToneTable 34, 203, 203, 203, 203
		.ImageToneTable 35, 205, 205, 205, 205
		.ImageToneTable 36, 208, 208, 208, 208
		.ImageToneTable 37, 210, 210, 210, 210
		.ImageToneTable 38, 212, 212, 212, 212
		.ImageToneTable 39, 214, 214, 214, 214
		.ImageToneTable 40, 216, 216, 216, 216
		.ImageToneTable 41, 218, 218, 218, 218
		.ImageToneTable 42, 220, 220, 220, 220
		.ImageToneTable 43, 221, 221, 221, 221
		.ImageToneTable 44, 222, 222, 222, 222
		.ImageToneTable 45, 223, 223, 223, 223
		.ImageToneTable 46, 224, 224, 224, 224
		.ImageToneTable 47, 224, 224, 224, 224
		.ImageToneTable 48, 224, 224, 224, 224
		.ImageToneTable 49, 223, 223, 223, 223
		.ImageToneTable 50, 223, 223, 223, 223
		.ImageToneTable 51, 222, 222, 222, 222
		.ImageToneTable 52, 220, 220, 220, 220
		.ImageToneTable 53, 218, 218, 218, 218
		.ImageToneTable 54, 216, 216, 216, 216
		.ImageToneTable 55, 213, 213, 213, 213
		.ImageToneTable 56, 210, 210, 210, 210
		.ImageToneTable 57, 206, 206, 206, 206
		.ImageToneTable 58, 202, 202, 202, 202
		.ImageToneTable 59, 197, 197, 197, 197
		.ImageToneTable 60, 192, 192, 192, 192
		.ImageToneTable 61, 186, 186, 186, 186
		.ImageToneTable 62, 181, 181, 181, 181
		.ImageToneTable 63, 174, 174, 174, 174
		.ImageToneTable 64, 168, 168, 168, 168
		.ImageToneTable 65, 161, 161, 161, 161
		.ImageToneTable 66, 153, 153, 153, 153
		.ImageToneTable 67, 146, 146, 146, 146
		.ImageToneTable 68, 138, 138, 138, 138
		.ImageToneTable 69, 131, 131, 131, 131
		.ImageToneTable 70, 123, 123, 123, 123
		.ImageToneTable 71, 115, 115, 115, 115
		.ImageToneTable 72, 108, 108, 108, 108
		.ImageToneTable 73, 100, 100, 100, 100
		.ImageToneTable 74, 93, 93, 93, 93
		.ImageToneTable 75, 86, 86, 86, 86
		.ImageToneTable 76, 79, 79, 79, 79
		.ImageToneTable 77, 73, 73, 73, 73
		.ImageToneTable 78, 67, 67, 67, 67
		.ImageToneTable 79, 62, 62, 62, 62
		.ImageToneTable 80, 57, 57, 57, 57
		.ImageToneTable 81, 52, 52, 52, 52
		.ImageToneTable 82, 48, 48, 48, 48
		.ImageToneTable 83, 44, 44, 44, 44
		.ImageToneTable 84, 41, 41, 41, 41
		.ImageToneTable 85, 38, 38, 38, 38
		.ImageToneTable 86, 36, 36, 36, 36
		.ImageToneTable 87, 33, 33, 33, 33
		.ImageToneTable 88, 32, 32, 32, 32
		.ImageToneTable 89, 30, 30, 30, 30
		.ImageToneTable 90, 29, 29, 29, 29
		.ImageToneTable 91, 29, 29, 29, 29
		.ImageToneTable 92, 28, 28, 28, 28
		.ImageToneTable 93, 28, 28, 28, 28
		.ImageToneTable 94, 28, 28, 28, 28
		.ImageToneTable 95, 29, 29, 29, 29
		.ImageToneTable 96, 30, 30, 30, 30
		.ImageToneTable 97, 31, 31, 31, 31
		.ImageToneTable 98, 32, 32, 32, 32
		.ImageToneTable 99, 34, 34, 34, 34
		.ImageToneTable 100, 35, 35, 35, 35
		.ImageToneTable 101, 37, 37, 37, 37
		.ImageToneTable 102, 40, 40, 40, 40
		.ImageToneTable 103, 42, 42, 42, 42
		.ImageToneTable 104, 45, 45, 45, 45
		.ImageToneTable 105, 48, 48, 48, 48
		.ImageToneTable 106, 51, 51, 51, 51
		.ImageToneTable 107, 55, 55, 55, 55
		.ImageToneTable 108, 58, 58, 58, 58
		.ImageToneTable 109, 62, 62, 62, 62
		.ImageToneTable 110, 66, 66, 66, 66
		.ImageToneTable 111, 70, 70, 70, 70
		.ImageToneTable 112, 74, 74, 74, 74
		.ImageToneTable 113, 78, 78, 78, 78
		.ImageToneTable 114, 83, 83, 83, 83
		.ImageToneTable 115, 87, 87, 87, 87
		.ImageToneTable 116, 92, 92, 92, 92
		.ImageToneTable 117, 96, 96, 96, 96
		.ImageToneTable 118, 101, 101, 101, 101
		.ImageToneTable 119, 105, 105, 105, 105
		.ImageToneTable 120, 110, 110, 110, 110
		.ImageToneTable 121, 115, 115, 115, 115
		.ImageToneTable 122, 119, 119, 119, 119
		.ImageToneTable 123, 124, 124, 124, 124
		.ImageToneTable 124, 128, 128, 128, 128
		.ImageToneTable 125, 133, 133, 133, 133
		.ImageToneTable 126, 137, 137, 137, 137
		.ImageToneTable 127, 142, 142, 142, 142
		.ImageToneTable 128, 146, 146, 146, 146
		.ImageToneTable 129, 150, 150, 150, 150
		.ImageToneTable 130, 154, 154, 154, 154
		.ImageToneTable 131, 158, 158, 158, 158
		.ImageToneTable 132, 162, 162, 162, 162
		.ImageToneTable 133, 165, 165, 165, 165
		.ImageToneTable 134, 169, 169, 169, 169
		.ImageToneTable 135, 172, 172, 172, 172
		.ImageToneTable 136, 175, 175, 175, 175
		.ImageToneTable 137, 178, 178, 178, 178
		.ImageToneTable 138, 181, 181, 181, 181
		.ImageToneTable 139, 184, 184, 184, 184
		.ImageToneTable 140, 186, 186, 186, 186
		.ImageToneTable 141, 189, 189, 189, 189
		.ImageToneTable 142, 191, 191, 191, 191
		.ImageToneTable 143, 194, 194, 194, 194
		.ImageToneTable 144, 196, 196, 196, 196
		.ImageToneTable 145, 198, 198, 198, 198
		.ImageToneTable 146, 200, 200, 200, 200
		.ImageToneTable 147, 201, 201, 201, 201
		.ImageToneTable 148, 203, 203, 203, 203
		.ImageToneTable 149, 205, 205, 205, 205
		.ImageToneTable 150, 207, 207, 207, 207
		.ImageToneTable 151, 208, 208, 208, 208
		.ImageToneTable 152, 210, 210, 210, 210
		.ImageToneTable 153, 211, 211, 211, 211
		.ImageToneTable 154, 212, 212, 212, 212
		.ImageToneTable 155, 214, 214, 214, 214
		.ImageToneTable 156, 215, 215, 215, 215
		.ImageToneTable 157, 216, 216, 216, 216
		.ImageToneTable 158, 217, 217, 217, 217
		.ImageToneTable 159, 219, 219, 219, 219
		.ImageToneTable 160, 220, 220, 220, 220
		.ImageToneTable 161, 221, 221, 221, 221
		.ImageToneTable 162, 222, 222, 222, 222
		.ImageToneTable 163, 223, 223, 223, 223
		.ImageToneTable 164, 223, 223, 223, 223
		.ImageToneTable 165, 224, 224, 224, 224
		.ImageToneTable 166, 225, 225, 225, 225
		.ImageToneTable 167, 226, 226, 226, 226
		.ImageToneTable 168, 226, 226, 226, 226
		.ImageToneTable 169, 227, 227, 227, 227
		.ImageToneTable 170, 227, 227, 227, 227
		.ImageToneTable 171, 227, 227, 227, 227
		.ImageToneTable 172, 228, 228, 228, 228
		.ImageToneTable 173, 228, 228, 228, 228
		.ImageToneTable 174, 228, 228, 228, 228
		.ImageToneTable 175, 227, 227, 227, 227
		.ImageToneTable 176, 227, 227, 227, 227
		.ImageToneTable 177, 226, 226, 226, 226
		.ImageToneTable 178, 226, 226, 226, 226
		.ImageToneTable 179, 225, 225, 225, 225
		.ImageToneTable 180, 224, 224, 224, 224
		.ImageToneTable 181, 222, 222, 222, 222
		.ImageToneTable 182, 221, 221, 221, 221
		.ImageToneTable 183, 219, 219, 219, 219
		.ImageToneTable 184, 217, 217, 217, 217
		.ImageToneTable 185, 215, 215, 215, 215
		.ImageToneTable 186, 212, 212, 212, 212
		.ImageToneTable 187, 210, 210, 210, 210
		.ImageToneTable 188, 207, 207, 207, 207
		.ImageToneTable 189, 203, 203, 203, 203
		.ImageToneTable 190, 200, 200, 200, 200
		.ImageToneTable 191, 196, 196, 196, 196
		.ImageToneTable 192, 192, 192, 192, 192
		.ImageToneTable 193, 188, 188, 188, 188
		.ImageToneTable 194, 184, 184, 184, 184
		.ImageToneTable 195, 179, 179, 179, 179
		.ImageToneTable 196, 174, 174, 174, 174
		.ImageToneTable 197, 169, 169, 169, 169
		.ImageToneTable 198, 164, 164, 164, 164
		.ImageToneTable 199, 159, 159, 159, 159
		.ImageToneTable 200, 154, 154, 154, 154
		.ImageToneTable 201, 149, 149, 149, 149
		.ImageToneTable 202, 143, 143, 143, 143
		.ImageToneTable 203, 138, 138, 138, 138
		.ImageToneTable 204, 132, 132, 132, 132
		.ImageToneTable 205, 127, 127, 127, 127
		.ImageToneTable 206, 121, 121, 121, 121
		.ImageToneTable 207, 116, 116, 116, 116
		.ImageToneTable 208, 111, 111, 111, 111
		.ImageToneTable 209, 106, 106, 106, 106
		.ImageToneTable 210, 101, 101, 101, 101
		.ImageToneTable 211, 96, 96, 96, 96
		.ImageToneTable 212, 91, 91, 91, 91
		.ImageToneTable 213, 87, 87, 87, 87
		.ImageToneTable 214, 83, 83, 83, 83
		.ImageToneTable 215, 79, 79, 79, 79
		.ImageToneTable 216, 75, 75, 75, 75
		.ImageToneTable 217, 72, 72, 72, 72
		.ImageToneTable 218, 69, 69, 69, 69
		.ImageToneTable 219, 66, 66, 66, 66
		.ImageToneTable 220, 63, 63, 63, 63
		.ImageToneTable 221, 61, 61, 61, 61
		.ImageToneTable 222, 59, 59, 59, 59
		.ImageToneTable 223, 57, 57, 57, 57
		.ImageToneTable 224, 56, 56, 56, 56
		.ImageToneTable 225, 55, 55, 55, 55
		.ImageToneTable 226, 54, 54, 54, 54
		.ImageToneTable 227, 53, 53, 53, 53
		.ImageToneTable 228, 53, 53, 53, 53
		.ImageToneTable 229, 53, 53, 53, 53
		.ImageToneTable 230, 53, 53, 53, 53
		.ImageToneTable 231, 54, 54, 54, 54
		.ImageToneTable 232, 55, 55, 55, 55
		.ImageToneTable 233, 56, 56, 56, 56
		.ImageToneTable 234, 57, 57, 57, 57
		.ImageToneTable 235, 59, 59, 59, 59
		.ImageToneTable 236, 61, 61, 61, 61
		.ImageToneTable 237, 62, 62, 62, 62
		.ImageToneTable 238, 64, 64, 64, 64
		.ImageToneTable 239, 66, 66, 66, 66
		.ImageToneTable 240, 69, 69, 69, 69
		.ImageToneTable 241, 71, 71, 71, 71
		.ImageToneTable 242, 73, 73, 73, 73
		.ImageToneTable 243, 75, 75, 75, 75
		.ImageToneTable 244, 77, 77, 77, 77
		.ImageToneTable 245, 79, 79, 79, 79
		.ImageToneTable 246, 82, 82, 82, 82
		.ImageToneTable 247, 84, 84, 84, 84
		.ImageToneTable 248, 86, 86, 86, 86
		.ImageToneTable 249, 87, 87, 87, 87
		.ImageToneTable 250, 89, 89, 89, 89
		.ImageToneTable 251, 90, 90, 90, 90
		.ImageToneTable 252, 91, 91, 91, 91
		.ImageToneTable 253, 92, 92, 92, 92
		.ImageToneTable 254, 93, 93, 93, 93
		.ImageToneTable 255, 94, 94, 94, 94
		.EndImageToneCurve 
	
	.MaskChannelToMask 0, 0
	.MaskInvert 
	.EditFill 0, 0, 100, 0, 0, 0, 0, 0, 0, 0, 0
		.FillSolid 5, 0, 0, 0, 0
		.EndEditFill 
	.MaskInvert 
	.MaskFeather 8, 0, 1
	.MaskInvert 
	.ImageInvert 
		.EndColorEffect 
	.MaskRemove 
	.MaskChannelToMask 0, 0
	.MaskFeather 45, 0, 1
	' Tried to add some Color to the text
	'.EffectPlugin "Fancy", "Julia Set Explorer 2.0...", 0, 0, ""
	.EditFill 0, 0, 100, 1, 0, 45, 18, 127, 4, 0, 0
		.FillSolid 5, 255, 255, 255, 0
		.EndEditFill 
END WITHOBJECT
