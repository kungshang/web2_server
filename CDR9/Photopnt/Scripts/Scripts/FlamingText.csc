REM Animated Fire Effect for PhotoPaint Version 9.
REM Created by Rob Wineck

' default Number of Frames, changes like to your liking
FRAMENUM = 10

label1 = "This script produces animated flaming text. Enter a text string to create a new 10 frame movie."
BEGIN DIALOG Dialog1 196, 74, "Flaming Text"
	GROUPBOX  2, 1, 191, 53
	TEXT  8, 12, 178, 37, label1
	OKBUTTON  109, 57, 40, 14
	CANCELBUTTON  153, 57, 40, 14
END DIALOG

ret = DIALOG(Dialog1)
' If Cancel is selected, stop the script
IF ret = 2 THEN STOP

' Get Text String
TxtStr$ = inputbox("Animated Flaming Text")

'Assigned another string data type to TxtStr - Hung Tran
TxtStr2$ = TxtStr$

if TxtStr$ = "" then
	message "This script needs a text string to work."
	stop
end if

WITHOBJECT "CorelPhotoPaint.Automation.9"
	.FileNew 320, 200, 1, 100, 100, FALSE, FALSE, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0
		
	' The Text String
	.TextTool 10, 100, FALSE, TRUE ' add Mask Mode 0
		.TextSetting "Fill", "255, 255, 255"
		.TextSetting "Font", "Futura XBlk BT"
		.TextSetting "TypeSize", "48.0" 
		.TextAppend TxtStr$
		.TextRender 
			
		
	.ObjectAlign 3, 3, FALSE, FALSE, FALSE, TRUE, FALSE, TRUE, TRUE
		.EndObject
	.ObjectMerge FALSE
		.EndObject

	
	' Create a Movie 	
	.MovieCreate
	.MovieInsertFrame 1, FRAMENUM - 1, 0, 1
	.MovieRewind
	
	' Set Default Angle
	Angle = 90
	' nRnd is used for Positive or negative angles
	Randomize
	For i = 1 to FRAMENUM
		
		' Create a little randomness
		nRnd# = RND()
		if nRnd > 0.5 then
			nSign = 2
		else
			nSign = -2
		end if
		
		Angle = Angle + nSign
		' These seem to work well
		.EffectScatter 6, 6
			.EndColorEffect
		.EffectGaussianBlur 3.0
		.EffectWind 50, 50, Angle	
		.EffectRipple 7, 1, Angle , TRUE, 0
		.MovieForwardOne
	next i

	' Now Convert to Paletted to get the Fire Effect		
	' First we convert to a Greyscale Palette.
	.ImageConvertPaletted 3, 256, 0, FALSE
		for i = 0 to 255
			.PaletteColor 5, i, i, i, 0, i
		next i
		.EndConvertPaletted 
	
' Now we convert to a Blackbody Palette
	.ImageColorTable 256
		.PaletteColor 5, 0, 0, 0, 0, 0
		.PaletteColor 5, 3, 0, 0, 0, 1
		.PaletteColor 5, 6, 0, 0, 0, 2
		.PaletteColor 5, 9, 0, 0, 0, 3
		.PaletteColor 5, 12, 0, 0, 0, 4
		.PaletteColor 5, 15, 0, 0, 0, 5
		.PaletteColor 5, 18, 0, 0, 0, 6
		.PaletteColor 5, 21, 0, 0, 0, 7
		.PaletteColor 5, 24, 0, 0, 0, 8
		.PaletteColor 5, 27, 0, 0, 0, 9
		.PaletteColor 5, 30, 0, 0, 0, 10
		.PaletteColor 5, 33, 0, 0, 0, 11
		.PaletteColor 5, 36, 0, 0, 0, 12
		.PaletteColor 5, 39, 0, 0, 0, 13
		.PaletteColor 5, 42, 0, 0, 0, 14
		.PaletteColor 5, 45, 0, 0, 0, 15
		.PaletteColor 5, 48, 0, 0, 0, 16
		.PaletteColor 5, 51, 0, 0, 0, 17
		.PaletteColor 5, 54, 0, 0, 0, 18
		.PaletteColor 5, 57, 0, 0, 0, 19
		.PaletteColor 5, 60, 0, 0, 0, 20
		.PaletteColor 5, 63, 0, 0, 0, 21
		.PaletteColor 5, 66, 0, 0, 0, 22
		.PaletteColor 5, 69, 0, 0, 0, 23
		.PaletteColor 5, 72, 0, 0, 0, 24
		.PaletteColor 5, 75, 0, 0, 0, 25
		.PaletteColor 5, 78, 0, 0, 0, 26
		.PaletteColor 5, 81, 0, 0, 0, 27
		.PaletteColor 5, 84, 0, 0, 0, 28
		.PaletteColor 5, 87, 0, 0, 0, 29
		.PaletteColor 5, 90, 0, 0, 0, 30
		.PaletteColor 5, 93, 0, 0, 0, 31
		.PaletteColor 5, 96, 0, 0, 0, 32
		.PaletteColor 5, 99, 0, 0, 0, 33
		.PaletteColor 5, 102, 0, 0, 0, 34
		.PaletteColor 5, 105, 0, 0, 0, 35
		.PaletteColor 5, 108, 0, 0, 0, 36
		.PaletteColor 5, 111, 0, 0, 0, 37
		.PaletteColor 5, 114, 0, 0, 0, 38
		.PaletteColor 5, 117, 0, 0, 0, 39
		.PaletteColor 5, 120, 0, 0, 0, 40
		.PaletteColor 5, 123, 0, 0, 0, 41
		.PaletteColor 5, 126, 0, 0, 0, 42
		.PaletteColor 5, 129, 0, 0, 0, 43
		.PaletteColor 5, 132, 0, 0, 0, 44
		.PaletteColor 5, 135, 0, 0, 0, 45
		.PaletteColor 5, 138, 0, 0, 0, 46
		.PaletteColor 5, 141, 0, 0, 0, 47
		.PaletteColor 5, 144, 0, 0, 0, 48
		.PaletteColor 5, 147, 0, 0, 0, 49
		.PaletteColor 5, 150, 0, 0, 0, 50
		.PaletteColor 5, 153, 0, 0, 0, 51
		.PaletteColor 5, 156, 0, 0, 0, 52
		.PaletteColor 5, 159, 0, 0, 0, 53
		.PaletteColor 5, 162, 0, 0, 0, 54
		.PaletteColor 5, 165, 0, 0, 0, 55
		.PaletteColor 5, 168, 0, 0, 0, 56
		.PaletteColor 5, 171, 0, 0, 0, 57
		.PaletteColor 5, 174, 0, 0, 0, 58
		.PaletteColor 5, 177, 0, 0, 0, 59
		.PaletteColor 5, 180, 0, 0, 0, 60
		.PaletteColor 5, 183, 0, 0, 0, 61
		.PaletteColor 5, 186, 0, 0, 0, 62
		.PaletteColor 5, 189, 0, 0, 0, 63
		.PaletteColor 5, 192, 0, 0, 0, 64
		.PaletteColor 5, 195, 0, 0, 0, 65
		.PaletteColor 5, 198, 0, 0, 0, 66
		.PaletteColor 5, 201, 0, 0, 0, 67
		.PaletteColor 5, 204, 0, 0, 0, 68
		.PaletteColor 5, 207, 0, 0, 0, 69
		.PaletteColor 5, 210, 0, 0, 0, 70
		.PaletteColor 5, 213, 0, 0, 0, 71
		.PaletteColor 5, 216, 0, 0, 0, 72
		.PaletteColor 5, 219, 0, 0, 0, 73
		.PaletteColor 5, 222, 0, 0, 0, 74
		.PaletteColor 5, 225, 0, 0, 0, 75
		.PaletteColor 5, 228, 0, 0, 0, 76
		.PaletteColor 5, 231, 0, 0, 0, 77
		.PaletteColor 5, 234, 0, 0, 0, 78
		.PaletteColor 5, 237, 0, 0, 0, 79
		.PaletteColor 5, 240, 0, 0, 0, 80
		.PaletteColor 5, 243, 0, 0, 0, 81
		.PaletteColor 5, 246, 0, 0, 0, 82
		.PaletteColor 5, 249, 0, 0, 0, 83
		.PaletteColor 5, 252, 0, 0, 0, 84
		.PaletteColor 5, 255, 0, 0, 0, 85
		.PaletteColor 5, 255, 3, 0, 0, 86
		.PaletteColor 5, 255, 6, 0, 0, 87
		.PaletteColor 5, 255, 9, 0, 0, 88
		.PaletteColor 5, 255, 12, 0, 0, 89
		.PaletteColor 5, 255, 15, 0, 0, 90
		.PaletteColor 5, 255, 18, 0, 0, 91
		.PaletteColor 5, 255, 21, 0, 0, 92
		.PaletteColor 5, 255, 24, 0, 0, 93
		.PaletteColor 5, 255, 27, 0, 0, 94
		.PaletteColor 5, 255, 30, 0, 0, 95
		.PaletteColor 5, 255, 33, 0, 0, 96
		.PaletteColor 5, 255, 36, 0, 0, 97
		.PaletteColor 5, 255, 39, 0, 0, 98
		.PaletteColor 5, 255, 42, 0, 0, 99
		.PaletteColor 5, 255, 45, 0, 0, 100
		.PaletteColor 5, 255, 48, 0, 0, 101
		.PaletteColor 5, 255, 51, 0, 0, 102
		.PaletteColor 5, 255, 54, 0, 0, 103
		.PaletteColor 5, 255, 57, 0, 0, 104
		.PaletteColor 5, 255, 60, 0, 0, 105
		.PaletteColor 5, 255, 63, 0, 0, 106
		.PaletteColor 5, 255, 66, 0, 0, 107
		.PaletteColor 5, 255, 69, 0, 0, 108
		.PaletteColor 5, 255, 72, 0, 0, 109
		.PaletteColor 5, 255, 75, 0, 0, 110
		.PaletteColor 5, 255, 78, 0, 0, 111
		.PaletteColor 5, 255, 81, 0, 0, 112
		.PaletteColor 5, 255, 84, 0, 0, 113
		.PaletteColor 5, 255, 87, 0, 0, 114
		.PaletteColor 5, 255, 90, 0, 0, 115
		.PaletteColor 5, 255, 93, 0, 0, 116
		.PaletteColor 5, 255, 96, 0, 0, 117
		.PaletteColor 5, 255, 99, 0, 0, 118
		.PaletteColor 5, 255, 102, 0, 0, 119
		.PaletteColor 5, 255, 105, 0, 0, 120
		.PaletteColor 5, 255, 108, 0, 0, 121
		.PaletteColor 5, 255, 111, 0, 0, 122
		.PaletteColor 5, 255, 114, 0, 0, 123
		.PaletteColor 5, 255, 117, 0, 0, 124
		.PaletteColor 5, 255, 120, 0, 0, 125
		.PaletteColor 5, 255, 123, 0, 0, 126
		.PaletteColor 5, 255, 126, 0, 0, 127
		.PaletteColor 5, 255, 129, 0, 0, 128
		.PaletteColor 5, 255, 132, 0, 0, 129
		.PaletteColor 5, 255, 135, 0, 0, 130
		.PaletteColor 5, 255, 138, 0, 0, 131
		.PaletteColor 5, 255, 141, 0, 0, 132
		.PaletteColor 5, 255, 144, 0, 0, 133
		.PaletteColor 5, 255, 147, 0, 0, 134
		.PaletteColor 5, 255, 150, 0, 0, 135
		.PaletteColor 5, 255, 153, 0, 0, 136
		.PaletteColor 5, 255, 156, 0, 0, 137
		.PaletteColor 5, 255, 159, 0, 0, 138
		.PaletteColor 5, 255, 162, 0, 0, 139
		.PaletteColor 5, 255, 165, 0, 0, 140
		.PaletteColor 5, 255, 168, 0, 0, 141
		.PaletteColor 5, 255, 171, 0, 0, 142
		.PaletteColor 5, 255, 174, 0, 0, 143
		.PaletteColor 5, 255, 177, 0, 0, 144
		.PaletteColor 5, 255, 180, 0, 0, 145
		.PaletteColor 5, 255, 183, 0, 0, 146
		.PaletteColor 5, 255, 186, 0, 0, 147
		.PaletteColor 5, 255, 189, 0, 0, 148
		.PaletteColor 5, 255, 192, 0, 0, 149
		.PaletteColor 5, 255, 195, 0, 0, 150
		.PaletteColor 5, 255, 198, 0, 0, 151
		.PaletteColor 5, 255, 201, 0, 0, 152
		.PaletteColor 5, 255, 204, 0, 0, 153
		.PaletteColor 5, 255, 207, 0, 0, 154
		.PaletteColor 5, 255, 210, 0, 0, 155
		.PaletteColor 5, 255, 213, 0, 0, 156
		.PaletteColor 5, 255, 216, 0, 0, 157
		.PaletteColor 5, 255, 219, 0, 0, 158
		.PaletteColor 5, 255, 222, 0, 0, 159
		.PaletteColor 5, 255, 225, 0, 0, 160
		.PaletteColor 5, 255, 228, 0, 0, 161
		.PaletteColor 5, 255, 231, 0, 0, 162
		.PaletteColor 5, 255, 234, 0, 0, 163
		.PaletteColor 5, 255, 237, 0, 0, 164
		.PaletteColor 5, 255, 240, 0, 0, 165
		.PaletteColor 5, 255, 243, 0, 0, 166
		.PaletteColor 5, 255, 246, 0, 0, 167
		.PaletteColor 5, 255, 249, 0, 0, 168
		.PaletteColor 5, 255, 252, 0, 0, 169
		.PaletteColor 5, 255, 255, 0, 0, 170
		.PaletteColor 5, 255, 255, 3, 0, 171
		.PaletteColor 5, 255, 255, 6, 0, 172
		.PaletteColor 5, 255, 255, 9, 0, 173
		.PaletteColor 5, 255, 255, 12, 0, 174
		.PaletteColor 5, 255, 255, 15, 0, 175
		.PaletteColor 5, 255, 255, 18, 0, 176
		.PaletteColor 5, 255, 255, 21, 0, 177
		.PaletteColor 5, 255, 255, 24, 0, 178
		.PaletteColor 5, 255, 255, 27, 0, 179
		.PaletteColor 5, 255, 255, 30, 0, 180
		.PaletteColor 5, 255, 255, 33, 0, 181
		.PaletteColor 5, 255, 255, 36, 0, 182
		.PaletteColor 5, 255, 255, 39, 0, 183
		.PaletteColor 5, 255, 255, 42, 0, 184
		.PaletteColor 5, 255, 255, 45, 0, 185
		.PaletteColor 5, 255, 255, 48, 0, 186
		.PaletteColor 5, 255, 255, 51, 0, 187
		.PaletteColor 5, 255, 255, 54, 0, 188
		.PaletteColor 5, 255, 255, 57, 0, 189
		.PaletteColor 5, 255, 255, 60, 0, 190
		.PaletteColor 5, 255, 255, 63, 0, 191
		.PaletteColor 5, 255, 255, 66, 0, 192
		.PaletteColor 5, 255, 255, 69, 0, 193
		.PaletteColor 5, 255, 255, 72, 0, 194
		.PaletteColor 5, 255, 255, 75, 0, 195
		.PaletteColor 5, 255, 255, 78, 0, 196
		.PaletteColor 5, 255, 255, 81, 0, 197
		.PaletteColor 5, 255, 255, 84, 0, 198
		.PaletteColor 5, 255, 255, 87, 0, 199
		.PaletteColor 5, 255, 255, 90, 0, 200
		.PaletteColor 5, 255, 255, 93, 0, 201
		.PaletteColor 5, 255, 255, 96, 0, 202
		.PaletteColor 5, 255, 255, 99, 0, 203
		.PaletteColor 5, 255, 255, 102, 0, 204
		.PaletteColor 5, 255, 255, 105, 0, 205
		.PaletteColor 5, 255, 255, 108, 0, 206
		.PaletteColor 5, 255, 255, 111, 0, 207
		.PaletteColor 5, 255, 255, 114, 0, 208
		.PaletteColor 5, 255, 255, 117, 0, 209
		.PaletteColor 5, 255, 255, 120, 0, 210
		.PaletteColor 5, 255, 255, 123, 0, 211
		.PaletteColor 5, 255, 255, 126, 0, 212
		.PaletteColor 5, 255, 255, 129, 0, 213
		.PaletteColor 5, 255, 255, 132, 0, 214
		.PaletteColor 5, 255, 255, 135, 0, 215
		.PaletteColor 5, 255, 255, 138, 0, 216
		.PaletteColor 5, 255, 255, 141, 0, 217
		.PaletteColor 5, 255, 255, 144, 0, 218
		.PaletteColor 5, 255, 255, 147, 0, 219
		.PaletteColor 5, 255, 255, 150, 0, 220
		.PaletteColor 5, 255, 255, 153, 0, 221
		.PaletteColor 5, 255, 255, 156, 0, 222
		.PaletteColor 5, 255, 255, 159, 0, 223
		.PaletteColor 5, 255, 255, 162, 0, 224
		.PaletteColor 5, 255, 255, 165, 0, 225
		.PaletteColor 5, 255, 255, 168, 0, 226
		.PaletteColor 5, 255, 255, 171, 0, 227
		.PaletteColor 5, 255, 255, 174, 0, 228
		.PaletteColor 5, 255, 255, 177, 0, 229
		.PaletteColor 5, 255, 255, 180, 0, 230
		.PaletteColor 5, 255, 255, 183, 0, 231
		.PaletteColor 5, 255, 255, 186, 0, 232
		.PaletteColor 5, 255, 255, 189, 0, 233
		.PaletteColor 5, 255, 255, 192, 0, 234
		.PaletteColor 5, 255, 255, 195, 0, 235
		.PaletteColor 5, 255, 255, 198, 0, 236
		.PaletteColor 5, 255, 255, 201, 0, 237
		.PaletteColor 5, 255, 255, 204, 0, 238
		.PaletteColor 5, 255, 255, 207, 0, 239
		.PaletteColor 5, 255, 255, 210, 0, 240
		.PaletteColor 5, 255, 255, 213, 0, 241
		.PaletteColor 5, 255, 255, 216, 0, 242
		.PaletteColor 5, 255, 255, 219, 0, 243
		.PaletteColor 5, 255, 255, 222, 0, 244
		.PaletteColor 5, 255, 255, 225, 0, 245
		.PaletteColor 5, 255, 255, 228, 0, 246
		.PaletteColor 5, 255, 255, 231, 0, 247
		.PaletteColor 5, 255, 255, 234, 0, 248
		.PaletteColor 5, 255, 255, 237, 0, 249
		.PaletteColor 5, 255, 255, 240, 0, 250
		.PaletteColor 5, 255, 255, 243, 0, 251
		.PaletteColor 5, 255, 255, 246, 0, 252
		.PaletteColor 5, 255, 255, 249, 0, 253
		.PaletteColor 5, 255, 255, 252, 0, 254
		.PaletteColor 5, 255, 255, 255, 0, 255
		.EndColorTable 
	
	' Create Text object that Floats above Fire Effect	
	.TextTool 10, 100, FALSE, TRUE ' add Mask Mode 0
		.TextSetting "Fill", "0, 0, 255"
		.TextSetting "Font", "Futura XBlk BT"
		.TextSetting "TypeSize", "48.0"

		'Modified - Hung Tran
		.TextAppend TxtStr2$
		.TextRender 

	.ObjectAlign 3, 3, FALSE, FALSE, FALSE, TRUE, FALSE, TRUE, TRUE
	.EndObject
	
	' Set Frame Delay to 1/10 of a second.
	.MovieFrameRate 10
	.MovieFrameDelay 1, 10, 100
	.EndMovieFrameRate
	
END WITHOBJECT
