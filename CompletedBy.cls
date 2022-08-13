VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Author: Athit Vue
'Created Date: 8/12/2022
'Last Modified: 8/12/2022
'
'Description:
'   This script and the macros: LinkAllCheckBoxesToCell and RemoveAllCheckBoxes helps with managing the Active X
'   checkboxes created on a particular column. When the checkboxes are ticked, the script will store the user's
'   name along with the date and time to the next column over. It will also highlight that cell green.
'   Unticking the checkbox will remove that user's name and turn the cell white.
'
'Prereqs:
'   1. You must first make one checkbox and resize it to your liking.
'   2. You must account for all the checkboxes "Checked" event functions...see below.
'   3. You must run the "LinkedAllCheckBoxesInCell" before and after you create more boxes.
'   4. Any additional checkboxes must be created manually or dragged-copy from the last checkbox and
'      the "LinkedAllCheckBoxesInCell" must be run if you have added many extras. If not, you can go
'      into the Developer tab=>Design Mode=>right-click the new checkbox=>Properties and assign the
'      appropriate cell or range for the LINKEDCELL property.
'   5. You must change the cbCol and completedByCol, column letters, for your appropriate checkbox and
'      "Completed By" columns.
'
'Steps to make this work:
'   1. Create the Active X checkboxes by going to the Developer tab in excel.
'   2. Click insert and select Active X checkboxes.
'   3. Put it in the cell of the column you want.
'   4. With the single checkbox from the top, drag it all the way done to the row you want it to end.
'   5. Below this Sub, add any click even and follow the naming convention.
'   6. Change the value of cbCol to the column that you have set for the checkboxes.
'   7. Now run the "LinkAllCheckBoxesToCell" macro. This will take a while.
'   8. Test the functionality. If the "Completed By" is not in the correct column, look at the offsets here
'      in the main RunCompletedBy Sub and also in the LinkAllCheckBoxesToCell Macro.
'   9. If you messed up with the checkbox index, you can run the RemoveAllCheckboxes and restart the process.
'      That way you don't need to look for a backup sheet.

Sub RunCompletedBy(cBox As Object)
    Dim curRow As String
    Dim checkBoxRange As String
    Dim completedByCol As String
    Dim cbCol As String
    
    cbCol = "E"
    completedByCol = "G" & (cBox.Index + 1)
    curRow = cBox.Index + 1
    
    'This is to ensure that even if we forgot to run the LinkAllCheckBoxesToCell macro,
    'it will still work. This way requires two clicks.
    If cBox.LinkedCell = "" Then
        cBox.LinkedCell = cbCol & (cBox.Index + 1)
    End If
 
    If cBox.LinkedCell = "" Then
        MsgBox ("Alert!!!" & vbCrLf & vbCrLf & "The checkbox you clicked did not linked to a cell properly!" & vbCrLf & vbCrLf & "Its index is: " & cBox.Index & "." & vbCrLf & vbCrLf & "Please enter your name manually in the completed by column.")
    End If
    
    If cBox.LinkedCell <> "" Then
        checkBoxRange = cBox.LinkedCell
        
        range(checkBoxRange).Select
        row = ActiveCell.row
        
        If cBox.Value = True Then
            range(ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 1)).Value = Application.UserName & " on " & Now
            'range(completedByCol).Value = Application.UserName & " on " & Now
            'Cells.Interior.ColorIndex = 0
            range(ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 1)).Interior.ColorIndex = 4
            'range(completedByCol).Interior.ColorIndex = 4
        End If
        
        If cBox.Value = False Then
            range(ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 1)).Value = ""
            'range(completedByCol).Value = ""
            'Cells.Interior.ColorIndex = 0
            range(ActiveCell.Offset(0, 0), ActiveCell.Offset(0, 1)).Interior.ColorIndex = 0
            'range(completedByCol).Interior.ColorIndex = 0
        End If
    
        'range("A1", "G1").Interior.ColorIndex = 46
    End If
End Sub

Private Sub CheckBox1_Click()
'CheckBox1.LinkedCell = "E2"
'range("E2").Select
RunCompletedBy CheckBox1
End Sub
Private Sub CheckBox2_Click()
RunCompletedBy CheckBox2
End Sub
Private Sub CheckBox3_Click()
RunCompletedBy CheckBox3
End Sub
Private Sub CheckBox4_Click()
RunCompletedBy CheckBox4
End Sub
Private Sub CheckBox5_Click()
RunCompletedBy CheckBox5
End Sub
Private Sub CheckBox6_Click()
RunCompletedBy CheckBox6
End Sub
Private Sub CheckBox7_Click()
RunCompletedBy CheckBox7
End Sub
Private Sub CheckBox8_Click()
RunCompletedBy CheckBox8
End Sub
Private Sub CheckBox9_Click()
RunCompletedBy CheckBox9
End Sub
Private Sub CheckBox10_Click()
RunCompletedBy CheckBox10
End Sub
Private Sub CheckBox11_Click()
RunCompletedBy CheckBox11
End Sub
Private Sub CheckBox12_Click()
RunCompletedBy CheckBox12
End Sub
Private Sub CheckBox13_Click()
RunCompletedBy CheckBox13
End Sub
Private Sub CheckBox14_Click()
RunCompletedBy CheckBox14
End Sub
Private Sub CheckBox15_Click()
RunCompletedBy CheckBox15
End Sub
Private Sub CheckBox16_Click()
RunCompletedBy CheckBox16
End Sub
Private Sub CheckBox17_Click()
RunCompletedBy CheckBox17
End Sub
Private Sub CheckBox18_Click()
RunCompletedBy CheckBox18
End Sub
Private Sub CheckBox19_Click()
RunCompletedBy CheckBox19
End Sub
Private Sub CheckBox20_Click()
RunCompletedBy CheckBox20
End Sub
Private Sub CheckBox21_Click()
RunCompletedBy CheckBox21
End Sub
Private Sub CheckBox22_Click()
RunCompletedBy CheckBox22
End Sub
Private Sub CheckBox23_Click()
RunCompletedBy CheckBox23
End Sub
Private Sub CheckBox24_Click()
RunCompletedBy CheckBox24
End Sub
Private Sub CheckBox25_Click()
RunCompletedBy CheckBox25
End Sub
Private Sub CheckBox26_Click()
RunCompletedBy CheckBox26
End Sub
Private Sub CheckBox27_Click()
RunCompletedBy CheckBox27
End Sub
Private Sub CheckBox28_Click()
RunCompletedBy CheckBox28
End Sub
Private Sub CheckBox29_Click()
RunCompletedBy CheckBox29
End Sub
Private Sub CheckBox30_Click()
RunCompletedBy CheckBox30
End Sub
Private Sub CheckBox31_Click()
RunCompletedBy CheckBox31
End Sub
Private Sub CheckBox32_Click()
RunCompletedBy CheckBox32
End Sub
Private Sub CheckBox33_Click()
RunCompletedBy CheckBox33
End Sub
Private Sub CheckBox34_Click()
RunCompletedBy CheckBox34
End Sub
Private Sub CheckBox35_Click()
RunCompletedBy CheckBox35
End Sub
Private Sub CheckBox36_Click()
RunCompletedBy CheckBox36
End Sub
Private Sub CheckBox37_Click()
RunCompletedBy CheckBox37
End Sub
Private Sub CheckBox38_Click()
RunCompletedBy CheckBox38
End Sub
Private Sub CheckBox39_Click()
RunCompletedBy CheckBox39
End Sub
Private Sub CheckBox40_Click()
RunCompletedBy CheckBox40
End Sub
Private Sub CheckBox41_Click()
RunCompletedBy CheckBox41
End Sub
Private Sub CheckBox42_Click()
RunCompletedBy CheckBox42
End Sub
Private Sub CheckBox43_Click()
RunCompletedBy CheckBox43
End Sub
Private Sub CheckBox44_Click()
RunCompletedBy CheckBox44
End Sub
Private Sub CheckBox45_Click()
RunCompletedBy CheckBox45
End Sub
Private Sub CheckBox46_Click()
RunCompletedBy CheckBox46
End Sub
Private Sub CheckBox47_Click()
RunCompletedBy CheckBox47
End Sub
Private Sub CheckBox48_Click()
RunCompletedBy CheckBox48
End Sub
Private Sub CheckBox49_Click()
RunCompletedBy CheckBox49
End Sub
Private Sub CheckBox50_Click()
RunCompletedBy CheckBox50
End Sub
Private Sub CheckBox51_Click()
RunCompletedBy CheckBox51
End Sub
Private Sub CheckBox52_Click()
RunCompletedBy CheckBox52
End Sub
Private Sub CheckBox53_Click()
RunCompletedBy CheckBox53
End Sub
Private Sub CheckBox54_Click()
RunCompletedBy CheckBox54
End Sub
Private Sub CheckBox55_Click()
RunCompletedBy CheckBox55
End Sub
Private Sub CheckBox56_Click()
RunCompletedBy CheckBox56
End Sub
Private Sub CheckBox57_Click()
RunCompletedBy CheckBox57
End Sub
Private Sub CheckBox58_Click()
RunCompletedBy CheckBox58
End Sub
Private Sub CheckBox59_Click()
RunCompletedBy CheckBox59
End Sub
Private Sub CheckBox60_Click()
RunCompletedBy CheckBox60
End Sub
Private Sub CheckBox61_Click()
RunCompletedBy CheckBox61
End Sub
Private Sub CheckBox62_Click()
RunCompletedBy CheckBox62
End Sub
Private Sub CheckBox63_Click()
RunCompletedBy CheckBox63
End Sub
Private Sub CheckBox64_Click()
RunCompletedBy CheckBox64
End Sub
Private Sub CheckBox65_Click()
RunCompletedBy CheckBox65
End Sub
Private Sub CheckBox66_Click()
RunCompletedBy CheckBox66
End Sub
Private Sub CheckBox67_Click()
RunCompletedBy CheckBox67
End Sub
Private Sub CheckBox68_Click()
RunCompletedBy CheckBox68
End Sub
Private Sub CheckBox69_Click()
RunCompletedBy CheckBox69
End Sub
Private Sub CheckBox70_Click()
RunCompletedBy CheckBox70
End Sub
Private Sub CheckBox71_Click()
RunCompletedBy CheckBox71
End Sub
Private Sub CheckBox72_Click()
RunCompletedBy CheckBox72
End Sub
Private Sub CheckBox73_Click()
RunCompletedBy CheckBox73
End Sub
Private Sub CheckBox74_Click()
RunCompletedBy CheckBox74
End Sub
Private Sub CheckBox75_Click()
RunCompletedBy CheckBox75
End Sub
Private Sub CheckBox76_Click()
RunCompletedBy CheckBox76
End Sub
Private Sub CheckBox77_Click()
RunCompletedBy CheckBox77
End Sub
Private Sub CheckBox78_Click()
RunCompletedBy CheckBox78
End Sub
Private Sub CheckBox79_Click()
RunCompletedBy CheckBox79
End Sub
Private Sub CheckBox80_Click()
RunCompletedBy CheckBox80
End Sub
Private Sub CheckBox81_Click()
RunCompletedBy CheckBox81
End Sub
Private Sub CheckBox82_Click()
RunCompletedBy CheckBox82
End Sub
Private Sub CheckBox83_Click()
RunCompletedBy CheckBox83
End Sub
Private Sub CheckBox84_Click()
RunCompletedBy CheckBox84
End Sub
Private Sub CheckBox85_Click()
RunCompletedBy CheckBox85
End Sub
Private Sub CheckBox86_Click()
RunCompletedBy CheckBox86
End Sub
Private Sub CheckBox87_Click()
RunCompletedBy CheckBox87
End Sub
Private Sub CheckBox88_Click()
RunCompletedBy CheckBox88
End Sub
Private Sub CheckBox89_Click()
RunCompletedBy CheckBox89
End Sub
Private Sub CheckBox90_Click()
RunCompletedBy CheckBox90
End Sub
Private Sub CheckBox91_Click()
RunCompletedBy CheckBox91
End Sub
Private Sub CheckBox92_Click()
RunCompletedBy CheckBox92
End Sub
Private Sub CheckBox93_Click()
RunCompletedBy CheckBox93
End Sub
Private Sub CheckBox94_Click()
RunCompletedBy CheckBox94
End Sub
Private Sub CheckBox95_Click()
RunCompletedBy CheckBox95
End Sub
Private Sub CheckBox96_Click()
RunCompletedBy CheckBox96
End Sub
Private Sub CheckBox97_Click()
RunCompletedBy CheckBox97
End Sub
Private Sub CheckBox98_Click()
RunCompletedBy CheckBox98
End Sub
Private Sub CheckBox99_Click()
RunCompletedBy CheckBox99
End Sub
Private Sub CheckBox100_Click()
RunCompletedBy CheckBox100
End Sub
Private Sub CheckBox101_Click()
RunCompletedBy CheckBox101
End Sub
Private Sub CheckBox102_Click()
RunCompletedBy CheckBox102
End Sub
Private Sub CheckBox103_Click()
RunCompletedBy CheckBox103
End Sub
Private Sub CheckBox104_Click()
RunCompletedBy CheckBox104
End Sub
Private Sub CheckBox105_Click()
RunCompletedBy CheckBox105
End Sub
Private Sub CheckBox106_Click()
RunCompletedBy CheckBox106
End Sub
Private Sub CheckBox107_Click()
RunCompletedBy CheckBox107
End Sub
Private Sub CheckBox108_Click()
RunCompletedBy CheckBox108
End Sub
Private Sub CheckBox109_Click()
RunCompletedBy CheckBox109
End Sub
Private Sub CheckBox110_Click()
RunCompletedBy CheckBox110
End Sub
Private Sub CheckBox111_Click()
RunCompletedBy CheckBox111
End Sub
Private Sub CheckBox112_Click()
RunCompletedBy CheckBox112
End Sub
Private Sub CheckBox113_Click()
RunCompletedBy CheckBox113
End Sub
Private Sub CheckBox114_Click()
RunCompletedBy CheckBox114
End Sub
Private Sub CheckBox115_Click()
RunCompletedBy CheckBox115
End Sub
Private Sub CheckBox116_Click()
RunCompletedBy CheckBox116
End Sub
Private Sub CheckBox117_Click()
RunCompletedBy CheckBox117
End Sub
Private Sub CheckBox118_Click()
RunCompletedBy CheckBox118
End Sub
Private Sub CheckBox119_Click()
RunCompletedBy CheckBox119
End Sub
Private Sub CheckBox120_Click()
RunCompletedBy CheckBox120
End Sub
Private Sub CheckBox121_Click()
RunCompletedBy CheckBox121
End Sub
Private Sub CheckBox122_Click()
RunCompletedBy CheckBox122
End Sub
Private Sub CheckBox123_Click()
RunCompletedBy CheckBox123
End Sub
Private Sub CheckBox124_Click()
RunCompletedBy CheckBox124
End Sub
Private Sub CheckBox125_Click()
RunCompletedBy CheckBox125
End Sub
Private Sub CheckBox126_Click()
RunCompletedBy CheckBox126
End Sub
Private Sub CheckBox127_Click()
RunCompletedBy CheckBox127
End Sub
Private Sub CheckBox128_Click()
RunCompletedBy CheckBox128
End Sub
Private Sub CheckBox129_Click()
RunCompletedBy CheckBox129
End Sub
Private Sub CheckBox130_Click()
RunCompletedBy CheckBox130
End Sub
Private Sub CheckBox131_Click()
RunCompletedBy CheckBox131
End Sub
Private Sub CheckBox132_Click()
RunCompletedBy CheckBox132
End Sub
Private Sub CheckBox133_Click()
RunCompletedBy CheckBox133
End Sub
Private Sub CheckBox134_Click()
RunCompletedBy CheckBox134
End Sub
Private Sub CheckBox135_Click()
RunCompletedBy CheckBox135
End Sub
Private Sub CheckBox136_Click()
RunCompletedBy CheckBox136
End Sub
Private Sub CheckBox137_Click()
RunCompletedBy CheckBox137
End Sub
Private Sub CheckBox138_Click()
RunCompletedBy CheckBox138
End Sub
Private Sub CheckBox139_Click()
RunCompletedBy CheckBox139
End Sub
Private Sub CheckBox140_Click()
RunCompletedBy CheckBox140
End Sub
Private Sub CheckBox141_Click()
RunCompletedBy CheckBox141
End Sub
Private Sub CheckBox142_Click()
RunCompletedBy CheckBox142
End Sub
Private Sub CheckBox143_Click()
RunCompletedBy CheckBox143
End Sub
Private Sub CheckBox144_Click()
RunCompletedBy CheckBox144
End Sub
Private Sub CheckBox145_Click()
RunCompletedBy CheckBox145
End Sub
Private Sub CheckBox146_Click()
RunCompletedBy CheckBox146
End Sub
Private Sub CheckBox147_Click()
RunCompletedBy CheckBox147
End Sub
Private Sub CheckBox148_Click()
RunCompletedBy CheckBox148
End Sub
Private Sub CheckBox149_Click()
RunCompletedBy CheckBox149
End Sub
Private Sub CheckBox150_Click()
RunCompletedBy CheckBox150
End Sub
Private Sub CheckBox151_Click()
RunCompletedBy CheckBox151
End Sub
Private Sub CheckBox152_Click()
RunCompletedBy CheckBox152
End Sub
Private Sub CheckBox153_Click()
RunCompletedBy CheckBox153
End Sub
Private Sub CheckBox154_Click()
RunCompletedBy CheckBox154
End Sub
Private Sub CheckBox155_Click()
RunCompletedBy CheckBox155
End Sub
Private Sub CheckBox156_Click()
RunCompletedBy CheckBox156
End Sub
Private Sub CheckBox157_Click()
RunCompletedBy CheckBox157
End Sub
Private Sub CheckBox158_Click()
RunCompletedBy CheckBox158
End Sub
Private Sub CheckBox159_Click()
RunCompletedBy CheckBox159
End Sub
Private Sub CheckBox160_Click()
RunCompletedBy CheckBox160
End Sub
Private Sub CheckBox161_Click()
RunCompletedBy CheckBox161
End Sub
Private Sub CheckBox162_Click()
RunCompletedBy CheckBox162
End Sub
Private Sub CheckBox163_Click()
RunCompletedBy CheckBox163
End Sub
Private Sub CheckBox164_Click()
RunCompletedBy CheckBox164
End Sub
Private Sub CheckBox165_Click()
RunCompletedBy CheckBox165
End Sub
Private Sub CheckBox166_Click()
RunCompletedBy CheckBox166
End Sub
Private Sub CheckBox167_Click()
RunCompletedBy CheckBox167
End Sub
Private Sub CheckBox168_Click()
RunCompletedBy CheckBox168
End Sub
Private Sub CheckBox169_Click()
RunCompletedBy CheckBox169
End Sub
Private Sub CheckBox170_Click()
RunCompletedBy CheckBox170
End Sub
Private Sub CheckBox171_Click()
RunCompletedBy CheckBox171
End Sub
Private Sub CheckBox172_Click()
RunCompletedBy CheckBox172
End Sub
Private Sub CheckBox173_Click()
RunCompletedBy CheckBox173
End Sub
Private Sub CheckBox174_Click()
RunCompletedBy CheckBox174
End Sub
Private Sub CheckBox175_Click()
RunCompletedBy CheckBox175
End Sub
Private Sub CheckBox176_Click()
RunCompletedBy CheckBox176
End Sub
Private Sub CheckBox177_Click()
RunCompletedBy CheckBox177
End Sub
Private Sub CheckBox178_Click()
RunCompletedBy CheckBox178
End Sub
Private Sub CheckBox179_Click()
RunCompletedBy CheckBox179
End Sub
Private Sub CheckBox180_Click()
RunCompletedBy CheckBox180
End Sub
Private Sub CheckBox181_Click()
RunCompletedBy CheckBox181
End Sub
Private Sub CheckBox182_Click()
RunCompletedBy CheckBox182
End Sub
Private Sub CheckBox183_Click()
RunCompletedBy CheckBox183
End Sub
Private Sub CheckBox184_Click()
RunCompletedBy CheckBox184
End Sub
Private Sub CheckBox185_Click()
RunCompletedBy CheckBox185
End Sub
Private Sub CheckBox186_Click()
RunCompletedBy CheckBox186
End Sub
Private Sub CheckBox187_Click()
RunCompletedBy CheckBox187
End Sub
Private Sub CheckBox188_Click()
RunCompletedBy CheckBox188
End Sub
Private Sub CheckBox189_Click()
RunCompletedBy CheckBox189
End Sub
Private Sub CheckBox190_Click()
RunCompletedBy CheckBox190
End Sub
Private Sub CheckBox191_Click()
RunCompletedBy CheckBox191
End Sub
Private Sub CheckBox192_Click()
RunCompletedBy CheckBox192
End Sub
Private Sub CheckBox193_Click()
RunCompletedBy CheckBox193
End Sub
Private Sub CheckBox194_Click()
RunCompletedBy CheckBox194
End Sub
Private Sub CheckBox195_Click()
RunCompletedBy CheckBox195
End Sub
Private Sub CheckBox196_Click()
RunCompletedBy CheckBox196
End Sub
Private Sub CheckBox197_Click()
RunCompletedBy CheckBox197
End Sub
Private Sub CheckBox198_Click()
RunCompletedBy CheckBox198
End Sub
Private Sub CheckBox199_Click()
RunCompletedBy CheckBox199
End Sub
Private Sub CheckBox200_Click()
RunCompletedBy CheckBox200
End Sub
Private Sub CheckBox201_Click()
RunCompletedBy CheckBox201
End Sub
Private Sub CheckBox202_Click()
RunCompletedBy CheckBox202
End Sub
Private Sub CheckBox203_Click()
RunCompletedBy CheckBox203
End Sub
Private Sub CheckBox204_Click()
RunCompletedBy CheckBox204
End Sub
Private Sub CheckBox205_Click()
RunCompletedBy CheckBox205
End Sub
Private Sub CheckBox206_Click()
RunCompletedBy CheckBox206
End Sub
Private Sub CheckBox207_Click()
RunCompletedBy CheckBox207
End Sub
Private Sub CheckBox208_Click()
RunCompletedBy CheckBox208
End Sub
Private Sub CheckBox209_Click()
RunCompletedBy CheckBox209
End Sub
Private Sub CheckBox210_Click()
RunCompletedBy CheckBox210
End Sub
Private Sub CheckBox211_Click()
RunCompletedBy CheckBox211
End Sub
Private Sub CheckBox212_Click()
RunCompletedBy CheckBox212
End Sub
Private Sub CheckBox213_Click()
RunCompletedBy CheckBox213
End Sub
Private Sub CheckBox214_Click()
RunCompletedBy CheckBox214
End Sub
Private Sub CheckBox215_Click()
RunCompletedBy CheckBox215
End Sub
Private Sub CheckBox216_Click()
RunCompletedBy CheckBox216
End Sub
Private Sub CheckBox217_Click()
RunCompletedBy CheckBox217
End Sub
Private Sub CheckBox218_Click()
RunCompletedBy CheckBox218
End Sub
Private Sub CheckBox219_Click()
RunCompletedBy CheckBox219
End Sub
Private Sub CheckBox220_Click()
RunCompletedBy CheckBox220
End Sub
Private Sub CheckBox221_Click()
RunCompletedBy CheckBox221
End Sub
Private Sub CheckBox222_Click()
RunCompletedBy CheckBox222
End Sub
Private Sub CheckBox223_Click()
RunCompletedBy CheckBox223
End Sub
Private Sub CheckBox224_Click()
RunCompletedBy CheckBox224
End Sub
Private Sub CheckBox225_Click()
RunCompletedBy CheckBox225
End Sub
Private Sub CheckBox226_Click()
RunCompletedBy CheckBox226
End Sub
Private Sub CheckBox227_Click()
RunCompletedBy CheckBox227
End Sub
Private Sub CheckBox228_Click()
RunCompletedBy CheckBox228
End Sub
Private Sub CheckBox229_Click()
RunCompletedBy CheckBox229
End Sub
Private Sub CheckBox230_Click()
RunCompletedBy CheckBox230
End Sub
Private Sub CheckBox231_Click()
RunCompletedBy CheckBox231
End Sub
Private Sub CheckBox232_Click()
RunCompletedBy CheckBox232
End Sub
Private Sub CheckBox233_Click()
RunCompletedBy CheckBox233
End Sub
Private Sub CheckBox234_Click()
RunCompletedBy CheckBox234
End Sub
Private Sub CheckBox235_Click()
RunCompletedBy CheckBox235
End Sub
Private Sub CheckBox236_Click()
RunCompletedBy CheckBox236
End Sub
Private Sub CheckBox237_Click()
RunCompletedBy CheckBox237
End Sub
Private Sub CheckBox238_Click()
RunCompletedBy CheckBox238
End Sub
Private Sub CheckBox239_Click()
RunCompletedBy CheckBox239
End Sub
Private Sub CheckBox240_Click()
RunCompletedBy CheckBox240
End Sub
Private Sub CheckBox241_Click()
RunCompletedBy CheckBox241
End Sub
Private Sub CheckBox242_Click()
RunCompletedBy CheckBox242
End Sub
Private Sub CheckBox243_Click()
RunCompletedBy CheckBox243
End Sub
Private Sub CheckBox244_Click()
RunCompletedBy CheckBox244
End Sub
Private Sub CheckBox245_Click()
RunCompletedBy CheckBox245
End Sub
Private Sub CheckBox246_Click()
RunCompletedBy CheckBox246
End Sub
Private Sub CheckBox247_Click()
RunCompletedBy CheckBox247
End Sub
Private Sub CheckBox248_Click()
RunCompletedBy CheckBox248
End Sub
Private Sub CheckBox249_Click()
RunCompletedBy CheckBox249
End Sub
Private Sub CheckBox250_Click()
RunCompletedBy CheckBox250
End Sub
Private Sub CheckBox251_Click()
RunCompletedBy CheckBox251
End Sub
Private Sub CheckBox252_Click()
RunCompletedBy CheckBox252
End Sub
Private Sub CheckBox253_Click()
RunCompletedBy CheckBox253
End Sub
Private Sub CheckBox254_Click()
RunCompletedBy CheckBox254
End Sub
Private Sub CheckBox255_Click()
RunCompletedBy CheckBox255
End Sub
Private Sub CheckBox256_Click()
RunCompletedBy CheckBox256
End Sub
Private Sub CheckBox257_Click()
RunCompletedBy CheckBox257
End Sub
Private Sub CheckBox258_Click()
RunCompletedBy CheckBox258
End Sub
Private Sub CheckBox259_Click()
RunCompletedBy CheckBox259
End Sub
Private Sub CheckBox260_Click()
RunCompletedBy CheckBox260
End Sub
Private Sub CheckBox261_Click()
RunCompletedBy CheckBox261
End Sub
Private Sub CheckBox262_Click()
RunCompletedBy CheckBox262
End Sub
Private Sub CheckBox263_Click()
RunCompletedBy CheckBox263
End Sub
Private Sub CheckBox264_Click()
RunCompletedBy CheckBox264
End Sub
Private Sub CheckBox265_Click()
RunCompletedBy CheckBox265
End Sub
Private Sub CheckBox266_Click()
RunCompletedBy CheckBox266
End Sub
Private Sub CheckBox267_Click()
RunCompletedBy CheckBox267
End Sub
Private Sub CheckBox268_Click()
RunCompletedBy CheckBox268
End Sub
Private Sub CheckBox269_Click()
RunCompletedBy CheckBox269
End Sub
Private Sub CheckBox270_Click()
RunCompletedBy CheckBox270
End Sub
Private Sub CheckBox271_Click()
RunCompletedBy CheckBox271
End Sub
Private Sub CheckBox272_Click()
RunCompletedBy CheckBox272
End Sub
Private Sub CheckBox273_Click()
RunCompletedBy CheckBox273
End Sub
Private Sub CheckBox274_Click()
RunCompletedBy CheckBox274
End Sub
Private Sub CheckBox275_Click()
RunCompletedBy CheckBox275
End Sub
Private Sub CheckBox276_Click()
RunCompletedBy CheckBox276
End Sub
Private Sub CheckBox277_Click()
RunCompletedBy CheckBox277
End Sub
Private Sub CheckBox278_Click()
RunCompletedBy CheckBox278
End Sub
Private Sub CheckBox279_Click()
RunCompletedBy CheckBox279
End Sub
Private Sub CheckBox280_Click()
RunCompletedBy CheckBox280
End Sub
Private Sub CheckBox281_Click()
RunCompletedBy CheckBox281
End Sub
Private Sub CheckBox282_Click()
RunCompletedBy CheckBox282
End Sub
Private Sub CheckBox283_Click()
RunCompletedBy CheckBox283
End Sub
Private Sub CheckBox284_Click()
RunCompletedBy CheckBox284
End Sub
Private Sub CheckBox285_Click()
RunCompletedBy CheckBox285
End Sub
Private Sub CheckBox286_Click()
RunCompletedBy CheckBox286
End Sub
Private Sub CheckBox287_Click()
RunCompletedBy CheckBox287
End Sub
Private Sub CheckBox288_Click()
RunCompletedBy CheckBox288
End Sub
Private Sub CheckBox289_Click()
RunCompletedBy CheckBox289
End Sub
Private Sub CheckBox290_Click()
RunCompletedBy CheckBox290
End Sub
Private Sub CheckBox291_Click()
RunCompletedBy CheckBox291
End Sub
Private Sub CheckBox292_Click()
RunCompletedBy CheckBox292
End Sub
Private Sub CheckBox293_Click()
RunCompletedBy CheckBox293
End Sub
Private Sub CheckBox294_Click()
RunCompletedBy CheckBox294
End Sub
Private Sub CheckBox295_Click()
RunCompletedBy CheckBox295
End Sub
Private Sub CheckBox296_Click()
RunCompletedBy CheckBox296
End Sub
Private Sub CheckBox297_Click()
RunCompletedBy CheckBox297
End Sub
Private Sub CheckBox298_Click()
RunCompletedBy CheckBox298
End Sub
Private Sub CheckBox299_Click()
RunCompletedBy CheckBox299
End Sub
Private Sub CheckBox300_Click()
RunCompletedBy CheckBox300
End Sub
Private Sub CheckBox301_Click()
RunCompletedBy CheckBox301
End Sub
Private Sub CheckBox302_Click()
RunCompletedBy CheckBox302
End Sub
Private Sub CheckBox303_Click()
RunCompletedBy CheckBox303
End Sub
Private Sub CheckBox304_Click()
RunCompletedBy CheckBox304
End Sub
Private Sub CheckBox305_Click()
RunCompletedBy CheckBox305
End Sub
Private Sub CheckBox306_Click()
RunCompletedBy CheckBox306
End Sub
Private Sub CheckBox307_Click()
RunCompletedBy CheckBox307
End Sub
Private Sub CheckBox308_Click()
RunCompletedBy CheckBox308
End Sub
Private Sub CheckBox309_Click()
RunCompletedBy CheckBox309
End Sub
Private Sub CheckBox310_Click()
RunCompletedBy CheckBox310
End Sub
Private Sub CheckBox311_Click()
RunCompletedBy CheckBox311
End Sub
Private Sub CheckBox312_Click()
RunCompletedBy CheckBox312
End Sub
Private Sub CheckBox313_Click()
RunCompletedBy CheckBox313
End Sub
Private Sub CheckBox314_Click()
RunCompletedBy CheckBox314
End Sub
Private Sub CheckBox315_Click()
RunCompletedBy CheckBox315
End Sub
Private Sub CheckBox316_Click()
RunCompletedBy CheckBox316
End Sub
Private Sub CheckBox317_Click()
RunCompletedBy CheckBox317
End Sub
Private Sub CheckBox318_Click()
RunCompletedBy CheckBox318
End Sub
Private Sub CheckBox319_Click()
RunCompletedBy CheckBox319
End Sub
Private Sub CheckBox320_Click()
RunCompletedBy CheckBox320
End Sub
Private Sub CheckBox321_Click()
RunCompletedBy CheckBox321
End Sub
Private Sub CheckBox322_Click()
RunCompletedBy CheckBox322
End Sub
Private Sub CheckBox323_Click()
RunCompletedBy CheckBox323
End Sub
Private Sub CheckBox324_Click()
RunCompletedBy CheckBox324
End Sub
Private Sub CheckBox325_Click()
RunCompletedBy CheckBox325
End Sub
Private Sub CheckBox326_Click()
RunCompletedBy CheckBox326
End Sub
Private Sub CheckBox327_Click()
RunCompletedBy CheckBox327
End Sub
Private Sub CheckBox328_Click()
RunCompletedBy CheckBox328
End Sub
Private Sub CheckBox329_Click()
RunCompletedBy CheckBox329
End Sub
Private Sub CheckBox330_Click()
RunCompletedBy CheckBox330
End Sub
Private Sub CheckBox331_Click()
RunCompletedBy CheckBox331
End Sub
Private Sub CheckBox332_Click()
RunCompletedBy CheckBox332
End Sub
Private Sub CheckBox333_Click()
RunCompletedBy CheckBox333
End Sub
Private Sub CheckBox334_Click()
RunCompletedBy CheckBox334
End Sub
Private Sub CheckBox335_Click()
RunCompletedBy CheckBox335
End Sub
Private Sub CheckBox336_Click()
RunCompletedBy CheckBox336
End Sub
Private Sub CheckBox337_Click()
RunCompletedBy CheckBox337
End Sub
Private Sub CheckBox338_Click()
RunCompletedBy CheckBox338
End Sub
Private Sub CheckBox339_Click()
RunCompletedBy CheckBox339
End Sub
Private Sub CheckBox340_Click()
RunCompletedBy CheckBox340
End Sub
Private Sub CheckBox341_Click()
RunCompletedBy CheckBox341
End Sub
Private Sub CheckBox342_Click()
RunCompletedBy CheckBox342
End Sub
Private Sub CheckBox343_Click()
RunCompletedBy CheckBox343
End Sub
Private Sub CheckBox344_Click()
RunCompletedBy CheckBox344
End Sub
Private Sub CheckBox345_Click()
RunCompletedBy CheckBox345
End Sub
Private Sub CheckBox346_Click()
RunCompletedBy CheckBox346
End Sub
Private Sub CheckBox347_Click()
RunCompletedBy CheckBox347
End Sub
Private Sub CheckBox348_Click()
RunCompletedBy CheckBox348
End Sub
Private Sub CheckBox349_Click()
RunCompletedBy CheckBox349
End Sub
Private Sub CheckBox350_Click()
RunCompletedBy CheckBox350
End Sub
Private Sub CheckBox351_Click()
RunCompletedBy CheckBox351
End Sub
Private Sub CheckBox352_Click()
RunCompletedBy CheckBox352
End Sub
Private Sub CheckBox353_Click()
RunCompletedBy CheckBox353
End Sub
Private Sub CheckBox354_Click()
RunCompletedBy CheckBox354
End Sub
Private Sub CheckBox355_Click()
RunCompletedBy CheckBox355
End Sub
Private Sub CheckBox356_Click()
RunCompletedBy CheckBox356
End Sub
Private Sub CheckBox357_Click()
RunCompletedBy CheckBox357
End Sub
Private Sub CheckBox358_Click()
RunCompletedBy CheckBox358
End Sub
Private Sub CheckBox359_Click()
RunCompletedBy CheckBox359
End Sub
Private Sub CheckBox360_Click()
RunCompletedBy CheckBox360
End Sub
Private Sub CheckBox361_Click()
RunCompletedBy CheckBox361
End Sub
Private Sub CheckBox362_Click()
RunCompletedBy CheckBox362
End Sub
Private Sub CheckBox363_Click()
RunCompletedBy CheckBox363
End Sub
Private Sub CheckBox364_Click()
RunCompletedBy CheckBox364
End Sub
Private Sub CheckBox365_Click()
RunCompletedBy CheckBox365
End Sub
Private Sub CheckBox366_Click()
RunCompletedBy CheckBox366
End Sub
Private Sub CheckBox367_Click()
RunCompletedBy CheckBox367
End Sub
Private Sub CheckBox368_Click()
RunCompletedBy CheckBox368
End Sub
Private Sub CheckBox369_Click()
RunCompletedBy CheckBox369
End Sub
Private Sub CheckBox370_Click()
RunCompletedBy CheckBox370
End Sub
Private Sub CheckBox371_Click()
RunCompletedBy CheckBox371
End Sub
Private Sub CheckBox372_Click()
RunCompletedBy CheckBox372
End Sub
Private Sub CheckBox373_Click()
RunCompletedBy CheckBox373
End Sub
Private Sub CheckBox374_Click()
RunCompletedBy CheckBox374
End Sub
Private Sub CheckBox375_Click()
RunCompletedBy CheckBox375
End Sub
Private Sub CheckBox376_Click()
RunCompletedBy CheckBox376
End Sub
Private Sub CheckBox377_Click()
RunCompletedBy CheckBox377
End Sub
Private Sub CheckBox378_Click()
RunCompletedBy CheckBox378
End Sub
Private Sub CheckBox379_Click()
RunCompletedBy CheckBox379
End Sub
Private Sub CheckBox380_Click()
RunCompletedBy CheckBox380
End Sub
Private Sub CheckBox381_Click()
RunCompletedBy CheckBox381
End Sub
Private Sub CheckBox382_Click()
RunCompletedBy CheckBox382
End Sub
Private Sub CheckBox383_Click()
RunCompletedBy CheckBox383
End Sub
Private Sub CheckBox384_Click()
RunCompletedBy CheckBox384
End Sub
Private Sub CheckBox385_Click()
RunCompletedBy CheckBox385
End Sub
Private Sub CheckBox386_Click()
RunCompletedBy CheckBox386
End Sub
Private Sub CheckBox387_Click()
RunCompletedBy CheckBox387
End Sub
Private Sub CheckBox388_Click()
RunCompletedBy CheckBox388
End Sub
Private Sub CheckBox389_Click()
RunCompletedBy CheckBox389
End Sub
Private Sub CheckBox390_Click()
RunCompletedBy CheckBox390
End Sub
Private Sub CheckBox391_Click()
RunCompletedBy CheckBox391
End Sub
Private Sub CheckBox392_Click()
RunCompletedBy CheckBox392
End Sub
Private Sub CheckBox393_Click()
RunCompletedBy CheckBox393
End Sub
Private Sub CheckBox394_Click()
RunCompletedBy CheckBox394
End Sub
Private Sub CheckBox395_Click()
RunCompletedBy CheckBox395
End Sub
Private Sub CheckBox396_Click()
RunCompletedBy CheckBox396
End Sub
Private Sub CheckBox397_Click()
RunCompletedBy CheckBox397
End Sub
Private Sub CheckBox398_Click()
RunCompletedBy CheckBox398
End Sub
Private Sub CheckBox399_Click()
RunCompletedBy CheckBox399
End Sub
Private Sub CheckBox400_Click()
RunCompletedBy CheckBox400
End Sub
Private Sub CheckBox401_Click()
RunCompletedBy CheckBox401
End Sub
Private Sub CheckBox402_Click()
RunCompletedBy CheckBox402
End Sub
Private Sub CheckBox403_Click()
RunCompletedBy CheckBox403
End Sub
