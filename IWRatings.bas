Attribute VB_Name = "IWRatings"
Public WB_Name As String
Public Tot_Plrs As Variant
Public TotPLib As Integer
Public NewPlayer As String
Public RChrt(10, 5)
Public RC_Num As Integer
Public PlrRslt_Row As Integer
Public League As String
Public Season As String
Public Card_Type As String
Public Rpt_Tm_Dot As Boolean
Public Chk_Plr_Clb As Boolean
Public Num_Divs As Integer
Public Divs(9, 3) As Variant
Public NR_Comp As String
Public Jump_Rpt As Integer
Public Fixt_Note As String
Public Nxt_Fix_Col As Integer
Public MacDelay As String
Public Const mdlu = "JanFebMarAprMayJunJulAugSepOctNovDec"
Public pd As Variant
Public GA1_1 As Variant
Public GA1_2 As Variant
Public GA1_3 As Variant
Public GA1_4 As Variant
Public GA1_5 As Variant
Public GA1_6 As Variant
Public GA2_1 As Variant
Public GA2_2 As Variant
Public GA2_3 As Variant
Public GA2_4 As Variant
Public GA2_5 As Variant
Public GA2_6 As Variant
Public GA3_1 As Variant
Public GA3_2 As Variant
Public GA3_3 As Variant
Public GA3_4 As Variant
Public GA3_5 As Variant
Public GA3_6 As Variant
Public GA4_1 As Variant
Public GA4_2 As Variant
Public GA4_3 As Variant
Public GA4_4 As Variant
Public GA4_5 As Variant
Public GA4_6 As Variant
Public GA5_1 As Variant
Public GA5_2 As Variant
Public GA5_3 As Variant
Public GA5_4 As Variant
Public GA5_5 As Variant
Public GA5_6 As Variant
Public GA6_1 As Variant
Public GA6_2 As Variant
Public GA6_3 As Variant
Public GA6_4 As Variant
Public GA6_5 As Variant
Public GA6_6 As Variant
Public GA7_1 As Variant
Public GA7_2 As Variant
Public GA7_3 As Variant
Public GA7_4 As Variant
Public GA7_5 As Variant
Public GA7_6 As Variant
Public GA8_1 As Variant
Public GA8_2 As Variant
Public GA8_3 As Variant
Public GA8_4 As Variant
Public GA8_5 As Variant
Public GA8_6 As Variant
Public GA9_1 As Variant
Public GA9_2 As Variant
Public GA9_3 As Variant
Public GA9_4 As Variant
Public GA9_5 As Variant
Public GA9_6 As Variant
Public GA10_1 As Variant
Public GA10_2 As Variant
Public GA10_3 As Variant
Public GA10_4 As Variant
Public GA10_5 As Variant
Public GA10_6 As Variant
Private Scratch_G1 As Boolean
Private Scratch_G2 As Boolean
Private Scratch_G3 As Boolean
Private Scratch_G4 As Boolean
Private Scratch_G5 As Boolean
Private Scratch_G6 As Boolean
Private Scratch_G7 As Boolean
Private Scratch_G8 As Boolean
Private Scratch_G9 As Boolean
Private Scratch_G10 As Boolean
Private oMatchHelper As MatchHelper
Private oTournament As TournamentHelper
Private oFixtureScraper As FixtureScraper
Private oFixtureHelper As FixtureHelper


Sub Plyr_Lib()
 showing = Sheets("Player Library").Visible
 plyr_txt = IIf(showing, "Show", "Hide") & " Player LIbrary"
 Sheets("Menu").Select
 ActiveSheet.Unprotect
 ActiveSheet.Shapes("Plyr_Lib").Select
 Selection.Characters.text = plyr_txt
 Sheets("Menu").Select
 ActiveSheet.Protect
 If showing Then
  Sheets("Player Library").Select
  Range("A2").Select
  ActiveWindow.SelectedSheets.Visible = False
  Sheets("Menu").Select
 Else
  Sheets("Player Library").Visible = True
  Sheets("Player Library").Select
  Range("A2").Select
  End If
End Sub
Sub V_Rpt1()
Sheets("R1_Jumpers").Select
ActiveWindow.SelectedSheets.PrintPreview
Sheets("Menu").Select
End Sub
Sub V_Rpt2()
Sheets("R2_RtgByPlyr").Select
ActiveWindow.SelectedSheets.PrintPreview
Sheets("Menu").Select
End Sub
Sub V_Rpt3()
Sheets("R3_RtgByRtg").Select
ActiveWindow.SelectedSheets.PrintPreview
Sheets("Menu").Select
End Sub
Sub V_Rpt4()
Sheets("R4_PlyrStats").Select
ActiveWindow.SelectedSheets.PrintPreview
Sheets("Menu").Select
End Sub

Sub Run_Reports()
tmp = GetParms(False, False)
tmp = Rem_Rpts

'R1_Jumpers

Sheets.Add.Name = "R1_Jumpers"
Sheets("R1_Jumpers").Tab.ColorIndex = 5
ActiveWindow.DisplayGridlines = False
Sheets("R1_Jumpers").Move After:=Sheets("Config")
Sheets("Players").Select
Columns("G:G").Select
Selection.Copy
Sheets("R1_Jumpers").Select
Columns("A:A").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("A1").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "MOVT"
Sheets("Players").Select
Columns("E:F").Select
Selection.Copy
Sheets("R1_Jumpers").Select
Columns("B:B").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("B1").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "START"
Range("C1").Select
ActiveCell.FormulaR1C1 = "CURR"
Sheets("Players").Select
Columns("B:B").Select
Selection.Copy
Sheets("R1_Jumpers").Select
Columns("D:D").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("D1").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "PLAYER"
Range("E1").Select
ActiveCell.FormulaR1C1 = "TEAM"
Range("E2").Select
Sheets("Players").Select
lastrow = ActiveSheet.UsedRange.Rows.Count
Sheets("R1_Jumpers").Select
For x = 2 To lastrow
If Rpt_Tm_Dot Then
  Cells(x, 5) = Left(Sheets("Players").Cells(x, 3), 3) & "." & Right(Sheets("Players").Cells(x, 3), 1) & ":" & Sheets("Players").Cells(x, 4)
 Else
  Cells(x, 5) = Sheets("Players").Cells(x, 3) & ":" & Sheets("Players").Cells(x, 4)
 End If
Next x
Sheets("Players").Select
Columns("O:R").Select
Selection.Copy
Sheets("R1_Jumpers").Select
Columns("F:I").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("F1").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "PLYD"
Range("G1").Select
ActiveCell.FormulaR1C1 = "WON"
Range("H1").Select
ActiveCell.FormulaR1C1 = "LOST"
Range("I1").Select
ActiveCell.FormulaR1C1 = "PCT"
Sheets("Players").Select
Columns("H:H").Select
Selection.Copy
Sheets("R1_Jumpers").Select
Columns("J:J").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("J1").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "PSEAS"
lastrow = ActiveSheet.UsedRange.Rows.Count
For x = 2 To lastrow
 If Trim(Cells(x, 10)) <> "Yes" Then
  Cells(x, 10) = "*DEL*"
 Else
  If Cells(x, 1) <= 0 Or Cells(x, 6) = 0 Then
   Cells(x, 10) = "*DEL*"
  Else
   Cells(x, 9).NumberFormat = "0.00%"
  End If
 End If
Next x
For x = lastrow To 2 Step -1
 If Cells(x, 10) = "*DEL*" Then
  Rows(x & ":" & x).Select
  Selection.Delete Shift:=xlUp
 End If
Next x
Columns("J:J").Select
Selection.Delete Shift:=xlToLeft
Range("A1").Select
Cells.Select
Selection.Sort Key1:=Range("A2"), Order1:=xlDescending, Key2:=Range("F2") _
                                                              , Order2:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:=False _
                                                                                                                               , Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:= _
               xlSortNormal
Rows("1:1").Select
Selection.Insert Shift:=xlDown
Range("B1").Select
ActiveCell.FormulaR1C1 = "RATING"
Range("F1").Select
ActiveCell.FormulaR1C1 = "STATS ALL GAMES PLAYED"
Range("B1:C1").Select
Selection.MergeCells = True
Range("F1:I1").Select
Selection.MergeCells = True
Range("A1:I2").Select
Selection.Font.Bold = True
Range("A2:I2").Select
Selection.Font.Underline = xlUnderlineStyleSingleAccounting
Columns("A:C").Select
Selection.HorizontalAlignment = xlCenter
Columns("E:H").Select
Selection.HorizontalAlignment = xlCenter
Columns("H:H").Select
Selection.HorizontalAlignment = xlRight
Range("D2").Select
Selection.HorizontalAlignment = xlCenter
Range("I2").Select
Selection.HorizontalAlignment = xlCenter
lastrow = ActiveSheet.UsedRange.Rows.Count
For x = lastrow To Jump_Rpt + 3 Step -1
 Rows(x & ":" & x).Select
 Selection.Delete Shift:=xlUp
Next x
Cells.Select
Selection.Columns.AutoFit
Columns("F:I").Select
Range("F2").Activate
Selection.ColumnWidth = 10
Range("A1").Select
lft_hdr = "&""Arial,Bold""&10 " & Season
rht_hdr = "&""Arial,Bold""&10&D"
ctr_hdr = "&""Arial,Bold""&12 " & League & Chr(10) & "&""Arial,Bold""&12&U " & IIf(lastrow < (Jump_Rpt + 2), "", "Top " & Jump_Rpt & " ") & "Rating Jumpers "
ctr_ftr = "&""Arial,Bold""&12PLAYERS FIRST RATED CURRENT SEASON NOT INCLUDED"
tmp = PrintSetup("$A$1:$I$", lft_hdr, ctr_hdr, rht_hdr, ctr_ftr, True)
Range("A1").Select
Sheets("Players").Select
Range("A1").Select

'R2_RtgByPlyr

Sheets.Add.Name = "R2_RtgByPlyr"
Sheets("R2_RtgByPlyr").Tab.ColorIndex = 5
ActiveWindow.DisplayGridlines = False
Sheets("R2_RtgByPlyr").Move After:=Sheets("R1_Jumpers")
Sheets("Players").Select
Columns("B:F").Select
Selection.Copy
Sheets("R2_RtgByPlyr").Select
Columns("A:A").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                                                :=False, Transpose:=False
Sheets("Players").Select
Range("A1").Select
Sheets("R2_RtgByPlyr").Select
Cells.Select
Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes, _
               OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
               DataOption1:=xlSortNormal
lastrow = ActiveSheet.UsedRange.Rows.Count
For x = 2 To lastrow
 If Rpt_Tm_Dot Then
  Cells(x, 4) = Left(Cells(x, 2), 3) & "." & Right(Cells(x, 2), 1) & ":" & Cells(x, 3)
 Else
  Cells(x, 4) = Cells(x, 2) & ":" & Cells(x, 3)
 End If
Next x
Columns("B:C").Select
Selection.Delete Shift:=xlToLeft
Cells(1, 1) = "Player"
Cells(1, 2) = "Team"
Cells(1, 3) = "Rtg"
Range("A1:C1").Font.Bold = True
Columns("B:B").HorizontalAlignment = xlCenter
Cells(1, 3).HorizontalAlignment = xlCenter
Range("A1:C" & lastrow).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range("A1:C" & lastrow).Borders(xlEdgeLeft).Weight = xlMedium
Range("A1:C" & lastrow).Borders(xlEdgeLeft).ColorIndex = 5
Range("A1:C" & lastrow).Borders(xlEdgeTop).LineStyle = xlContinuous
Range("A1:C" & lastrow).Borders(xlEdgeTop).Weight = xlMedium
Range("A1:C" & lastrow).Borders(xlEdgeTop).ColorIndex = 5
Range("A1:C" & lastrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A1:C" & lastrow).Borders(xlEdgeBottom).Weight = xlMedium
Range("A1:C" & lastrow).Borders(xlEdgeBottom).ColorIndex = 5
Range("A1:C" & lastrow).Borders(xlEdgeRight).LineStyle = xlContinuous
Range("A1:C" & lastrow).Borders(xlEdgeRight).Weight = xlMedium
Range("A1:C" & lastrow).Borders(xlEdgeRight).ColorIndex = 5
Range("A1:C" & lastrow).Borders(xlInsideVertical).LineStyle = xlContinuous
Range("A1:C" & lastrow).Borders(xlInsideVertical).Weight = xlThin
Range("A1:C" & lastrow).Borders(xlInsideVertical).ColorIndex = 5
Range("A1:C1").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A1:C1").Borders(xlEdgeBottom).Weight = xlMedium
Range("A1:C1").Borders(xlEdgeBottom).ColorIndex = 5
Range("A1").Select
Select Case lastrow
 Case Is <= 60
  numcols = 1
  prange = "$A$1:$C$"
 Case Is <= 120
  numcols = 2
  prange = "$A$1:$G$"
  Range("A1:C1").Copy
  Range("E1:G1").Select
  ActiveSheet.Paste
  Application.CutCopyMode = False
 Case Is <= 225
  numcols = 3
  prange = "$A$1:$K$"
  Range("A1:C1").Copy
  Range("E1:G1").Select
  ActiveSheet.Paste
  Range("I1:K1").Select
  ActiveSheet.Paste
  Application.CutCopyMode = False
 Case Else
  numcols = 4
  prange = "$A$1:$O$"
  Range("A1:C1").Copy
  Range("E1:G1").Select
  ActiveSheet.Paste
  Range("I1:K1").Select
  ActiveSheet.Paste
  Range("M1:O1").Select
  ActiveSheet.Paste
  Application.CutCopyMode = False
End Select
Range("A1").Select
SplitRow = Round(((lastrow) / numcols) + 1, 0) + 1
Range("A" & SplitRow + 1 & ":C" & lastrow).Cut
Range("E2:G2").Select
ActiveSheet.Paste
Range("A" & SplitRow & ":C" & SplitRow).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A" & SplitRow & ":C" & SplitRow).Borders(xlEdgeBottom).Weight = xlMedium
Range("A" & SplitRow & ":C" & SplitRow).Borders(xlEdgeBottom).ColorIndex = 5
lastrow = ActiveSheet.UsedRange.Rows.Count
If lastrow > SplitRow Then
 Range("E" & SplitRow + 1 & ":G" & lastrow).Cut
 Range("I2:K2").Select
 ActiveSheet.Paste
 Range("E" & SplitRow & ":G" & SplitRow).Borders(xlEdgeBottom).LineStyle = xlContinuous
 Range("E" & SplitRow & ":G" & SplitRow).Borders(xlEdgeBottom).Weight = xlMedium
 Range("E" & SplitRow & ":G" & SplitRow).Borders(xlEdgeBottom).ColorIndex = 5
End If
lastrow = ActiveSheet.UsedRange.Rows.Count
If lastrow > SplitRow Then
 Range("I" & SplitRow + 1 & ":K" & lastrow).Cut
 Range("M2:O2").Select
 ActiveSheet.Paste
 Range("I" & SplitRow & ":K" & SplitRow).Borders(xlEdgeBottom).LineStyle = xlContinuous
 Range("I" & SplitRow & ":K" & SplitRow).Borders(xlEdgeBottom).Weight = xlMedium
 Range("I" & SplitRow & ":K" & SplitRow).Borders(xlEdgeBottom).ColorIndex = 5
End If
Range("A1").Select
lft_hdr = "&""Arial,Bold""&10 " & Season
rht_hdr = "&""Arial,Bold""&10&D"
ctr_hdr = "&""Arial,Bold""&12 " & League & Chr(10) & "&""Arial,Bold""&12&U " & "Ratings by Player"
ctr_ftr = ""
tmp = PrintSetup(prange, lft_hdr, ctr_hdr, rht_hdr, ctr_ftr, True)
Range("A1").Select

'R3_RtgByRtg

Sheets.Add.Name = "R3_RtgByRtg"
Sheets("R3_RtgByRtg").Tab.ColorIndex = 5
ActiveWindow.DisplayGridlines = False
Sheets("R3_RtgByRtg").Move After:=Sheets("R2_RtgByPlyr")
Sheets("Players").Select
Columns("B:F").Select
Selection.Copy
Sheets("R3_RtgByRtg").Select
Columns("A:A").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                                                :=False, Transpose:=False
Sheets("Players").Select
Range("A1").Select
Sheets("R3_RtgByRtg").Select
Columns("E:E").Cut
Columns("A:A").Insert Shift:=xlToRight
Cells.Select
Selection.Sort Key1:=Range("A2"), Order1:=xlDescending, Header:=xlYes, _
               OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
               DataOption1:=xlSortNormal
lastrow = ActiveSheet.UsedRange.Rows.Count
For x = 2 To lastrow
 If Rpt_Tm_Dot Then
  Cells(x, 5) = Left(Cells(x, 3), 3) & "." & Right(Cells(x, 3), 1) & ":" & Cells(x, 4)
 Else
  Cells(x, 5) = Cells(x, 3) & ":" & Cells(x, 4)
 End If
Next x
Columns("C:D").Select
Selection.Delete Shift:=xlToLeft
Cells(1, 1) = "Rtg"
Cells(1, 2) = "Player"
Cells(1, 3) = "Team"
Range("A1:C1").Font.Bold = True
Columns("C:C").HorizontalAlignment = xlCenter
Cells(1, 1).HorizontalAlignment = xlCenter
Range("A1:C" & lastrow).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range("A1:C" & lastrow).Borders(xlEdgeLeft).Weight = xlMedium
Range("A1:C" & lastrow).Borders(xlEdgeLeft).ColorIndex = 5
Range("A1:C" & lastrow).Borders(xlEdgeTop).LineStyle = xlContinuous
Range("A1:C" & lastrow).Borders(xlEdgeTop).Weight = xlMedium
Range("A1:C" & lastrow).Borders(xlEdgeTop).ColorIndex = 5
Range("A1:C" & lastrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A1:C" & lastrow).Borders(xlEdgeBottom).Weight = xlMedium
Range("A1:C" & lastrow).Borders(xlEdgeBottom).ColorIndex = 5
Range("A1:C" & lastrow).Borders(xlEdgeRight).LineStyle = xlContinuous
Range("A1:C" & lastrow).Borders(xlEdgeRight).Weight = xlMedium
Range("A1:C" & lastrow).Borders(xlEdgeRight).ColorIndex = 5
Range("A1:C" & lastrow).Borders(xlInsideVertical).LineStyle = xlContinuous
Range("A1:C" & lastrow).Borders(xlInsideVertical).Weight = xlThin
Range("A1:C" & lastrow).Borders(xlInsideVertical).ColorIndex = 5
Range("A1:C1").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A1:C1").Borders(xlEdgeBottom).Weight = xlMedium
Range("A1:C1").Borders(xlEdgeBottom).ColorIndex = 5
Range("A1").Select
Select Case lastrow
 Case Is <= 60
  numcols = 1
  prange = "$A$1:$C$"
 Case Is <= 120
  numcols = 2
  prange = "$A$1:$G$"
  Range("A1:C1").Copy
  Range("E1:G1").Select
  ActiveSheet.Paste
  Application.CutCopyMode = False
 Case Is <= 225
  numcols = 3
  prange = "$A$1:$K$"
  Range("A1:C1").Copy
  Range("E1:G1").Select
  ActiveSheet.Paste
  Range("I1:K1").Select
  ActiveSheet.Paste
  Application.CutCopyMode = False
 Case Else
  numcols = 4
  prange = "$A$1:$O$"
  Range("A1:C1").Copy
  Range("E1:G1").Select
  ActiveSheet.Paste
  Range("I1:K1").Select
  ActiveSheet.Paste
  Range("M1:O1").Select
  ActiveSheet.Paste
  Application.CutCopyMode = False
End Select
Range("A1").Select
SplitRow = Round(((lastrow) / numcols) + 1, 0) + 1
Range("A" & SplitRow + 1 & ":C" & lastrow).Cut
Range("E2:G2").Select
ActiveSheet.Paste
Range("A" & SplitRow & ":C" & SplitRow).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A" & SplitRow & ":C" & SplitRow).Borders(xlEdgeBottom).Weight = xlMedium
Range("A" & SplitRow & ":C" & SplitRow).Borders(xlEdgeBottom).ColorIndex = 5
lastrow = ActiveSheet.UsedRange.Rows.Count
If lastrow > SplitRow Then
 Range("E" & SplitRow + 1 & ":G" & lastrow).Cut
 Range("I2:K2").Select
 ActiveSheet.Paste
 Range("E" & SplitRow & ":G" & SplitRow).Borders(xlEdgeBottom).LineStyle = xlContinuous
 Range("E" & SplitRow & ":G" & SplitRow).Borders(xlEdgeBottom).Weight = xlMedium
 Range("E" & SplitRow & ":G" & SplitRow).Borders(xlEdgeBottom).ColorIndex = 5
End If
lastrow = ActiveSheet.UsedRange.Rows.Count
If lastrow > SplitRow Then
 Range("I" & SplitRow + 1 & ":K" & lastrow).Cut
 Range("M2:O2").Select
 ActiveSheet.Paste
 Range("I" & SplitRow & ":K" & SplitRow).Borders(xlEdgeBottom).LineStyle = xlContinuous
 Range("I" & SplitRow & ":K" & SplitRow).Borders(xlEdgeBottom).Weight = xlMedium
 Range("I" & SplitRow & ":K" & SplitRow).Borders(xlEdgeBottom).ColorIndex = 5
End If
Range("A1").Select
lft_hdr = "&""Arial,Bold""&10 " & Season
rht_hdr = "&""Arial,Bold""&10&D"
ctr_hdr = "&""Arial,Bold""&12 " & League & Chr(10) & "&""Arial,Bold""&12&U " & "Ratings by Rating"
ctr_ftr = ""
tmp = PrintSetup(prange, lft_hdr, ctr_hdr, rht_hdr, ctr_ftr, True)
Range("A1").Select

'R4_PlyrStats

Sheets.Add.Name = "R4_PlyrStats"
Sheets("R4_PlyrStats").Tab.ColorIndex = 5
ActiveWindow.DisplayGridlines = False
Sheets("R4_PlyrStats").Move After:=Sheets("R3_RtgByRtg")
Sheets("Players").Select
Columns("B:R").Select
Selection.Copy
Sheets("R4_PlyrStats").Select
Columns("A:A").Select
ActiveSheet.Paste
Selection.Interior.ColorIndex = xlNone
Columns("D:D").Select
Application.CutCopyMode = False
Selection.Insert Shift:=xlToRight
Sheets("Players").Select
Range("A1").Select
Sheets("R4_PlyrStats").Select
Cells.Select
Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes, _
               OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
               DataOption1:=xlSortNormal
lastrow = ActiveSheet.UsedRange.Rows.Count
For x = 2 To lastrow
 If Rpt_Tm_Dot Then
  Cells(x, 4) = Left(Cells(x, 2), 3) & "." & Right(Cells(x, 2), 1) & ":" & Cells(x, 3)
 Else
  Cells(x, 4) = Cells(x, 2) & ":" & Cells(x, 3)
 End If
Next x
Cells(1, 1) = "Player"
Cells(1, 4) = "Team"
Range("A1").Select
Columns("Q:Q").Delete Shift:=xlToLeft
Columns("N:N").Delete Shift:=xlToLeft
Columns("L:L").Delete Shift:=xlToLeft
Columns("H:I").Delete Shift:=xlToLeft
Columns("B:C").Delete Shift:=xlToLeft
Range("A1:K" & lastrow).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range("A1:K" & lastrow).Borders(xlEdgeLeft).Weight = xlMedium
Range("A1:K" & lastrow).Borders(xlEdgeLeft).ColorIndex = 5
Range("A1:K" & lastrow).Borders(xlEdgeTop).LineStyle = xlContinuous
Range("A1:K" & lastrow).Borders(xlEdgeTop).Weight = xlMedium
Range("A1:K" & lastrow).Borders(xlEdgeTop).ColorIndex = 5
Range("A1:K" & lastrow).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A1:K" & lastrow).Borders(xlEdgeBottom).Weight = xlMedium
Range("A1:K" & lastrow).Borders(xlEdgeBottom).ColorIndex = 5
Range("A1:K" & lastrow).Borders(xlEdgeRight).LineStyle = xlContinuous
Range("A1:K" & lastrow).Borders(xlEdgeRight).Weight = xlMedium
Range("A1:K" & lastrow).Borders(xlEdgeRight).ColorIndex = 5
Range("A1:K" & lastrow).Borders(xlInsideVertical).LineStyle = xlContinuous
Range("A1:K" & lastrow).Borders(xlInsideVertical).Weight = xlThin
Range("A1:K" & lastrow).Borders(xlInsideVertical).ColorIndex = 5
Range("A1:K" & lastrow).Borders(xlInsideHorizontal).LineStyle = xlContinuous
Range("A1:K" & lastrow).Borders(xlInsideHorizontal).Weight = xlThin
Range("A1:K" & lastrow).Borders(xlInsideHorizontal).ColorIndex = 5
Range("A1:K1").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A1:K1").Borders(xlEdgeBottom).Weight = xlMedium
Range("A1:K1").Borders(xlEdgeBottom).ColorIndex = 5
Range("A1").Select
With ActiveSheet.PageSetup
 .PrintTitleRows = "$1:$1"
 .PrintTitleColumns = ""
End With
lft_hdr = "&""Arial,Bold""&10 " & Season
rht_hdr = "&""Arial,Bold""&10&D"
ctr_hdr = "&""Arial,Bold""&12 " & League & Chr(10) & "&""Arial,Bold""&12&U " & "Player Statistics"
ctr_ftr = ""
tmp = PrintSetup("$A$1:$K$", lft_hdr, ctr_hdr, rht_hdr, ctr_ftr, False)
ActiveSheet.PageSetup.PrintArea = "$A$1:$K$" & lastrow
Range("A1").Select


Sheets("Players").Select
Range("A1").Select
Sheets("Menu").Select
ActiveSheet.Unprotect
ActiveSheet.Shapes("V_Rpt1_Box").TextFrame.Characters.Font.ColorIndex = 1
ActiveSheet.Shapes("V_Rpt1_Box").OnAction = "V_Rpt1"
ActiveSheet.Shapes("V_Rpt2_Box").TextFrame.Characters.Font.ColorIndex = 1
ActiveSheet.Shapes("V_Rpt2_Box").OnAction = "V_Rpt2"
ActiveSheet.Shapes("V_Rpt3_Box").TextFrame.Characters.Font.ColorIndex = 1
ActiveSheet.Shapes("V_Rpt3_Box").OnAction = "V_Rpt3"
ActiveSheet.Shapes("V_Rpt4_Box").TextFrame.Characters.Font.ColorIndex = 1
ActiveSheet.Shapes("V_Rpt4_Box").OnAction = "V_Rpt4"
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
tmp = Tidy_Up("Run_Reports_Info")
End Sub

Sub Get_Results()
tmp = GetParms(True, True)
' Get Results
tmp = Rem_Rpts
Application.DisplayAlerts = False
For Each ws In Worksheets
 If ws.Name = "To_Be_Rated" Then
  Sheets("To_Be_Rated").Delete
 End If
Next ws
Application.DisplayAlerts = True
Sheets("PlyrRslt").Select
lastrow = ActiveSheet.UsedRange.Rows.Count
If lastrow > 1 Then
 Rows("2:" & lastrow).Select
 Application.CutCopyMode = False
 Selection.Delete Shift:=xlUp
End If
Range("A2").Select
Sheets("Fixtures").Select
Cells.Select
Application.CutCopyMode = False
Selection.Clear
Range("A3").Select
Sheets("Results").Select
tot_res = ActiveSheet.UsedRange.Rows.Count
rsel = "A2:Z" & tot_res
Range(rsel).Interior.ColorIndex = 36
Range("A2").Select
Nxt_Fix_Col = 1
Rslt_Txt = ""

Set oFixtureScraper = New FixtureScraper
Set oFixtureHelper = New FixtureHelper
Set oMatchHelper = New MatchHelper
Dim fixtureSheet As Worksheet
Set fixtureSheet = Worksheets("Fixtures")
fixtureSheet.Cells.Clear

Dim divFixtures As Collection

ProgressDialog.Show
ProgressDialog.SetProgress "Processing...", 0

For x = 1 To Num_Divs
' tmp = GetFixts(Divs(x, 1), Divs(x, 3))
' lastrow = ActiveSheet.UsedRange.Rows.Count
' Nxt_Fix_Col = PutFixts(Divs(x, 1), lastrow, False)

 ProgressDialog.SetProgress "Processing Division " + CStr(Divs(x, 1)), x / Num_Divs
 Set divFixtures = oFixtureScraper.GetFixtures(CStr(Divs(x, 3)))
 lastrow = divFixtures.Count + 2
 oFixtureHelper.AppendFixtures fixtureSheet, CInt(Divs(x, 1)), divFixtures
 
 For Each oFixture In divFixtures
  lookup_week = IIf(Len(oFixture.WeekNumber) = 1, "0", "") & oFixture.WeekNumber
  lookup_val = lookup_week & "-" & Divs(x, 1) & "-" & Team(oFixture.HomeTeam) & "-" & Team(oFixture.AwayTeam)

  If oFixture.MatchCardUrl <> "" Then
   
   ' Check if result needed
   For Z = 2 To tot_res
    If Sheets("Results").Cells(Z, 1) = lookup_val Then
     If oFixture.MatchScore <> Trim(Sheets("Results").Cells(Z, 8)) Then
      Sheets("Results").Cells(Z, 1).Interior.ColorIndex = 4
      rslt_os = True
     Else
      rslt_os = False
     End If
     
     If rslt_os Then
      restype = IIf(Trim(Sheets("Results").Cells(Z, 8)) = "", "New : ", "Amended : ")
      Rslt_Txt = Rslt_Txt & restype & lookup_val & vbCr
      
      Dim oMatchScraper As MatchScraper
      Set oMatchScraper = New MatchScraper
        
      Dim url As String
      url = oFixture.MatchCardUrl
        
      Dim oMatchResult As MatchResult
      Set oMatchResult = oMatchScraper.GetMatchResults(url)
      oMatchResult.Key = lookup_val
      oMatchResult.WeekNumber = oFixture.WeekNumber
      oMatchResult.Division = Divs(x, 1)
      
      If oMatchHelper.ValidateMatch(oFixture.MatchScore, oMatchResult) Then
        Call oMatchHelper.UpdateMatchResult(oMatchResult, Z)
      ElseIf MsgBox("Invalid Result Detected: " & oMatchResult.Key & vbNewLine & " Result has not been added. Continue processing?", vbYesNo) = vbNo Then
        Unload ProgressDialog
        Exit Sub
      End If
     End If
    End If
   Next Z
  End If
 Next
Next x

Unload ProgressDialog

Call UpdateTournamentResults
Sheets("Results").Select
Cells.Select
Selection.Columns.AutoFit
Range("A1").Select
' Check for Awaiting Match Results
Call Await_Res

Call Chk_Plyrs
Sheets("Results").Select
Range("A1").Select
Sheets("Menu").Select
ActiveSheet.Unprotect
ActiveSheet.Shapes("Upd_Stats_Info").Select
Selection.Characters.text = " *** NEEDS UPDATING ***"
ActiveSheet.Shapes("Upd_Stats_Box").TextFrame.Characters.Font.ColorIndex = 1
ActiveSheet.Shapes("Upd_Stats_Box").OnAction = "Upd_Stats"
ActiveSheet.Shapes("Run_Reports_Info").Select
Selection.Characters.text = "*RUN UPDATE STATS 1st*"
ActiveSheet.Shapes("Run_Reports_Box").TextFrame.Characters.Font.ColorIndex = 16
ActiveSheet.Shapes("Run_Reports_Box").OnAction = ""
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
tmp = Tidy_Up("OS_Res_Info")
' check Fixtures
Call Chk_fixts
' Message if Results found
If Rslt_Txt = "" Then
 Rslt_Txt = "No New/Amended Results Detected"
Else
 Rslt_Txt = "New/Amended Results Detected" & vbCr & vbCr & Rslt_Txt
End If
 MsgBox (Rslt_Txt)
' Message if new players
If NewPlayer <> "" Then
 Sheets("Players").Select
 Range("A2").Select
 NewPlayer = "New Player(s) Found:-" & vbCr & vbCr & NewPlayer & vbCr & "Update Players Worksheet"
 MsgBox (NewPlayer)
End If
End Sub

Sub UpdateTournamentResults()
    Set oTournament = New TournamentHelper
    Call oTournament.CopyResults
End Sub


Sub Chk_fixts()
 Application.DisplayAlerts = False
 For Each ws In Worksheets
  Select Case ws.Name
   Case "Fxt_Check"
    Sheets("Fxt_Check").Delete
  End Select
 Next ws
 Application.DisplayAlerts = True
 tmp = GetParms(True, False)
 shw_msg = False
 Sheets("tmp").Select
 ActiveWindow.DisplayGridlines = True
 Range("A1").Select
 ActiveCell.FormulaR1C1 = "KEY"
 Range("B1").Select
 ActiveCell.FormulaR1C1 = "WkNo"
 Range("C1").Select
 ActiveCell.FormulaR1C1 = "FORMAT"
 Range("D1").Select
 ActiveCell.FormulaR1C1 = "DIVISION"
 Range("E1").Select
 ActiveCell.FormulaR1C1 = "MATCH_DATE"
 Range("F1").Select
 ActiveCell.FormulaR1C1 = "HOME_TEAM"
 Range("G1").Select
 ActiveCell.FormulaR1C1 = "AWAY_TEAM"
 Range("H1").Select
 ActiveCell.FormulaR1C1 = "RES"
 Range("I1").Select
 ActiveCell.FormulaR1C1 = "COMMENT"
 Range("J1").Select
 ActiveCell.FormulaR1C1 = "? REASON"
 tmprow = 1
 Sheets("Results").Select
 Columns("BY:BY").Select
 Selection.ClearContents
 Range("A2").Select
 tot_res = ActiveSheet.UsedRange.Rows.Count
 lu_range = "Results!A1:A" & tot_res
 Range("A2").Select
 Sheets("Fixtures").Select
 For x = 1 To Num_Divs
  fxtcol = ((x - 1) * 7) + 1
  fxtrow = 3
  While Trim(Cells(fxtrow, fxtcol)) <> ""
   txt = ""
   w_val = Trim(Cells(fxtrow, fxtcol))
   WkNo = IIf(Len(w_val) = 1, "0", "") & w_val
   h_val = Trim(Cells(fxtrow, fxtcol + 1))
   a_val = Trim(Cells(fxtrow, fxtcol + 3))
   d_val = DateValue(Trim(Cells(fxtrow, fxtcol + 4)))
   r_val = Trim(Cells(fxtrow, fxtcol + 5))
   lookup_val = WkNo & "-" & Divs(x, 1) & "-" & Team(h_val) & "-" & Team(a_val)
   lu_row = Application.Match(lookup_val, Range(lu_range), 0)
   If IsError(lu_row) Then
    lu_row = 0
    txt = "Match not in Results Tab"
   Else
    Sheets("Results").Cells(lu_row, 77) = "OK"
    If r_val = "" And Trim(Sheets("Results").Cells(lu_row, 8)) <> "" Then
     txt = "Results Tab Row: " & lu_row & " has a Result - But does not in Fixtures Tab"
     r_val = Trim(Sheets("Results").Cells(lu_row, 8))
    End If
   End If
   If txt <> "" Then
    If lu_row <> 0 Then
     Sheets("Results").Cells(lu_row, 1).Interior.ColorIndex = 38
    End If
    Sheets("Fixtures").Range(Cells(fxtrow, fxtcol), Cells(fxtrow, fxtcol + 4)).Interior.ColorIndex = 38
    tmprow = tmprow + 1
    Sheets("tmp").Cells(tmprow, 1) = lookup_val
    Sheets("tmp").Cells(tmprow, 2) = w_val
    Sheets("tmp").Cells(tmprow, 3) = Divs(x, 2)
    Sheets("tmp").Cells(tmprow, 4) = Divs(x, 1)
    Sheets("tmp").Cells(tmprow, 5) = d_val
    Sheets("tmp").Cells(tmprow, 6) = h_val
    Sheets("tmp").Cells(tmprow, 7) = a_val
    Sheets("tmp").Cells(tmprow, 8) = r_val
    Sheets("tmp").Cells(tmprow, 9) = txt
   End If
    fxtrow = fxtrow + 1
  Wend
 Next x
 Sheets("Results").Select
 For x = 2 To tot_res
  If Cells(x, 77) <> "OK" Then
   If Cells(x, 4) <> "T" Then 'Tounament result
    Cells(x, 1).Interior.ColorIndex = 38
    tmprow = tmprow + 1
    Sheets("tmp").Cells(tmprow, 1) = Cells(x, 1)
    Sheets("tmp").Cells(tmprow, 2) = Cells(x, 2)
    Sheets("tmp").Cells(tmprow, 3) = Cells(x, 3)
    Sheets("tmp").Cells(tmprow, 4) = Cells(x, 4)
    Sheets("tmp").Cells(tmprow, 5) = Cells(x, 5)
    Sheets("tmp").Cells(tmprow, 6) = Cells(x, 6)
    Sheets("tmp").Cells(tmprow, 7) = Cells(x, 7)
    Sheets("tmp").Cells(tmprow, 8) = Cells(x, 8)
    Sheets("tmp").Cells(tmprow, 9) = "Results Tab Row: " & x & " not in Fixtures Tab"
   End If
  End If
 Next x
 Columns("BY:BY").Select
 Selection.ClearContents
 Range("A2").Select
 If tmprow > 1 Then
  For fxc = 2 To tmprow
   ht_wd = False
   at_wd = False
   h_val = Trim(Sheets("tmp").Cells(fxc, 6))
   ht_wd = WD_Team(h_val)
   a_val = Trim(Sheets("tmp").Cells(fxc, 7))
   at_wd = WD_Team(a_val)
   If ht_wd Then
     Sheets("tmp").Cells(fxc, 10) = Sheets("tmp").Cells(fxc, 10) & " (" & h_val & " WITHDRAWN)"
    End If
    If at_wd Then
     Sheets("tmp").Cells(fxc, 10) = Sheets("tmp").Cells(fxc, 10) & " (" & a_val & " WITHDRAWN)"
    End If
   If Not (ht_wd Or at_wd) Then
     shw_msg = True
    End If
  
  Next fxc
 
 
 
 
 
 
 
 
 
 
  Sheets("tmp").Select
  Sheets("tmp").Name = "Fxt_Check"
  Sheets("Fxt_Check").Select
  Sheets("Fxt_Check").Tab.ColorIndex = 3
  Range("A2").Select
  sel_rng = "A1:J" & tmprow
  Range(sel_rng).Borders(xlEdgeLeft).LineStyle = xlContinuous
  Range(sel_rng).Borders(xlEdgeTop).LineStyle = xlContinuous
  Range(sel_rng).Borders(xlEdgeBottom).LineStyle = xlContinuous
  Range(sel_rng).Borders(xlEdgeRight).LineStyle = xlContinuous
  Range(sel_rng).Borders(xlInsideVertical).LineStyle = xlContinuous
  Range(sel_rng).Borders(xlInsideHorizontal).LineStyle = xlContinuous
  Columns("E:E").Select
  Selection.NumberFormat = "d-mmm-yy"
  Columns("A:E").HorizontalAlignment = xlCenter
  Columns("H:H").HorizontalAlignment = xlCenter
  Range("A1:J1").Font.Bold = True
  Selection.Font.Bold = True
  Range("A1:J1").HorizontalAlignment = xlCenter
  Range("A1:J1").Interior.ColorIndex = 15
  Range("A2").Select
  ActiveWindow.FreezePanes = True
  Cells.Select
  Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes, _
                    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                    DataOption1:=xlSortNormal
  Selection.Columns.AutoFit
  Range("A2").Select
  Sheets("Menu").Select
  Range("A1").Select
  If shw_msg Then
   MsgBox ("There are Anomolies between the Results & Fixtures Tabs" & vbCr & vbCr & "Refer to the Fxt_Check tab and amend Results Tab where necessary")
  End If
 Else
  Application.DisplayAlerts = False
  Sheets("tmp").Delete
  Application.DisplayAlerts = True
 End If
  Sheets("Menu").Select
  Range("A1").Select
End Sub




Sub Chk_Plyrs()
Sheets("Players").Select
Tot_Plrs = ActiveSheet.UsedRange.Rows.Count
For x = 2 To Tot_Plrs
 Cells(x, 6) = Cells(x, 5)
 Cells(x, 7) = 0
 Cells(x, 9) = 0
 Cells(x, 10) = 0
 Cells(x, 11) = 0
 Cells(x, 12) = 0
 Cells(x, 13) = 0
 Cells(x, 14) = 0
 Cells(x, 15) = 0
 Cells(x, 16) = 0
 Cells(x, 17) = 0
 Cells(x, 18) = 0
Next x
TotPLib = Sheets("Player Library").UsedRange.Rows.Count
Sheets("Results").Select
tot_res = ActiveSheet.UsedRange.Rows.Count
NewPlayer = ""
For x = 2 To tot_res
 If Trim(Cells(x, 8)) <> "" Then
  NewPlayer = NewPlayer & Pl_Exist(Cells(x, 9), Mid(Cells(x, 1), 6, 4), Mid(Cells(x, 1), 4, 1))
  NewPlayer = NewPlayer & Pl_Exist(Cells(x, 10), Mid(Cells(x, 1), 6, 4), Mid(Cells(x, 1), 4, 1))
  NewPlayer = NewPlayer & Pl_Exist(Cells(x, 11), Mid(Cells(x, 1), 6, 4), Mid(Cells(x, 1), 4, 1))
  NewPlayer = NewPlayer & Pl_Exist(Cells(x, 12), Mid(Cells(x, 1), 11, 4), Mid(Cells(x, 1), 4, 1))
  NewPlayer = NewPlayer & Pl_Exist(Cells(x, 13), Mid(Cells(x, 1), 11, 4), Mid(Cells(x, 1), 4, 1))
  NewPlayer = NewPlayer & Pl_Exist(Cells(x, 14), Mid(Cells(x, 1), 11, 4), Mid(Cells(x, 1), 4, 1))
 End If
Next x
Sheets("Players").Select
Cells.Select
Selection.Sort Key1:=Range("B2"), Order1:=xlAscending, Header:=xlYes, _
                OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                DataOption1:=xlSortNormal
Selection.Columns.AutoFit
Columns("E:R").Select
Selection.ColumnWidth = 10
Tot_Plrs = ActiveSheet.UsedRange.Rows.Count
For c = 2 To Tot_Plrs
 Cells(c, 4) = MV(Cells(c, 4), "V")
Next c
Range("A2").Select
End Sub
Function MV(v As Variant, t As String)
' t=V for Value
' t=T or anything else for Text
If t = "V" And (Val(Trim(v)) * 2 <> 0) Then
 MV = Val(Trim(v))
Else
 MV = IIf(VarType(v) = vbString, v, CStr(Val(Trim(v))))
End If

End Function
Sub Await_Res()
' Check for Awaiting Match Results
Sheets("tmp").Select
ActiveWindow.FreezePanes = False
Cells.Select
Selection.Clear
Columns("D:D").Select
Selection.NumberFormat = "d-mmm-yy"
Range("A1").Select
Sheets("Results").Select
tot_res = ActiveSheet.UsedRange.Rows.Count
cnt = 0
For x = 2 To tot_res
 If Cells(x, 5) <= Date Then
  If Trim(Cells(x, 8)) = "" Then
   cnt = cnt + 1
   Sheets("tmp").Cells(cnt, 1) = Trim(Cells(x, 4)) & Trim(Cells(x, 2)) & Trim(Cells(x, 6)) & Trim(Cells(x, 7))
   Sheets("tmp").Cells(cnt, 2) = Cells(x, 4)
   Sheets("tmp").Cells(cnt, 3) = Cells(x, 2)
   Sheets("tmp").Cells(cnt, 4) = Cells(x, 5)
   Sheets("tmp").Cells(cnt, 5) = Cells(x, 6)
   Sheets("tmp").Cells(cnt, 6) = Cells(x, 7)
  End If
 End If
Next x
Sheets("tmp").Select
Cells.Select
Selection.Columns.AutoFit
Selection.Sort Key1:=Range("B1"), Order1:=xlAscending, Key2:=Range("D1") _
                                                             , Order2:=xlAscending, Key3:=Range("E1"), Order3:=xlAscending, Header:= _
               xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
               DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, DataOption3:= _
               xlSortNormal

For x = 1 To cnt
 If Cells(x, 4) = Date Then
  Range("B" & x & ":G" & x).Interior.ColorIndex = 35
 Else
  Range("B" & x & ":G" & x).Interior.ColorIndex = 36
 End If
Next x
If cnt > 0 Then
Range("B1:G" & cnt).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range("B1:G" & cnt).Borders(xlEdgeLeft).Weight = xlThin
Range("B1:G" & cnt).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
Range("B1:G" & cnt).Borders(xlEdgeTop).LineStyle = xlContinuous
Range("B1:G" & cnt).Borders(xlEdgeTop).Weight = xlThin
Range("B1:G" & cnt).Borders(xlEdgeTop).ColorIndex = xlAutomatic
Range("B1:G" & cnt).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("B1:G" & cnt).Borders(xlEdgeBottom).Weight = xlThin
Range("B1:G" & cnt).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
Range("B1:G" & cnt).Borders(xlEdgeRight).LineStyle = xlContinuous
Range("B1:G" & cnt).Borders(xlEdgeRight).Weight = xlThin
Range("B1:G" & cnt).Borders(xlEdgeRight).ColorIndex = xlAutomatic
Range("B1:G" & cnt).Borders(xlInsideVertical).LineStyle = xlContinuous
Range("B1:G" & cnt).Borders(xlInsideVertical).Weight = xlThin
Range("B1:G" & cnt).Borders(xlInsideVertical).ColorIndex = xlAutomatic
If cnt > 1 Then
Range("B1:G" & cnt).Borders(xlInsideHorizontal).LineStyle = xlContinuous
Range("B1:G" & cnt).Borders(xlInsideHorizontal).Weight = xlThin
Range("B1:G" & cnt).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
End If
End If
Range("A1").Select
If Trim(Cells(1, 1)) = "" Then
 tot_res = 0
Else
 tot_res = ActiveSheet.UsedRange.Rows.Count
End If
cnt = 5
While Trim(Sheets("Menu").Cells(cnt, 10)) <> ""
 lur = Trim(Sheets("Menu").Cells(cnt, 10)) _
       & Trim(Sheets("Menu").Cells(cnt, 11)) & Trim(Sheets("Menu").Cells(cnt, 13)) _
       & Trim(Sheets("Menu").Cells(cnt, 14))
 For x = 1 To tot_res
  If lur = Cells(x, 1) Then
   Cells(x, 7) = Sheets("Menu").Cells(cnt, 15)
   If InStr(UCase(Cells(x, 7)), "PLAYED") = 1 Then
    Range("B" & x & ":G" & x).Interior.ColorIndex = 38
   End If
   Exit For
  End If
 Next x
 cnt = cnt + 1
Wend
Sheets("tmp").Select
Cells.Select
Selection.Columns.AutoFit
Range("A1").Select
Sheets("Menu").Select
Sheets("Menu").Unprotect
If cnt > 5 Then
 Range("J5:O" & (cnt - 1)).Delete Shift:=xlUp
End If
Sheets("tmp").Select
If tot_res > 0 Then
 Range("B1:G" & tot_res).Copy
 Sheets("Menu").Select
 Range("J5").Select
 ActiveSheet.Paste
Else
 Sheets("Menu").Select
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
x = 5
While Trim(Cells(x, 10)) <> ""
 osm_com = Cells(x, 15)
 com_add = ""
 fxtcol = 1
 For y = 1 To Num_Divs
  If Cells(x, 10) & "~" = Divs(y, 1) & "~" Then
   fxtcol = 1 + ((y - 1) * 7)
   Exit For
  End If
 Next y
 y = 3
 Do While Trim(Sheets("Fixtures").Cells(y, fxtcol)) <> ""
  If Cells(x, 11) = Sheets("Fixtures").Cells(y, fxtcol) And _
     Cells(x, 13) = Sheets("Fixtures").Cells(y, fxtcol + 1) And _
     Cells(x, 14) = Sheets("Fixtures").Cells(y, fxtcol + 3) _
  Then
   Select Case Sheets("Fixtures").Cells(y, fxtcol + 2)
    Case "P"
     com_add = "[Postponed] "
    Case "R"
     com_add = "[Re-Arranged] "
   End Select
   Exit Do
  End If
  y = y + 1
 Loop
 osm_com = Trim(Replace(osm_com, Trim(com_add), ""))
 Cells(x, 15) = Trim(com_add & osm_com)
 x = x + 1
Wend
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Columns("J:O").Select
Selection.Columns.AutoFit
Range("J3:O3").MergeCells = False
Range("A1").Select
Columns("O:O").Select
Selection.ColumnWidth = 50
Columns("J:L").HorizontalAlignment = xlCenter
Range("J3:O3").MergeCells = True

Range("O5:O" & tot_res + 4).Select
Selection.Locked = False
Selection.FormulaHidden = False
Range("A1").Select
Sheets("Menu").Protect

Sheets("Fixtures").Select
For y = 1 To Num_Divs
   fxtcol = 1 + ((y - 1) * 7)
 Z = 3
 While Trim(Cells(Z, fxtcol)) <> ""
  If Cells(Z, fxtcol + 2) <> "vs" Then
   Cells(Z, fxtcol + 2).Interior.ColorIndex = 8
  End If
  Z = Z + 1
 Wend
Next y
Range("A1").Select
Sheets("Menu").Select
End Sub
Sub Upd_Stats()
tmp = GetParms(True, False)
' Update Ratings
tmp = Rem_Rpts
Application.DisplayAlerts = False
For Each ws In Worksheets
 If ws.Name = "To_Be_Rated" Then
  Sheets("To_Be_Rated").Delete
 End If
Next ws
Application.DisplayAlerts = True
Sheets.Add.Name = "To_Be_Rated"
Sheets("To_Be_Rated").Tab.ColorIndex = 13
ActiveWindow.DisplayGridlines = False
Sheets("To_Be_Rated").Move After:=Sheets("Players")
Cells(1, 1) = "Div & Team Player Criteria: " & IIf(NR_Comp = "PLAYED", "Only Players that have Played", "All Players whether Played or Not")
Range("A1:M1").HorizontalAlignment = xlCenter
Range("A1:M1").VerticalAlignment = xlTop
Range("A1:M1").MergeCells = True
Range("A1:M1").Font.Name = "Arial"
Range("A1:M1").Font.FontStyle = "Bold"
Range("A1:M1").Font.Size = 16
Range("A1:M1").Font.Underline = xlUnderlineStyleSingle
Range("A1:M1").Font.ColorIndex = 3
Sheets("Results").Select
'Cells.Select
Columns("A:BV").Select
Selection.Copy
Sheets("tmp").Select
Cells.Select
ActiveSheet.Paste
Application.CutCopyMode = False
Sheets("Results").Select
Range("A1").Select
Sheets("tmp").Select





Range("A1").Select
Cells.Select
Selection.Sort Key1:=Range("E2"), Order1:=xlAscending, Key2:=Range("A2") _
    , Order2:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:= _
    False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
    :=xlSortNormal
Range("A2").Select
Sheets("Rtg_Chart").Select
lastrow = ActiveSheet.UsedRange.Rows.Count
RC_Num = 0
For x = 3 To lastrow
 RC_Num = RC_Num + 1
 RChrt(RC_Num, 1) = Cells(x, 1)
 RChrt(RC_Num, 2) = Cells(x, 2)
 RChrt(RC_Num, 3) = Cells(x, 3)
 RChrt(RC_Num, 4) = Cells(x, 4)
 RChrt(RC_Num, 5) = Cells(x, 5)
Next x
Sheets("PlyrRslt").Select
lastrow = ActiveSheet.UsedRange.Rows.Count
If lastrow > 1 Then
 Rows("2:" & lastrow).Select
 Application.CutCopyMode = False
 Selection.Delete Shift:=xlUp
End If
Range("A1").Select
PlrRslt_Row = 0
Sheets("Players").Select
Columns("D:D").Select
Selection.NumberFormat = "@"
Cells.Select
Selection.Sort Key1:=Range("B2"), Order1:=xlAscending, Header:=xlYes, _
               OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
               DataOption1:=xlSortNormal
Selection.Columns.AutoFit
Columns("E:R").Select
Selection.ColumnWidth = 10
Range("A1").Select
Tot_Plrs = ActiveSheet.UsedRange.Rows.Count
For x = 2 To Tot_Plrs
 Cells(x, 6) = Cells(x, 5)
 Cells(x, 7) = 0
 Cells(x, 9) = 0
 Cells(x, 10) = 0
 Cells(x, 11) = 0
 Cells(x, 12) = 0
 Cells(x, 13) = 0
 Cells(x, 14) = 0
 Cells(x, 15) = 0
 Cells(x, 16) = 0
 Cells(x, 17) = 0
 Cells(x, 18) = 0
Next x
Sheets("tmp").Select
tot_res = ActiveSheet.UsedRange.Rows.Count
For x = 2 To tot_res
 If Trim(Cells(x, 8)) <> "" Then
  Rows(x & ":" & x).Select
  dv = Cells(x, 4)
  matdat = Cells(x, 5)
  hte = Team(Cells(x, 6))
  ate = Team(Cells(x, 7))
  home_cl = Mid(Cells(x, 1), 6, 3)
  home_te = Mid(Cells(x, 1), 9, 1)
  away_cl = Mid(Cells(x, 1), 11, 3)
  away_te = Mid(Cells(x, 1), 14, 1)
  PLYR_A = Cells(x, 9)
  PLYR_B = Cells(x, 10)
  PLYR_C = Cells(x, 11)
  PLYR_X = Cells(x, 12)
  PLYR_Y = Cells(x, 13)
  PLYR_Z = Cells(x, 14)
  GA1_1 = Cells(x, 15)
  GA1_2 = Cells(x, 16)
  GA1_3 = Cells(x, 17)
  GA1_4 = Cells(x, 18)
  GA1_5 = Cells(x, 19)
  hw1 = Cells(x, 20)
  GA2_1 = Cells(x, 21)
  GA2_2 = Cells(x, 22)
  GA2_3 = Cells(x, 23)
  GA2_4 = Cells(x, 24)
  GA2_5 = Cells(x, 25)
  hw2 = Cells(x, 26)
  GA3_1 = Cells(x, 27)
  GA3_2 = Cells(x, 28)
  GA3_3 = Cells(x, 29)
  GA3_4 = Cells(x, 30)
  GA3_5 = Cells(x, 31)
  hw3 = Cells(x, 32)
  GA4_1 = Cells(x, 33)
  GA4_2 = Cells(x, 34)
  GA4_3 = Cells(x, 35)
  GA4_4 = Cells(x, 36)
  GA4_5 = Cells(x, 37)
  hw4 = Cells(x, 38)
  GA5_1 = Cells(x, 39)
  GA5_2 = Cells(x, 40)
  GA5_3 = Cells(x, 41)
  GA5_4 = Cells(x, 42)
  GA5_5 = Cells(x, 43)
  hw5 = Cells(x, 44)
  GA6_1 = Cells(x, 45)
  GA6_2 = Cells(x, 46)
  GA6_3 = Cells(x, 47)
  GA6_4 = Cells(x, 48)
  GA6_5 = Cells(x, 49)
  hw6 = Cells(x, 50)
  GA7_1 = Cells(x, 51)
  GA7_2 = Cells(x, 52)
  GA7_3 = Cells(x, 53)
  GA7_4 = Cells(x, 54)
  GA7_5 = Cells(x, 55)
  hw7 = Cells(x, 56)
  GA8_1 = Cells(x, 57)
  GA8_2 = Cells(x, 58)
  GA8_3 = Cells(x, 59)
  GA8_4 = Cells(x, 60)
  GA8_5 = Cells(x, 61)
  hw8 = Cells(x, 62)
  GA9_1 = Cells(x, 63)
  GA9_2 = Cells(x, 64)
  GA9_3 = Cells(x, 65)
  GA9_4 = Cells(x, 66)
  GA9_5 = Cells(x, 67)
  hw9 = Cells(x, 68)
  Pl_Row_A = PRow(PLYR_A, home_cl, home_te)
  Pl_Row_B = PRow(PLYR_B, home_cl, home_te)
  Pl_Row_C = PRow(PLYR_C, home_cl, home_te)
  Pl_Row_X = PRow(PLYR_X, away_cl, away_te)
  Pl_Row_Y = PRow(PLYR_Y, away_cl, away_te)
  Pl_Row_Z = PRow(PLYR_Z, away_cl, away_te)
  If Pl_Row_A > 1 Then
   PLYR_A = Sheets("Players").Cells(Pl_Row_A, 2)
   rtg_a = Sheets("Players").Cells(Pl_Row_A, 6)
  Else
   rtg_a = 0
  End If
  If Pl_Row_B > 1 Then
   PLYR_B = Sheets("Players").Cells(Pl_Row_B, 2)
   rtg_b = Sheets("Players").Cells(Pl_Row_B, 6)
  Else
   rtg_b = 0
  End If
  If Pl_Row_C > 1 Then
   PLYR_C = Sheets("Players").Cells(Pl_Row_C, 2)
   rtg_c = Sheets("Players").Cells(Pl_Row_C, 6)
  Else
   rtg_c = 0
  End If
  If Pl_Row_X > 1 Then
   PLYR_X = Sheets("Players").Cells(Pl_Row_X, 2)
   rtg_x = Sheets("Players").Cells(Pl_Row_X, 6)
  Else
   rtg_x = 0
  End If
  If Pl_Row_Y > 1 Then
   PLYR_Y = Sheets("Players").Cells(Pl_Row_Y, 2)
   rtg_y = Sheets("Players").Cells(Pl_Row_Y, 6)
  Else
   rtg_y = 0
  End If
  If Pl_Row_Z > 1 Then
   PLYR_Z = Sheets("Players").Cells(Pl_Row_Z, 2)
   rtg_z = Sheets("Players").Cells(Pl_Row_Z, 6)
  Else
   rtg_z = 0
  End If
  
  
  
  
 ' ******************************************************
  
  
  
Select Case Cells(x, 3)
 Case "6S3D"
   'AvY Game 2
   tmp = PlRslt_Bef(PLYR_A, PLYR_Y, matdat, hte, dv, ate, rtg_y, GA2_1, GA2_2, GA2_3, GA2_4, GA2_5, rtg_a, hw2, Pl_Row_A, Pl_Row_Y, 1, 1)
   If tmp Then
    tmph = rtg_a
    tmpa = rtg_y
    rtg_a = CalcRtg(tmph, tmpa, IIf(hw2 = "Home", True, False))
    rtg_y = CalcRtg(tmpa, tmph, IIf(hw2 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_a, rtg_y, tmph, tmpa)
   End If
   'BvZ Game 3
   tmp = PlRslt_Bef(PLYR_B, PLYR_Z, matdat, hte, dv, ate, rtg_z, GA3_1, GA3_2, GA3_3, GA3_4, GA3_5, rtg_b, hw3, Pl_Row_B, Pl_Row_Z, 1, 1)
   If tmp Then
    tmph = rtg_b
    tmpa = rtg_z
    rtg_b = CalcRtg(tmph, tmpa, IIf(hw3 = "Home", True, False))
    rtg_z = CalcRtg(tmpa, tmph, IIf(hw3 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_b, rtg_z, tmph, tmpa)
   End If
   'CvY Game 4
   tmp = PlRslt_Bef(PLYR_C, PLYR_Y, matdat, hte, dv, ate, rtg_y, GA4_1, GA4_2, GA4_3, GA4_4, GA4_5, rtg_c, hw4, Pl_Row_C, Pl_Row_Y, 1, 2)
   If tmp Then
    tmph = rtg_c
    tmpa = rtg_y
    rtg_c = CalcRtg(tmph, tmpa, IIf(hw4 = "Home", True, False))
    rtg_y = CalcRtg(tmpa, tmph, IIf(hw4 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_c, rtg_y, tmph, tmpa)
   End If
   'AvX Game 5
   tmp = PlRslt_Bef(PLYR_A, PLYR_X, matdat, hte, dv, ate, rtg_x, GA5_1, GA5_2, GA5_3, GA5_4, GA5_5, rtg_a, hw5, Pl_Row_A, Pl_Row_X, 2, 1)
   If tmp Then
    tmph = rtg_a
    tmpa = rtg_x
    rtg_a = CalcRtg(tmph, tmpa, IIf(hw5 = "Home", True, False))
    rtg_x = CalcRtg(tmpa, tmph, IIf(hw5 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_a, rtg_x, tmph, tmpa)
   End If
   'CvZ Game 7
   tmp = PlRslt_Bef(PLYR_C, PLYR_Z, matdat, hte, dv, ate, rtg_z, GA7_1, GA7_2, GA7_3, GA7_4, GA7_5, rtg_c, hw7, Pl_Row_C, Pl_Row_Z, 2, 2)
   If tmp Then
    tmph = rtg_c
    tmpa = rtg_z
    rtg_c = CalcRtg(tmph, tmpa, IIf(hw7 = "Home", True, False))
    rtg_z = CalcRtg(tmpa, tmph, IIf(hw7 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_c, rtg_z, tmph, tmpa)
   End If
   'BvX Game 8
   tmp = PlRslt_Bef(PLYR_B, PLYR_X, matdat, hte, dv, ate, rtg_x, GA8_1, GA8_2, GA8_3, GA8_4, GA8_5, rtg_b, hw8, Pl_Row_B, Pl_Row_X, 2, 2)
   If tmp Then
    tmph = rtg_b
    tmpa = rtg_x
    rtg_b = CalcRtg(tmph, tmpa, IIf(hw8 = "Home", True, False))
    rtg_x = CalcRtg(tmpa, tmph, IIf(hw8 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_b, rtg_x, tmph, tmpa)
   End If
Case "4S1D"
   'AvX Game1
   tmp = PlRslt_Bef(PLYR_A, PLYR_X, matdat, hte, dv, ate, rtg_x, GA1_1, GA1_2, GA1_3, GA1_4, GA1_5, rtg_a, hw1, Pl_Row_A, Pl_Row_X, 1, 1)
   If tmp Then
    tmph = rtg_a
    tmpa = rtg_x
    rtg_a = CalcRtg(tmph, tmpa, IIf(hw1 = "Home", True, False))
    rtg_x = CalcRtg(tmpa, tmph, IIf(hw1 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_a, rtg_x, tmph, tmpa)
   End If
   'BvY Game 2
   tmp = PlRslt_Bef(PLYR_B, PLYR_Y, matdat, hte, dv, ate, rtg_y, GA2_1, GA2_2, GA2_3, GA2_4, GA2_5, rtg_b, hw2, Pl_Row_B, Pl_Row_Y, 1, 1)
   If tmp Then
    tmph = rtg_b
    tmpa = rtg_y
    rtg_b = CalcRtg(tmph, tmpa, IIf(hw2 = "Home", True, False))
    rtg_y = CalcRtg(tmpa, tmph, IIf(hw2 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_b, rtg_y, tmph, tmpa)
   End If
   'AvY Game 4
   tmp = PlRslt_Bef(PLYR_A, PLYR_Y, matdat, hte, dv, ate, rtg_y, GA4_1, GA4_2, GA4_3, GA4_4, GA4_5, rtg_a, hw4, Pl_Row_A, Pl_Row_Y, 2, 2)
   If tmp Then
    tmph = rtg_a
    tmpa = rtg_y
    rtg_a = CalcRtg(tmph, tmpa, IIf(hw4 = "Home", True, False))
    rtg_y = CalcRtg(tmpa, tmph, IIf(hw4 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_a, rtg_y, tmph, tmpa)
   End If
   'BvX Game 5
   tmp = PlRslt_Bef(PLYR_B, PLYR_X, matdat, hte, dv, ate, rtg_x, GA5_1, GA5_2, GA5_3, GA5_4, GA5_5, rtg_b, hw5, Pl_Row_B, Pl_Row_X, 2, 2)
   If tmp Then
    tmph = rtg_b
    tmpa = rtg_x
    rtg_b = CalcRtg(tmph, tmpa, IIf(hw5 = "Home", True, False))
    rtg_x = CalcRtg(tmpa, tmph, IIf(hw5 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_b, rtg_x, tmph, tmpa)
   End If
Case Else
  'AvX Game1
   tmp = PlRslt_Bef(PLYR_A, PLYR_X, matdat, hte, dv, ate, rtg_x, GA1_1, GA1_2, GA1_3, GA1_4, GA1_5, rtg_a, hw1, Pl_Row_A, Pl_Row_X, 1, 1)
   If tmp Then
    tmph = rtg_a
    tmpa = rtg_x
    rtg_a = CalcRtg(tmph, tmpa, IIf(hw1 = "Home", True, False))
    rtg_x = CalcRtg(tmpa, tmph, IIf(hw1 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_a, rtg_x, tmph, tmpa)
   End If
   'BvY Game 2
   tmp = PlRslt_Bef(PLYR_B, PLYR_Y, matdat, hte, dv, ate, rtg_y, GA2_1, GA2_2, GA2_3, GA2_4, GA2_5, rtg_b, hw2, Pl_Row_B, Pl_Row_Y, 1, 1)
   If tmp Then
    tmph = rtg_b
    tmpa = rtg_y
    rtg_b = CalcRtg(tmph, tmpa, IIf(hw2 = "Home", True, False))
    rtg_y = CalcRtg(tmpa, tmph, IIf(hw2 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_b, rtg_y, tmph, tmpa)
   End If
   'CvZ Game 3
   tmp = PlRslt_Bef(PLYR_C, PLYR_Z, matdat, hte, dv, ate, rtg_z, GA3_1, GA3_2, GA3_3, GA3_4, GA3_5, rtg_c, hw3, Pl_Row_C, Pl_Row_Z, 1, 1)
   If tmp Then
    tmph = rtg_c
    tmpa = rtg_z
    rtg_c = CalcRtg(tmph, tmpa, IIf(hw3 = "Home", True, False))
    rtg_z = CalcRtg(tmpa, tmph, IIf(hw3 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_c, rtg_z, tmph, tmpa)
   End If
   'BvX Game 4
   tmp = PlRslt_Bef(PLYR_B, PLYR_X, matdat, hte, dv, ate, rtg_x, GA4_1, GA4_2, GA4_3, GA4_4, GA4_5, rtg_b, hw4, Pl_Row_B, Pl_Row_X, 2, 2)
   If tmp Then
    tmph = rtg_b
    tmpa = rtg_x
    rtg_b = CalcRtg(tmph, tmpa, IIf(hw4 = "Home", True, False))
    rtg_x = CalcRtg(tmpa, tmph, IIf(hw4 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_b, rtg_x, tmph, tmpa)
   End If
   'AvZ Game 5
   tmp = PlRslt_Bef(PLYR_A, PLYR_Z, matdat, hte, dv, ate, rtg_z, GA5_1, GA5_2, GA5_3, GA5_4, GA5_5, rtg_a, hw5, Pl_Row_A, Pl_Row_Z, 2, 2)
   If tmp Then
    tmph = rtg_a
    tmpa = rtg_z
    rtg_a = CalcRtg(tmph, tmpa, IIf(hw5 = "Home", True, False))
    rtg_z = CalcRtg(tmpa, tmph, IIf(hw5 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_a, rtg_z, tmph, tmpa)
   End If
   'CvY Game 6
   tmp = PlRslt_Bef(PLYR_C, PLYR_Y, matdat, hte, dv, ate, rtg_y, GA6_1, GA6_2, GA6_3, GA6_4, GA6_5, rtg_c, hw6, Pl_Row_C, Pl_Row_Y, 2, 2)
   If tmp Then
    tmph = rtg_c
    tmpa = rtg_y
    rtg_c = CalcRtg(tmph, tmpa, IIf(hw6 = "Home", True, False))
    rtg_y = CalcRtg(tmpa, tmph, IIf(hw6 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_c, rtg_y, tmph, tmpa)
   End If
   'BvZ Game 7
   tmp = PlRslt_Bef(PLYR_B, PLYR_Z, matdat, hte, dv, ate, rtg_z, GA7_1, GA7_2, GA7_3, GA7_4, GA7_5, rtg_b, hw7, Pl_Row_B, Pl_Row_Z, 3, 3)
   If tmp Then
    tmph = rtg_b
    tmpa = rtg_z
    rtg_b = CalcRtg(tmph, tmpa, IIf(hw7 = "Home", True, False))
    rtg_z = CalcRtg(tmpa, tmph, IIf(hw7 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_b, rtg_z, tmph, tmpa)
   End If
   'CvX Game 8
   tmp = PlRslt_Bef(PLYR_C, PLYR_X, matdat, hte, dv, ate, rtg_x, GA8_1, GA8_2, GA8_3, GA8_4, GA8_5, rtg_c, hw8, Pl_Row_C, Pl_Row_X, 3, 3)
   If tmp Then
    tmph = rtg_c
    tmpa = rtg_x
    rtg_c = CalcRtg(tmph, tmpa, IIf(hw8 = "Home", True, False))
    rtg_x = CalcRtg(tmpa, tmph, IIf(hw8 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_c, rtg_x, tmph, tmpa)
   End If
   'AvY Game 9
   tmp = PlRslt_Bef(PLYR_A, PLYR_Y, matdat, hte, dv, ate, rtg_y, GA9_1, GA9_2, GA9_3, GA9_4, GA9_5, rtg_a, hw9, Pl_Row_A, Pl_Row_Y, 3, 3)
   If tmp Then
    tmph = rtg_a
    tmpa = rtg_y
    rtg_a = CalcRtg(tmph, tmpa, IIf(hw9 = "Home", True, False))
    rtg_y = CalcRtg(tmpa, tmph, IIf(hw9 = "Home", False, True))
    tmp = PlRslt_Aft(rtg_a, rtg_y, tmph, tmpa)
   End If
End Select



  
    
 ' ******************************************************
  
  
  
  If Pl_Row_A > 0 Then
   Sheets("Players").Cells(Pl_Row_A, 6) = rtg_a
  End If
  If Pl_Row_B > 0 Then
   Sheets("Players").Cells(Pl_Row_B, 6) = rtg_b
  End If
  If Pl_Row_C > 0 Then
   Sheets("Players").Cells(Pl_Row_C, 6) = rtg_c
  End If
  If Pl_Row_X > 0 Then
   Sheets("Players").Cells(Pl_Row_X, 6) = rtg_x
  End If
  If Pl_Row_Y > 0 Then
   Sheets("Players").Cells(Pl_Row_Y, 6) = rtg_y
  End If
  If Pl_Row_Z > 0 Then
   Sheets("Players").Cells(Pl_Row_Z, 6) = rtg_z
  End If
 End If
Next x
Range("A1").Select
Sheets("PlyrRslt").Select
lastrow = ActiveSheet.UsedRange.Rows.Count
If lastrow > 1 Then
 Range("A2:R" & lastrow).Select
 With Selection.Interior
  .ColorIndex = 36
  .Pattern = xlSolid
 End With
 Selection.Borders(xlDiagonalDown).LineStyle = xlNone
 Selection.Borders(xlDiagonalUp).LineStyle = xlNone
 With Selection.Borders(xlEdgeLeft)
  .LineStyle = xlContinuous
  .Weight = xlThin
  .ColorIndex = xlAutomatic
 End With
 With Selection.Borders(xlEdgeTop)
  .LineStyle = xlContinuous
  .Weight = xlThin
  .ColorIndex = xlAutomatic
 End With
 With Selection.Borders(xlEdgeBottom)
  .LineStyle = xlContinuous
  .Weight = xlThin
  .ColorIndex = xlAutomatic
 End With
 With Selection.Borders(xlEdgeRight)
  .LineStyle = xlContinuous
  .Weight = xlThin
  .ColorIndex = xlAutomatic
 End With
 With Selection.Borders(xlInsideVertical)
  .LineStyle = xlContinuous
  .Weight = xlThin
  .ColorIndex = xlAutomatic
 End With
 With Selection.Borders(xlInsideHorizontal)
  .LineStyle = xlContinuous
  .Weight = xlThin
  .ColorIndex = xlAutomatic
 End With
End If
Cells.Select
Selection.Columns.AutoFit
Range("A1").Select
Cells.Select
Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Key2:=Range("C2") _
                                                             , Order2:=xlAscending, Key3:=Range("R2"), Order3:=xlAscending, Header:= _
               xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
               DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, DataOption3:= _
               xlSortNormal
Range("A1").Select
Sheets("Players").Select
lastrow = ActiveSheet.UsedRange.Rows.Count
For x = 2 To lastrow
 Cells(x, 7) = Cells(x, 6) - Cells(x, 5)
Next x
Range("A1").Select
' Do to be Rated Sheet
Sheets("Teams").Select
Num_Teams = ActiveSheet.UsedRange.Rows.Count
Dim rtms(100, 5) As Variant
For x = 1 To Num_Teams
 rtms(x, 1) = Cells(x, 2) ' Team Mnemonic
 rtms(x, 2) = 1
 rtms(x, 3) = 0
 rtms(x, 4) = 1
 rtms(x, 5) = 9999
Next x
Dim rdivs(9, 5) As Variant
For x = 1 To Num_Divs
 rdivs(x, 1) = Divs(x, 1) ' Divsion Number or letter should be Premier
 rdivs(x, 2) = 1 ' reserved for Highest rated player row
 rdivs(x, 3) = 0 ' reserved for Highest rated player's rating
 rdivs(x, 4) = 1 ' reserved for Lowest rated player row
 rdivs(x, 5) = 9999 ' reserved for Lowest rated player's rating
Next x
Sheets("Players").Select
Tot_Plrs = ActiveSheet.UsedRange.Rows.Count
For x = 2 To Tot_Plrs

 If Cells(x, 6) > 0 Then ' Only do if Player has a rating
  If Cells(x, 15) > 0 Or NR_Comp = "ALL" Then ' Only do if Player has played
   ' Get Highest and Lowest Rated Player per Division
   For y = 1 To Num_Divs
    If Trim(Cells(x, 4) & " ") = rdivs(y, 1) Then
     rdv = y
     Exit For
    End If
   Next y
   If Cells(x, 6) > rdivs(rdv, 3) Then
    rdivs(rdv, 2) = x
    rdivs(rdv, 3) = Cells(x, 6)
   End If
   If Cells(x, 6) < rdivs(rdv, 5) Then
    rdivs(rdv, 4) = x
    rdivs(rdv, 5) = Cells(x, 6)
   End If
' Get Highest and Lowest Rated Player per team
   For y = 1 To Num_Teams
    If Cells(x, 3) = rtms(y, 1) Then
     rtm = y
     Exit For
    End If
   Next y
   If Cells(x, 6) > rtms(rtm, 3) Then
    rtms(rtm, 2) = x
    rtms(rtm, 3) = Cells(x, 6)
   End If
   If Cells(x, 6) < rtms(rtm, 5) Then
    rtms(rtm, 4) = x
    rtms(rtm, 5) = Cells(x, 6)
   End If
  End If
 End If
Next x
Sheets("PlyrRslt").Select
lastrow = ActiveSheet.UsedRange.Rows.Count
tbr_cnt = 1
Sheets("To_Be_Rated").Select
cplr = "****"
For x = 2 To lastrow
 If Sheets("PlyrRslt").Cells(x, 15) = 0 Then
  If cplr <> Sheets("PlyrRslt").Cells(x, 1) Then
   cplr = Sheets("PlyrRslt").Cells(x, 1)
   tbr_cnt = tbr_cnt + 1
   cplr = Sheets("PlyrRslt").Cells(x, 1)
   plyr_row = Application.Match(Sheets("PlyrRslt").Cells(x, 1), Range("Players!B1:B" & Tot_Plrs), 0)
   Rows(tbr_cnt & ":" & tbr_cnt).Select
   rdiv = Sheets("Players").Cells(plyr_row, 4)
   rteam = Trim(Sheets("Players").Cells(plyr_row, 3))
   Cells(tbr_cnt, 1) = Trim(Sheets("Players").Cells(plyr_row, 2)) & " (" & rteam & ":" & rdiv & ")"
   rnge = "A" & tbr_cnt & ":M" & tbr_cnt
   tmp = CellForm(rnge, "C", True, True, 5, 14, False, False, 0)
   Cells(tbr_cnt + 1, 1) = "Reg Div - Plyd: " & Sheets("Players").Cells(plyr_row, 10) & " Won: " & _
                           Sheets("Players").Cells(plyr_row, 11) & " Pct: " & Round((Sheets("Players").Cells(plyr_row, 13) * 100), 2) & "%" & _
                           "  All Div - Plyd: " & Sheets("Players").Cells(plyr_row, 15) & " Won: " & _
                           Sheets("Players").Cells(plyr_row, 16) & " Pct: " & Round((Sheets("Players").Cells(plyr_row, 18) * 100), 2) & "%"
   rnge = "A" & tbr_cnt + 1 & ":M" & tbr_cnt + 1
   tmp = CellForm(rnge, "C", True, True, 1, 12, True, True, 37)
   rbig_win = 0
   nbig_win = ""
   tbig_win = ""
   sbig_win = ""
   Cells(tbr_cnt + 2, 1) = "Biggest Win:"
   Cells(tbr_cnt + 2, 2) = nbig_win
   Cells(tbr_cnt + 2, 5) = tbig_win
   Cells(tbr_cnt + 2, 6) = sbig_win
   Cells(tbr_cnt + 2, 12) = "Ratg:"
   Cells(tbr_cnt + 2, 13) = rbig_win
   rnge = "A" & tbr_cnt + 2 & ":M" & tbr_cnt + 2
   tmp = CellForm(rnge, "N", False, True, 0, 10, True, True, 4)
   Range("A" & tbr_cnt + 2 & ":A" & tbr_cnt + 2).HorizontalAlignment = xlRight
   Range("B" & tbr_cnt + 2 & ":D" & tbr_cnt + 2).MergeCells = True
   Range("F" & tbr_cnt + 2 & ":K" & tbr_cnt + 2).MergeCells = True
   Range("E" & tbr_cnt + 2 & ":K" & tbr_cnt + 2).HorizontalAlignment = xlCenter
   rlow_loss = 9999
   nlow_loss = ""
   tlow_loss = ""
   slow_loss = ""
   Cells(tbr_cnt + 3, 1) = "Lowest Loss:"
   Cells(tbr_cnt + 3, 2) = nlow_loss
   Cells(tbr_cnt + 3, 5) = tlow_loss
   Cells(tbr_cnt + 3, 6) = slow_loss
   Cells(tbr_cnt + 3, 12) = "Ratg:"
   Cells(tbr_cnt + 3, 13) = rlow_loss
   rnge = "A" & tbr_cnt + 3 & ":M" & tbr_cnt + 3
   tmp = CellForm(rnge, "N", False, True, 0, 10, True, True, 36)
   Range("A" & tbr_cnt + 3 & ":A" & tbr_cnt + 3).HorizontalAlignment = xlRight
   Range("B" & tbr_cnt + 3 & ":D" & tbr_cnt + 3).MergeCells = True
   Range("F" & tbr_cnt + 3 & ":K" & tbr_cnt + 3).MergeCells = True
   Range("E" & tbr_cnt + 3 & ":K" & tbr_cnt + 3).HorizontalAlignment = xlCenter
  ' sug_rtg = "????"
   'Cells(tbr_cnt + 4, 1) = "Suggested Rating: " & sug_rtg
   'rnge = "A" & tbr_cnt + 4 & ":M" & tbr_cnt + 4
   'tmp = CellForm(rnge, "C", True, True, 1, 12, True, True, 33)
   Cells(tbr_cnt + 4, 1) = "Div " & rdiv & " Highest Rated Player:"
   For y = 1 To Num_Divs
    If Trim(rdiv & "") = rdivs(y, 1) Then
     rdv = rdivs(y, 2)
     Exit For
    End If
   Next y
   If rdv = 1 Then
    Cells(tbr_cnt + 4, 2) = "No Rated Player"
    Cells(tbr_cnt + 4, 5) = "N/A"
    Cells(tbr_cnt + 4, 6) = "N/A"
    Cells(tbr_cnt + 4, 12) = "Ratg:"
    Cells(tbr_cnt + 4, 13) = "N/A"
   Else
    Cells(tbr_cnt + 4, 2) = Sheets("Players").Cells(rdv, 2)
    Cells(tbr_cnt + 4, 5) = Sheets("Players").Cells(rdv, 3)
    Cells(tbr_cnt + 4, 6) = "Plyd: " & Sheets("Players").Cells(rdv, 15) & " Won: " & _
                           Sheets("Players").Cells(rdv, 16) & " Pct: " & Round((Sheets("Players").Cells(rdv, 18) * 100), 2)
    Cells(tbr_cnt + 4, 12) = "Ratg:"
    Cells(tbr_cnt + 4, 13) = Sheets("Players").Cells(rdv, 6)
   End If
   rnge = "A" & tbr_cnt + 4 & ":M" & tbr_cnt + 4
   tmp = CellForm(rnge, "N", False, True, 0, 10, True, True, 34)
   Range("A" & tbr_cnt + 4 & ":A" & tbr_cnt + 4).HorizontalAlignment = xlRight
   Range("B" & tbr_cnt + 4 & ":D" & tbr_cnt + 4).MergeCells = True
   Range("F" & tbr_cnt + 4 & ":K" & tbr_cnt + 4).MergeCells = True
   Range("E" & tbr_cnt + 4 & ":K" & tbr_cnt + 4).HorizontalAlignment = xlCenter
   Cells(tbr_cnt + 5, 1) = "Div " & rdiv & " Lowest Rated Player:"
   For y = 1 To Num_Divs
    If Trim(rdiv & "") = rdivs(y, 1) Then
     rdv = rdivs(y, 4)
     Exit For
    End If
   Next y
   If rdv = 1 Then
    Cells(tbr_cnt + 5, 2) = "No Rated Player"
    Cells(tbr_cnt + 5, 5) = "N/A"
    Cells(tbr_cnt + 5, 6) = "N/A"
    Cells(tbr_cnt + 5, 12) = "Ratg:"
    Cells(tbr_cnt + 5, 13) = "N/A"
   Else
    Cells(tbr_cnt + 5, 2) = Sheets("Players").Cells(rdv, 2)
    Cells(tbr_cnt + 5, 5) = Sheets("Players").Cells(rdv, 3)
    Cells(tbr_cnt + 5, 6) = "Plyd: " & Sheets("Players").Cells(rdv, 15) & " Won: " & _
                            Sheets("Players").Cells(rdv, 16) & " Pct: " & Round((Sheets("Players").Cells(rdv, 18) * 100), 2)
    Cells(tbr_cnt + 5, 12) = "Ratg:"
    Cells(tbr_cnt + 5, 13) = Sheets("Players").Cells(rdv, 6)
   End If
   rnge = "A" & tbr_cnt + 5 & ":M" & tbr_cnt + 5
   tmp = CellForm(rnge, "N", False, True, 0, 10, True, True, 34)
   Range("A" & tbr_cnt + 5 & ":A" & tbr_cnt + 5).HorizontalAlignment = xlRight
   Range("B" & tbr_cnt + 5 & ":D" & tbr_cnt + 5).MergeCells = True
   Range("F" & tbr_cnt + 5 & ":K" & tbr_cnt + 5).MergeCells = True
   Range("E" & tbr_cnt + 5 & ":K" & tbr_cnt + 5).HorizontalAlignment = xlCenter
   Cells(tbr_cnt + 6, 1) = rteam & " Highest Rated Player:"
   For y = 1 To Num_Teams
    If rteam = rtms(y, 1) Then
     rtm = rtms(y, 2)
     Exit For
    End If
   Next y
   If rtm = 1 Then
    Cells(tbr_cnt + 6, 2) = "No Rated Player"
    Cells(tbr_cnt + 6, 5) = "N/A"
    Cells(tbr_cnt + 6, 6) = "N/A"
    Cells(tbr_cnt + 6, 12) = "Ratg:"
    Cells(tbr_cnt + 6, 13) = "N/A"
   Else
    Cells(tbr_cnt + 6, 2) = Sheets("Players").Cells(rtm, 2)
    Cells(tbr_cnt + 6, 5) = Sheets("Players").Cells(rtm, 3)
    Cells(tbr_cnt + 6, 6) = "Plyd: " & Sheets("Players").Cells(rtm, 15) & " Won: " & _
                            Sheets("Players").Cells(rtm, 16) & " Pct: " & Round((Sheets("Players").Cells(rtm, 18) * 100), 2)
    Cells(tbr_cnt + 6, 12) = "Ratg:"
    Cells(tbr_cnt + 6, 13) = Sheets("Players").Cells(rtm, 6)
   End If
   rnge = "A" & tbr_cnt + 6 & ":M" & tbr_cnt + 6
   tmp = CellForm(rnge, "N", False, True, 0, 10, True, True, 40)
   Range("A" & tbr_cnt + 6 & ":A" & tbr_cnt + 6).HorizontalAlignment = xlRight
   Range("B" & tbr_cnt + 6 & ":D" & tbr_cnt + 6).MergeCells = True
   Range("F" & tbr_cnt + 6 & ":K" & tbr_cnt + 6).MergeCells = True
   Range("E" & tbr_cnt + 6 & ":K" & tbr_cnt + 6).HorizontalAlignment = xlCenter
   Cells(tbr_cnt + 7, 1) = rteam & " Lowest Rated Player:"
   For y = 1 To Num_Teams
    If rteam = rtms(y, 1) Then
     rtm = rtms(y, 4)
     Exit For
    End If
   Next y
   If rtm = 1 Then
    Cells(tbr_cnt + 7, 2) = "No Rated Player"
    Cells(tbr_cnt + 7, 5) = "N/A"
    Cells(tbr_cnt + 7, 6) = "N/A"
    Cells(tbr_cnt + 7, 12) = "Ratg:"
    Cells(tbr_cnt + 7, 13) = "N/A"
   Else
    Cells(tbr_cnt + 7, 2) = Sheets("Players").Cells(rtm, 2)
    Cells(tbr_cnt + 7, 5) = Sheets("Players").Cells(rtm, 3)
    Cells(tbr_cnt + 7, 6) = "Plyd: " & Sheets("Players").Cells(rtm, 15) & " Won: " & _
                            Sheets("Players").Cells(rtm, 16) & " Pct: " & Round((Sheets("Players").Cells(rtm, 18) * 100), 2)
    Cells(tbr_cnt + 7, 12) = "Ratg:"
    Cells(tbr_cnt + 7, 13) = Sheets("Players").Cells(rtm, 6)
   End If
   rnge = "A" & tbr_cnt + 7 & ":M" & tbr_cnt + 7
   tmp = CellForm(rnge, "N", False, True, 0, 10, True, True, 40)
   Range("A" & tbr_cnt + 7 & ":A" & tbr_cnt + 7).HorizontalAlignment = xlRight
   Range("B" & tbr_cnt + 7 & ":D" & tbr_cnt + 7).MergeCells = True
   Range("F" & tbr_cnt + 7 & ":K" & tbr_cnt + 7).MergeCells = True
   Range("E" & tbr_cnt + 7 & ":K" & tbr_cnt + 7).HorizontalAlignment = xlCenter
   Cells(tbr_cnt + 8, 1) = "OPPOSITION"
   Cells(tbr_cnt + 8, 2) = "MATCH" & Chr(10) & "DATE"
   Cells(tbr_cnt + 8, 3) = "PLAY" & Chr(10) & "TEAM"
   Cells(tbr_cnt + 8, 4) = "DIV"
   Cells(tbr_cnt + 8, 5) = "OPP" & Chr(10) & "TEAM"
   Cells(tbr_cnt + 8, 6) = "HOME" & Chr(10) & "AWAY"
   Cells(tbr_cnt + 8, 7) = "GA1"
   Cells(tbr_cnt + 8, 8) = "GA2"
   Cells(tbr_cnt + 8, 9) = "GA3"
   Cells(tbr_cnt + 8, 10) = "GA4"
   Cells(tbr_cnt + 8, 11) = "GA5"
   Cells(tbr_cnt + 8, 12) = "WON" & Chr(10) & "LOST"
   Cells(tbr_cnt + 8, 13) = "OPP" & Chr(10) & "RTG"
   rnge = "A" & tbr_cnt + 8 & ":M" & tbr_cnt + 8
   tmp = CellForm(rnge, "C", False, True, 0, 10, True, True, 33)
  End If
  rcnt = 0
  While cplr = Sheets("PlyrRslt").Cells(x, 1)
   If Sheets("PlyrRslt").Cells(x, 14) > rbig_win And Sheets("PlyrRslt").Cells(x, 13) = "Won" Then
    rbig_win = Sheets("PlyrRslt").Cells(x, 14)
    nbig_win = Sheets("PlyrRslt").Cells(x, 2)
    tbig_win = Sheets("PlyrRslt").Cells(x, 6)
    plyr_row = Application.Match(nbig_win, Range("Players!B1:B" & Tot_Plrs), 0)
    sbig_win = "Plyd: " & Sheets("Players").Cells(plyr_row, 15) & " Won: " & _
               Sheets("Players").Cells(plyr_row, 16) & " Pct: " & Round((Sheets("Players").Cells(plyr_row, 18) * 100), 2)
   End If
   If Sheets("PlyrRslt").Cells(x, 14) < rlow_loss And Sheets("PlyrRslt").Cells(x, 13) = "Lost" And Sheets("PlyrRslt").Cells(x, 14) > 0 Then
    rlow_loss = Sheets("PlyrRslt").Cells(x, 14)
    nlow_loss = Sheets("PlyrRslt").Cells(x, 2)
    tlow_loss = Sheets("PlyrRslt").Cells(x, 6)
    plyr_row = Application.Match(nlow_loss, Range("Players!B1:B" & Tot_Plrs), 0)
    slow_loss = "Plyd: " & Sheets("Players").Cells(plyr_row, 15) & " Won: " & _
                Sheets("Players").Cells(plyr_row, 16) & " Pct: " & Round((Sheets("Players").Cells(plyr_row, 18) * 100), 2)
   End If
   rcnt = rcnt + 1
   srow = tbr_cnt + 8 + rcnt
   Cells(srow, 1) = Sheets("PlyrRslt").Cells(x, 2)
   Cells(srow, 2) = Sheets("PlyrRslt").Cells(x, 3)
   Cells(srow, 3) = Sheets("PlyrRslt").Cells(x, 4)
   Cells(srow, 4) = Sheets("PlyrRslt").Cells(x, 5)
   Cells(srow, 5) = Sheets("PlyrRslt").Cells(x, 6)
   Cells(srow, 6) = Sheets("PlyrRslt").Cells(x, 7)
   Cells(srow, 7) = Sheets("PlyrRslt").Cells(x, 8)
   Cells(srow, 8) = Sheets("PlyrRslt").Cells(x, 9)
   Cells(srow, 9) = Sheets("PlyrRslt").Cells(x, 10)
   Cells(srow, 10) = Sheets("PlyrRslt").Cells(x, 11)
   Cells(srow, 11) = Sheets("PlyrRslt").Cells(x, 12)
   Cells(srow, 12) = Sheets("PlyrRslt").Cells(x, 13)
   Cells(srow, 13) = Sheets("PlyrRslt").Cells(x, 14)
   Cells(srow, 2).NumberFormat = "d-mmm-yy"
   rnge = "A" & srow & ":M" & srow
   If Cells(srow, 13) = 0 Then
    Range(rnge).Interior.ColorIndex = 15
   Else
    Range(rnge).Interior.ColorIndex = IIf(Cells(srow, 12) = "Won", 4, 36)
   End If
   x = x + 1
  Wend
  rnge = "A" & tbr_cnt + 9 & ":M" & tbr_cnt + 8 + rcnt
  tmp = CellForm(rnge, "C", False, True, 0, 10, True, False, 4)
  rnge = "A" & tbr_cnt + 9 & ":A" & tbr_cnt + 8 + rcnt
  tmp = CellForm(rnge, "L", False, True, 0, 10, True, False, 4)
  x = x - 1
  Cells(tbr_cnt + 2, 2) = nbig_win
  Cells(tbr_cnt + 2, 5) = tbig_win
  Cells(tbr_cnt + 2, 6) = sbig_win
  Cells(tbr_cnt + 2, 13) = IIf(rbig_win = 0, "", rbig_win)
  Cells(tbr_cnt + 3, 2) = nlow_loss
  Cells(tbr_cnt + 3, 5) = tlow_loss
  Cells(tbr_cnt + 3, 6) = slow_loss
  Cells(tbr_cnt + 3, 13) = IIf(rlow_loss = 9999, "", rlow_loss)
  tbr_cnt = tbr_cnt + 9 + rcnt
 End If
Next x
Cells.Select
Selection.Columns.AutoFit
Range("A1").Select
Application.DisplayAlerts = True
Range("A1").Select
Sheets("Menu").Select
ActiveSheet.Unprotect
ActiveSheet.Shapes("Run_Reports_Info").Select
Selection.Characters.text = " *** NEEDS UPDATING ***"
ActiveSheet.Shapes("Run_Reports_Box").TextFrame.Characters.Font.ColorIndex = 1
ActiveSheet.Shapes("Run_Reports_Box").OnAction = "Run_Reports"
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
tmp = Tidy_Up("Upd_Stats_Info")
End Sub
Sub Init_Sys()
Msg = "This action will Rewrite the WorkSheets: Results,Fixtures & PlyrRslt" & vbCr & vbCr
Msg = Msg & "          Players Stats and Ratings will be Reset to Start of Season" & vbCr & vbCr
Msg = Msg & "                              All Report Sheets will be Deleted" & vbCr & vbCr
Msg = Msg & "              Following Sheets will remain untouched :" & vbCr & vbCr
Msg = Msg & "   Teams, Rtg_Chart, Config & Plyr_Lib which is hidden " & vbCr & vbCr
Msg = Msg & vbCr & "Do you want to continue ?"
Style = vbYesNo + vbExclamation + vbDefaultButton2
title = "You are about to Initialize the System"
response = MsgBox(Msg, Style, title)
If response = vbYes Then GoTo Continue Else GoTo Abort
Continue:
tmp = GetParms(True, True)
' Initialize All Sheets - Completly Rewritten: Results, PlyrRslt - Stats Zeroed Players - Untouched: Config,Teams,Rtg_Chart
' All Report Pages deleted
tmp = Rem_Rpts
Sheets("Fixtures").Select
Cells.Select
Application.CutCopyMode = False
Selection.Clear
Range("A3").Select
Sheets("Results").Select
lastrow = ActiveSheet.UsedRange.Rows.Count
If lastrow > 1 Then
 Rows("2:" & lastrow).Select
 Application.CutCopyMode = False
 Selection.Delete Shift:=xlUp
End If
Range("A2").Select
ActiveWindow.FreezePanes = True
Range("A1").Select
tot_res = 1
Nxt_Fix_Col = 1

Set oFixtureScraper = New FixtureScraper
Set oFixtureHelper = New FixtureHelper
Set oMatchHelper = New MatchHelper
Dim fixtureSheet As Worksheet
Set fixtureSheet = Worksheets("Fixtures")
fixtureSheet.Cells.Clear

Dim divFixtures As Collection

For x = 1 To Num_Divs
' tmp = GetFixts(Divs(x, 1), Divs(x, 3))
' lastrow = ActiveSheet.UsedRange.Rows.Count
' Nxt_Fix_Col = PutFixts(Divs(x, 1), lastrow, False)
 Debug.Print "Getting Fixtures for Division " + CStr(Divs(x, 1))
 Set divFixtures = oFixtureScraper.GetFixtures(CStr(Divs(x, 3)))
 lastrow = divFixtures.Count + 2
 oFixtureHelper.AppendFixtures fixtureSheet, CInt(Divs(x, 1)), divFixtures

 Sheets("Results").Select
 For Each oFixture In divFixtures
  lookup_week = IIf(Len(oFixture.WeekNumber) = 1, "0", "") & oFixture.WeekNumber
  lookup_val = lookup_week & "-" & Divs(x, 1) & "-" & Team(oFixture.HomeTeam) & "-" & Team(oFixture.AwayTeam)
  Cells(tot_res + 1, 1) = lookup_val
  Cells(tot_res + 1, 2) = oFixture.WeekNumber
  Cells(tot_res + 1, 3) = Divs(x, 2)
  Cells(tot_res + 1, 4) = Divs(x, 1)
  Cells(tot_res + 1, 5) = oFixture.MatchDate
  Cells(tot_res + 1, 6) = oFixture.HomeTeam
  Cells(tot_res + 1, 7) = oFixture.AwayTeam
    
  tot_res = tot_res + 1
 Next
Next x
Sheets("tmp").Select
Cells.Select
Selection.Clear
Range("A1").Select
Sheets("tmp1").Select
Cells.Select
Selection.Clear
Range("A1").Select
Sheets("Results").Select
Cells.Select
Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes, _
               OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
               DataOption1:=xlSortNormal
Range("A1").Select
Cells.Select
Selection.Interior.ColorIndex = xlNone
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
Selection.Borders(xlEdgeLeft).LineStyle = xlNone
Selection.Borders(xlEdgeTop).LineStyle = xlNone
Selection.Borders(xlEdgeBottom).LineStyle = xlNone
Selection.Borders(xlEdgeRight).LineStyle = xlNone
Selection.Borders(xlInsideVertical).LineStyle = xlNone
Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
Range("A1:BV1").Select
With Selection.Interior
 .ColorIndex = 33
 .Pattern = xlSolid
End With
dest_range = "A2:BV" & tot_res
Range(dest_range).Select
With Selection.Interior
 .ColorIndex = 36
 .Pattern = xlSolid
End With
dest_range = "A1:BV" & tot_res
Range(dest_range).Select
With Selection.Borders(xlEdgeLeft)
 .LineStyle = xlContinuous
 .Weight = xlThin
 .ColorIndex = xlAutomatic
End With
With Selection.Borders(xlEdgeTop)
 .LineStyle = xlContinuous
 .Weight = xlThin
 .ColorIndex = xlAutomatic
End With
With Selection.Borders(xlEdgeBottom)
 .LineStyle = xlContinuous
 .Weight = xlThin
 .ColorIndex = xlAutomatic
End With
With Selection.Borders(xlEdgeRight)
 .LineStyle = xlContinuous
 .Weight = xlThin
 .ColorIndex = xlAutomatic
End With
With Selection.Borders(xlInsideVertical)
 .LineStyle = xlContinuous
 .Weight = xlThin
 .ColorIndex = xlAutomatic
End With
With Selection.Borders(xlInsideHorizontal)
 .LineStyle = xlContinuous
 .Weight = xlThin
 .ColorIndex = xlAutomatic
End With
Range("A1").Select
Cells.Select
Selection.Columns.AutoFit
Range("A1").Select
Sheets("Players").Select
Tot_Plrs = ActiveSheet.UsedRange.Rows.Count
For x = 2 To Tot_Plrs
 Cells(x, 6) = Cells(x, 5)
 Cells(x, 7) = 0
 Cells(x, 9) = 0
 Cells(x, 10) = 0
 Cells(x, 11) = 0
 Cells(x, 12) = 0
 Cells(x, 13) = 0
 Cells(x, 14) = 0
 Cells(x, 15) = 0
 Cells(x, 16) = 0
 Cells(x, 17) = 0
 Cells(x, 18) = 0
Next x
Range("A1").Select
Sheets("PlyrRslt").Select
lastrow = ActiveSheet.UsedRange.Rows.Count
If lastrow > 1 Then
 Rows("2:" & lastrow).Select
 Application.CutCopyMode = False
 Selection.Delete Shift:=xlUp
End If
Range("A2").Select
Application.DisplayAlerts = False
For Each ws In Worksheets
 If ws.Name = "To_Be_Rated" Then
  Sheets("To_Be_Rated").Delete
 End If
Next ws
Call Await_Res
Application.DisplayAlerts = True
Range("A1").Select
Sheets("Menu").Select
ActiveSheet.Unprotect
ActiveSheet.Shapes("OS_Res_Info").Select
Selection.Characters.text = " *** NEEDS UPDATING ***"
ActiveSheet.Shapes("Upd_Stats_Info").Select
Selection.Characters.text = "* GET O/S RESULTS 1st *"
ActiveSheet.Shapes("Upd_Stats_Box").TextFrame.Characters.Font.ColorIndex = 16
ActiveSheet.Shapes("Upd_Stats_Box").OnAction = ""
ActiveSheet.Shapes("Run_Reports_Info").Select
Selection.Characters.text = "*RUN UPDATE STATS 1st*"
ActiveSheet.Shapes("Run_Reports_Box").TextFrame.Characters.Font.ColorIndex = 16
ActiveSheet.Shapes("Run_Reports_Box").OnAction = ""
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
tmp = Tidy_Up("Init_Sys_Info")
Abort:
End Sub
Function CalcRtg(var1 As Variant, var2 As Variant, var3 As Boolean)
If var1 = 0 Or var2 = 0 Then
 CalcRtg = var1
Else
 If var1 >= var2 Then
  If var3 Then
   rc_col = 2
  Else
   rc_col = 5
  End If
 Else
  If var3 Then
   rc_col = 4
  Else
   rc_col = 3
  End If
 End If
 diff = Abs(var1 - var2)
 For r = 1 To RC_Num
  If diff < RChrt(r, 1) Then
   CalcRtg = var1 + RChrt(r, rc_col)
   Exit For
  End If
 Next r
End If
End Function
Function PRow(var1 As Variant, var2 As Variant, var3 As Variant)
lu_range = "Players!A1:A" & Tot_Plrs
plyr_row = Application.Match(var1, Range(lu_range), 0)
If IsError(plyr_row) Then
 plyr_row = 0
End If
If Chk_Plr_Clb Then
 If plyr_row > 0 Then
  ' Check if Player found is right Club and Team
  pl_confirmed = False
  pl_cnt = plyr_row
  While Sheets("Players").Cells(pl_cnt, 1) = var1
   If Sheets("Players").Cells(pl_cnt, 3) = var2 & var3 Then
    plyr_row = pl_cnt
    pl_confirmed = True
   End If
   pl_cnt = pl_cnt + 1
  Wend
  If Not pl_confirmed Then
   ' If not confirmed right Club and Team - just check for Club
   pl_cnt = plyr_row
   While Sheets("Players").Cells(pl_cnt, 1) = var1
    If Left(Sheets("Players").Cells(pl_cnt, 3), 3) = var2 Then
     plyr_row = pl_cnt
     pl_confirmed = True
    End If
    pl_cnt = pl_cnt + 1
   Wend
  End If
  If Not pl_confirmed Then
   ' Player not even a Club Player
    plyr_row = 0
    dummy = MsgBox(var1 & " Not registered for Club: " & var2, , "Procedure Aborted")
    End
  End If
 End If
End If
If plyr_row = 0 Then
 PRow = 0
Else
 PRow = plyr_row
End If
End Function
Function Team(var1 As Variant) As String
Team = CStr(Application.VLookup(var1, Range("Teams!A:B"), 2, 0))
End Function
Function WD_Team(var1 As Variant) As Boolean
WD_Team = CBool(Application.VLookup(var1, Range("Teams!A:E"), 5, 0))
End Function
Function PlRslt_Bef(v_homp, v_awap, v_matdat, v_hte, v_dv, v_ate, v_rtga, v_ga1, v_ga2, v_ga3, v_ga4, v_ga5, v_rtgh, v_wl, v_plrowh, v_plrowa, v_oph, v_opa)
If v_oph = 1 And v_plrowh > 1 Then
 Sheets("Players").Cells(v_plrowh, 14) = Sheets("Players").Cells(v_plrowh, 14) + 1
 If Trim(Sheets("Players").Cells(v_plrowh, 4)) = Trim(v_dv) Then
  Sheets("Players").Cells(v_plrowh, 9) = Sheets("Players").Cells(v_plrowh, 9) + 1
 End If
End If
If v_opa = 1 And v_plrowa > 1 Then
 Sheets("Players").Cells(v_plrowa, 14) = Sheets("Players").Cells(v_plrowa, 14) + 1
 If Trim(Sheets("Players").Cells(v_plrowa, 4)) = Trim(v_dv) Then
  Sheets("Players").Cells(v_plrowa, 9) = Sheets("Players").Cells(v_plrowa, 9) + 1
 End If
End If
If v_homp = "Forfeit" Or v_awap = "Forfeit" Or v_ga2 = "Walk" Then
 PlRslt_Bef = False
Else
 Sheets("Players").Cells(v_plrowh, 15) = Sheets("Players").Cells(v_plrowh, 15) + 1
 If v_wl = "Home" Then
  Sheets("Players").Cells(v_plrowh, 16) = Sheets("Players").Cells(v_plrowh, 16) + 1
 Else
  Sheets("Players").Cells(v_plrowh, 17) = Sheets("Players").Cells(v_plrowh, 17) + 1
 End If
 Sheets("Players").Cells(v_plrowh, 18) = Sheets("Players").Cells(v_plrowh, 16) / Sheets("Players").Cells(v_plrowh, 15)
 If Trim(Sheets("Players").Cells(v_plrowh, 4)) = Trim(v_dv) Then
  Sheets("Players").Cells(v_plrowh, 10) = Sheets("Players").Cells(v_plrowh, 10) + 1
  If v_wl = "Home" Then
   Sheets("Players").Cells(v_plrowh, 11) = Sheets("Players").Cells(v_plrowh, 11) + 1
  Else
   Sheets("Players").Cells(v_plrowh, 12) = Sheets("Players").Cells(v_plrowh, 12) + 1
  End If
  Sheets("Players").Cells(v_plrowh, 13) = Sheets("Players").Cells(v_plrowh, 11) / Sheets("Players").Cells(v_plrowh, 10)
 End If
 Sheets("Players").Cells(v_plrowa, 15) = Sheets("Players").Cells(v_plrowa, 15) + 1
 If v_wl = "Home" Then
  Sheets("Players").Cells(v_plrowa, 17) = Sheets("Players").Cells(v_plrowa, 17) + 1
 Else
  Sheets("Players").Cells(v_plrowa, 16) = Sheets("Players").Cells(v_plrowa, 16) + 1
 End If
 Sheets("Players").Cells(v_plrowa, 18) = Sheets("Players").Cells(v_plrowa, 16) / Sheets("Players").Cells(v_plrowa, 15)
 If Trim(Sheets("Players").Cells(v_plrowa, 4)) = Trim(v_dv) Then
  Sheets("Players").Cells(v_plrowa, 10) = Sheets("Players").Cells(v_plrowa, 10) + 1
  If v_wl = "Home" Then
   Sheets("Players").Cells(v_plrowa, 12) = Sheets("Players").Cells(v_plrowa, 12) + 1
  Else
   Sheets("Players").Cells(v_plrowa, 11) = Sheets("Players").Cells(v_plrowa, 11) + 1
  End If
  Sheets("Players").Cells(v_plrowa, 13) = Sheets("Players").Cells(v_plrowa, 11) / Sheets("Players").Cells(v_plrowa, 10)
 End If
 PlrRslt_Row = PlrRslt_Row + 2
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 1) = v_homp
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 2) = v_awap
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 3) = v_matdat
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 4) = v_hte
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 5) = v_dv
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 6) = v_ate
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 7) = "Home"
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 8) = v_ga1
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 9) = v_ga2
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 10) = v_ga3
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 11) = v_ga4
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 12) = v_ga5
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 13) = IIf(v_wl = "Home", "Won", "Lost")
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 14) = v_rtga
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 15) = v_rtgh
 Sheets("PlyrRslt").Cells(PlrRslt_Row, 18) = v_oph
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 1) = v_awap
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 2) = v_homp
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 3) = v_matdat
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 4) = v_ate
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 5) = v_dv
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 6) = v_hte
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 7) = "Away"
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 8) = RvScore(v_ga1)
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 9) = RvScore(v_ga2)
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 10) = RvScore(v_ga3)
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 11) = RvScore(v_ga4)
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 12) = RvScore(v_ga5)
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 13) = IIf(v_wl = "Home", "Lost", "Won")
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 14) = v_rtgh
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 15) = v_rtga
 Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 18) = v_opa
 PlRslt_Bef = True
End If
End Function
Function RvScore(var1)
RvScore = Right(var1, 2) & Mid(var1, 3, 1) & Left(var1, 2)
End Function
Function PlRslt_Aft(v_rtgh, v_rtga, v_tmph, v_tmpa)
Sheets("PlyrRslt").Cells(PlrRslt_Row, 17) = v_rtgh
Sheets("PlyrRslt").Cells(PlrRslt_Row, 16) = v_rtgh - v_tmph
Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 17) = v_rtga
Sheets("PlyrRslt").Cells(PlrRslt_Row + 1, 16) = v_rtga - v_tmpa
PlRslt_Aft = True
End Function
Function TrimPlyr(varpl As String) As String
 TrimPlyr = Trim(Replace(varpl, "(S)", " "))
End Function
Function score(var As Variant, var2 As Variant, var3 As Variant, var4 As Variant, var5 As Variant)
 var1 = Trim(var)
 del_char = 0
 If var2 <> "" Then
  del_char = del_char + Len(Val(Left(var2, 2)) & "-" & Val(Right(var2, 2)))
 End If
 If var3 <> "" Then
  del_char = del_char + Len(Val(Left(var3, 2)) & "-" & Val(Right(var3, 2)))
 End If
 If var4 <> "" Then
  del_char = del_char + Len(Val(Left(var4, 2)) & "-" & Val(Right(var4, 2)))
 End If
 If var5 <> "" Then
  del_char = del_char + Len(Val(Left(var5, 2)) & "-" & Val(Right(var5, 2)))
 End If
 If del_char = Len(var1) Then
  score = ""
 Else
  var1 = Trim(Mid(var1, del_char + 1, 99)) & "19-49"
  ls = Left(var1, InStr(var1, "-") - 1)
  Select Case Val(ls)
   Case Is < 10
    Rs = Mid(var1, 3, 2)
   Case 11
    Select Case InStr(4, var1, "-")
     Case 6
      Rs = Mid(var1, 4, 1)
     Case 8
      Rs = Mid(var1, 4, 2)
     Case Else
      If Mid(var1, 4, 2) = "13" Then
       Rs = "13"
      Else
       Rs = Mid(var1, 4, 1)
      End If
    End Select
   Case Else
    Rs = Mid(var1, 4, 2)
  End Select
  ls = IIf(Val(ls) < 10, "0", "") & ls
  Rs = IIf(Val(Rs) < 10, "0", "") & Rs
  score = ls & "~" & Rs
 End If
End Function
Function whowin(var1 As Variant)
ls = Left(var1, 1)
Rs = Right(var1, 1)
whowin = ""
If ls > Rs Then
 whowin = "Home"
End If
If Rs > ls Then
 whowin = "Away"
End If
End Function
Function ChkWO(var123 As String, Scr_Gam As Boolean, GameNo As String) As Boolean
If var123 = "11~0011~0011~00" Or var123 = "00~1100~1100~11" Then
  If Scr_Gam Then
   ChkWO = True
  Else
   Sheets("Tmp1").Select
   Msg = "Game scores of " & Mid(var123, 1, 5) & " " & Mid(var123, 6, 5) & " " & Mid(var123, 11, 5) & " detected" & vbCr
   Msg = Msg & "  in the " & GameNo & " Game with both players available" & vbCr & vbCr
   Msg = Msg & vbCr & "Is this Game a WalkOver ?"
   Style = vbYesNo + vbExclamation + vbDefaultButton2
   title = "Checking if a Game is a Walk Over"
   Ask_If_WO = MsgBox(Msg, Style, title)
   ChkWO = IIf(Ask_If_WO = vbYes, True, False)
   Sheets("Results").Select
  End If
Else
 ChkWO = False
End If
End Function
Function detilda(var1 As Variant)
detilda = Replace(var1, "~", "-")
End Function
Function Pl_Exist(var1 As Variant, var2 As Variant, var3 As Variant)
If Trim(var1) = "Forfeit" Then
 Pl_Exist = ""
Else
 lu_range = "Players!A1:A" & Tot_Plrs
 plyr_row = Application.Match(var1, Range(lu_range), 0)
 If IsError(plyr_row) Then
  plyr_row = 0
 End If
 If plyr_row = 0 Then
  Sheets("Players").Select
  Rows("2:2").Select
  Selection.Copy
  Selection.Insert Shift:=xlDown
  Range("A1").Select
  Tot_Plrs = Tot_Plrs + 1
  Cells(2, 1) = Trim(var1)
  Cells(2, 3) = Trim(var2)
  Cells(2, 4) = Trim(var3)
  Cells(2, 5) = 0
  Cells(2, 6) = 0
  Cells(2, 7) = 0
  Cells(2, 8) = ""
  Cells(2, 9) = 0
  Cells(2, 10) = 0
  Cells(2, 11) = 0
  Cells(2, 12) = 0
  Cells(2, 13) = 0
  Cells(2, 14) = 0
  Cells(2, 15) = 0
  Cells(2, 16) = 0
  Cells(2, 17) = 0
  Cells(2, 18) = 0
  lu_range = "'Player Library'!A1:A" & TotPLib
  lib_row = Application.Match(var1, Range(lu_range), 0)
  If IsError(lib_row) Then
   lib_row = 0
  End If
  If lib_row > 0 Then
   Sheets("Player Library").Rows(lib_row & ":" & lib_row).Font.ColorIndex = 3
   Cells(2, 2) = "*PLYR LIB - " & Sheets("Player Library").Cells(lib_row, 2)
   Cells(2, 5) = Sheets("Player Library").Cells(lib_row, 6)
   Cells(2, 6) = Cells(2, 5)
   If Cells(2, 5) > 0 Then
    Cells(2, 8) = "LIB->"
    Cells(2, 9) = Sheets("Player Library").Cells(lib_row, 9)
    Cells(2, 10) = Sheets("Player Library").Cells(lib_row, 10)
    Cells(2, 11) = Sheets("Player Library").Cells(lib_row, 11)
    Cells(2, 12) = Sheets("Player Library").Cells(lib_row, 12)
    Cells(2, 13) = Sheets("Player Library").Cells(lib_row, 13)
    Cells(2, 14) = Sheets("Player Library").Cells(lib_row, 14)
    Cells(2, 15) = Sheets("Player Library").Cells(lib_row, 15)
    Cells(2, 16) = Sheets("Player Library").Cells(lib_row, 16)
    Cells(2, 17) = Sheets("Player Library").Cells(lib_row, 17)
    Cells(2, 18) = Sheets("Player Library").Cells(lib_row, 18)
   End If
  Else
   Cells(2, 2) = "*NEW PLYR - " & Trim(Mid(UCase(Cells(2, 1)), InStr(Cells(2, 1), " ") + 1, 99)) & ", " & Left(Cells(2, 1), InStr(Cells(2, 1), " ") - 1)
  End If
  Lenb2 = Len("B2") - 12
  Range("B2").Select
  With ActiveCell.Characters(start:=1, Length:=12).Font
       .Name = "Arial"
       .FontStyle = "Bold"
       .Size = 10
       .ColorIndex = 3
  End With
  With ActiveCell.Characters(start:=13, Length:=Lenb2).Font
       .Name = "Arial"
       .FontStyle = "Regular"
       .Size = 10
       .ColorIndex = xlAutomatic
  End With
  Pl_Exist = "Div: " & Trim(var3) & " " & Trim(var2) & " " & Trim(var1) & vbCr
  Sheets("results").Select
 Else
  Pl_Exist = ""
 End If
End If
End Function
Function CellForm(v_Range, v_Align, v_Merge, v_Bold, v_FontClr, v_FontSize, v_Borders, v_Fill, v_FillClr)
Select Case v_Align
 Case "C"
  Range(v_Range).HorizontalAlignment = xlCenter
 Case "L"
  Range(v_Range).HorizontalAlignment = xlLeft
 Case "R"
  Range(v_Range).HorizontalAlignment = xlRight
End Select
Range(v_Range).MergeCells = v_Merge
Range(v_Range).Font.Bold = v_Bold
Range(v_Range).Font.ColorIndex = v_FontClr
Range(v_Range).Font.Size = v_FontSize
If v_Borders Then
 Range(v_Range).Borders(xlDiagonalDown).LineStyle = xlNone
 Range(v_Range).Borders(xlDiagonalUp).LineStyle = xlNone
 Range(v_Range).Borders(xlEdgeLeft).LineStyle = xlContinuous
 Range(v_Range).Borders(xlEdgeLeft).Weight = xlThin
 Range(v_Range).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
 Range(v_Range).Borders(xlEdgeTop).LineStyle = xlContinuous
 Range(v_Range).Borders(xlEdgeTop).Weight = xlThin
 Range(v_Range).Borders(xlEdgeTop).ColorIndex = xlAutomatic
 Range(v_Range).Borders(xlEdgeBottom).LineStyle = xlContinuous
 Range(v_Range).Borders(xlEdgeBottom).Weight = xlThin
 Range(v_Range).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
 Range(v_Range).Borders(xlEdgeRight).LineStyle = xlContinuous
 Range(v_Range).Borders(xlEdgeRight).Weight = xlThin
 Range(v_Range).Borders(xlEdgeRight).ColorIndex = xlAutomatic
 If Range(v_Range).Borders(xlInsideVertical).LineStyle <> 241 Then
  Range(v_Range).Borders(xlInsideVertical).LineStyle = xlContinuous
  Range(v_Range).Borders(xlInsideVertical).Weight = xlThin
  Range(v_Range).Borders(xlInsideVertical).ColorIndex = xlAutomatic
 End If
 If Range(v_Range).Borders(xlInsideHorizontal).LineStyle <> 241 Then
  Range(v_Range).Borders(xlInsideHorizontal).LineStyle = xlContinuous
  Range(v_Range).Borders(xlInsideHorizontal).Weight = xlThin
  Range(v_Range).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
 End If
 If v_Fill Then
  Range(v_Range).Interior.ColorIndex = v_FillClr
 End If
End If
End Function
Function PutFixts(Div_Var As Variant, tmp_cnt As Variant, Clear_Played As Boolean)
If Nxt_Fix_Col > 26 Then
 Nxt_Fix_Let = Chr(Int((Nxt_Fix_Col - 1) / 26) + 64) & Chr(((Nxt_Fix_Col - 1) Mod 26) + 65)
Else
 Nxt_Fix_Let = Chr(Nxt_Fix_Col + 64)
End If
If Nxt_Fix_Col + 5 > 26 Then
 Lst_Fix_Let = Chr(Int((Nxt_Fix_Col + 5 - 1) / 26) + 64) & Chr(((Nxt_Fix_Col + 5 - 1) Mod 26) + 65)
Else
 Lst_Fix_Let = Chr(Nxt_Fix_Col + 5 + 64)
End If
Sheets("Fixtures").Select
Cells(1, Nxt_Fix_Col) = "Division " & Div_Var & " (Rslt is a Hyperlink to Match Card)"
frng = Nxt_Fix_Let & "1:" & Lst_Fix_Let & "1"
Range(frng).Select
Application.CutCopyMode = False
Range(frng).HorizontalAlignment = xlCenter
Range(frng).MergeCells = True
Range(frng).Font.Name = "Arial"
Range(frng).Font.FontStyle = "Bold"
Range(frng).Font.Size = 12
Range(frng).Font.Underline = xlUnderlineStyleSingle
Range(frng).Font.ColorIndex = 3
Range("A2").Select
Sheets("tmp").Select
trng = "A2:F" & tmp_cnt
Range(trng).Select
Selection.Copy
Sheets("Fixtures").Select
frng = Nxt_Fix_Let & "2"
Range(frng).Select
ActiveSheet.Paste
Range("A3").Select
For f = 3 To tmp_cnt
 If Clear_Played Then
  Cells(f, (Nxt_Fix_Col - 1) + 6) = ""
 Else
  If Trim(Cells(f, (Nxt_Fix_Col - 1) + 6)) <> "" Then
   sco_txt = Cells(f, (Nxt_Fix_Col - 1) + 6)
   url_txt = Sheets("tmp").Cells(f, 13)
   frng = Lst_Fix_Let & f
   ActiveSheet.Hyperlinks.Add Anchor:=Range(frng), Address:=url_txt, TextToDisplay:=sco_txt
  End If
 End If
Next f
Range("A3").Select
Cells.Select
Selection.Columns.AutoFit
If Nxt_Fix_Col + 6 > 26 Then
 Blnk_Col = Chr(Int((Nxt_Fix_Col + 6 - 1) / 26) + 64) & Chr(((Nxt_Fix_Col + 6 - 1) Mod 26) + 65)
Else
 Blnk_Col = Chr(Nxt_Fix_Col + 6 + 64)
End If
frng = Blnk_Col & ":" & Blnk_Col
Columns(frng).Select
Selection.ColumnWidth = 3
Range("A3").Select
Sheets("Tmp").Select
PutFixts = (Nxt_Fix_Col + 7)
End Function
Function GetFixts(Div_Var As Variant, URL_Page As Variant)
new_rawfixt_format = False
Sheets("tmp").Select
 Cells.Select
 Selection.Clear
 ActiveWindow.FreezePanes = False
 Range("A1:K1").Select
 With Selection
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlBottom
  .WrapText = False
  .Orientation = 0
  .AddIndent = False
  .IndentLevel = 0
  .ShrinkToFit = False
  .ReadingOrder = xlContext
  .MergeCells = True
 End With
 With Selection.Font
  .Name = "Arial"
  .FontStyle = "Bold"
  .Size = 16
  .Strikethrough = False
  .Superscript = False
  .Subscript = False
  .OutlineFont = False
  .Shadow = False
  .Underline = xlUnderlineStyleSingle
  .ColorIndex = 3
 End With
 Cells(1, 1) = "DIVISION " & Div_Var & " FIXTURES FROM TT365 WEB"
 Cells(2, 1) = "WkNo"
 Cells(2, 2) = "Home Team"
 Cells(2, 3) = "VPR"
 Cells(2, 4) = "Away Team"
 Cells(2, 5) = "Match Date"
 Cells(2, 6) = "Rslt"
 Cells(2, 7) = "Plyr A"
 Cells(2, 8) = "Plyr B"
 Cells(2, 9) = "Plyr C"
 Cells(2, 10) = "Plyr X"
 Cells(2, 11) = "Plyr Y"
 Cells(2, 12) = "Plyr Z"
 Cells(2, 13) = "Result Card URL (HyperLink)"
 Rows("2:2").Select
 With Selection
       .HorizontalAlignment = xlCenter
       .VerticalAlignment = xlBottom
       .WrapText = False
       .Orientation = 0
       .AddIndent = False
       .IndentLevel = 0
       .ShrinkToFit = False
       .ReadingOrder = xlContext
       .MergeCells = False
 End With
 Range("M2:M2").HorizontalAlignment = xlLeft
 Selection.Font.Bold = True
 Range("A3").Select
 ActiveWindow.FreezePanes = True
 Range("A1").Select
 Sheets("tmp1").Select
 Cells.Select
 Selection.Clear
 ActiveWindow.FreezePanes = False
 Columns("A:E").Select
 Range("A2").Activate
 Selection.ColumnWidth = 20
 Range("A1:E1").Select
 With Selection
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlBottom
  .WrapText = False
  .Orientation = 0
  .AddIndent = False
  .IndentLevel = 0
  .ShrinkToFit = False
  .ReadingOrder = xlContext
  .MergeCells = True
 End With
 With Selection.Font
  .Name = "Arial"
  .FontStyle = "Bold"
  .Size = 16
  .Strikethrough = False
  .Superscript = False
  .Subscript = False
  .OutlineFont = False
  .Shadow = False
  .Underline = xlUnderlineStyleSingle
  .ColorIndex = 3
 End With
 Cells(1, 1) = "GETTING DIVISION " & Div_Var & " FIXTURES FROM TT365 WEB"
 Range("A2").Select
 ActiveWindow.FreezePanes = True
 Range("A1").Select
 With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;" & URL_Page, Destination:=Range("A2"))
        .Name = "Division_" & Div_Var
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingAll
        .WebPreFormattedTextToColumns = False
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = True
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
  lastrow = ActiveSheet.UsedRange.Rows.Count
  Sheets("tmp").Select
  row_num = 2
  wn = 0
  If new_rawfixt_format Then 'If new_rawfixt_format
   For Z = 2 To lastrow
    If Left(Sheets("tmp1").Cells(Z, 1), Len(Fixt_Note)) = Fixt_Note And Len(Sheets("tmp1").Cells(Z, 1)) > 20 Then
     wn = Mid(Sheets("tmp1").Cells(Z, 1), Len(Fixt_Note) + 1, InStr(Sheets("tmp1").Cells(Z, 1), " -") - (Len(Fixt_Note) + 1))
    Else
     If wn > 0 Then
      chk_cell = Trim(Sheets("tmp1").Cells(Z, 1))
      If InStr(chk_cell, " ") > 0 Then
       fixt_date = IIf(InStr("MondayTuesdayWednesdayThursdayFriday", Left(chk_cell, InStr(chk_cell, " ") - 1)) = 0, False, True)
       If fixt_date Then
        row_num = row_num + 1
        Cells(row_num, 1) = wn 'WkNo
        md = Sheets("tmp1").Cells(Z, 1)
        sp1 = InStr(1, md, " ")
        sp2 = InStr(sp1 + 1, md, " ")
        sp3 = InStr(sp2 + 1, md, " ")
        Cells(row_num, 5) = Left(md, 3) & " " _
                             & IIf(Mid(md, sp1 + 1, sp2 - sp1 - 3) < 10, "0", "") & Mid(md, sp1 + 1, sp2 - sp1 - 3) & "/" _
                             & IIf(Int(InStr(mdlu, Mid(md, sp2 + 1, 3)) / 3) + 1 < 10, "0", "") & Int(InStr(mdlu, Mid(md, sp2 + 1, 3)) / 3) + 1 & "/" _
                             & Mid(md, sp3 + 1, 4) 'Match Date
        If Trim(Sheets("tmp1").Cells(Z + 7, 1)) = "vs" Then
         Cells(row_num, 2) = Sheets("tmp1").Cells(Z + 3, 1) 'Home Team
         Cells(row_num, 3) = Sheets("tmp1").Cells(Z + 7, 1) 'vs Flag
         Cells(row_num, 4) = Sheets("tmp1").Cells(Z + 9, 1)  'Away Team
         Cells(row_num, 6) = Sheets("tmp1").Cells(Z + 2, 1) & "~" & Sheets("tmp1").Cells(Z + 8, 1) 'Rslt
         Cells(row_num, 7) = Trim_Plyr(Sheets("tmp1").Cells(Z + 4, 1)) 'plyA
         Cells(row_num, 8) = Trim_Plyr(Sheets("tmp1").Cells(Z + 5, 1)) 'plyrB
         Cells(row_num, 9) = Trim_Plyr(Sheets("tmp1").Cells(Z + 6, 1)) 'plyrC
         Cells(row_num, 10) = Trim_Plyr(Sheets("tmp1").Cells(Z + 10, 1)) 'plyrX
         Cells(row_num, 11) = Trim_Plyr(Sheets("tmp1").Cells(Z + 11, 1)) 'plyrY
         Cells(row_num, 12) = Trim_Plyr(Sheets("tmp1").Cells(Z + 12, 1)) 'plyrZ
         Cells(row_num, 13) = Sheets("tmp1").Cells(Z + 2, 1) 'Result Card URL
         Cells(row_num, 13) = Sheets("tmp1").Cells(Z + 2, 1).Hyperlinks.Item(1).Address
         url_txt = Cells(row_num, 13)
         Rng = "M" & row_num
         ActiveSheet.Hyperlinks.Add Anchor:=Range(Rng), Address:=url_txt, TextToDisplay:=url_txt
        Else
         If Trim(Sheets("tmp1").Cells(Z + 3, 1)) = "P" Or Trim(Sheets("tmp1").Cells(Z + 3, 1)) = "R" Then
          Cells(row_num, 2) = Sheets("tmp1").Cells(Z + 2, 1) 'Home Team
          Cells(row_num, 3) = Sheets("tmp1").Cells(Z + 3, 1) 'vs Flag P or R Flag
          Cells(row_num, 4) = Sheets("tmp1").Cells(Z + 4, 1) 'Away Team
         Else
          Cells(row_num, 2) = Sheets("tmp1").Cells(Z + 2, 1) 'Home Team
          Cells(row_num, 3) = "" 'vs Flag No Flag
          Cells(row_num, 4) = Sheets("tmp1").Cells(Z + 3, 1) 'Away Team
         End If
        End If
       End If
      End If
     End If
    End If
   Next Z
  End If 'End of If new_rawfixt_format
  If Not new_rawfixt_format Then ' If not new_rawfixt_format
   For Z = 2 To lastrow
    If Left(Sheets("tmp1").Cells(Z, 1), 17) = "Fixture List View" Then
     Exit For
    End If
    If Left(Sheets("tmp1").Cells(Z, 1), Len(Fixt_Note)) = Fixt_Note _
    And (Left(Sheets("tmp1").Cells(Z + 1, 1), 4) = "Free" _
    Or Left(Sheets("tmp1").Cells(Z + 1, 1), 12) = "KO Competion" _
    Or Left(Sheets("tmp1").Cells(Z + 1, 1), 14) = "KO Competition" _
    Or Left(Sheets("tmp1").Cells(Z + 1, 1), 13) = "Turner Trophy" _
    Or Left(Sheets("tmp1").Cells(Z + 1, 1), 9) = "Xmas Holz" _
    Or Left(Sheets("tmp1").Cells(Z + 1, 1), 13) = "New Year Holz" _
    Or Left(Sheets("tmp1").Cells(Z + 1, 1), 15) = "Christmas Break" _
    Or Left(Sheets("tmp1").Cells(Z + 1, 1), 16) = "Competition Week" _
    Or Left(Sheets("tmp1").Cells(Z + 1, 1), 15) = "**Competition**" _
    Or Left(Sheets("tmp1").Cells(Z + 1, 1), 16) = "**Winter Break**" _
    Or Left(Sheets("tmp1").Cells(Z + 1, 1), 16) = "** Free" _
    ) Then
     Z = Z + 1
    Else
    If Left(Sheets("tmp1").Cells(Z, 1), Len(Fixt_Note)) = Fixt_Note And Len(Sheets("tmp1").Cells(Z, 1)) > 20 Then
     wn = Mid(Sheets("tmp1").Cells(Z, 1), Len(Fixt_Note) + 1, InStr(Sheets("tmp1").Cells(Z, 1), " -") - (Len(Fixt_Note) + 1))
    Else
     If wn > 0 Then
      If InStr("~vs~P~R~", "~" & Sheets("tmp1").Cells(Z + 3, 1) & "~") > 0 Then ' Match Not Played
        vpr = Sheets("tmp1").Cells(Z + 3, 1)
        row_num = row_num + 1
        md = Sheets("tmp1").Cells(Z, 1)
        sp1 = InStr(1, md, " ")
        sp2 = InStr(sp1 + 1, md, " ")
        sp3 = InStr(sp2 + 1, md, " ")
        Cells(row_num, 1) = wn 'WkNo
        Cells(row_num, 2) = Sheets("tmp1").Cells(Z + 2, 1) 'Home Team
        Cells(row_num, 3) = vpr 'VPR
        Cells(row_num, 4) = Sheets("tmp1").Cells(Z + 4, 1) 'Away Team
        Cells(row_num, 5) = Left(md, 3) & " " _
                            & IIf(Mid(md, sp1 + 1, sp2 - sp1 - 3) < 10, "0", "") & Mid(md, sp1 + 1, sp2 - sp1 - 3) & "/" _
                            & IIf(Int(InStr(mdlu, Mid(md, sp2 + 1, 3)) / 3) + 1 < 10, "0", "") & Int(InStr(mdlu, Mid(md, sp2 + 1, 3)) / 3) + 1 & "/" _
                            & Mid(md, sp3 + 1, 4) 'Match Date
         Z = Z + 4
       ElseIf Sheets("tmp1").Cells(Z + 3, 1).text = "V" Then ' Void Match
         Z = Z + 4
       ElseIf InStr("~vs~P~R~", "~" & Sheets("tmp1").Cells(Z + 7 - IIf(Plyrs2, 1, 0), 1) & "~") > 0 Then ' Match Played
        vpr = Sheets("tmp1").Cells(Z + 7 - IIf(Plyrs2, 1, 0), 1)
        row_num = row_num + 1
        md = Sheets("tmp1").Cells(Z, 1)
        sp1 = InStr(1, md, " ")
        sp2 = InStr(sp1 + 1, md, " ")
        sp3 = InStr(sp2 + 1, md, " ")
        Cells(row_num, 1) = wn 'WkNo
        Cells(row_num, 2) = Sheets("tmp1").Cells(Z + 3, 1) 'Home Team
        Cells(row_num, 3) = vpr 'VPR
        Cells(row_num, 4) = Sheets("tmp1").Cells(Z + 9 - IIf(Plyrs2, 1, 0), 1) 'Away Team
        Cells(row_num, 5) = Left(md, 3) & " " _
                            & IIf(Mid(md, sp1 + 1, sp2 - sp1 - 3) < 10, "0", "") & Mid(md, sp1 + 1, sp2 - sp1 - 3) & "/" _
                            & IIf(Int(InStr(mdlu, Mid(md, sp2 + 1, 3)) / 3) + 1 < 10, "0", "") & Int(InStr(mdlu, Mid(md, sp2 + 1, 3)) / 3) + 1 & "/" _
                            & Mid(md, sp3 + 1, 4) 'Match Date
        Cells(row_num, 6) = Sheets("tmp1").Cells(Z + 2, 1) & "~" & Sheets("tmp1").Cells(Z + 8 - IIf(Plyrs2, 1, 0), 1) 'Rslt
        Cells(row_num, 7) = Trim_Plyr(Sheets("tmp1").Cells(Z + 4, 1)) 'plyA
        Cells(row_num, 8) = Trim_Plyr(Sheets("tmp1").Cells(Z + 5, 1)) 'plyrB
         Cells(row_num, 9) = Trim_Plyr(Sheets("tmp1").Cells(Z + 6, 1)) 'plyrC
         Cells(row_num, 10) = Trim_Plyr(Sheets("tmp1").Cells(Z + 10, 1)) 'plyrX
        Cells(row_num, 11) = Trim_Plyr(Sheets("tmp1").Cells(Z + 11, 1)) 'plyrY
        
        Cells(row_num, 12) = Trim_Plyr(Sheets("tmp1").Cells(Z + 12, 1)) 'plyrC

        
        
        Cells(row_num, 13) = Sheets("tmp1").Cells(Z + 2, 1) 'Result Card URL
        Cells(row_num, 13) = Sheets("tmp1").Cells(Z + 2, 1).Hyperlinks.Item(1).Address
        url_txt = Cells(row_num, 13)
        Rng = "M" & row_num
        ActiveSheet.Hyperlinks.Add Anchor:=Range(Rng), Address:=url_txt, TextToDisplay:=url_txt
        Z = Z + 12
       ElseIf Trim(Sheets("tmp1").Cells(Z + 4, 1)) = "vs" Then ''Match has been awarded dont record Match Fixture data this will become apparent in anomolies sheet
        vpr = "Awd"
        row_num = row_num + 1
        md = Sheets("tmp1").Cells(Z, 1)
        sp1 = InStr(1, md, " ")
        sp2 = InStr(sp1 + 1, md, " ")
        sp3 = InStr(sp2 + 1, md, " ")
        Cells(row_num, 1) = wn 'WkNo
        Cells(row_num, 2) = Sheets("tmp1").Cells(Z + 3, 1) 'Home Team
        Cells(row_num, 3) = vpr 'VPR
        Cells(row_num, 4) = Sheets("tmp1").Cells(Z + 6, 1) 'Away Team
        Cells(row_num, 5) = Left(md, 3) & " " _
                            & IIf(Mid(md, sp1 + 1, sp2 - sp1 - 3) < 10, "0", "") & Mid(md, sp1 + 1, sp2 - sp1 - 3) & "/" _
                            & IIf(Int(InStr(mdlu, Mid(md, sp2 + 1, 3)) / 3) + 1 < 10, "0", "") & Int(InStr(mdlu, Mid(md, sp2 + 1, 3)) / 3) + 1 & "/" _
                            & Mid(md, sp3 + 1, 4) 'Match Date
        Cells(row_num, 6) = Sheets("tmp1").Cells(Z + 2, 1) & "~" & Sheets("tmp1").Cells(Z + 5, 1) 'Rslt
        Cells(row_num, 7) = "Forfeit" 'plyA
        Cells(row_num, 8) = "Forfeit" 'plyB
        Cells(row_num, 9) = "Forfeit" 'plyC
        Cells(row_num, 10) = "Forfeit" 'plyX
        Cells(row_num, 11) = "Forfeit" 'plyY
        Cells(row_num, 12) = "Forfeit" 'plyZ
        Cells(row_num, 13) = Sheets("tmp1").Cells(Z + 2, 1) 'Result Card URL
        Cells(row_num, 13) = Sheets("tmp1").Cells(Z + 2, 1).Hyperlinks.Item(1).Address
        url_txt = Cells(row_num, 13)
        Rng = "M" & row_num
        ActiveSheet.Hyperlinks.Add Anchor:=Range(Rng), Address:=url_txt, TextToDisplay:=url_txt
        Z = Z + 6
       End If
      End If
     End If
    End If
    Next Z
   End If 'End of If not new_rawfixt_format
   Range("A2:M" & row_num).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 5
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 5
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 5
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 5
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 5
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 5
    End With
   Cells.Select
   Selection.Columns.AutoFit
   Range("A3").Select
End Function


Function GetParms(Tmp_WS As Boolean, Tmp1_WS As Boolean)
' Get Config Parameters and Create tmp & tmp1 Worksheets
WB_Name = ActiveWorkbook.Name
Sheets("Config").Select
lastrow = ActiveSheet.UsedRange.Rows.Count
Rpt_Tm_Dot = True ' Default if Parameter Missing
Card_Type = "A" ' Default if Parameter Missing
Chk_Plr_Clb = False ' Default if Parameter Missing
For x = 1 To lastrow
 Select Case Trim(Cells(x, 1))
  Case "League"
   League = Trim(Cells(x, 2))
  Case "Season"
   Season = Trim(Cells(x, 2))
  Case "Divisions"
   Num_Divs = Len(Trim(Cells(x, 2)))
   For y = 1 To Num_Divs
    Divs(y, 1) = Mid(Cells(x, 2), y, 1)
   Next y
  Case "Division Formats"
   For y = 1 To Num_Divs
    Select Case Mid(Cells(x, 2), y, 1)
     Case "S"
      Divs(y, 2) = "9S0D"
     Case "D"
      Divs(y, 2) = "9S1D"
     Case "N"
      Divs(y, 2) = "6S3D"
     End Select
   Next y
  Case "Card_Format_Type"
   Card_Type = UCase(Trim(Cells(x, 2)))
  Case "Rpts_Team_Dots"
   Rpt_Tm_Dot = IIf(UCase(Trim(Cells(x, 2))) = "YES", True, False)
  Case "Chk_Plyr_Club"
   Chk_Plr_Clb = IIf(UCase(Trim(Cells(x, 2))) = "YES", True, False)
  Case "FIXURL_1st"
   Divs(1, 3) = Replace(Trim(Cells(x, 2)), "#", "")
  Case "FIXURL_2nd"
   Divs(2, 3) = Replace(Trim(Cells(x, 2)), "#", "")
  Case "FIXURL_3rd"
   Divs(3, 3) = Replace(Trim(Cells(x, 2)), "#", "")
  Case "FIXURL_4th"
   Divs(4, 3) = Replace(Trim(Cells(x, 2)), "#", "")
  Case "FIXURL_5th"
   Divs(5, 3) = Replace(Trim(Cells(x, 2)), "#", "")
  Case "FIXURL_6th"
   Divs(6, 3) = Replace(Trim(Cells(x, 2)), "#", "")
  Case "FIXURL_7th"
   Divs(7, 3) = Replace(Trim(Cells(x, 2)), "#", "")
  Case "FIXURL_8th"
   Divs(8, 3) = Replace(Trim(Cells(x, 2)), "#", "")
  Case "FIXURL_9th"
   Divs(9, 3) = Replace(Trim(Cells(x, 2)), "#", "")
  Case "Fixture Notation"
   Fixt_Note = Trim(Cells(x, 2))
  Case "Non_Rated_Comp"
   NR_Comp = UCase(Trim(Cells(x, 2)))
   If NR_Comp <> "PLAYED" Then
    NR_Comp = "ALL"
   End If
  Case "R1_Jumpers"
   Jump_Rpt = Cells(x, 2)
  Case "Macro Delay"
   MacDelay = Trim(Cells(x, 2))
 End Select
Next x
Application.DisplayAlerts = False
For Each ws In Worksheets
 Select Case ws.Name
  Case "tmp"
   Sheets("tmp").Delete
  Case "tmp1"
   Sheets("tmp1").Delete
 End Select
Next ws
Application.DisplayAlerts = True
If Tmp1_WS Then
 Sheets.Add.Name = "tmp1"
 Sheets("tmp1").Tab.ColorIndex = 15
 'ActiveWindow.DisplayGridlines = False
 Sheets("tmp1").Move After:=Sheets("Config")
End If
If Tmp_WS Then
 Sheets.Add.Name = "tmp"
 Sheets("tmp").Tab.ColorIndex = 15
 'ActiveWindow.DisplayGridlines = False
 Sheets("tmp").Move After:=Sheets("Config")
End If
End Function
Function Tidy_Up(ShapeID As Variant)
Application.DisplayAlerts = False
For Each ws In Worksheets
 Select Case ws.Name
  Case "tmp"
   Sheets("tmp").Delete
  Case "tmp1"
   Sheets("tmp1").Delete
 End Select
Next ws
Application.DisplayAlerts = True
Sheets("Menu").Select
If Trim(ShapeID) <> "" Then
 ActiveSheet.Unprotect
 ActiveSheet.Shapes(ShapeID).Select
 Selection.Characters.text = "Last Run: " & Format(Now(), "ddmmmYY HH:MM:SS")
 ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End If
Range("A1").Select
End Function
Function PrintSetup(ColRange As Variant, LHtxt As Variant, CHtxt As Variant, RHtxt As Variant, CFtxt As Variant, SinglePage As Boolean)
ActiveSheet.Cells.Columns.AutoFit
ActiveSheet.Range("A2").Select
SelRange = ColRange & Trim(ActiveSheet.UsedRange.Rows.Count)
ActiveSheet.PageSetup.PrintArea = SelRange
ActiveSheet.PageSetup.LeftHeader = LHtxt
ActiveSheet.PageSetup.CenterHeader = CHtxt
ActiveSheet.PageSetup.RightHeader = RHtxt
ActiveSheet.PageSetup.CenterFooter = CFtxt
ActiveSheet.PageSetup.LeftMargin = Application.InchesToPoints(0.354330708661417)
ActiveSheet.PageSetup.RightMargin = Application.InchesToPoints(0.354330708661417)
ActiveSheet.PageSetup.TopMargin = Application.InchesToPoints(0.984251968503937)
ActiveSheet.PageSetup.BottomMargin = Application.InchesToPoints(0.393700787401575)
ActiveSheet.PageSetup.HeaderMargin = Application.InchesToPoints(0.393700787401575)
ActiveSheet.PageSetup.FooterMargin = Application.InchesToPoints(0.118110236220472)
ActiveSheet.PageSetup.CenterHorizontally = True
ActiveSheet.PageSetup.Orientation = xlPortrait
ActiveSheet.PageSetup.PaperSize = xlPaperA4
ActiveSheet.PageSetup.BlackAndWhite = False
If SinglePage Then
 ActiveSheet.PageSetup.Zoom = False
 ActiveSheet.PageSetup.FitToPagesWide = 1
 ActiveSheet.PageSetup.FitToPagesTall = 1
Else
 ActiveSheet.PageSetup.Zoom = False
 ActiveSheet.PageSetup.FitToPagesWide = 1
 ActiveSheet.PageSetup.FitToPagesTall = 20
End If
End Function
Function Rem_Rpts()
Application.DisplayAlerts = False
For Each ws In Worksheets
 If ws.Name = "R1_Jumpers" Then
  Sheets("R1_Jumpers").Delete
 End If
Next ws
For Each ws In Worksheets
 If ws.Name = "R2_RtgByPlyr" Then
  Sheets("R2_RtgByPlyr").Delete
 End If
Next ws
For Each ws In Worksheets
 If ws.Name = "R3_RtgByRtg" Then
  Sheets("R3_RtgByRtg").Delete
 End If
Next ws
For Each ws In Worksheets
 If ws.Name = "R4_PlyrStats" Then
  Sheets("R4_PlyrStats").Delete
 End If
Next ws
Application.DisplayAlerts = True
Sheets("Menu").Select
ActiveSheet.Unprotect
ActiveSheet.Shapes("V_Rpt1_Box").TextFrame.Characters.Font.ColorIndex = 16
ActiveSheet.Shapes("V_Rpt1_Box").OnAction = ""
ActiveSheet.Shapes("V_Rpt2_Box").TextFrame.Characters.Font.ColorIndex = 16
ActiveSheet.Shapes("V_Rpt2_Box").OnAction = ""
ActiveSheet.Shapes("V_Rpt3_Box").TextFrame.Characters.Font.ColorIndex = 16
ActiveSheet.Shapes("V_Rpt3_Box").OnAction = ""
ActiveSheet.Shapes("V_Rpt4_Box").TextFrame.Characters.Font.ColorIndex = 16
ActiveSheet.Shapes("V_Rpt4_Box").OnAction = ""
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Function

Function WaitFor(tvar As Variant)
' Delay for tvar which is whole or fraction of seconds
If UCase(MacDelay) = "YES" Then
 TimeNow = Timer
 While Timer - TimeNow < tvar
 Wend
End If
End Function

Function Delay(tvar As Variant)
' Delay for tvar which is whole or fraction of seconds
 TimeNow = Timer
 While Timer - TimeNow < tvar
 Wend
End Function

Function Trim_Plyr(pvar)
 len_pvar = Len(Trim(pvar))
 pvar = Left(Trim(pvar), len_pvar - 3)
'Par_Start = InStr(pvar, "(")
'If Par_Start > 0 Then
' pvar = Left(pvar, Par_Start - 1)
'End If
Trim_Plyr = Trim(pvar)
End Function


Function Card_TypeA(div_format As Variant)
' Type A Card Format in TT365 (Normal used in NWK and Gravesend)
start_row = 1
While Trim(Cells(start_row, 1)) <> "Fixture Details"
 start_row = start_row + 1
Wend
pd = Cells(start_row + 3, 1)

' As from 11Apr18 TT365 Prefixed Date Played with Match Date and suffixed with reason if postponment
pd = Replace(pd, "Match Date: ", "")
pd = Left(pd, 11)



While Trim(Cells(start_row, 1)) <> "Home Player"
 start_row = start_row + 1
Wend
Delay = WaitFor(2)
rst_cnt = start_row + 1
tmp = Rmv_Spc(Cells(rst_cnt, 3))
GA1_1 = score(tmp, "", "", "", "")
GA1_2 = score(tmp, GA1_1, "", "", "")
GA1_3 = score(tmp, GA1_1, GA1_2, "", "")
GA1_4 = score(tmp, GA1_1, GA1_2, GA1_3, "")
GA1_5 = score(tmp, GA1_1, GA1_2, GA1_3, GA1_4)
GA1_6 = whowin(Cells(rst_cnt, 4))
rst_cnt = rst_cnt + 1
While Cells(rst_cnt, 4) = ""
 rst_cnt = rst_cnt + 1
Wend
tmp = Rmv_Spc(Cells(rst_cnt, 3))
GA2_1 = score(tmp, "", "", "", "")
GA2_2 = score(tmp, GA2_1, "", "", "")
GA2_3 = score(tmp, GA2_1, GA2_2, "", "")
GA2_4 = score(tmp, GA2_1, GA2_2, GA2_3, "")
GA2_5 = score(tmp, GA2_1, GA2_2, GA2_3, GA2_4)
GA2_6 = whowin(Cells(rst_cnt, 4))
rst_cnt = rst_cnt + 1
While Cells(rst_cnt, 4) = ""
 rst_cnt = rst_cnt + 1
Wend
tmp = Rmv_Spc(Cells(rst_cnt, 3))
GA3_1 = score(tmp, "", "", "", "")
GA3_2 = score(tmp, GA3_1, "", "", "")
GA3_3 = score(tmp, GA3_1, GA3_2, "", "")
GA3_4 = score(tmp, GA3_1, GA3_2, GA3_3, "")
GA3_5 = score(tmp, GA3_1, GA3_2, GA3_3, GA3_4)
GA3_6 = whowin(Cells(rst_cnt, 4))
rst_cnt = rst_cnt + 1
While Cells(rst_cnt, 4) = ""
 rst_cnt = rst_cnt + 1
Wend
tmp = Rmv_Spc(Cells(rst_cnt, 3))
GA4_1 = score(tmp, "", "", "", "")
GA4_2 = score(tmp, GA4_1, "", "", "")
GA4_3 = score(tmp, GA4_1, GA4_2, "", "")
GA4_4 = score(tmp, GA4_1, GA4_2, GA4_3, "")
GA4_5 = score(tmp, GA4_1, GA4_2, GA4_3, GA4_4)
GA4_6 = whowin(Cells(rst_cnt, 4))
rst_cnt = rst_cnt + 1
While Cells(rst_cnt, 4) = ""
 rst_cnt = rst_cnt + 1
Wend
tmp = Rmv_Spc(Cells(rst_cnt, 3))
GA5_1 = score(tmp, "", "", "", "")
GA5_2 = score(tmp, GA5_1, "", "", "")
GA5_3 = score(tmp, GA5_1, GA5_2, "", "")
GA5_4 = score(tmp, GA5_1, GA5_2, GA5_3, "")
GA5_5 = score(tmp, GA5_1, GA5_2, GA5_3, GA5_4)
GA5_6 = whowin(Cells(rst_cnt, 4))
rst_cnt = rst_cnt + 1
While Cells(rst_cnt, 4) = ""
 rst_cnt = rst_cnt + 1
Wend
tmp = Rmv_Spc(Cells(rst_cnt, 3))
GA6_1 = score(tmp, "", "", "", "")
GA6_2 = score(tmp, GA6_1, "", "", "")
GA6_3 = score(tmp, GA6_1, GA6_2, "", "")
GA6_4 = score(tmp, GA6_1, GA6_2, GA6_3, "")
GA6_5 = score(tmp, GA6_1, GA6_2, GA6_3, GA6_4)
GA6_6 = whowin(Cells(rst_cnt, 4))
rst_cnt = rst_cnt + 1
While Cells(rst_cnt, 4) = ""
 rst_cnt = rst_cnt + 1
Wend
tmp = Rmv_Spc(Cells(rst_cnt, 3))
GA7_1 = score(tmp, "", "", "", "")
GA7_2 = score(tmp, GA7_1, "", "", "")
GA7_3 = score(tmp, GA7_1, GA7_2, "", "")
GA7_4 = score(tmp, GA7_1, GA7_2, GA7_3, "")
GA7_5 = score(tmp, GA7_1, GA7_2, GA7_3, GA7_4)
GA7_6 = whowin(Cells(rst_cnt, 4))
rst_cnt = rst_cnt + 1
While Cells(rst_cnt, 4) = ""
 rst_cnt = rst_cnt + 1
Wend
tmp = Rmv_Spc(Cells(rst_cnt, 3))
GA8_1 = score(tmp, "", "", "", "")
GA8_2 = score(tmp, GA8_1, "", "", "")
GA8_3 = score(tmp, GA8_1, GA8_2, "", "")
GA8_4 = score(tmp, GA8_1, GA8_2, GA8_3, "")
GA8_5 = score(tmp, GA8_1, GA8_2, GA8_3, GA8_4)
GA8_6 = whowin(Cells(rst_cnt, 4))
rst_cnt = rst_cnt + 1
While Cells(rst_cnt, 4) = ""
 rst_cnt = rst_cnt + 1
Wend
tmp = Rmv_Spc(Cells(rst_cnt, 3))
GA9_1 = score(tmp, "", "", "", "")
GA9_2 = score(tmp, GA9_1, "", "", "")
GA9_3 = score(tmp, GA9_1, GA9_2, "", "")
GA9_4 = score(tmp, GA9_1, GA9_2, GA9_3, "")
GA9_5 = score(tmp, GA9_1, GA9_2, GA9_3, GA9_4)
GA9_6 = whowin(Cells(rst_cnt, 4))
If div_format = "9S1D" Then
 rst_cnt = rst_cnt + 1
 While Cells(rst_cnt, 4) = ""
  rst_cnt = rst_cnt + 1
 Wend
 tmp = Rmv_Spc(Cells(rst_cnt, 3))
 GA10_1 = score(tmp, "", "", "", "")
 GA10_2 = score(tmp, GA10_1, "", "", "")
 GA10_3 = score(tmp, GA10_1, GA10_2, "", "")
 GA10_4 = score(tmp, GA10_1, GA10_2, GA10_3, "")
 GA10_5 = score(tmp, GA10_1, GA10_2, GA10_3, GA10_4)
 GA10_6 = whowin(Cells(rst_cnt, 4))
End If
End Function

Function Rmv_Spc(var_score)
''' Remove spaces in score strings - TT365 changed with spaces 20Dec17
''' Putting back to example: 11-811-131-111012 from 11-8 11-13 1-11 10-12
Rmv_Spc = Replace(var_score, " ", "")
End Function

Function Card_TypeB(div_format As Variant)
' Type B Card Format in TT365 (Used in Isle of Wight)
Crd_Rws = ActiveSheet.UsedRange.Rows.Count
start_row = 1
Dim tmp_sc(1 To 60) As String
For ts = 1 To 60
 tmp_sc(ts) = ""
Next ts
ts = 0
last_set = 0
hwg = 0
awg = 0
While Trim(Cells(start_row, 1)) <> "Fixture Details"
 start_row = start_row + 1
Wend
Delay = WaitFor(2)
pd = Mid(Cells(start_row + 3, 1), 13)
For rst_cnt = start_row + 4 To Crd_Rws
 tmp = Trim(Cells(rst_cnt, 1))
 If tmp <> "" Then
  If Left(tmp, 11) = "Recorded By" Or Left(tmp, 12) = "Submitted By" Then
   Exit For
  End If
  If Asc(Left(tmp, 1)) > 47 And Asc(Left(tmp, 1)) < 58 Then
   ts = ts + 1
   dpos = InStr(tmp, "-")
   ls = Left(tmp, dpos - 1)
   Rs = Right(tmp, Len(tmp) - dpos)
   hwg = hwg + IIf(Val(ls) > Val(Rs), 1, 0)
   awg = awg + IIf(Val(Rs) > Val(ls), 1, 0)
   ls = IIf(Val(ls) < 10, "0", "") & ls
   Rs = IIf(Val(Rs) < 10, "0", "") & Rs
   tmp_sc(ts) = ls & "~" & Rs
   If awg = 3 Or hwg = 3 Then
    While ts / 6 <> Int(ts / 6)
     ts = ts + 1
     tmp_sc(ts) = ""
    Wend
    tmp_sc(ts) = IIf(hwg > awg, "Home", "Away")
    hwg = 0
    awg = 0
   End If
  End If
 End If
Next rst_cnt
GA1_1 = tmp_sc(1)
GA1_2 = tmp_sc(2)
GA1_3 = tmp_sc(3)
GA1_4 = tmp_sc(4)
GA1_5 = tmp_sc(5)
GA1_6 = tmp_sc(6)
GA2_1 = tmp_sc(25)
GA2_2 = tmp_sc(26)
GA2_3 = tmp_sc(27)
GA2_4 = tmp_sc(28)
GA2_5 = tmp_sc(29)
GA2_6 = tmp_sc(30)
GA3_1 = tmp_sc(49)
GA3_2 = tmp_sc(50)
GA3_3 = tmp_sc(51)
GA3_4 = tmp_sc(52)
GA3_5 = tmp_sc(53)
GA3_6 = tmp_sc(54)
GA4_1 = tmp_sc(19)
GA4_2 = tmp_sc(20)
GA4_3 = tmp_sc(21)
GA4_4 = tmp_sc(22)
GA4_5 = tmp_sc(23)
GA4_6 = tmp_sc(24)
GA5_1 = tmp_sc(13)
GA5_2 = tmp_sc(14)
GA5_3 = tmp_sc(15)
GA5_4 = tmp_sc(16)
GA5_5 = tmp_sc(17)
GA5_6 = tmp_sc(18)
GA6_1 = tmp_sc(43)
GA6_2 = tmp_sc(44)
GA6_3 = tmp_sc(45)
GA6_4 = tmp_sc(46)
GA6_5 = tmp_sc(47)
GA6_6 = tmp_sc(48)
GA7_1 = tmp_sc(31)
GA7_2 = tmp_sc(32)
GA7_3 = tmp_sc(33)
GA7_4 = tmp_sc(34)
GA7_5 = tmp_sc(35)
GA7_6 = tmp_sc(36)
GA8_1 = tmp_sc(37)
GA8_2 = tmp_sc(38)
GA8_3 = tmp_sc(39)
GA8_4 = tmp_sc(40)
GA8_5 = tmp_sc(41)
GA8_6 = tmp_sc(42)
GA9_1 = tmp_sc(7)
GA9_2 = tmp_sc(8)
GA9_3 = tmp_sc(9)
GA9_4 = tmp_sc(10)
GA9_5 = tmp_sc(11)
GA9_6 = tmp_sc(12)
If div_format = "9S1D" Then
 GA10_1 = tmp_sc(55)
 GA10_2 = tmp_sc(56)
 GA10_3 = tmp_sc(57)
 GA10_4 = tmp_sc(58)
 GA10_5 = tmp_sc(59)
 GA10_6 = tmp_sc(60)
End If
End Function

Public Sub TestProcessFixtures()
    Dim oFixtureHelper As FixtureHelper
    Set oFixtureHelper = New FixtureHelper
    Dim divisions As Collection
    Set divisions = New Collection
    divisions.Add "https://www.tabletennis365.com/IsleOfWight/Fixtures/Winter_2018-19/Division_1"
    divisions.Add "https://www.tabletennis365.com/IsleOfWight/Fixtures/Winter_2018-19/Division_2"
    divisions.Add "https://www.tabletennis365.com/IsleOfWight/Fixtures/Winter_2018-19/Division_3"
    
    oFixtureHelper.ProcessFixtures divisions
    
End Sub
