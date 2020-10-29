# hello-world
first github
Hello
I'm about to become the best coder.

   
Sub Trade()
shtadd ("Cash")
shtadd ("CM_UL")
shtadd ("pivot")
shtadd ("FO_UL")
shtadd ("Future")
shtadd ("Opption")
shtadd ("pivot2")
shtadd ("Summery")
spath = ThisWorkbook.Path & "\input\"
    ''''''''''''''''cm zip''''''''''''''
    cv = GetFullFileName(spath, "CM_UL_")
    
    If cv <> Empty Or cv <> "" Then
       Set obj1 = Workbooks.Open(Filename:=spath & cv)
    End If
    shtnname = ActiveSheet.Name
    
    rcct = Sheets(shtnname).UsedRange.Rows.Count
    ccct = Sheets(shtnname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Selection.Copy
    ThisWorkbook.Worksheets("CM_UL").Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveSheet.Range("A1:" & celv).AutoFilter Field:=6, Criteria1:="UBVL", Operator:=xlFilterValues
    
    rcct = Sheets("CM_UL").UsedRange.Rows.Count
    ccct = Sheets("CM_UL").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    
    Range("D1:D" & rcct).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Cash").Activate
    
    Range("A2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ThisWorkbook.Worksheets("CM_UL").Activate
    Range("G1:G" & rcct).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Cash").Activate
    
    Range("B2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ThisWorkbook.Worksheets("CM_UL").Activate
    ActiveSheet.Range("A1:" & celv).AutoFilter Field:=6, Criteria1:="USVL", Operator:=xlFilterValues
    
    Range("G1:G" & rcct).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Cash").Activate
    Range("C2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
    Selection.Value = "TerminalID"
    Range("B1").Select
    Selection.Value = "UBVL"
    Range("C1").Select
    Selection.Value = "USVL"
    Rows("2:2").Select
    Selection.Delete
    Range("D1").Select
    Selection.Value = "ONLNBVL"
    Range("E1").Select
    Selection.Value = "ONLNSVL"
    Range("F1").Select
    Selection.Value = "Buy%"
    Range("G1").Select
    Selection.Value = "Sell%"
    Workbooks(obj1.Name).Close savechanges:=False
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    shtadd ("NSECash")
    spath = ThisWorkbook.Path & "\input\"
    ''''''''''''''''cm zip''''''''''''''
    cv = GetFullFileName(spath, "HAWK NSE CASH_")
    
    If cv <> Empty Or cv <> "" Then
       Set obj1 = Workbooks.Open(Filename:=spath & cv)
    End If
    shtnname = ActiveSheet.Name
    
    rcct = Sheets(shtnname).UsedRange.Rows.Count
    ccct = Sheets(shtnname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Selection.Copy
    ThisWorkbook.Worksheets("NSECash").Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    Workbooks(obj1.Name).Close savechanges:=False
    rcct = Sheets("NSECash").UsedRange.Rows.Count
    alk1 = returncolumnnumber("NSECash", "TradePrice", 1)
    colchar1 = ColLtr(alk1 + 1)
    colchar3 = ColLtr(alk1)
    colchar2 = ColLtr(alk1 - 1)
    Columns(colchar1).Select
    Selection.Insert Shift:=xlToRight
    Range(colchar1 & 1).Select
    Selection.Value = "value"
    Range(colchar1 & 2).Select
    ActiveCell.Formula = "=" & colchar2 & "2*" & colchar3 & "2"
    Selection.AutoFill Destination:=Range(colchar1 & "2:" & colchar1 & rcct)
    Range(colchar1 & "2:" & colchar1 & rcct).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ThisWorkbook.Worksheets("Cash").Activate
    rctt = Worksheets("Cash").Cells(1, 1).CurrentRegion.Rows.Count - 2
    ReDim criteria_val(rctt) As String
    Dim iRow As Integer
    'Erase criteria_val
    iRow = 2
    While Sheets("Cash").Cells(iRow, 1) <> ""
       criteria_val(iRow - 2) = Sheets("Cash").Cells(iRow, 1)
       iRow = iRow + 1
       Wend
    ThisWorkbook.Worksheets("NSECash").Activate
    
    alook1 = returncolumnnumber("NSECash", "TerminalID", 1)
    
    ActiveSheet.Range("A1:" & celv).AutoFilter Field:=alook1, Criteria1:=criteria_val, Operator:=xlFilterValues
  ''''''''''''''''''''''''''''''''''''''''''''''''''pivot'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rcct = Sheets("NSECash").UsedRange.Rows.Count
    ccct = Sheets("NSECash").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "NSECash!R1C1:R" & rcct & "C" & ccct, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="pivot!R1C1", TableName:="PivotTable4", DefaultVersion _
        :=xlPivotTableVersion14
    Sheets("pivot").Select
    
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("TerminalID")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable4").PivotFields("TerminalID").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable4").PivotFields("TerminalID").LayoutForm = _
        xlTabular
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("BuySell")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables("PivotTable4").PivotFields("value"), "Sum of value", xlSum
    
    ActiveWorkbook.ShowPivotTableFieldList = False
    Range("B1").Select
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("BuySell")
        .PivotItems("SELL").Visible = False
    End With
    ThisWorkbook.Worksheets("Cash").Activate
    rcct = Sheets("Cash").UsedRange.Rows.Count
    ccct = Sheets("Cash").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Cells(2, ccct - 3).Select
    Vllookcopypastefromotherbook 2, ccct - 3, ThisWorkbook.Name, ThisWorkbook.Name, "Cash", "pivot", "TerminalID", "Row Labels", "Sum of value", "ONLNBVL", 1, 0
    ThisWorkbook.Worksheets("pivot").Activate
    Range("B1").Select
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("BuySell")
        .PivotItems("SELL").Visible = True
        .PivotItems("BUY").Visible = False
    End With
    ThisWorkbook.Worksheets("Cash").Activate
      rcct = Sheets("Cash").UsedRange.Rows.Count
    ccct = Sheets("Cash").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Cells(2, ccct - 2).Select
    Vllookcopypastefromotherbook 2, ccct - 2, ThisWorkbook.Name, ThisWorkbook.Name, "Cash", "pivot", "TerminalID", "Row Labels", "Sum of value", "ONLNSVL", 1, 0
    ThisWorkbook.Worksheets("Cash").Activate
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "100000"
    Selection.Copy
    Range("B" & 2 & ":C" & rcct).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, _
    SkipBlanks:=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("B" & 2 & ":E" & rcct).Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Cells(2, ccct - 1).Select
   
    ActiveCell.FormulaR1C1 = "=(RC[-2]/RC[-4])*100"
    Selection.AutoFill Destination:=Range("F2:F" & rcct), Type:=xlFillDefault
    Range("F2:F" & rcct).Select
    Selection.Copy
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("F2:F" & rcct).Select
    Selection.NumberFormat = "0.00"
    alk1 = returncolumnnumber("Cash", "Buy%", 1)
    colchar1 = ColLtr(alk1)
    Cells(2, ccct).Select
   
    ActiveCell.FormulaR1C1 = "=(RC[-2]/RC[-4])*100"
    Selection.AutoFill Destination:=Range("G2:G" & rcct), Type:=xlFillDefault
    Range("G2:G" & rcct).Select
    Selection.Copy
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("G2:G" & rcct).Select
    Selection.NumberFormat = "0.00"
    alk2 = returncolumnnumber("Cash", "Sell%", 1)
    colchar2 = ColLtr(alk2)
    
    colchar3 = ColLtr(ccct + 1)
    Range(colchar3 & 1).Value = "Buy%+Sell%"
   Range(colchar3 & 2).Select
   ActiveCell.FormulaR1C1 = "=MAX(RC[-2],RC[-1])"
   Selection.AutoFill Destination:=Range(colchar3 & "2:" & colchar3 & rcct), Type:=xlFillDefault
    'Cells(2, alk1).Select
    'ActiveCell.Formula = "=IF(" & colchar & ">60,""0"")"
    ccct = Sheets("Cash").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
     Range(colchar3 & 2).Select
     
    ActiveWorkbook.Worksheets("Cash").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Cash").Sort.SortFields.Add Key:=Range("H2"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Cash").Sort
        .SetRange Range("A1:" & celv)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
     alk3 = returnrownumber("Cash", "TerminalID", 1)
     
    
    ActiveSheet.Range("A" & alk3 & ":" & celv).AutoFilter Field:=8, Criteria1:=">60", _
    Operator:=xlAnd
    Range(colchar3 & alk3 & ":" & colchar3 & rcct).Select
    If ActiveSheet.Range(colchar3 & alk3 & ":" & colchar3 & rcct).Cells(xlCellTypeVisible).Count = 0 Then
    Rows(alk3).Select
    Selection.AutoFilter
    Else
      Selection.SpecialCells(xlCellTypeVisible).Select
      Call font
      Range(colchar3 & alk3).Select
      Call font1
      
      Rows(alk3).Select
      Selection.AutoFilter
      End If
      Range("I1").Select
    Selection.Value = "Remark"
    Range("I2").Select
    Selection.Value = "Cash"
    Selection.AutoFill Destination:=Range("I2:I" & rcct), Type:=xlFillDefault
      '''''''''''''''''''''''''''''''''''''''''''''FUTURE & OPP.'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
      
      cv = GetFullFileName(spath, "FO_UL_")
    
    If cv <> Empty Or cv <> "" Then
       Set obj1 = Workbooks.Open(Filename:=spath & cv)
    End If
    shtnname = ActiveSheet.Name
    
    rcct = Sheets(shtnname).UsedRange.Rows.Count
    ccct = Sheets(shtnname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Selection.Copy
    ThisWorkbook.Worksheets("FO_UL").Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveSheet.Range("A1:" & celv).AutoFilter Field:=6, Criteria1:="UFBVL", Operator:=xlFilterValues
    
    rcct = Sheets("FO_UL").UsedRange.Rows.Count
    ccct = Sheets("FO_UL").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    
    Range("D1:D" & rcct).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Future").Activate
    
    Range("A2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ThisWorkbook.Worksheets("FO_UL").Activate
    Range("G1:G" & rcct).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Future").Activate
    
    Range("B2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ThisWorkbook.Worksheets("FO_UL").Activate
    ActiveSheet.Range("A1:" & celv).AutoFilter Field:=6, Criteria1:="UFSVL", Operator:=xlFilterValues
    
    Range("G1:G" & rcct).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Future").Activate
    Range("C2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
    Selection.Value = "TerminalID"
    Range("B1").Select
    Selection.Value = "UFBVL"
    Range("C1").Select
    Selection.Value = "UFSVL"
    Rows("2:2").Select
    Selection.Delete
    Range("D1").Select
    Selection.Value = "ONLNFBVL"
    Range("E1").Select
    Selection.Value = "ONLNFSVL"
    Range("F1").Select
    Selection.Value = "Buy%"
    Range("G1").Select
    Selection.Value = "Sell%"
   
 ''''''''''''''''''''''''''''
     ThisWorkbook.Worksheets("FO_UL").Activate
       ActiveSheet.Range("A1:" & celv).AutoFilter Field:=6, Criteria1:="UOBVL", Operator:=xlFilterValues
    
    rcct = Sheets("FO_UL").UsedRange.Rows.Count
    ccct = Sheets("FO_UL").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    
    Range("D1:D" & rcct).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Opption").Activate
    
    Range("A2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ThisWorkbook.Worksheets("FO_UL").Activate
    Range("G1:G" & rcct).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Opption").Activate
    
    Range("B2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ThisWorkbook.Worksheets("FO_UL").Activate
    ActiveSheet.Range("A1:" & celv).AutoFilter Field:=6, Criteria1:="UOSVL", Operator:=xlFilterValues
    
    Range("G1:G" & rcct).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Opption").Activate
    Range("C2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
    Selection.Value = "TerminalID"
    Range("B1").Select
    Selection.Value = "UOBVL"
    Range("C1").Select
    Selection.Value = "UOSVL"
    Rows("2:2").Select
    Selection.Delete
    Range("D1").Select
    Selection.Value = "ONLNOBVL"
    Range("E1").Select
    Selection.Value = "ONLNOSVL"
    Range("F1").Select
    Selection.Value = "Buy%"
    Range("G1").Select
    Selection.Value = "Sell%"
    Workbooks(obj1.Name).Close savechanges:=False
 ''''''''''''''''''''''''''''''''''''''''''''''''''''FNO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
 shtadd ("NSEFNO")
    spath = ThisWorkbook.Path & "\input\"
    ''''''''''''''''cm zip''''''''''''''
    cv = GetFullFileName(spath, "HAWK NSE F&O_")
    
    If cv <> Empty Or cv <> "" Then
       Set obj1 = Workbooks.Open(Filename:=spath & cv)
    End If
    shtnname = ActiveSheet.Name
    
    rcct = Sheets(shtnname).UsedRange.Rows.Count
    ccct = Sheets(shtnname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Selection.Copy
    ThisWorkbook.Worksheets("NSEFNO").Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    Workbooks(obj1.Name).Close savechanges:=False
    rcct = Sheets("NSEFNO").UsedRange.Rows.Count
    alk1 = returncolumnnumber("NSEFNO", "TradePrice", 1)
    colchar1 = ColLtr(alk1 + 1)
    colchar3 = ColLtr(alk1)
    colchar2 = ColLtr(alk1 - 1)
    Columns(colchar1).Select
    Selection.Insert Shift:=xlToRight
    Range(colchar1 & 1).Select
    Selection.Value = "value"
    Range(colchar1 & 2).Select
    ActiveCell.Formula = "=" & colchar2 & "2*" & colchar3 & "2"
    Selection.AutoFill Destination:=Range(colchar1 & "2:" & colchar1 & rcct)
    Range(colchar1 & "2:" & colchar1 & rcct).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    colchar4 = ColLtr(alk1 + 2)
    Columns(colchar4).Select
    Selection.Insert Shift:=xlToRight
    Range(colchar4 & 1).Select
    Selection.Value = "value2"
     alk2 = returncolumnnumber("NSEFNO", "StrikePrice", 1)
     colchar5 = ColLtr(alk2)
    Range(colchar4 & 2).Select
    ActiveCell.Formula = "=" & colchar1 & "2+" & colchar5 & "2"
     Selection.AutoFill Destination:=Range(colchar4 & "2:" & colchar4 & rcct)
    Range(colchar4 & "2:" & colchar4 & rcct).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    ThisWorkbook.Worksheets("Future").Activate
    rctt = Worksheets("Future").Cells(1, 1).CurrentRegion.Rows.Count - 2
    ReDim criteria_val(rctt) As String
    Dim iRow1 As Integer
    'Erase criteria_val
    iRow1 = 2
    While Sheets("Future").Cells(iRow1, 1) <> ""
       criteria_val(iRow1 - 2) = Sheets("Future").Cells(iRow1, 1)
       iRow1 = iRow1 + 1
       Wend
    ThisWorkbook.Worksheets("NSEFNO").Activate
    
    alook1 = returncolumnnumber("NSEFNO", "TerminalID", 1)
    
    ActiveSheet.Range("A1:" & celv).AutoFilter Field:=alook1, Criteria1:=criteria_val, Operator:=xlFilterValues
 
  rcct = Sheets("NSEFNO").UsedRange.Rows.Count
    ccct = Sheets("NSEFNO").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "NSEFNO!R1C1:R" & rcct & "C" & ccct, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="pivot2!R1C1", TableName:="PivotTable4", DefaultVersion _
        :=xlPivotTableVersion14
    Sheets("pivot2").Select
  Cells(1, 1).Select
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("TerminalID")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("BuySell")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable4").PivotFields("TerminalID").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable4").PivotFields("TerminalID").LayoutForm = _
        xlTabular
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Series")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables("PivotTable4").PivotFields("BuySell").LayoutForm = _
        xlTabular
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("value"), "Sum of value", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("value2"), "Sum of value2", xlSum
    Range("B2").Select
    ActiveSheet.PivotTables("PivotTable4").PivotFields("BuySell").PivotItems("BUY") _
        .ShowDetail = False
    Range("B3").Select
    ActiveSheet.PivotTables("PivotTable4").PivotFields("BuySell").PivotItems("SELL" _
        ).ShowDetail = False
   With ActiveSheet.PivotTables("PivotTable4").PivotFields("Series")
        .PivotItems("CE").Visible = False
        .PivotItems("PE").Visible = False
    End With
 Range("B1").Select
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("BuySell")
        .PivotItems("BUY").Visible = True
        .PivotItems("SELL").Visible = False
    End With
  ThisWorkbook.Worksheets("Future").Activate
    rcct = Sheets("Future").UsedRange.Rows.Count
    ccct = Sheets("Future").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Cells(2, ccct - 3).Select
    Vllookcopypastefromotherbook 2, ccct - 3, ThisWorkbook.Name, ThisWorkbook.Name, "Future", "pivot2", "TerminalID", "Row Labels", "Sum of value", "ONLNFBVL", 1, 0
    ThisWorkbook.Worksheets("pivot2").Activate
    Range("B1").Select
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("BuySell")
        .PivotItems("SELL").Visible = True
        .PivotItems("BUY").Visible = False
    End With
  
   ThisWorkbook.Worksheets("Future").Activate
    rcct = Sheets("Future").UsedRange.Rows.Count
    ccct = Sheets("Future").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Cells(2, ccct - 2).Select
    Vllookcopypastefromotherbook 2, ccct - 2, ThisWorkbook.Name, ThisWorkbook.Name, "Future", "pivot2", "TerminalID", "Row Labels", "Sum of value", "ONLNFSVL", 1, 0
 ThisWorkbook.Worksheets("Future").Activate
  Range("J2").Select
    ActiveCell.FormulaR1C1 = "100000"
    Selection.Copy
    Range("B" & 2 & ":C" & rcct).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, _
    SkipBlanks:=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    Selection.ClearContents
      Cells(2, ccct - 1).Select
   
    ActiveCell.FormulaR1C1 = "=(RC[-2]/RC[-4])*100"
    Selection.AutoFill Destination:=Range("F2:F" & rcct), Type:=xlFillDefault
    Range("F2:F" & rcct).Select
    Selection.Copy
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("F2:F" & rcct).Select
    Selection.NumberFormat = "0.00"
    alk1 = returncolumnnumber("Future", "Buy%", 1)
    colchar1 = ColLtr(alk1)
    Cells(2, ccct).Select
   
    ActiveCell.FormulaR1C1 = "=(RC[-2]/RC[-4])*100"
    Selection.AutoFill Destination:=Range("G2:G" & rcct), Type:=xlFillDefault
    Range("G2:G" & rcct).Select
    Selection.Copy
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("G2:G" & rcct).Select
    Selection.NumberFormat = "0.00"
    alk2 = returncolumnnumber("Future", "Sell%", 1)
    colchar2 = ColLtr(alk2)
    
    colchar3 = ColLtr(ccct + 1)
    Range(colchar3 & 1).Value = "Buy%+Sell%"
   Range(colchar3 & 2).Select
   ActiveCell.FormulaR1C1 = "=MAX(RC[-2],RC[-1])"
   Selection.AutoFill Destination:=Range(colchar3 & "2:" & colchar3 & rcct), Type:=xlFillDefault
    'Cells(2, alk1).Select
    'ActiveCell.Formula = "=IF(" & colchar & ">60,""0"")"
    ccct = Sheets("Future").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
     Range(colchar3 & 2).Select
     
    ActiveWorkbook.Worksheets("Future").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Future").Sort.SortFields.Add Key:=Range("H2"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Future").Sort
        .SetRange Range("A1:" & celv)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
     alk3 = returnrownumber("Future", "TerminalID", 1)
    ActiveSheet.Range("A" & alk3 & ":" & celv).AutoFilter Field:=8, Criteria1:=">60", _
    Operator:=xlAnd
    Range(colchar3 & alk3 & ":" & colchar3 & rcct).Select
    If ActiveSheet.Range(colchar3 & alk3 & ":" & colchar3 & rcct).Cells(xlCellTypeVisible).Count = 0 Then
    Rows(alk3).Select
    Selection.AutoFilter
    Else
      Selection.SpecialCells(xlCellTypeVisible).Select
      Call font
      Range(colchar3 & alk3).Select
      Call font1
      
      Rows(alk3).Select
      Selection.AutoFilter
      End If
          rcct = Sheets("Future").UsedRange.Rows.Count
    Range("I1").Select
    Selection.Value = "Remark"
    Range("I2").Select
    Selection.Value = "Future"
    Selection.AutoFill Destination:=Range("I2:I" & rcct), Type:=xlFillDefault

  
  ThisWorkbook.Worksheets("pivot2").Activate
 
    Range("C1").Select
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Series")
        .PivotItems("CE").Visible = True
        .PivotItems("PE").Visible = True
        .PivotItems("(blank)").Visible = False
     End With
   Range("B1").Select
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("BuySell")
        .PivotItems("BUY").Visible = True
        .PivotItems("SELL").Visible = False
    End With
 
 ThisWorkbook.Worksheets("Opption").Activate
    rcct = Sheets("Opption").UsedRange.Rows.Count
    ccct = Sheets("Opption").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Cells(2, ccct - 3).Select
    Vllookcopypastefromotherbook 2, ccct - 3, ThisWorkbook.Name, ThisWorkbook.Name, "Opption", "pivot2", "TerminalID", "Row Labels", "Sum of value", "ONLNOBVL", 1, 0
 
 ThisWorkbook.Worksheets("pivot2").Activate
 Range("B1").Select
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("BuySell")
        .PivotItems("SELL").Visible = True
        .PivotItems("BUY").Visible = False
    End With
 
 ThisWorkbook.Worksheets("Opption").Activate
    rcct = Sheets("Opption").UsedRange.Rows.Count
    ccct = Sheets("Opption").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Cells(2, ccct - 2).Select
    Vllookcopypastefromotherbook 2, ccct - 2, ThisWorkbook.Name, ThisWorkbook.Name, "Opption", "pivot2", "TerminalID", "Row Labels", "Sum of value2", "ONLNOSVL", 1, 0
ThisWorkbook.Worksheets("Opption").Activate
 Range("J2").Select
    ActiveCell.FormulaR1C1 = "100000"
    Selection.Copy
    Range("B" & 2 & ":C" & rcct).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, _
    SkipBlanks:=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    Selection.ClearContents
 
    Cells(2, ccct - 1).Select
   
    ActiveCell.FormulaR1C1 = "=(RC[-2]/RC[-4])*100"
    Selection.AutoFill Destination:=Range("F2:F" & rcct), Type:=xlFillDefault
    Range("F2:F" & rcct).Select
    Selection.Copy
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("F2:F" & rcct).Select
    Selection.NumberFormat = "0.00"
    alk1 = returncolumnnumber("Opption", "Buy%", 1)
    colchar1 = ColLtr(alk1)
    Cells(2, ccct).Select
   
    ActiveCell.FormulaR1C1 = "=(RC[-2]/RC[-4])*100"
    Selection.AutoFill Destination:=Range("G2:G" & rcct), Type:=xlFillDefault
    Range("G2:G" & rcct).Select
    Selection.Copy
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("G2:G" & rcct).Select
    Selection.NumberFormat = "0.00"
    alk2 = returncolumnnumber("Opption", "Sell%", 1)
    colchar2 = ColLtr(alk2)
    
    colchar3 = ColLtr(ccct + 1)
    Range(colchar3 & 1).Value = "Buy%+Sell%"
   Range(colchar3 & 2).Select
   ActiveCell.FormulaR1C1 = "=MAX(RC[-2],RC[-1])"
   Selection.AutoFill Destination:=Range(colchar3 & "2:" & colchar3 & rcct), Type:=xlFillDefault
    'Cells(2, alk1).Select
    'ActiveCell.Formula = "=IF(" & colchar & ">60,""0"")"
    ccct = Sheets("Opption").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
     Range(colchar3 & 2).Select
     
    ActiveWorkbook.Worksheets("Opption").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Opption").Sort.SortFields.Add Key:=Range("H2"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Opption").Sort
        .SetRange Range("A1:" & celv)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
 
     alk3 = returnrownumber("Opption", "TerminalID", 1)
    ActiveSheet.Range("A" & alk3 & ":" & celv).AutoFilter Field:=8, Criteria1:=">60", _
    Operator:=xlAnd
    Range(colchar3 & alk3 & ":" & colchar3 & rcct).Select
    If ActiveSheet.Range(colchar3 & alk3 & ":" & colchar3 & rcct).Cells(xlCellTypeVisible).Count = 0 Then
    Rows(alk3).Select
    Selection.AutoFilter
    Else
      Selection.SpecialCells(xlCellTypeVisible).Select
      Call font
      Range(colchar3 & alk3).Select
      Call font1
      
      Rows(alk3).Select
      Selection.AutoFilter
      End If
    rcct = Sheets("Opption").UsedRange.Rows.Count
    Range("I1").Select
    Selection.Value = "Remark"
    Range("I2").Select
    Selection.Value = "Opption"
    Selection.AutoFill Destination:=Range("I2:I" & rcct), Type:=xlFillDefault
    
'''''''''''''''''''''''''''Summery'''''''''''''''''''''''''''''''''''''''''''
  ThisWorkbook.Worksheets("Summery").Activate
   Range("A1").Select
    Selection.Value = "TerminalID"
    Range("B1").Select
    Selection.Value = "BuyVaule"
    Range("C1").Select
    Selection.Value = "SellValue"
    Rows("2:2").Select
    Selection.Delete
    Range("D1").Select
    Selection.Value = "OnlineBuyValue"
    Range("E1").Select
    Selection.Value = "OnllineSellVaule"
    Range("F1").Select
    Selection.Value = "Buy%"
    Range("G1").Select
    Selection.Value = "Sell%"
    Range("H1").Select
    Selection.Value = "Buy%+Sell%"
    Range("I1").Select
    Selection.Value = "Remark"
     Range("A1").Select
     Call wrap
    ThisWorkbook.Worksheets("Opption").Activate
    ActiveSheet.Range("A" & alk3 & ":" & celv).AutoFilter Field:=8, Criteria1:=RGB(255, 0 _
    , 0), Operator:=xlFilterFontColor
     Range("A" & alk3).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Summery").Activate
    Range("A2").Select
    ActiveSheet.Paste
    Rows("2:2").Select
    Selection.Delete
    ThisWorkbook.Worksheets("Opption").Activate
    ActiveSheet.Range("A" & alk3 & ":" & celv).AutoFilter
    ThisWorkbook.Worksheets("Summery").Activate
      rcct = Sheets("Summery").UsedRange.Rows.Count
    ccct = Sheets("Summery").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    
    ThisWorkbook.Worksheets("Future").Activate
      rcct = Sheets("Future").UsedRange.Rows.Count
    ccct = Sheets("Future").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    alk3 = returnrownumber("Future", "TerminalID", 1)
    ActiveSheet.Range("A" & alk3 & ":" & celv).AutoFilter Field:=8, Criteria1:=RGB(255, 0 _
    , 0), Operator:=xlFilterFontColor
    
    Range("A" & alk3).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Summery").Activate
     rcct = Sheets("Summery").UsedRange.Rows.Count
    Range("A" & rcct + 1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Rows(rcct + 1).Select
    Selection.Delete
    ThisWorkbook.Worksheets("Future").Activate
    ActiveSheet.Range("A" & alk3 & ":" & celv).AutoFilter
    ''''''''''''cashsummery''''''''''''''
    ThisWorkbook.Worksheets("Summery").Activate
      rcct = Sheets("Summery").UsedRange.Rows.Count
    ccct = Sheets("Summery").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    
    ThisWorkbook.Worksheets("Cash").Activate
      rcct = Sheets("Cash").UsedRange.Rows.Count
    ccct = Sheets("Cash").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    alk3 = returnrownumber("Cash", "TerminalID", 1)
    ActiveSheet.Range("A" & alk3 & ":" & celv).AutoFilter Field:=8, Criteria1:=RGB(255, 0 _
    , 0), Operator:=xlFilterFontColor
    
    Range("A" & alk3).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Summery").Activate
     rcct = Sheets("Summery").UsedRange.Rows.Count
    Range("A" & rcct + 1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Rows(rcct + 1).Select
    Selection.Delete
    ThisWorkbook.Worksheets("Cash").Activate
    ActiveSheet.Range("A" & alk3 & ":" & celv).AutoFilter
    
      
'''''''''''''''''''''''''''''''''''''''''''SAVE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      Workbooks.Add
pn = ActiveWorkbook.Name
Workbooks(pn).Activate

      ThisWorkbook.Worksheets("Cash").Activate
Sheets(Array("Cash", "Opption", "Future", "Summery")).Select

Sheets(Array("Cash", "Opption", "Future", "Summery")).Copy Before:=Workbooks(pn).Sheets(1)
 For Each Sheet In ActiveWorkbook.Sheets
        If Sheet.Name = "Sheet1" Or Sheet.Name = "Sheet2" Or Sheet.Name = "Sheet3" Then
           Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        Else
        
        End If
Next Sheet
Worksheets("Summery").Activate
Cells(1, 1).Select
      
   Set fso = CreateObject("Scripting.FileSystemObject")
    spath = ThisWorkbook.Path & "\Report"
             If fso.FolderExists(ThisWorkbook.Path & "\report") Then
                 New_file_name = ThisWorkbook.Path & "\report\" & combo1 & " Report_" & Format(Now, "ddmmyyyy_h_m_s") & ".xlsx"
                 Workbooks(pn).SaveAs Filename:=New_file_name
                 fname = combo1 & " Report_" & Format(Now, "ddmmyyyy_h_m_s")
             Else
                 fso.createfolder (ThisWorkbook.Path & "\report")
                 New_file_name = ThisWorkbook.Path & "\report\" & combo1 & " Report_" & Format(Now, "ddmmyyyy_h_m_s") & ".xlsx"
                 Workbooks(pn).SaveAs Filename:=New_file_name
                 fname = combo1 & " Report_" & Format(Now, "ddmmyyyy_h_m_s")
            End If
   
    
    
    End Sub
    Sub shtadd(shtnname)
       exists = False
       For i = 1 To ThisWorkbook.Worksheets.Count
               If ThisWorkbook.Worksheets(i).Name = shtnname Then
                   exists = True
               End If
       Next i
       If Not exists Then
           ThisWorkbook.Worksheets.Add(After:=Worksheets("HOME")).Name = shtnname
       Else
           cler (shtnname)
       End If
    End Sub
    Function cler(SheetName)
    ThisWorkbook.Worksheets(SheetName).Activate
       Cells.Select
       Cells.Clear
       Cells.Select
       Cells.Delete
       Cells(1, 1).Select
    End Function
    Function returnrownumber(parasheet, rowname, icol)
    Sheets(parasheet).Activate
    checkstatus = ""
    R = Sheets(parasheet).UsedRange.Rows.Count
    For i = 1 To R
      findval = Sheets(parasheet).Cells(i, icol).Value
    If Trim(UCase(rowname)) = Trim(UCase(findval)) Then
      checkstatus = "found"
       Exit For
    End If
    Next
    If checkstatus = " " Then
    ''MsgBox "no column found"
    End If
    returnrownumber = i
    End Function
    Function returncolumnnumber(parasheet, columnname, iRow)
    
    Sheets(parasheet).Activate
    tolst = Sheets(parasheet).Cells(iRow, 1).CurrentRegion.Columns.Count
    checkstatus = ""
    For i = 1 To tolst
      findval = Sheets(parasheet).Cells(iRow, i).Value
    If Trim(UCase(columnname)) = Trim(UCase(findval)) Then
       checkstatus = "found"
       Exit For
    End If
    Next
    If checkstatus = " " Then
    ' MsgBox "no column found"
    End If
    returncolumnnumber = i
    End Function
    Function ColLtr(icol)
    If icol > 0 And icol <= Columns.Count Then
    ColLtr = Evaluate("substitute(address(1, " & icol & ", 4), ""1"", """")")
    End If
    End Function
    Function UnzipFile(ByVal sZipFile As String, ByVal sDestFolder As String, vbn)
    
    Dim objApp As Object
    Dim objArchive As Object
    Dim objDest As Object
    Dim vDestFolder As Variant
    Dim vZipFile As Variant
    
    Set objApp = CreateObject("Shell.Application")
    
    vZipFile = sZipFile
    vDestFolder = sDestFolder
    
    If Dir$(sDestFolder, vbDirectory) = "" Then MkDir sDestFolder
    
    
    For Each oFile In objApp.Namespace(vZipFile).Items
    
    vbn = oFile.Name
    
    Next
    objApp.Namespace(vDestFolder).CopyHere objApp.Namespace(vZipFile).Items
    objApp.Namespace(vZipFile).Items
    
    
    
    End Function
    
    Function GetFullFileName(strfilepath, strFileNamePartial)
    
    Dim objFS As Variant
    Dim objFolder As Variant
    Dim objFile As Variant
    Dim intLengthOfPartialName As Integer
    Dim strfilenamefull As String
    
    Set objFS = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFS.getfolder(strfilepath)
    
    'work out how long the partial file name is
    intLengthOfPartialName = Len(strFileNamePartial)
    
    For Each objFile In objFolder.Files
    
    'Test to see if the file matches the partial file name
    If Left(objFile.Name, intLengthOfPartialName) = strFileNamePartial Then
    
    'get the full file name
    strfilenamefull = objFile.Name
    Exit For
    
    Else
    
    End If
    
    Next objFile
    
    'Return the full file name as the function's value
    GetFullFileName = strfilenamefull
    
    End Function
    Function Vllookcopypastefromotherbook(ivrow, ivcol, lookupbook, lookinbook, lookupsheet, lookinsheet, lookupcolnm, lookincolnm, getcol, putcol, iRow, err)
    inseclocde = returncolumnnumber(lookupsheet, lookupcolnm, 1)
    colchar = ColLtr(inseclocde) & ivrow
    insemo = returncolumnnumber(lookupsheet, putcol, iRow)
     plstrow = Sheets(lookupsheet).Cells(iRow, inseclocde).CurrentRegion.Rows.Count
     pasterng1 = ColLtr(ivcol) & ivrow
    pasterng2 = ColLtr(insemo) & plstrow
    
    ThisWorkbook.Worksheets(lookinsheet).Activate
    idetclocde = returncolumnnumber(lookinsheet, lookincolnm, 1)
    idetmo = returncolumnnumber(lookinsheet, getcol, 1)
    ilastrow = Sheets(lookinsheet).UsedRange.Rows.Count
    Rng1 = Sheets(lookinsheet).Cells(1, idetclocde).Address
    
    rng2 = Sheets(lookinsheet).Cells(ilastrow, idetmo).Address
    search_table_range = Sheets(lookinsheet).Range(Rng1 & ":" & rng2).Address
    colmn = idetmo - idetclocde + 1
    'Sheets(lookupsheet).Cells(2, insemo).Value = "=IFERROR(VLOOKUP(" & colchar & ",'" & lookinsheet & "'!" & search_table_range & "," & colmn & ",0)," & err & ")"
    Workbooks(lookinbook).Activate
'    If err = 0 Then
'        Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & ",'[" & lookinbook & lookinsheet & "]" & "'!" & search_table_range & " ," & colmn & ",0)," & err & ")"
'    Else
'
'      Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & ",'[" & lookinbook & "]" & lookinsheet & "'!" & search_table_range & " ," & colmn & ",0)," & Chr(34) & err & Chr(34) & ")"
'    End If
 Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & "," & lookinsheet & "!" & search_table_range & " ," & colmn & ",0)," & Chr(34) & err & Chr(34) & ")"
    Sheets(lookupsheet).Cells(ivrow, ivcol).Copy
    Sheets(lookupsheet).Activate
    Sheets(lookupsheet).Range(pasterng1 & ":" & pasterng2).Select
    ActiveSheet.Paste
    Sheets(lookupsheet).Range(pasterng1 & ":" & pasterng2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets(lookupsheet).Range(pasterng1).Select

End Function
Sub font()
 With Selection.font
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
        
    End With
 End Sub
 
 Sub font1()

 
        With Selection.font
        .Color = RGB(0, 0, 0)
        .TintAndShade = 0
   End With
 End Sub
 Sub wrap()
  Range(Selection, Selection.End(xlToRight)).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
 End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''fnoban''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public to1 As Variant
Public celv As Variant


Sub ban()
 ThisWorkbook.Worksheets("HOME").Activate
 shtadd ("Banscript")
 shtadd ("Report")
 shtadd ("Pivot")
 spath = ThisWorkbook.Path & "\Input\"
  cv = GetFullFileName(spath, "Ban script file")
 If cv <> "" Then
    Set obj1 = Workbooks.Open(Filename:=spath & "\" & cv)
    Sheets("Sheet1").Activate
    shtname = ActiveSheet.Name
    rc = ActiveWorkbook.Worksheets(shtname).UsedRange.Rows.Count
    CC = ActiveWorkbook.Worksheets(shtname).UsedRange.Columns.Count
    Cells(rc, CC).Select
    cv = Selection.Address(False, False)
    Range("A1:" & cv).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Banscript").Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    End If
    Workbooks(obj1.Name).Close savechanges:=False
    ThisWorkbook.Worksheets("Banscript").Activate
    rcct = Sheets("Banscript").UsedRange.Rows.Count
    ccct = Sheets("Banscript").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    alk = returncolumnnumber("Banscript", "BOD", 1)
    alk2 = returncolumnnumber("Banscript", "NetQty", 1)
    colchar3 = ColLtr(alk)
    colchar4 = ColLtr(alk2)
    alk1 = ccct + 1
    colchar1 = ColLtr(alk1)
     colchar2 = ColLtr(alk1 + 1)
    Range(colchar1 & 1).Select
    Selection.Value = "Absolute BOD Qty"
    Range(colchar2 & 1).Select
    Selection.Value = "Absolute NET Qty"
    Range(colchar1 & 2).Select
   ActiveCell.Formula = "=ABS(" & colchar3 & 2 & ")"
 Selection.AutoFill Destination:=Range(colchar1 & "2:" & colchar1 & rcct)
 Range(colchar1 & "2:" & colchar1 & rcct).Copy
Range(colchar1 & 2).Select
 Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 Range(colchar2 & 2).Select
   ActiveCell.Formula = "=ABS(" & colchar4 & 2 & ")"
 Selection.AutoFill Destination:=Range(colchar2 & "2:" & colchar2 & rcct)
 Range(colchar2 & "2:" & colchar2 & rcct).Copy
Range(colchar2 & 2).Select
 Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    rcct = Sheets("Banscript").UsedRange.Rows.Count
    ccct = Sheets("Banscript").UsedRange.Columns.Count
 Sheets("Pivot").Select
 Range("A1").Select
ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Banscript!R1C1:R" & rcct & "C" & ccct, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="Pivot!R1C1", TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion14
    Sheets("Pivot").Select
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("NSECode")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Absolute BOD Qty"), "Sum of Absolute BOD Qty", _
        xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Absolute NET Qty"), "Sum of Absolute NET Qty", _
       xlSum
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("ClientCode")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("NSECode").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("NSECode").LayoutForm = _
        xlTabular
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
 
 Sheets("Pivot").Select
    rc = Sheets("Pivot").UsedRange.Rows.Count
    CC = Sheets("Pivot").UsedRange.Columns.Count
     Cells(rc, CC).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Selection.Copy
    Sheets("Report").Select
      Range("A1").Select
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
 Range("A1").Select
 Selection.Value = "NSECode "
 Range("C1").Select
 Selection.Value = " Absolute BOD Qty "
 Range("D1").Select
 Selection.Value = " Absolute NET Qty"
 Range("E1").Select
 Selection.Value = " Violation Qty"
 Range("E2").Select
ActiveCell.Formula = "=IF(RC[-1]>RC[-2],RC[-1]-RC[-2],0)"
 rc = Sheets("Report").UsedRange.Rows.Count
 Selection.AutoFill Destination:=Range("E2:E" & rc)
 Range("E2:E" & rc).Copy
 Range("E2").Select
 Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    CC = Sheets("Report").UsedRange.Columns.Count
     Cells(rc, CC).Select
    celv = Selection.Address(False, False)

   Range("A1:E1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:" & celv).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1:E1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = True
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
'   Sheets("Report").Select
'    Range("E1").Select
'    Selection.AutoFilter
'    ActiveSheet.Range("A1:" & celv).AutoFilter Field:=5, Criteria1:=RGB(250, _
'        191, 143), Operator:=xlFilterCellColor
'
'  Range("A1:" & celv).SpecialCells (xlCellTypeVisible)
'    Selection.Copy
   ThisWorkbook.Worksheets("Home").Activate
           Range("I1").Select
           to1 = Selection.Value
    If to1 <> "" Then
                    Call SendEmailUsingOutlook(New_file_name)
    Else
    End If
  '''''''''''''''''''''''save_report''''''''''''''''''''''''''''''''''''''''
  
  Workbooks.Add
pn = ActiveWorkbook.Name
Workbooks(pn).Activate

ThisWorkbook.Worksheets("Report").Activate
Sheets(Array("Banscript", "Report")).Select
Sheets(Array("Banscript", "Report")).Copy Before:=Workbooks(pn).Sheets(1)
 For Each Sheet In ActiveWorkbook.Sheets
        If Sheet.Name = "Sheet1" Or Sheet.Name = "Sheet2" Or Sheet.Name = "Sheet3" Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        Else
        
        End If
Next Sheet
Worksheets("Report").Activate
Cells(1, 1).Select
 
Set fso = CreateObject("Scripting.FileSystemobject")
 spath = ThisWorkbook.Path & "\report"
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\report\" & Year(Now)
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\report\" & Year(Now) & "\" & MonthName(Month(Now))
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\report\" & Year(Now) & "\" & MonthName(Month(Now)) & "\" & Day(Now)
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 savereport = spath
 
       Set fso = CreateObject("Scripting.filesystemobject")
             If fso.FolderExists(savereport) Then
                 New_file_name = savereport & "\" & " Report_" & Format(Now, "ddmmyyyy_h_m_s") & ".xlsx"
                 pn = ActiveWorkbook.Name
                 Workbooks(pn).SaveAs Filename:=New_file_name
                 fname = "Report_" & Format(Now, "ddmmyyyy_h_m_s")
             Else
                 fso.createfolder (savereport)
                 
                 New_file_name = savereport & "\" & "Report_" & Format(Now, "ddmmyyyy_h_m_s") & ".xlsx"
                 pn = ActiveWorkbook.Name
                 Workbooks(pn).SaveAs Filename:=New_file_name
                 fname = "Report_" & Format(Now, "ddmmyyyy_h_m_s")
            End If
    
    
    ActiveWorkbook.Close savechanges = False

ThisWorkbook.Worksheets("Home").Activate
  
  
  
    
    
    
End Sub

    Function FolderExists(strFolderPath)
    On Error Resume Next
    FolderExists = (GetAttr(strFolderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
    
    
    End Function
    
    Function GetFullFileName(strfilepath, strFileNamePartial)
    
    Dim objFS As Variant
    Dim objFolder As Variant
    Dim objFile As Variant
    Dim intLengthOfPartialName As Integer
    Dim strfilenamefull As String
    
    Set objFS = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFS.getfolder(strfilepath)
    
    'work out how long the partial file name is
    intLengthOfPartialName = Len(strFileNamePartial)
    
    For Each objFile In objFolder.Files
    
    'Test to see if the file matches the partial file name
    If Left(objFile.Name, intLengthOfPartialName) = strFileNamePartial Then
    
    'get the full file name
    strfilenamefull = objFile.Name
    Exit For
    
    Else
    
    End If
    
    Next objFile
    
    'Return the full file name as the function's value
    GetFullFileName = strfilenamefull
    
    End Function
    
    Function ColLtr(icol)
    If icol > 0 And icol <= Columns.Count Then
    ColLtr = Evaluate("substitute(address(1, " & icol & ", 4), ""1"", """")")
    End If
    
    End Function
    
    
    Sub shtadd(shtnname)
    exists = False
    For I = 1 To ActiveWorkbook.Worksheets.Count
    If ActiveWorkbook.Worksheets(I).Name = shtnname Then
    exists = True
    End If
    Next I
    If Not exists Then
    ActiveWorkbook.Worksheets.Add(After:=Worksheets(1)).Name = shtnname
    Else
    cler (shtnname)
    End If
    End Sub
    Function cler(SheetName)
    ActiveWorkbook.Worksheets(SheetName).Activate
    Cells.Select
    Cells.Clear
    Cells.Select
    Cells.Delete
    Cells(1, 1).Select
    End Function
    
    Function returncolumnnumber(parasheet, columnname, iRow)
    Sheets(parasheet).Activate
    tolst = Sheets(parasheet).UsedRange.Rows.Columns.Count + 5
    checkstatus = ""
    For I = 1 To tolst
    findval = Sheets(parasheet).Cells(iRow, I).Value
    If Trim(UCase(columnname)) = Trim(UCase(findval)) Then
    checkstatus = "found"
    Exit For
    End If
    Next
    If checkstatus = "" Then
    returncolumnnumber = 0
    Exit Function
    End If
    returncolumnnumber = I
    End Function
    Function Vllookcopypastefromotherbook(ivrow, ivcol, lookupbook, lookinbook, lookupsheet, lookinsheet, lookupcolnm, lookincolnm, getcol, putcol, iRow, err)
    inseclocde = returncolumnnumber(lookupsheet, lookupcolnm, 1)
    colchar = ColLtr(inseclocde) & ivrow
    insemo = returncolumnnumber(lookupsheet, putcol, iRow)
    plstrow = Sheets(lookupsheet).Cells(iRow, inseclocde).CurrentRegion.Rows.Count
    pasterng1 = ColLtr(ivcol) & ivrow
    pasterng2 = ColLtr(insemo) & plstrow
    
    ThisWorkbook.Worksheets(lookinsheet).Activate
    idetclocde = returncolumnnumber(lookinsheet, lookincolnm, 1)
    idetmo = returncolumnnumber(lookinsheet, getcol, 1)
    ilastrow = Sheets(lookinsheet).UsedRange.Rows.Count
    Rng1 = Sheets(lookinsheet).Cells(1, idetclocde).Address
    
    rng2 = Sheets(lookinsheet).Cells(ilastrow, idetmo).Address
    search_table_range = Sheets(lookinsheet).Range(Rng1 & ":" & rng2).Address
    colmn = idetmo - idetclocde + 1
    'Sheets(lookupsheet).Cells(2, insemo).Value = "=IFERROR(VLOOKUP(" & colchar & ",'" & lookinsheet & "'!" & search_table_range & "," & colmn & ",0)," & err & ")"
    Workbooks(lookinbook).Activate
    '    If err = 0 Then
    '        Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & ",'[" & lookinbook & lookinsheet & "]" & "'!" & search_table_range & " ," & colmn & ",0)," & err & ")"
    '    Else
    '
    '      Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & ",'[" & lookinbook & "]" & lookinsheet & "'!" & search_table_range & " ," & colmn & ",0)," & Chr(34) & err & Chr(34) & ")"
    '    End If
    Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & "," & lookinsheet & "!" & search_table_range & " ," & colmn & ",0)," & Chr(34) & err & Chr(34) & ")"
    Sheets(lookupsheet).Cells(ivrow, ivcol).Copy
    Sheets(lookupsheet).Activate
    Sheets(lookupsheet).Range(pasterng1 & ":" & pasterng2).Select
    ActiveSheet.Paste
    Sheets(lookupsheet).Range(pasterng1 & ":" & pasterng2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets(lookupsheet).Range(pasterng1).Select
    
    End Function
Function SendEmailUsingOutlook(New_file_name)

Dim OlApp As New Outlook.Application
Dim myNameSp As Outlook.Namespace
Dim myInbox As Outlook.MAPIFolder
Dim myExplorer As Outlook.Explorer
Dim NewMail As Outlook.MailItem
Dim OutOpen As Boolean
  
  Dim rng As Range
  Dim rng2 As Range
  Dim rng3 As Range
  Set rng = Nothing
  Set rng2 = Nothing
  Set rng3 = Nothing
  Sheets("Report").Select
    Range("E1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A1:" & celv).AutoFilter Field:=5, Criteria1:=RGB(250, _
        191, 143), Operator:=xlFilterCellColor
  
'  Range("A1:" & celv).SpecialCells (xlCellTypeVisible)
   
            Set rng = Sheets("Report").Range("A1:" & celv).SpecialCells(xlCellTypeVisible)
     cnt = rng.Count
     
If cnt > 5 Then
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    ' Check to see if there's an explorer window open
    ' If not then open up a new one
    OutOpen = True
    Set myExplorer = OlApp.ActiveExplorer
    If TypeName(myExplorer) = "Nothing" Then
        OutOpen = False
        Set myNameSp = OlApp.GetNamespace("MAPI")
        Set myInbox = myNameSp.GetDefaultFolder(olFolderInbox)
        Set myExplorer = myInbox.GetExplorer
    End If
        SigString = Environ("appdata") & "\Microsoft\Signatures\Fenil patel.htm"
   
    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If
    Set NewMail = OlApp.CreateItem(olMailItem)
   
            With NewMail
                '.Display ' You don't have to show the e-mail to send it
                .Display
                .Subject = "FNO Banscrip  Violation on " & Now()
                .To = to1
                '.CC = cc1
                .HTMLBody = RangetoHTML(rng) & "<br>" & Signature
                
                '.Attachments.Add (New_file_name)
            End With
   
    
    With Application
    .EnableEvents = True
    .ScreenUpdating = True
  End With
    Application.DisplayAlerts = True
     NewMail.Send
    If Not OutOpen Then OlApp.Quit
    'Release memory.
    Set OlApp = Nothing
    Set myNameSp = Nothing
    Set myInbox = Nothing
    Set myExplorer = Nothing
    Set NewMail = Nothing
 Else
 End If
    
    
End Function
Function GetBoiler(ByVal SigString As String) As String
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(SigString).OpenAsTextStream(1, -2)
    GetBoiler = ts.ReadAll
    ts.Close
End Function
Function RangetoHTML(rng As Range)
' By Ron de Bruin.
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
    'Copy the range and create a new workbook to past the data in
    
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(5).PasteSpecial Paste:=8
        .Cells(5).PasteSpecial xlPasteValues, , False, False
        .Cells(5).PasteSpecial xlPasteFormats, , False, False
        .Cells(5).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        '.DrawingObjects.Delete
        On Error GoTo 0
    End With
    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")
    'Close TempWB
    TempWB.Close savechanges:=False
    'Delete the htm file we used in this function
    Kill TempFile
   
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function
'''''''''''''''''''''dhvalrbi''''''''''''''''''''''''''
Public spath As String
Sub Rbi()
''''''''''''''''''''''''''''''''''''''''''''''PATH'''''''''''''''''''''''''''''''''''''''''''''''''''''''
ThisWorkbook.Worksheets("HOME").Activate
    path1 = Range("G1").Value
    path2 = Range("G2").Value
   spath = Range("G3").Value
    strtdate = Range("E1").Value
'''''''''''''''''''''sheetADD''''''''''''''''''''''''''''''''''''''''''''
shtadd ("MTF_report")
shtadd ("Group1")
shtadd ("Aproved stock")
shtadd ("ClientMFHoldingReport")
shtadd ("StockHoldingReport")
shtadd ("RBI_Report")
''''''''''''''''''''''''MTF_Report_copy''''''''''''''''''''''''''''''''''''''''
path3 = path2 & Format(strtdate, "ddmmyy") & "\"
cv = GetFullFileName(path3, "MTF_Report_")
 If cv <> Empty Or cv <> "" Then
     Set obj1 = Workbooks.Open(Filename:=path3 & cv)
    
   shtname = ActiveSheet.Name
    End If
    rcct = Sheets(shtname).UsedRange.Rows.Count
    ccct = Sheets(shtname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Selection.Copy
    ThisWorkbook.Worksheets("MTF_report").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Workbooks(obj1.Name).Close savechanges:=False
    rcct1 = Sheets("MTF_report").UsedRange.Rows.Count
    ccct1 = Sheets("MTF_report").UsedRange.Columns.Count
     Cells(rcct1, ccct1).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
   ' alook = Module_function.returncolumnnumber("MTF_report", "Category", 1)
    alk1 = returncolumnnumber("MTF_report", "Category", 1)
    ActiveSheet.Range("A1:" & celv).AutoFilter Field:=alk1, Criteria1:="UNSECURED LOAN", Operator:=xlAnd
    Range("A2:" & celv).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete xlUp
    Rows(1).Select
    Selection.AutoFilter
     ActiveSheet.Range("A1:" & celv).AutoFilter Field:=alk1, Criteria1:= _
        "=*p loan", Operator:=xlAnd
    Range("A2:" & celv).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete xlUp
    Rows(1).Select
    Selection.AutoFilter
    
    rcct1 = Sheets("MTF_report").UsedRange.Rows.Count
    ccct1 = Sheets("MTF_report").UsedRange.Columns.Count
''''''''''''''''''''''''''''''''Group1''''''''''''''''''''''
cv = GetFullFileName(path1, "Group1")
 If cv <> Empty Or cv <> "" Then
     Set obj1 = Workbooks.Open(Filename:=path1 & cv)
    
    shtname = ActiveSheet.Name
    End If
    rcct = Sheets(shtname).UsedRange.Rows.Count
    ccct = Sheets(shtname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Group1").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Workbooks(obj1.Name).Close savechanges:=False

''''''''''''''''''''''''''''''''Aproved stock''''''''''''''''''''''
cv = GetFullFileName(path1, "Aproved stock")
 If cv <> Empty Or cv <> "" Then
     Set obj1 = Workbooks.Open(Filename:=path1 & cv)
    shtname = ActiveSheet.Name
    End If
    rcct = Sheets(shtname).UsedRange.Rows.Count
    ccct = Sheets(shtname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Aproved stock").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Workbooks(obj1.Name).Close savechanges:=False
''''''''''''''''''''''''''ClientMFHoldingReport'''''''''''''''''''''''''''''''''
fld = Format(strtdate, "ddmmyyyy_") & "Margin Shortfall"

spath1 = spath & fld & "\"
cv = GetFullFileName(spath1, "ClientMFHoldingReport_")
 If cv <> Empty Or cv <> "" Then
 Set obj1 = Workbooks.Open(Filename:=spath1 & cv)
    shtname = ActiveSheet.Name
    End If
    Rows(1).Select
    Selection.AutoFilter
    rcct = Sheets(shtname).UsedRange.Rows.Count
    ccct = Sheets(shtname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Selection.Copy
    ThisWorkbook.Worksheets("ClientMFHoldingReport").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Workbooks(obj1.Name).Close savechanges:=False
    ''''''''''''''''''''''''''StockHoldingReport'''''''''''''''''''''''''''''''''
'spath = ThisWorkbook.Path & "\Input\"
cv = GetFullFileName(spath1, "StockHoldingReport_")
 If cv <> Empty Or cv <> "" Then
     Set obj1 = Workbooks.Open(Filename:=spath1 & cv)
    shtname = ActiveSheet.Name
    End If
    Rows(1).Select
    Selection.AutoFilter
    rcct = Sheets(shtname).UsedRange.Rows.Count
    ccct = Sheets(shtname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Selection.Copy
    ThisWorkbook.Worksheets("StockHoldingReport").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Workbooks(obj1.Name).Close savechanges:=False

''''''''''''''''''''''''''''''''''''''''''RBI REPORT''''''''''''''''''''''''''''''''''''''''''''''''''''
ThisWorkbook.Worksheets("RBI_Report").Activate
Sheets("RBI_Report").Cells(1, 1).Select
Selection.Value = "ClientCode"
ilstcol = Sheets("RBI_Report").Cells(1, 1).CurrentRegion.Columns.Count + 1
Module_function.Addcolummn "RBI_Report", "Group 1 Stock ", ilstcol
Module_function.Addcolummn "RBI_Report", "Approved Stock", ilstcol
Module_function.Addcolummn "RBI_Report", "Stock value", ilstcol
Module_function.Addcolummn "RBI_Report", "Rate", ilstcol
Module_function.Addcolummn "RBI_Report", "Qty", ilstcol
Module_function.Addcolummn "RBI_Report", "Final Category", ilstcol
Module_function.Addcolummn "RBI_Report", "Group 1", ilstcol
Module_function.Addcolummn "RBI_Report", "Approved Category", ilstcol
Module_function.Addcolummn "RBI_Report", "Category", ilstcol
Module_function.Addcolummn "RBI_Report", "ISIN", ilstcol
Module_function.Addcolummn "RBI_Report", "Scrip name", ilstcol
Module_function.Addcolummn "RBI_Report", "ClientName", ilstcol
Module_function.Addcolummn "RBI_Report", "FamilyName", ilstcol

Module_function.copycolmn "StockHoldingReport", "A/C Number", "RBI_Report", "ClientCode"
Module_function.copycolmn "StockHoldingReport", "ScripName ", "RBI_Report", "Scrip name"
Module_function.copycolmn "StockHoldingReport", "ISIN / ICIN", "RBI_Report", "ISIN"
Module_function.copycolmn "StockHoldingReport", "Logical Balance", "RBI_Report", "Qty"
Module_function.copycolmn "StockHoldingReport", "Total Holding Value", "RBI_Report", "Stock value"
Module_function.copycolmn "StockHoldingReport", "Category", "RBI_Report", "Category"
    rc = Sheets("RBI_Report").UsedRange.Rows.Count
    Rows(rc).Select
    Selection.Delete
    rc = Sheets("RBI_Report").UsedRange.Rows.Count
    rc1 = rc + 1
    Sheets("ClientMFHoldingReport").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("RBI_Report").Select
    Range("A" & rc1).Select
    ActiveSheet.Paste
    Sheets("ClientMFHoldingReport").Select
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("RBI_Report").Select
    Range("D" & rc1).Select
    ActiveSheet.Paste
    Sheets("ClientMFHoldingReport").Select
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("RBI_Report").Select
    Range("E" & rc1).Select
    ActiveSheet.Paste
    Sheets("ClientMFHoldingReport").Select
    Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("RBI_Report").Select
    Range("J" & rc1).Select
    ActiveSheet.Paste
    Sheets("ClientMFHoldingReport").Select
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("RBI_Report").Select
    Range("K" & rc1).Select
    ActiveSheet.Paste
    Range("F" & rc1).Select
    Selection.Value = "MF"
    Range("G" & rc1).Select
    Selection.Value = "MF"
    Range("H" & rc1).Select
    Selection.Value = "MutualFund"
    Range("I" & rc1).Select
    Selection.Value = "MF"
    rct = Sheets("RBI_Report").UsedRange.Rows.Count
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],MTF_report!C:C[2],3,0)"
    Range("B2").Select
    Selection.Copy
    Range("B2:B" & rct).Select
    ActiveSheet.Paste
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],MTF_report!C[-1]:C[1],2,0)"
    Range("C2").Select
    Selection.Copy
    Range("C2:C" & rct).Select
    ActiveSheet.Paste
    Range("F" & rc1 & ":" & "I" & rc1).Select
     Selection.Copy
    Range("F" & rc1 & ":" & "I" & rct).Select
    ActiveSheet.Paste
    Range("G2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-2],'Aproved stock'!C[-3]:C[-2],2,0),""Z"")"
    Range("G2").Select
    Selection.Copy
    Range("G2:G" & rc).Select
    ActiveSheet.Paste
    Range("G2:G" & rc).Select
    Selection.Copy
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],Group1!C[-5]:C[-4],2,0),""NotAppoved"")"
    Range("H2").Select
    Selection.Copy
    Range("H2:H" & rc).Select
    ActiveSheet.Paste
     Range("H2:H" & rc).Select
    Selection.Copy
    Range("H2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("I2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]=""A"",""A"",IF(RC[-3]=""B"",""B"",IF(RC[-3]=""C"",""C"",IF(RC[-3]=""D"",""D"",IF(RC[-3]=""E"",""E"",IF(RC[-3]=""F"",""F"",IF(RC[-3]=""G"",""G"",IF(RC[-3]=""H"",""H"",IF(RC[-3]=""T"",""T"",(IF(RC[-3]=""M"",""M"",IF(RC[-2]<>""Z"",RC[-2],IF(AND(RC[-2]=""Z"",RC[-1]=""Group 1""),""Y"",""Z"")))))))))))))"
    Range("I2").Select
    Selection.Copy
    Range("I2:I" & rc).Select
    ActiveSheet.Paste
    Range("I2:I" & rc).Select
    Selection.Copy
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=RC[1]/RC[-1]"
     Range("K2").Select
    Selection.Copy
    Range("K2:K" & rc).Select
    ActiveSheet.Paste
    Range("K2:K" & rc).Select
    Selection.Copy
    Range("K2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-4]=""Y"",RC[-4]=""Z""),0,RC[-1])"
    Range("M2").Select
    Selection.Copy
    Range("M2:M" & rct).Select
    ActiveSheet.Paste
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-5]=""Z"",0,RC[-2])"
    Range("N2").Select
    Selection.Copy
    Range("N2:N" & rct).Select
    ActiveSheet.Paste
    Range("L" & rc1).Select
    ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
     Range("L" & rc1).Select
    Selection.Copy
     Range("L" & rc1 & ":" & "L" & rct).Select
    ActiveSheet.Paste
    ilstcol = Sheets("RBI_Report").Cells(1, 1).CurrentRegion.Columns.Count + 1
    Module_function.Addcolummn "RBI_Report", "HairCut", ilstcol
    Module_function.copycolmn "StockHoldingReport", "HairCut", "RBI_Report", "HairCut"
    char2 = ColLtr(ilstcol + 1)
    char1 = ColLtr(ilstcol + 2)
    
     Range(char2 & 1).Select
    Selection.Value = "Margin Requirement %"
    Range(char2 & 2).Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-15],'MTF_report'!C[-14]:C[-5],10,0),""50"")"
        
    Range(char2 & 2).Select
    Selection.Copy
    Range(char2 & "2:" & char2 & rct).Select
    ActiveSheet.Paste
     Range("P" & rc1).Select
    Selection.Value = "0.01"
    Range("P" & rc1).Select
    Selection.Copy
    Range("P" & rc1 & ":" & "P" & rct).Select
    ActiveSheet.Paste
    Range(char1 & 1).Select
    Selection.Value = "Approved Stock (AHC)"
    Range(char1 & 2).Select
    ActiveCell.FormulaR1C1 = "=RC[-4]*(100-RC[-1])%"
    Range(char1 & 2).Select
    Selection.Copy
    Range(char1 & "2:" & char1 & rct).Select
    ActiveSheet.Paste
    cct = Sheets("RBI_Report").UsedRange.Columns.Count
     shtadd ("Pivot")
    shtadd ("FamilyWise")
 '''''''''''''''''''''''''''''''''''''''''PIVOT''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Pivot").Select
    Range("A2").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "RBI_Report!R1C1:R" & rct & "C" & cct, Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:="Pivot!R2C1", TableName:="PivotTable1", _
        DefaultVersion:=xlPivotTableVersion14
    Sheets("Pivot").Select
    Cells(2, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("FamilyName")
        .Orientation = xlRowField
        .position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Approved Stock (AHC)"), _
        "Sum of Approved Stock (AHC)", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Group 1 Stock "), "Sum of Group 1 Stock ", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Margin Requirement %"), _
        "Sum of Margin Requirement %", xlSum
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Sum of Margin Requirement %")
        .Caption = "Max of Margin Requirement %"
        .Function = xlMax
    End With
    ActiveWorkbook.ShowPivotTableFieldList = False
    rcct = Sheets("Pivot").UsedRange.Rows.Count
    ccct = Sheets("Pivot").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A2:" & celv).Select
    Selection.Copy
     Sheets("FamilyWise").Select
    Range("A4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
     Application.CutCopyMode = False
      Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "Loan"
    Range("B5").Select
    '''''''''''''''''''Pivot2'''''''''''''''''''''''''
     shtadd ("Pvt1")
     Sheets("Pvt1").Select
    Range("A2").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "MTF_Report!R1C1:R" & rcct1 & "C" & ccct1, Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:="Pvt1!R2C1", TableName:="PivotTable1", _
        DefaultVersion:=xlPivotTableVersion14
    Sheets("Pvt1").Select
    Cells(2, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Family Name ")
        .Orientation = xlRowField
        .position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Logical Loan Balance "), _
        "Sum of Logical Loan Balance ", xlSum
    ''''''''''''''''''''''''''''''''family wise''''''''''''''''''''''''''
    ThisWorkbook.Worksheets("FamilyWise").Activate
    rc = Sheets("FamilyWise").UsedRange.Rows.Count + 3
    cc = Sheets("FamilyWise").UsedRange.Columns.Count
    Cells(rc, cc).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],'Pvt1'!C[-1]:C,2,0)"
     Range("B5").Select
     Selection.Copy
      Range("B5:B" & rc).Select
    ActiveSheet.Paste
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]-RC[-4]"
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]*RC[-2]%"
    Range("H5").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
    Range("I5").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-3]/RC[-6]*100,0)"
    Range("J5").Select
    ActiveCell.FormulaR1C1 = "=RC[-6]-RC[-8]"
    Range("K5").Select
    ActiveCell.FormulaR1C1 = "=RC[-7]*RC[-6]%"
    Range("L5").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
    Range("M5").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-3]/RC[-9]*100,0)"
    Range("I5").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-3]/RC[-6]*100,0)"
    Range("M5").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-3]/RC[-9]*100,0)"
    Range("N5").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-13],MTF_report!C[-10]:C[35],46,0)"
    Range("F5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    rc = Sheets("FamilyWise").UsedRange.Rows.Count + 3
    cc = Sheets("FamilyWise").UsedRange.Columns.Count
    Cells(rc, cc).Select
    celv = Selection.Address(False, False)
    Range("F6:" & celv).Select
    ActiveSheet.Paste
    Range("F4").Select
    Selection.Value = "Available Margin on approved stock"
    Range("G4").Select
    Selection.Value = "Margin Required on Approved stock"
    Range("H4").Select
    Selection.Value = "Shortfall / Surplus"
    Range("I4").Select
    Selection.Value = "LVT % on Approved Stock"
    Range("J4").Select
    Selection.Value = "Available Margin on Group1 stock"
    Range("K4").Select
    Selection.Value = "Margin Required on Group1 stock"
    Range("L4").Select
    Selection.Value = "Shortfall / Surplus"
    Range("M4").Select
    Selection.Value = "LVT % on Group1 Stock"
    Range("N4").Select
    Selection.Value = "Category"
    Range("A4").Select
    Selection.Value = "Family Name"
    Range("C4").Select
    Selection.Value = "Approved Stock (AHC)"
    Range("D4").Select
    Selection.Value = "Group 1 Stock"
    Range("E4").Select
    Selection.Value = "Margin Requirement %"
    ''''''''''''''''''''''Family Report formart'''''''''''''''''''''''''''''''''''''''''''''''''''''
    rcct1 = Sheets("FamilyWise").UsedRange.Rows.Count
    ccct1 = Sheets("FamilyWise").UsedRange.Columns.Count
     Cells(rcct1, ccct1).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
   ' alook = Module_function.returncolumnnumber("MTF_report", "Category", 1)
    alk1 = returncolumnnumber("FamilyWise", "Category", 4)
    ActiveSheet.Range("A4:" & celv).AutoFilter Field:=alk1, Criteria1:="FID", Operator:=xlAnd
    Range("A5:" & celv).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete xlUp
    Rows(4).Select
    Selection.AutoFilter
    alk2 = returncolumnnumber("FamilyWise", "Loan", 4)
   ActiveSheet.Range("A4:" & celv).AutoFilter Field:=alk2, Criteria1:="<0", Operator:=xlAnd
    Range("A5:" & celv).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete xlUp
    Rows(4).Select
    Selection.AutoFilter
    rc1 = Sheets("FamilyWise").UsedRange.Rows.Count + 3
     cc1 = Sheets("FamilyWise").UsedRange.Columns.Count
    Cells(rc1, cc1).Select
    celv = Selection.Address(False, False)
    Range("A4:" & celv).Select
    Selection.Copy
     Range("A4").Select
     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A4:" & celv).Select
    Call Border
    ActiveWorkbook.Worksheets("FamilyWise").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("FamilyWise").Sort.SortFields.Add Key:=Range( _
        "B5:B" & rc1), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("FamilyWise").Sort
        .SetRange Range("A4:" & celv)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
     Range("B5").Select
     Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "#,##_);[Red](#,##)"
    Range("C5:D5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "#,##_);[Red](#,##)"
    Range("F5:H5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "#,##0_);[Red](#,##0)"
    Range("J5:L5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "#,##0_);[Red](#,##0)"
     Range("I5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0.00_);[Red](0.00)"
    Range("M5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0.00_);[Red](0.00)"
    Range("A4:" & celv).Select
    Call font
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Call Wrap
''''''''''''''''''''''''''''''''''''savereport''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Workbooks.Add
pn = ActiveWorkbook.Name
Workbooks(pn).Activate

ThisWorkbook.Worksheets("FamilyWise").Activate
Sheets(Array("FamilyWise", "RBI_Report", "StockHoldingReport", "ClientMFHoldingReport", "MTF_report")).Select
Sheets(Array("FamilyWise", "RBI_Report", "StockHoldingReport", "ClientMFHoldingReport", "MTF_report")).Copy Before:=Workbooks(pn).Sheets(1)
 For Each Sheet In ActiveWorkbook.Sheets
        If Sheet.Name = "Sheet1" Or Sheet.Name = "Sheet2" Or Sheet.Name = "Sheet3" Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        Else
        
        End If
Next Sheet
Worksheets("FamilyWise").Activate
Cells(1, 1).Select
 
Set fso = CreateObject("Scripting.FileSystemobject")
 spath = ThisWorkbook.Path & "\report"
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\report\" & Year(Now)
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\report\" & Year(Now) & "\" & MonthName(Month(Now))
' If Not fso.FolderExists(spath) Then fso.createfolder (spath)
' spath = ThisWorkbook.Path & "\report\" & Year(Now) & "\" & MonthName(Month(Now)) & "\" & Day(Now)
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 savereport = spath
 
        Set fso = CreateObject("Scripting.filesystemobject")
             If fso.FolderExists(savereport) Then
                 New_file_name = savereport & "\" & "FamilyWiseRBIReport_" & Format(strtdate, "ddmmyyyy") & ".xlsx"
                 Workbooks(pn).SaveAs Filename:=New_file_name
                 fname = "FamilyWiseRBIReport_" & Format(Now, "ddmmyyyy")
             Else
                 fso.createfolder (savereport)
                 
                 New_file_name = savereport & "\" & "FamilyWiseRBIReport_" & Format(strtdate, "ddmmyyyy") & ".xlsx"
                 Workbooks(pn).SaveAs Filename:=New_file_name
                 fname = "FamilyWiseRBIReport_" & Format(strtdate, "ddmmyyyy")
            End If
    
    ActiveWorkbook.Close savechanges = False
    ThisWorkbook.Worksheets("Home").Activate
'''''''''''''''''''''''''''SUMMARY'''''''''''''''''''''''''''''''''''''''''''''''''
 spath1 = ThisWorkbook.Path & "\"
  cv = GetFullFileName(spath1, "Summary_Report")
 If cv <> Empty Or cv <> "" Then
     Set obj1 = Workbooks.Open(Filename:=spath1 & cv)
   shtnname = ActiveSheet.Name
    End If
  Range("A2").Select
  Selection.Value = "Family Name"
    Range("B2").Select
  Selection.Value = "Category"
    Range("D2").Select
  Selection.Value = "Count"
  Range("C2").Select
  Selection.Value = "MarginRequirement %"
  Columns(5).Select
  Selection.Insert Shift:=xlToRight
    Range("E2").Select
  Selection.Value = "LVT % on Group1 Stock"
  workbook2 = ActiveWorkbook.Name
  ThisWorkbook.Worksheets("FamilyWise").Activate
   rc1 = Sheets("FamilyWise").UsedRange.Rows.Count + 3
    'Range("A4:" & celv).Select
   Range("O5").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-14],[Summary_Report.xlsx]Sheet1!C1,1,0)"
    Range("O5").Select
    Selection.Copy
    Range("O5:O" & rc1).Select
    ActiveSheet.Paste
    Range("A4").Select
    cc1 = Sheets("FamilyWise").UsedRange.Columns.Count
     Cells(rc1, cc1).Select
    celv = Selection.Address(False, False)
    ActiveSheet.Range("A4:" & celv).AutoFilter Field:=15, Criteria1:="#N/A"
     Range("A5:A" & rc1).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Windows(obj1.Name).Activate
     shtnname = ActiveSheet.Name
     rcct = Sheets(shtnname).UsedRange.Rows.Count
    ccct = Sheets(shtnname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
     Range("A" & rcct + 1).Select
     ActiveSheet.Paste
     rcct = Sheets(shtnname).UsedRange.Rows.Count
    ccct = Sheets(shtnname).UsedRange.Columns.Count
     Range("B3").Select
    
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],[RBI_Macro.xlsm]MTF_report!C[2]:C[41],40,0)"
    Range("B3").Select
    Selection.Copy
     Range("B3:B" & rcct).Select
    ActiveSheet.Paste
    Range("B3:B" & rcct).Select
    Selection.Copy
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    '''''''''''''''
     Range("C3").Select
   ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-2],[RBI_Macro.xlsm]FamilyWise!C1:C5,5,0),50)"
    Range("C3").Select
    Selection.Copy
     Range("C3:C" & rcct).Select
    ActiveSheet.Paste
    Range("C3:C" & rcct).Select
    Selection.Copy
    Range("C3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ''''''''''''''''''
    Range("E3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-4],[RBI_Macro.xlsm]FamilyWise!C1:C13,13,0),0)"
    Selection.Copy
     Range("E3:E" & rcct).Select
    ActiveSheet.Paste
    Range("E3:E" & rcct).Select
    Selection.Copy
    Range("E3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
     ThisWorkbook.Worksheets("MTF_report").Activate
    Range("A2").Select
    Selection.Copy
    Workbooks(obj1.Name).Activate
    Range("E1").Select
    ActiveSheet.Paste
     Range("E3:E" & rcct).Select
     Selection.NumberFormat = "0.00"
 '''''''''''''''''''''''''''''''''''''''''''''COUNT''''''''''''''''''''''''''''''''''''
     Workbooks(obj1.Name).Activate
     shtname = ActiveSheet.Name
      rcount = Sheets(shtname).UsedRange.Rows.Count
ccty = Sheets(shtname).UsedRange.Columns.Count
    For i = 3 To rcount
Count = 0

For j = 5 To ccty
      
        rbi_val1 = Cells(i, j).Value
      rbi_val2 = Cells(i, j + 1).Value

       If rbi_val1 > 50# Or rbi_val1 = 0# Then
        Cells(i, 4).Value = Count
         Exit For
           Else
                If rbi_val1 < 50# And rbi_val2 >= 50# Then
                Count = Count + 1
                Cells(i, 4).Value = Count
                Exit For
                End If
                    If rbi_val1 < 50# Then
                    Count = Count + 1
                        For k = j + 1 To ccty
                        Cells(i, k).Select
                        rbi_val3 = Cells(i, k).Value
                        rbi_val4 = Cells(i, k + 1).Value
                            If rbi_val3 < 50# And rbi_val4 >= 50# Then
                            Count = Count + 1
                            Cells(i, 4).Value = Count
                            Exit For
                            Else
                        If rbi_val3 < 50# Then
                        Count = Count + 1
                    k = k
                    If k = ccty Then
                    Cells(i, 4).Value = Count
                    j = k
                Exit For
                End If
                Cells(i, 4).Value = Count
                End If

             End If
            Next

        End If
      
      End If
  Exit For

Next

    Next
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Workbooks(obj1.Name).Close savechanges:=True
 ThisWorkbook.Worksheets("Home").Activate
    For Each sht In ThisWorkbook.Sheets
    If sht.Name = "Home" Or sht.Name = "Sheet1" Or sht.Name = "Sheet2" Or sht.Name = "Sheet3" Then
    Else
    sht.Activate
    Application.DisplayAlerts = False
    sht.Delete
    Application.DisplayAlerts = True
    End If
    Next
End Sub

 Sub shtadd(shtnname)
    exists = False
    For i = 1 To ActiveWorkbook.Worksheets.Count
    If ActiveWorkbook.Worksheets(i).Name = shtnname Then
        exists = True
    End If
    Next i
    If Not exists Then
    ActiveWorkbook.Worksheets.Add(After:=Worksheets(1)).Name = shtnname
    Else
    cler (shtnname)
    End If
    End Sub
    Function cler(SheetName)
    ActiveWorkbook.Worksheets(SheetName).Activate
    Cells.Select
    Cells.Clear
    Cells.Select
    Cells.Delete
    Cells(1, 1).Select
    End Function
 Function GetFullFileName(strfilepath, strFileNamePartial)
    
    Dim objFS As Variant
    Dim objFolder As Variant
    Dim objFile As Variant
    Dim intLengthOfPartialName As Integer
    Dim strfilenamefull As String
    
    Set objFS = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFS.getfolder(strfilepath)
    
    'work out how long the partial file name is
    intLengthOfPartialName = Len(strFileNamePartial)
    
    For Each objFile In objFolder.Files
    
    'Test to see if the file matches the partial file name
    If Left(objFile.Name, intLengthOfPartialName) = strFileNamePartial Then
    
    'get the full file name
    strfilenamefull = objFile.Name
    Exit For
    
    Else
    
    End If
    
    Next objFile
    
    'Return the full file name as the function's value
    GetFullFileName = strfilenamefull
    
    End Function
    
  Function ColLtr(icol)
    If icol > 0 And icol <= Columns.Count Then
        ColLtr = Evaluate("substitute(address(1, " & icol & ", 4), ""1"", """")")
    End If
    
End Function
 Function returncolumnnumber(parasheet, columnname, iRow)
    Sheets(parasheet).Activate
    tolst = Sheets(parasheet).UsedRange.Rows.Columns.Count + 5
    checkstatus = ""
    For i = 1 To tolst
    findval = Sheets(parasheet).Cells(iRow, i).Value
    If Trim(UCase(columnname)) = Trim(UCase(findval)) Then
    checkstatus = "found"
    Exit For
    End If
    Next
    If checkstatus = "" Then
    returncolumnnumber = 0
    Exit Function
    End If
    returncolumnnumber = i
    End Function
    Function Border()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    End Function

Function Wrap()

    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

End Function
Function font()
 With Selection.font
        .Name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
End Function
Function FolderExists(strFolderPath)
    On Error Resume Next
    FolderExists = (GetAttr(strFolderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0

  
End Function
Function returnrownumber(parasheet, rowname, icol)
    Sheets(parasheet).Activate
    tolst = Sheets(parasheet).UsedRange.Columns.Rows.Count + 100
    checkstatus = ""
    For i = 1 To tolst
    findval = Sheets(parasheet).Cells(i, icol).Value
    If Trim(UCase(rowname)) = Trim(UCase(findval)) Then
    checkstatus = "found"
    Exit For
    End If
    Next
    If checkstatus = "" Then
    returnrownumber = 0
    Exit Function
    End If
  
    returnrownumber = i
End Function
Function Vllookcopypastefromotherbook2(ivrow, ivcol, lookupbook, lookinbook, lookupsheet, lookinsheet, lookupcolnm, lookincolnm, getcol, putcol, iRow, err)
    inseclocde = returncolumnnumber(lookupsheet, lookupcolnm, 4)
    colchar = ColLtr(inseclocde) & ivrow
    insemo = returncolumnnumber(lookupsheet, putcol, iRow)
     plstrow = Sheets(lookupsheet).Cells(iRow, inseclocde).CurrentRegion.Rows.Count
     plstrow1 = plstrow + 5
     pasterng1 = ColLtr(ivcol) & ivrow
    pasterng2 = ColLtr(ivcol) & plstrow1 - 1
    
    Workbooks(lookinbook).Worksheets(lookinsheet).Activate
    idetclocde = returncolumnnumber(lookinsheet, lookincolnm, 2)
    idetmo = returncolumnnumber(lookinsheet, getcol, 2)
    ilastrow = Sheets(lookinsheet).UsedRange.Rows.Count
    Rng1 = Sheets(lookinsheet).Cells(1, idetclocde).Address
    
    rng2 = Sheets(lookinsheet).Cells(ilastrow, idetmo).Address
    search_table_range = Sheets(lookinsheet).Range(Rng1 & ":" & rng2).Address
    colmn = idetmo - idetclocde + 1
    'Sheets(lookupsheet).Cells(2, insemo).Value = "=IFERROR(VLOOKUP(" & colchar & ",'" & lookinsheet & "'!" & search_table_range & "," & colmn & ",0)," & err & ")"
    Workbooks(lookupbook).Activate
    If err = 0 Then
        Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & ",'[" & lookinbook & "]" & lookinsheet & "'!" & search_table_range & " ," & colmn & ",0)," & err & ")"
    Else

      Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & ",'[" & lookinbook & "]" & lookinsheet & "'!" & search_table_range & " ," & colmn & ",0)," & Chr(34) & err & Chr(34) & ")"
    End If
 'Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & "," & lookinsheet & "!" & search_table_range & " ," & colmn & ",0)," & Chr(34) & err & Chr(34) & ")"
    Sheets(lookupsheet).Cells(ivrow, ivcol).Copy
    Sheets(lookupsheet).Activate
    Sheets(lookupsheet).Range(pasterng1 & ":" & pasterng2).Select
    ActiveSheet.Paste
    Sheets(lookupsheet).Range(pasterng1 & ":" & pasterng2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets(lookupsheet).Range(pasterng1).Select

End Function
''''''''''''''''''rbiclientwise''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub clientwise()
 
ThisWorkbook.Worksheets("HOME").Activate
path1 = Range("G1").Value
strtdate = Range("E1").Value
shtadd ("MTF_report")
shtadd ("RBI_Report")
shtadd ("pivot2")
 shtadd ("pivot1")
 shtadd ("Report1")
d = Format(strtdate, "ddmmyyyy")
path2 = path1 & Format(strtdate, "yyyy") & "\" & Format(strtdate, "mmmm") & "\"
cv = GetFullFileName(path2, "FamilyWiseRBIReport_" & d)
 If cv <> Empty Or cv <> "" Then
     Set obj1 = Workbooks.Open(Filename:=path2 & cv)
      Sheets("MTF_report").Activate
      shtname = ActiveSheet.Name
    End If
   rcct = Sheets(shtname).UsedRange.Rows.Count
    ccct = Sheets(shtname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Selection.Copy
    ThisWorkbook.Worksheets("MTF_report").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Call deletecol
    rcct1 = Sheets("MTF_report").UsedRange.Rows.Count
    ccct1 = Sheets("MTF_report").UsedRange.Columns.Count
     Cells(rcct1, ccct1).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    alk1 = returncolumnnumber("MTF_report", "Department", 1)
    ActiveSheet.Range("A1:" & celv).AutoFilter Field:=alk1, Criteria1:="FID", Operator:=xlAnd
    Range("A2:" & celv).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete xlUp
    Rows(1).Select
    Selection.AutoFilter
 '''''''''''''''''RBIReport''''''''''''''''''''''''''
 Workbooks(obj1.Name).Activate
  Sheets("RBI_Report").Activate
  shtname = ActiveSheet.Name
 rcct = Sheets(shtname).UsedRange.Rows.Count
    ccct = Sheets(shtname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Selection.Copy
    ThisWorkbook.Worksheets("RBI_Report").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Call deletecol
    rcct1 = Sheets("RBI_Report").UsedRange.Rows.Count
    ccct1 = Sheets("RBI_Report").UsedRange.Columns.Count
     Cells(rcct1, ccct1).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    
    ''''''''''''pivot1'''''
    ThisWorkbook.Worksheets("RBI_Report").Activate
rcct1 = Sheets("RBI_Report").UsedRange.Rows.Count
     ccct1 = Sheets("RBI_Report").UsedRange.Columns.Count
ThisWorkbook.Worksheets("pivot1").Activate
  Range("A1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "RBI_Report!R1C1:R" & rcct1 & "C" & ccct1, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="pivot1!R1C1", TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion14
    Sheets("pivot1").Select
    Cells(1, 1).Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("ClientCode")
        .Orientation = xlRowField
        .position = 1
    End With
     
   ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Approved Stock (AHC)"), _
        "Sum of Approved Stock (AHC)", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Group 1 Stock "), "Sum of Group 1 Stock ", xlSum
    
Workbooks(obj1.Name).Close savechanges:=False
ThisWorkbook.Worksheets("MTF_report").Activate
Range("J1").Select
    ActiveCell.FormulaR1C1 = "Approved Stock (AHC)"
Range("K1").Select
    ActiveCell.FormulaR1C1 = "Group 1 Stock "
Range("L1").Select
    ActiveCell.FormulaR1C1 = "Available Margin on approved stock"
Range("M1").Select
    ActiveCell.FormulaR1C1 = "Available Margin on Group1 stock"

Range("N1").Select
    ActiveCell.FormulaR1C1 = "Margin Required on Approved Stock "
Range("O1").Select
    ActiveCell.FormulaR1C1 = "Margin Required on Group1 stock"

Range("P1").Select
    ActiveCell.FormulaR1C1 = "Shortfall / Surplus of approved stock"
Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Shortfall / Surplus of Group1"
  Range("A1").Select
    Selection.Copy
    Range("J1:Q1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Rows("1:1").Select
    With Selection
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
  Range("J2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-8],pivot1!C[-9]:C[-7],2,0),0)"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-9],pivot1!C[-10]:C[-8],3,0),0)"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-6]"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-7]"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]*RC[-7]%"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]*RC[-8]%"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]-RC[-2]"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]-RC[-2]"
    Range("J2:Q2").Select
    Selection.Copy
     rcct1 = Sheets("MTF_report").UsedRange.Rows.Count
     ccct1 = Sheets("MTF_report").UsedRange.Columns.Count
    Range("J" & "2:Q" & rcct1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
     Cells(rcct1, ccct1).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
        Sheets("pivot2").Select
    Range("A1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "MTF_report!R1C1:R" & rcct1 & "C" & ccct1, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="pivot2!R1C1", TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion14
    Sheets("pivot2").Select
    Cells(1, 1).Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Family Name ")
        .Orientation = xlRowField
        .position = 1
    End With
With ActiveSheet.PivotTables("PivotTable1").PivotFields("Department")
        .Orientation = xlRowField
        .position = 2
    End With
  ActiveSheet.PivotTables("PivotTable1").PivotFields("Family Name ").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Family Name ").LayoutForm _
        = xlTabular
    
'    ActiveSheet.PivotTables("PivotTable1").PivotFields("Family Name ").Subtotals = _
'        Array(False, False, False, False, False, False, False, False, False, False, False, False)
'    ActiveSheet.PivotTables("PivotTable1").PivotFields("Family Name ").LayoutForm _
'        = xlTabular
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Shortfall / Surplus of Group1"), _
        "Sum of Shortfall / Surplus of Group1", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Group 1 Stock "), "Sum of Group 1 Stock ", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Logical Loan Balance "), _
        "Sum of Logical Loan Balance ", xlSum
    ActiveSheet.PivotTables("PivotTable1").DataPivotField.PivotItems( _
        "Sum of Shortfall / Surplus of Group1").Caption = _
        "  Shortfall / Surplus of Group1"
    ActiveSheet.PivotTables("PivotTable1").DataPivotField.PivotItems( _
        "Sum of Group 1 Stock ").Caption = " Group 1 Stock "
    ActiveSheet.PivotTables("PivotTable1").DataPivotField.PivotItems( _
        "Sum of Logical Loan Balance ").Caption = " Logical Loan Balance "
    Sheets("pivot2").Select
    rc = Sheets("pivot2").UsedRange.Rows.Count
    cc = Sheets("pivot2").UsedRange.Columns.Count
     Cells(rc, cc).Select
    celv = Selection.Address(False, False)
    Range("A1:" & celv).Select
    Selection.Copy
     Sheets("Report1").Select
      Range("A1").Select
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     Range("A1").Select
    ActiveCell.FormulaR1C1 = "Family Name"
     Sheets("pivot2").Delete
     rc = Sheets("Report1").UsedRange.Rows.Count
    cc = Sheets("Report1").UsedRange.Columns.Count
     Cells(rc, cc).Select
    celv = Selection.Address(False, False)
    char1 = ColLtr(cc + 1)
    Range(char1 & 1).Select
    Selection.Value = "LTV%"
    Range(char1 & 2).Select
    ActiveCell.FormulaR1C1 = "=IFERROR((RC[-2]-RC[-1])/RC[-2]*100,0) "
    Selection.AutoFill Destination:=Range(char1 & "2:" & char1 & rc)
 Range(char1 & "2:" & char1 & rc).Copy
Range(char1 & 2).Select
 Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   Range(char1 & "2:" & char1 & rc).Select
 Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
   Range("B2:D" & rc).Select
    Selection.NumberFormat = "#,##0_);[Red](#,##0)"
     rc = Sheets("Report1").UsedRange.Rows.Count
    cc = Sheets("Report1").UsedRange.Columns.Count
     Cells(rc, cc).Select
    celv = Selection.Address(False, False)
    ActiveWorkbook.Worksheets("Report1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Report1").Sort.SortFields.Add Key:=Range( _
        "C2:C" & rc), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Report1").Sort
        .SetRange Range("A1:" & celv)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
  Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
     With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
        
   Range("A1:" & celv).Select
       Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
        
    '''''''''''''''''''''''''''''''''''save report'''''''''''''''''''''''''''''''''''''''''''''
Workbooks.Add
pn = ActiveWorkbook.Name
Workbooks(pn).Activate

ThisWorkbook.Worksheets("Report1").Activate
Sheets(Array("RBI_Report", "MTF_report", "Report1")).Select
Sheets(Array("RBI_Report", "MTF_report", "Report1")).Copy Before:=Workbooks(pn).Sheets(1)
 For Each Sheet In ActiveWorkbook.Sheets
        If Sheet.Name = "Sheet1" Or Sheet.Name = "Sheet2" Or Sheet.Name = "Sheet3" Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        Else
        
        End If
Next Sheet
Worksheets("Report1").Activate
Cells(1, 1).Select
 
Set fso = CreateObject("Scripting.FileSystemobject")
 spath = ThisWorkbook.Path & "\ClientwiseRBIReport"
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\ClientwiseRBIReport\" & Year(Now)
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\ClientwiseRBIReport\" & Year(Now) & "\" & MonthName(Month(Now))
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\ClientwiseRBIReport\" & Year(Now) & "\" & MonthName(Month(Now)) & "\" & Day(Now)
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 savereport = spath
     Set fso = CreateObject("Scripting.filesystemobject")
             If fso.FolderExists(savereport) Then
                 New_file_name = savereport & "\" & "ClientwiseRBIReport_" & Format(strtdate, "ddmmyyyy") & ".xlsx"
                 Workbooks(pn).SaveAs Filename:=New_file_name
                 fname = "ClientwiseRBIReport_" & Format(Now, "ddmmyyyy")
             Else
                 fso.createfolder (savereport)
                 
                 New_file_name = savereport & "\" & "ClientwiseRBIReport_" & Format(strtdate, "ddmmyyyy") & ".xlsx"
                 Workbooks(pn).SaveAs Filename:=New_file_name
                 fname = "ClientwiseRBIReport_" & Format(strtdate, "ddmmyyyy")
            End If
    
    ActiveWorkbook.Close savechanges = False
    ThisWorkbook.Worksheets("Home").Activate
'''''''''''''''''''''''''''''''''''''sumerry client wise'''''''''''''''''''''''''''''''''''
spath1 = ThisWorkbook.Path & "\"
  cv = GetFullFileName(spath1, "Summary_Clientwise")
 If cv <> Empty Or cv <> "" Then
     Set obj1 = Workbooks.Open(Filename:=spath1 & cv)
   shtnname = ActiveSheet.Name
    End If
  Range("A2").Select
  Selection.Value = "Family Name"
  Range("B2").Select
  Selection.Value = "LTV%"
    Range("D2").Select
  Selection.Value = "Count"
  workbook2 = ActiveWorkbook.Name
  ThisWorkbook.Worksheets("Report1").Activate
  rc1 = Sheets("Report1").UsedRange.Rows.Count
   Range("G2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-6],[Summary_Clientwise.xlsx]Sheet1!C1,1,0)"
  Range("G2").Select
    Selection.Copy
    Range("G2:G" & rc1).Select
    ActiveSheet.Paste
   Range("A4").Select
    cc1 = Sheets("Report1").UsedRange.Columns.Count
     Cells(rc1 - 1, cc1).Select
    celv = Selection.Address(False, False)
    ActiveSheet.Range("A1:" & celv).AutoFilter Field:=7, Criteria1:="#N/A"
     Range("A2:A" & rc1).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Windows(obj1.Name).Activate
     shtnname = ActiveSheet.Name
     rcct = Sheets(shtnname).UsedRange.Rows.Count + 1
    ccct = Sheets(shtnname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
     Range("A" & rcct).Select
     ActiveSheet.Paste
     ThisWorkbook.Worksheets("Report1").Activate
     Rows(1).Select
     Selection.AutoFilter
     Windows(obj1.Name).Activate
      shtnname = ActiveSheet.Name
     rcct = Sheets(shtnname).UsedRange.Rows.Count
    ccct = Sheets(shtnname).UsedRange.Columns.Count
    Range("B3").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],[RBI_ClientWise.xlsm]Report1!C1:C6,6,0)"
    Range("B3").Select
        Selection.Copy
     Range("B3:B" & rcct).Select
    ActiveSheet.Paste
    Range("B3:B" & rcct).Select
    Selection.Copy
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
     Selection.NumberFormat = "0.00"
     Range("C2").Select
  Selection.Value = "Department"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-2],[RBI_ClientWise.xlsm]Report1!C1:C2,2,0)"
    Range("C3").Select
        Selection.Copy
     Range("C3:C" & rcct).Select
    ActiveSheet.Paste
    Range("C3:C" & rcct).Select
    Selection.Copy
    Range("C3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
     Columns(5).Select
  Selection.Insert Shift:=xlToRight
      Range("E2").Select
  Selection.Value = "  Shortfall/Surplus of Group1"
       Range("E3").Select
     ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-4],[RBI_ClientWise.xlsm]Report1!C1:C3,3,0)"
     Range("E3").Select
        Selection.Copy
     Range("E3:E" & rcct).Select
    ActiveSheet.Paste
    Range("E3:E" & rcct).Select
    Selection.Copy
    Range("E3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
     Range("E1").Select
   Selection.Value = Format(strtdate, "ddmmmyy")
   Range("E3:E" & rcct).Select
   Selection.NumberFormat = "#,##0_);[Red](#,##0)"
   
    rcct = Sheets(shtnname).UsedRange.Rows.Count
    ccct = Sheets(shtnname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    ActiveSheet.Range("A1:" & celv).AutoFilter Field:=3, Criteria1:="=0", _
        Operator:=xlOr, Criteria2:="=(blank)"
    If Range("A2:" & celv).SpecialCells(xlCellTypeVisible).Count > 1 Then
    Range("A3:" & celv).SpecialCells(xlCellTypeVisible).Select
    Selection.EntireRow.Delete
    Rows(1).Select
    Selection.AutoFilter
    Else
    End If
     rcct = Sheets(shtnname).UsedRange.Rows.Count
    ccct = Sheets(shtnname).UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    ActiveSheet.Range("A2:" & celv).AutoFilter Field:=2, Criteria1:="=#N/A"
    If Range("A3:" & celv).SpecialCells(xlCellTypeVisible).Count > 0 Then
    Range("A3:" & celv).SpecialCells(xlCellTypeVisible).Select
    Selection.EntireRow.Delete
    Rows(2).Select
    Selection.AutoFilter
    Else
    End If
   
   ''''''''''''''''''count'''''''''''''''''''''''''''''
   Workbooks(obj1.Name).Activate
     shtname = ActiveSheet.Name
      rcount = Sheets(shtname).UsedRange.Rows.Count
ccty = Sheets(shtname).UsedRange.Columns.Count
    For i = 3 To rcount
Count = 0

For j = 5 To ccty
      
        rbi_val1 = Cells(i, j).Value
      rbi_val2 = Cells(i, j + 1).Value

       If rbi_val1 > 0# Or rbi_val1 = 0# Then
        Cells(i, 4).Value = Count
         Exit For
           Else
                If rbi_val1 < 0# And rbi_val2 >= 0# Then
                Count = Count + 1
                Cells(i, 4).Value = Count
                Exit For
                End If
                    If rbi_val1 < 0# Then
                    Count = Count + 1
                        For k = j + 1 To ccty
                        Cells(i, k).Select
                        rbi_val3 = Cells(i, k).Value
                        rbi_val4 = Cells(i, k + 1).Value
                            If rbi_val3 < 0# And rbi_val4 >= 0# Then
                            Count = Count + 1
                            Cells(i, 4).Value = Count
                            Exit For
                            Else
                        If rbi_val3 < 0# Then
                        Count = Count + 1
                    k = k
                    If k = ccty Then
                    Cells(i, 4).Value = Count
                    j = k
                Exit For
                End If
                Cells(i, 4).Value = Count
                End If

             End If
            Next

        End If
      
      End If
  Exit For

Next
    Next
    
    
  
       
  Workbooks(obj1.Name).Close savechanges:=True
  '''''''''''''''''''''''''''''''''''''''''''
 ThisWorkbook.Worksheets("Home").Activate
    For Each sht In ThisWorkbook.Sheets
    If sht.Name = "Home" Or sht.Name = "Sheet1" Or sht.Name = "Sheet2" Or sht.Name = "Sheet3" Then
    Else
    sht.Activate
    Application.DisplayAlerts = False
    sht.Delete
    Application.DisplayAlerts = True
    End If
    Next
   ThisWorkbook.Worksheets("Home").Activate

    
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
Sub shtadd(shtnname)
    exists = False
    For i = 1 To ActiveWorkbook.Worksheets.Count
    If ActiveWorkbook.Worksheets(i).Name = shtnname Then
        exists = True
    End If
    Next i
    If Not exists Then
    ActiveWorkbook.Worksheets.Add(After:=Worksheets(1)).Name = shtnname
    Else
    cler (shtnname)
    End If
    End Sub
    Function cler(SheetName)
    ActiveWorkbook.Worksheets(SheetName).Activate
    Cells.Select
    Cells.Clear
    Cells.Select
    Cells.Delete
    Cells(1, 1).Select
    End Function
 Function GetFullFileName(strfilepath, strFileNamePartial)
    
    Dim objFS As Variant
    Dim objFolder As Variant
    Dim objFile As Variant
    Dim intLengthOfPartialName As Integer
    Dim strfilenamefull As String
    
    Set objFS = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFS.getfolder(strfilepath)
    
    'work out how long the partial file name is
    intLengthOfPartialName = Len(strFileNamePartial)
    
    For Each objFile In objFolder.Files
    
    'Test to see if the file matches the partial file name
    If Left(objFile.Name, intLengthOfPartialName) = strFileNamePartial Then
    
    'get the full file name
    strfilenamefull = objFile.Name
    Exit For
    
    Else
    
    End If
    
    Next objFile
    
    'Return the full file name as the function's value
    GetFullFileName = strfilenamefull
    
    End Function
    
  Function ColLtr(icol)
    If icol > 0 And icol <= Columns.Count Then
        ColLtr = Evaluate("substitute(address(1, " & icol & ", 4), ""1"", """")")
    End If
    
End Function
 Function returncolumnnumber(parasheet, columnname, iRow)
    Sheets(parasheet).Activate
    tolst = Sheets(parasheet).UsedRange.Rows.Columns.Count + 5
    checkstatus = ""
    For i = 1 To tolst
    findval = Sheets(parasheet).Cells(iRow, i).Value
    If Trim(UCase(columnname)) = Trim(UCase(findval)) Then
    checkstatus = "found"
    Exit For
    End If
    Next
    If checkstatus = "" Then
    returncolumnnumber = 0
    Exit Function
    End If
    returncolumnnumber = i
    End Function
    Function Border()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    End Function

Function Wrap()

    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

End Function
Function font()
 With Selection.font
        .Name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
End Function
Function FolderExists(strFolderPath)
    On Error Resume Next
    FolderExists = (GetAttr(strFolderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0

  
End Function
Function returnrownumber(parasheet, rowname, icol)
    Sheets(parasheet).Activate
    tolst = Sheets(parasheet).UsedRange.Columns.Rows.Count + 100
    checkstatus = ""
    For i = 1 To tolst
    findval = Sheets(parasheet).Cells(i, icol).Value
    If Trim(UCase(rowname)) = Trim(UCase(findval)) Then
    checkstatus = "found"
    Exit For
    End If
    Next
    If checkstatus = "" Then
    returnrownumber = 0
    Exit Function
    End If
  
    returnrownumber = i
End Function
Function Vllookcopypastefromotherbook2(ivrow, ivcol, lookupbook, lookinbook, lookupsheet, lookinsheet, lookupcolnm, lookincolnm, getcol, putcol, iRow, err)
    inseclocde = returncolumnnumber(lookupsheet, lookupcolnm, 4)
    colchar = ColLtr(inseclocde) & ivrow
    insemo = returncolumnnumber(lookupsheet, putcol, iRow)
     plstrow = Sheets(lookupsheet).Cells(iRow, inseclocde).CurrentRegion.Rows.Count
     plstrow1 = plstrow + 5
     pasterng1 = ColLtr(ivcol) & ivrow
    pasterng2 = ColLtr(ivcol) & plstrow1 - 1
    
    Workbooks(lookinbook).Worksheets(lookinsheet).Activate
    idetclocde = returncolumnnumber(lookinsheet, lookincolnm, 2)
    idetmo = returncolumnnumber(lookinsheet, getcol, 2)
    ilastrow = Sheets(lookinsheet).UsedRange.Rows.Count
    Rng1 = Sheets(lookinsheet).Cells(1, idetclocde).Address
    
    rng2 = Sheets(lookinsheet).Cells(ilastrow, idetmo).Address
    search_table_range = Sheets(lookinsheet).Range(Rng1 & ":" & rng2).Address
    colmn = idetmo - idetclocde + 1
    'Sheets(lookupsheet).Cells(2, insemo).Value = "=IFERROR(VLOOKUP(" & colchar & ",'" & lookinsheet & "'!" & search_table_range & "," & colmn & ",0)," & err & ")"
    Workbooks(lookupbook).Activate
    If err = 0 Then
        Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & ",'[" & lookinbook & "]" & lookinsheet & "'!" & search_table_range & " ," & colmn & ",0)," & err & ")"
    Else

      Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & ",'[" & lookinbook & "]" & lookinsheet & "'!" & search_table_range & " ," & colmn & ",0)," & Chr(34) & err & Chr(34) & ")"
    End If
 'Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & "," & lookinsheet & "!" & search_table_range & " ," & colmn & ",0)," & Chr(34) & err & Chr(34) & ")"
    Sheets(lookupsheet).Cells(ivrow, ivcol).Copy
    Sheets(lookupsheet).Activate
    Sheets(lookupsheet).Range(pasterng1 & ":" & pasterng2).Select
    ActiveSheet.Paste
    Sheets(lookupsheet).Range(pasterng1 & ":" & pasterng2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets(lookupsheet).Range(pasterng1).Select

End Function
Sub deletecol()
ThisWorkbook.Worksheets("MTF_report").Activate
 Module_function.Delcolummn "MTF_report", "Margin Coverage After Haircut %"
Module_function.Delcolummn "MTF_report", "Margin Coverage After Haircut Amount %"
Module_function.Delcolummn "MTF_report", " Portfolio At Mkt Price Before Haircut "
Module_function.Delcolummn "MTF_report", "Portfolio At Mkt Price After Haircut Approved"
Module_function.Delcolummn "MTF_report", "Min Margin Requirement"
Module_function.Delcolummn "MTF_report", "Amount % Margin Surplus / (Shortfall) after Haircut"
Module_function.Delcolummn "MTF_report", "Bank Balance + cash margin"
Module_function.Delcolummn "MTF_report", "Branch Name"
Module_function.Delcolummn "MTF_report", "Risk Manager Remarks"
Module_function.Delcolummn "MTF_report", "Net Margin Surplus / (Shortfall) % "
Module_function.Delcolummn "MTF_report", "Loan Balance As On Date"
Module_function.Delcolummn "MTF_report", "Additional Loan Balance"
Module_function.Delcolummn "MTF_report", "Unsettled Obligation"
Module_function.Delcolummn "MTF_report", "Uncollected Outstanding Interest"
Module_function.Delcolummn "MTF_report", "Client Bank Balance"
Module_function.Delcolummn "MTF_report", "Add On Loan Sanctioned"
Module_function.Delcolummn "MTF_report", "Pan No"
Module_function.Delcolummn "MTF_report", "TDS Credit Receivable"
Module_function.Delcolummn "MTF_report", "Applied Normal haircut"
Module_function.Delcolummn "MTF_report", "Applied Additional Haircut"
Module_function.Delcolummn "MTF_report", "Applied Concentration Haircut"
Module_function.Delcolummn "MTF_report", "Applied Total Haircut "
Module_function.Delcolummn "MTF_report", "Margin Coverage Before Haircut %"
Module_function.Delcolummn "MTF_report", "Margin Coverage Before haircut Amount"
Module_function.Delcolummn "MTF_report", "Margin Coverage Mezanine Level %"
Module_function.Delcolummn "MTF_report", "Margin Coverage Mezanine Level Amount "
Module_function.Delcolummn "MTF_report", "Uncleared Bank Balance"
Module_function.Delcolummn "MTF_report", "Third Party Stock Value"
Module_function.Delcolummn "MTF_report", "Third Party Cheque Value"
Module_function.Delcolummn "MTF_report", "Probable Margin % after unclear Chq /Thirdparty Stock"
Module_function.Delcolummn "MTF_report", "Probable Margin Amount after unclear Chq /Thirdparty Stock"
Module_function.Delcolummn "MTF_report", "Net Available Loan Balance"
Module_function.Delcolummn "MTF_report", "Additional Exposure Available over Unutilised Margin"
Module_function.Delcolummn "MTF_report", "Sub Broker"
Module_function.Delcolummn "MTF_report", "RM"
Module_function.Delcolummn "MTF_report", "Risk Manager"
Module_function.Delcolummn "MTF_report", "Cash Margin"
Module_function.Delcolummn "MTF_report", "Relationship Manager Remarks"
Module_function.Delcolummn "MTF_report", "RBI Category"
Module_function.Delcolummn "MTF_report", "Loan Policy Category "
'Module_function.Delcolummn "MTF_report", "Department"
Module_function.Delcolummn "MTF_report", "Risk Category "
Module_function.Delcolummn "MTF_report", "Loan Account Status"
Module_function.Delcolummn "MTF_report", "Min Margin Requirement Amount %"
Module_function.Delcolummn "MTF_report", "Margin Surplus / (Shortfall) after Haircut "
'Module_function.Delcolummn "MTF_report", "Category"
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public startdt1 As Variant
Public enddt As Variant
Public dtdf As Variant
Public fldrpt As String
Public dt As Variant

Sub Summary()
'''''''''''''''''shtadd'''''''''''''''''''''''
    shtadd ("Summary")
    shtadd ("Loan")
    shtadd ("JM approved")
    shtadd ("Group1")
    shtadd ("Margin%")
     shtadd ("Margin% on group 1")
''''''''''''''''''HOME'''''''''''''''''''''''''
    ThisWorkbook.Worksheets("HOME").Activate
    path1 = Range("G1").Value
    Enterdate.Show
    strtdt = Enterdate.Startdate.Text
    startdt1 = strtdt
    ThisWorkbook.Worksheets("HOME").Activate
    Range("G2").Value = startdt1
    enddt = Enterdate.Enddate.Text
    Range("G3").Value = enddt
    Range("G4").Select
    dtdf = Selection.Value
    yr = Format(startdt1, "yyyy")
    mon = Format(startdt1, "mmmm")
    dt = Format(startdt1, "d")
    fldrpt = path1 & "\" & yr & "\" & mon & "\" & dt
    
        
''''''''''''''''''''''''''''''''For''''''''''''''''''''''''''''''''''''''''''
 For i = 0 To dtdf

    dt = Format(startdt1, "d")
    yr = Format(startdt1, "yyyy")
    mon = Format(startdt1, "mmmm")
    dt = Format(startdt1, "d")
    fldrpt = path1 & "\" & yr & "\" & mon & "\" & dt
    aa = FolderExists(fldrpt)
    If aa <> "" Then
        cv = GetFullFileName(fldrpt, "MTF_Book_Riview_" & Format(startdt1, "d"))
        If cv = "" Then
        MsgBox "Not found"
   
    Else
        Set obj1 = Workbooks.Open(Filename:=fldrpt & "\" & cv)
        pn = ActiveWorkbook.Name
        ActiveWorkbook.Worksheets("Family_Wise").Select
        shtnname = ActiveSheet.Name
        rcct = ActiveWorkbook.Worksheets("Family_Wise").UsedRange.Rows.Count
        ccct = ActiveWorkbook.Worksheets("Family_Wise").UsedRange.Columns.Count
        Cells(rcct, ccct).Select
        celv = Selection.Address(False, False)
        alk1 = returncolumnnumber("Family_Wise", "Familyname", 6)
        colchar1 = ColLtr(alk1)
        colchar = ColLtr(ccct + 2)
        rcct = ActiveWorkbook.Worksheets("Family_Wise").UsedRange.Rows.Count
        ccct = ActiveWorkbook.Worksheets("Family_Wise").UsedRange.Columns.Count
        Cells(rcct - 1, ccct + 2).Select
        celv = Selection.Address(False, False)
        Cells(6, ccct + 2).Select
        Selection.Value = "fmlynm"
        Cells(7, ccct + 2).Select
      Vllookcopypastefromotherbook2 7, ccct + 2, pn, ThisWorkbook.Name, shtnname, "Familyname", "Family Name", "Familyname", " Familyname", "fmlynm", 6, 0
        Cells(6, ccct + 2).Select
        ActiveSheet.Range("A6:" & celv).AutoFilter Field:=ccct + 1, Criteria1:="0", Operator:=xlFilterValues
        Range("B6").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
        ThisWorkbook.Worksheets("Familyname").Activate
        Range("A1").Select
        rcct = Sheets("Familyname").UsedRange.Rows.Count
        
        Range("A" & rcct + 1).Select
        ActiveSheet.Paste
        Range("A" & rcct + 1).Select
        Selection.Delete
        rcct = Sheets("Familyname").UsedRange.Rows.Count
          If (Range("A" & rcct).Value) = "" Then
            Range("A" & rcct).Select
            Selection.Delete
  End If
''''''''''''''''''''''Familyname'''''''''''''''''''''''
        ThisWorkbook.Worksheets("Familyname").Activate
        rcct = Sheets("Familyname").UsedRange.Rows.Count
        Range("A1:A" & rcct).Select
        Selection.Copy
        ThisWorkbook.Worksheets("Summary").Activate
        Range("A5").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        ThisWorkbook.Worksheets("Loan").Activate
        Range("A5").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        ThisWorkbook.Worksheets("JM approved").Activate
        Range("A5").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        ThisWorkbook.Worksheets("Group1").Activate
        Range("A5").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        ThisWorkbook.Worksheets("Margin%").Activate
        Range("A5").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        ThisWorkbook.Worksheets("Margin% on group 1").Activate
        Range("A5").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
        
        
        
'''''''''''''''Summary'''''''''''''''''''''''
    ThisWorkbook.Worksheets("Summary").Activate
        rcct = ActiveWorkbook.Worksheets("Summary").UsedRange.Rows.Count
        ccct = ActiveWorkbook.Worksheets("Summary").UsedRange.Columns.Count
        Cells(rcct, ccct).Select
    alk1 = returncolumnnumber("Summary", "Familyname", 5)
        colchar1 = ColLtr(ccct + 1)
        colchar2 = ColLtr(ccct + 2)
        colchar3 = ColLtr(ccct + 3)
         colchar4 = ColLtr(ccct + 4)
          colchar5 = ColLtr(ccct + 5)
        Range(colchar1 & 5).Value = "Family Loan Amt in Cr."
        Range(colchar2 & 5).Value = "Jm Approved Stock in (Cr.)"
        Range(colchar3 & 5).Value = " Group 1 Stock in (Cr.)"
        Range(colchar4 & 5).Value = "Margin % on Approved stock"
        Range(colchar5 & 5).Value = "Margin% on group 1"
     ccct = ActiveWorkbook.Worksheets("Summary").UsedRange.Columns.Count
   Vllookcopypastefromotherbook 6, ccct - 4, ThisWorkbook.Name, pn, "Summary", shtnname, "Familyname", "Family Name", " Family Loan Amt in Cr.", "Family Loan Amt in Cr.", 5, 0
   Vllookcopypastefromotherbook 6, ccct - 3, ThisWorkbook.Name, pn, "Summary", shtnname, "Familyname", "Family Name", " Jm Approved Stock in (Cr.)", "Jm Approved Stock in (Cr.)", 5, 0
   Vllookcopypastefromotherbook 6, ccct - 2, ThisWorkbook.Name, pn, "Summary", shtnname, "Familyname", "Family Name", " Group 1 Stock in (Cr.)", " Group 1 Stock in (Cr.)", 5, 0
   Vllookcopypastefromotherbook 6, ccct - 1, ThisWorkbook.Name, pn, "Summary", shtnname, "Familyname", "Family Name", "Margin % on Approved stock", " Margin % on Approved stock", 5, 0
   Vllookcopypastefromotherbook 6, ccct, ThisWorkbook.Name, pn, "Summary", shtnname, "Familyname", "Family Name", "Margin% on group 1", "Margin% on group 1", 5, 0

 Range(colchar1 & "6:" & colchar5 & rcct).Select
 Selection.Style = "Comma"
   Call Macro2
       Range(colchar1 & "4:" & colchar5 & 4).Select
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
    Selection.Merge
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range(colchar1 & "4:" & colchar5 & 4).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(colchar1 & "4:" & colchar5 & 4).Select
    startdt1 = Format(startdt1, "dd-mmm-yy")
    Selection.Value = startdt1
     
    
''''''''''''''''''''''''''''''''Loan'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ThisWorkbook.Worksheets("Loan").Activate
        rcct = ActiveWorkbook.Worksheets("Loan").UsedRange.Rows.Count
        ccct = ActiveWorkbook.Worksheets("Loan").UsedRange.Columns.Count
        Cells(rcct, ccct).Select
        alk1 = returncolumnnumber("Loan", "Familyname", 5)
        colchar1 = ColLtr(ccct + 1)
        Range(colchar1 & 5).Value = "Family Loan Amt in Cr."
   ccct = ActiveWorkbook.Worksheets("Loan").UsedRange.Columns.Count
 Vllookcopypastefromotherbook 6, ccct, ThisWorkbook.Name, pn, "Loan", shtnname, "Familyname", "Family Name", " Family Loan Amt in Cr.", "Family Loan Amt in Cr.", 5, 0
  Range(colchar1 & "6:" & colchar1 & rcct).Select
 Selection.Style = "Comma"
  Call Macro3
  Range(colchar1 & 4).Select
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
    Selection.Merge
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range(colchar1 & 4).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(colchar1 & 4).Select
    Selection.Value = startdt1
     startdt1 = Format(startdt1, "dd-mmm-yy")
    Columns(colchar1).ColumnWidth = 10
    ''''''''''''''''''''''JM approved''''''''''''''''''''''''''
  ThisWorkbook.Worksheets("JM approved").Activate
        rcct = ActiveWorkbook.Worksheets("JM approved").UsedRange.Rows.Count
        ccct = ActiveWorkbook.Worksheets("JM approved").UsedRange.Columns.Count
        Cells(rcct, ccct).Select
        colchar1 = ColLtr(ccct + 1)
        Range(colchar1 & 5).Value = "Jm Approved Stock in (Cr.)"
   ccct = ActiveWorkbook.Worksheets("JM approved").UsedRange.Columns.Count
 Vllookcopypastefromotherbook 6, ccct, ThisWorkbook.Name, pn, "JM approved", shtnname, "Familyname", "Family Name", " Jm Approved Stock in (Cr.)", "Jm Approved Stock in (Cr.)", 5, 0
   Range(colchar1 & "6:" & colchar1 & rcct).Select
 Selection.Style = "Comma"
   Call Macro3
  Range(colchar1 & 4).Select
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
    Selection.Merge
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range(colchar1 & 4).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(colchar1 & 4).Select
    Selection.Value = startdt1
     startdt1 = Format(startdt1, "dd-mmm-yy")
    Columns(colchar1).ColumnWidth = 10
    '''''''''''''''''''''''''''''Group'''''''''''''''''''''''''''''''''
  ThisWorkbook.Worksheets("Group1").Activate
        rcct = ActiveWorkbook.Worksheets("Group1").UsedRange.Rows.Count
        ccct = ActiveWorkbook.Worksheets("Group1").UsedRange.Columns.Count
        Cells(rcct, ccct).Select
        alk1 = returncolumnnumber("Group1", "Familyname", 5)
        colchar1 = ColLtr(ccct + 1)
        Range(colchar1 & 5).Value = "Group 1 Stock in (Cr.)"
   ccct = ActiveWorkbook.Worksheets("Group1").UsedRange.Columns.Count
  Vllookcopypastefromotherbook 6, ccct, ThisWorkbook.Name, pn, "Group1", shtnname, "Familyname", "Family Name", " Group 1 Stock in (Cr.)", " Group 1 Stock in (Cr.)", 5, 0
  Range(colchar1 & "6:" & colchar1 & rcct).Select
 Selection.Style = "Comma"
  Call Macro3
  Range(colchar1 & 4).Select
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
    Selection.Merge
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range(colchar1 & 4).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(colchar1 & 4).Select
    Selection.Value = startdt1
     startdt1 = Format(startdt1, "dd-mmm-yy")
  Columns(colchar1).ColumnWidth = 10
  ''''''''''''''''''''''''''''''''''''''''''''''''%margin'''''''''''''''''''''''''''''''''''''''
  
   ThisWorkbook.Worksheets("Margin%").Activate
        rcct = ActiveWorkbook.Worksheets("Margin%").UsedRange.Rows.Count
        ccct = ActiveWorkbook.Worksheets("Margin%").UsedRange.Columns.Count
        Cells(rcct + 4, ccct).Select
        
        colchar1 = ColLtr(ccct + 1)
        Range(colchar1 & 5).Value = "Margin % on Approved stock"
   ccct = ActiveWorkbook.Worksheets("Margin%").UsedRange.Columns.Count
  Vllookcopypastefromotherbook 6, ccct, ThisWorkbook.Name, pn, "Margin%", shtnname, "Familyname", "Family Name", "Margin % on Approved stock", " Margin % on Approved stock", 5, 0
  Range(colchar1 & "6:" & colchar1 & rcct).Select
 Selection.Style = "Comma"
 Call Macro3
  Range(colchar1 & 4).Select
    Selection.Value = startdt1
     startdt1 = Format(startdt1, "dd-mmm-yy")
  Columns(colchar1).ColumnWidth = 10
  
''''''''''''''''''''''''''''''''''''Margin% on group 1''''''''''''''''''''''''''''''''''''''''''''''''
  
 
   ThisWorkbook.Worksheets("Margin% on group 1").Activate
        rcct = ActiveWorkbook.Worksheets("Margin% on group 1").UsedRange.Rows.Count
        ccct = ActiveWorkbook.Worksheets("Margin% on group 1").UsedRange.Columns.Count
        Cells(rcct, ccct).Select
        alk1 = returncolumnnumber("Margin% on group 1", "Familyname", 5)
        colchar1 = ColLtr(ccct + 1)
        Range(colchar1 & 5).Value = "Margin% on group 1"
   ccct = ActiveWorkbook.Worksheets("Margin% on group 1").UsedRange.Columns.Count
  Vllookcopypastefromotherbook 6, ccct, ThisWorkbook.Name, pn, "Margin% on group 1", shtnname, "Familyname", "Family Name", "Margin% on group 1", "Margin% on group 1", 5, 0
  Range(colchar1 & "6:" & colchar1 & rcct).Select
 Selection.Style = "Comma"
 Call Macro3
  Range(colchar1 & 4).Select
    Selection.Value = startdt1
     startdt1 = Format(startdt1, "dd-mmm-yy")
  Columns(colchar1).ColumnWidth = 10
  
  
  Workbooks(obj1.Name).Close savechanges:=False
  End If
  
Else


End If
startdt1 = DateSerial(Year(startdt1), Month(startdt1), Day(startdt1) + 1)
Next




''''''''''''''''''''''''''''''''''save''''''''''''''''''''''''''''''''''''''
Workbooks.Add
pn = ActiveWorkbook.Name
Workbooks(pn).Activate

ThisWorkbook.Worksheets("Summary").Activate
Sheets(Array("Summary", "Loan", "JM approved", "Group1", "Margin%", "Margin% on group 1")).Select
Sheets(Array("Summary", "Loan", "JM approved", "Group1", "Margin%", "Margin% on group 1")).Copy Before:=Workbooks(pn).Sheets(1)
 For Each Sheet In ActiveWorkbook.Sheets
        If Sheet.Name = "Sheet1" Or Sheet.Name = "Sheet2" Or Sheet.Name = "Sheet3" Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        Else
        
        End If
Next Sheet
Worksheets("Summary").Activate
Cells(1, 1).Select
 
Set fso = CreateObject("Scripting.FileSystemobject")
 spath = ThisWorkbook.Path & "\report"
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\report\" & Year(Now)
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\report\" & Year(Now) & "\" & MonthName(Month(Now))
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\report\" & Year(Now) & "\" & MonthName(Month(Now)) & "\" & Day(Now)
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 savereport = spath
 
        Set fso = CreateObject("Scripting.filesystemobject")
             If fso.FolderExists(savereport) Then
                 New_file_name = savereport & "\" & " MTFReport_" & Format(Now, "ddmmyyyy_") & Format(strtdt, "dd-mmm-yyyyTo") & Format(enddt, "dd-mmm-yyyy") & ".xlsx"
                 Workbooks(pn).SaveAs Filename:=New_file_name
                 fname = "MTFReport_" & Format(Now, "ddmmyyyy_") & Format(strtdt, "dd-mmm-yyTo") & Format(enddt, "dd-mmm-yyyy")
             Else
                 fso.createfolder (savereport)
                 
                 New_file_name = savereport & "\" & "MTFReport_" & Format(Now, "ddmmyyyy_") & Format(strtdt, "dd-mmm-yyyyTo") & Format(enddt, "dd-mmm-yyyy") & ".xlsx"
                 Workbooks(pn).SaveAs Filename:=New_file_name
                 fname = "MTFReport_" & Format(Now, "ddmmyyyy_") & Format(strtdt, "dd-mmm-yyyyTo") & Format(enddt, "dd-mmm-yyyy")
            End If
    
    ActiveWorkbook.Close savechanges = False

ThisWorkbook.Worksheets("Home").Activate
End Sub

Sub shtadd(shtnname)
            exists = False
            For i = 1 To ThisWorkbook.Worksheets.Count
                    If ThisWorkbook.Worksheets(i).Name = shtnname Then
                        exists = True
                    End If
            Next i
            If Not exists Then
                ThisWorkbook.Worksheets.Add(After:=Worksheets("HOME")).Name = shtnname
            Else
                cler (shtnname)
            End If
End Sub
Function cler(SheetName)
        ThisWorkbook.Worksheets(SheetName).Activate
            Cells.Select
            Cells.Clear
            Cells.Select
            Cells.Delete
            Cells(1, 1).Select
End Function
Function returncolumnnumber(parasheet, columnname, iRow)
Sheets(parasheet).Activate
tolst = Sheets(parasheet).UsedRange.Rows.Columns.Count + 5
checkstatus = ""
For i = 1 To tolst
      findval = Sheets(parasheet).Cells(iRow, i).Value
   If Trim(UCase(columnname)) = Trim(UCase(findval)) Then
       checkstatus = "found"
       Exit For
    End If
Next
If checkstatus = "" Then
  returncolumnnumber = 0
  Exit Function
End If
returncolumnnumber = i
End Function
Function FolderExists(strFolderPath)
    On Error Resume Next
    FolderExists = (GetAttr(strFolderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0

  
End Function
Function GetFullFileName(strfilepath, strFileNamePartial)

Dim objFS As Variant
Dim objFolder As Variant
Dim objFile As Variant
Dim intLengthOfPartialName As Integer
Dim strfilenamefull As String

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFS.getfolder(strfilepath)

'work out how long the partial file name is
intLengthOfPartialName = Len(strFileNamePartial)

For Each objFile In objFolder.Files

'Test to see if the file matches the partial file name
If Left(objFile.Name, intLengthOfPartialName) = strFileNamePartial Then

'get the full file name
strfilenamefull = objFile.Name
Exit For

Else

End If

Next objFile

'Return the full file name as the function's value
GetFullFileName = strfilenamefull

End Function
Function ColLtr(icol)
If icol > 0 And icol <= Columns.Count Then
   ColLtr = Evaluate("substitute(address(1, " & icol & ", 4), ""1"", """")")
End If

End Function
Function Vllookcopypastefromotherbook(ivrow, ivcol, lookupbook, lookinbook, lookupsheet, lookinsheet, lookupcolnm, lookincolnm, getcol, putcol, iRow, err)
    inseclocde = returncolumnnumber(lookupsheet, lookupcolnm, 5)
    colchar = ColLtr(inseclocde) & ivrow
    insemo = returncolumnnumber(lookupsheet, putcol, iRow)
     plstrow = Sheets(lookupsheet).Cells(iRow, inseclocde).CurrentRegion.Rows.Count
     plstrow1 = plstrow + 4
     pasterng1 = ColLtr(ivcol) & ivrow
    pasterng2 = ColLtr(ivcol) & plstrow1
    
    Workbooks(lookinbook).Worksheets(lookinsheet).Activate
    idetclocde = returncolumnnumber(lookinsheet, lookincolnm, 6)
    idetmo = returncolumnnumber(lookinsheet, getcol, 6)
    ilastrow = Sheets(lookinsheet).UsedRange.Rows.Count
    Rng1 = Sheets(lookinsheet).Cells(1, idetclocde).Address
    
    rng2 = Sheets(lookinsheet).Cells(ilastrow, idetmo).Address
    search_table_range = Sheets(lookinsheet).Range(Rng1 & ":" & rng2).Address
    colmn = idetmo - idetclocde + 1
    'Sheets(lookupsheet).Cells(2, insemo).Value = "=IFERROR(VLOOKUP(" & colchar & ",'" & lookinsheet & "'!" & search_table_range & "," & colmn & ",0)," & err & ")"
    Workbooks(lookupbook).Activate
    If err = 0 Then
        Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & ",'[" & lookinbook & "]" & lookinsheet & "'!" & search_table_range & " ," & colmn & ",0)," & err & ")"
    Else

      Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & ",'[" & lookinbook & "]" & lookinsheet & "'!" & search_table_range & " ," & colmn & ",0)," & Chr(34) & err & Chr(34) & ")"
    End If
 'Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & "," & lookinsheet & "!" & search_table_range & " ," & colmn & ",0)," & Chr(34) & err & Chr(34) & ")"
    Sheets(lookupsheet).Cells(ivrow, ivcol).Copy
    Sheets(lookupsheet).Activate
    Sheets(lookupsheet).Range(pasterng1 & ":" & pasterng2).Select
    ActiveSheet.Paste
    Sheets(lookupsheet).Range(pasterng1 & ":" & pasterng2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets(lookupsheet).Range(pasterng1).Select

End Function
Sub Macro3()
''
    Range("A5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Color = -4165632
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 12611584
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Columns("A:A").ColumnWidth = 45.29
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End Sub
Function Vllookcopypastefromotherbook2(ivrow, ivcol, lookupbook, lookinbook, lookupsheet, lookinsheet, lookupcolnm, lookincolnm, getcol, putcol, iRow, err)
    inseclocde = returncolumnnumber(lookupsheet, lookupcolnm, 6)
    colchar = ColLtr(inseclocde) & ivrow
    insemo = returncolumnnumber(lookupsheet, putcol, iRow)
     plstrow = Sheets(lookupsheet).Cells(iRow, inseclocde).CurrentRegion.Rows.Count
     plstrow1 = plstrow + 5
     pasterng1 = ColLtr(ivcol) & ivrow
    pasterng2 = ColLtr(ivcol) & plstrow1 - 1
    
    Workbooks(lookinbook).Worksheets(lookinsheet).Activate
    idetclocde = returncolumnnumber(lookinsheet, lookincolnm, 1)
    idetmo = returncolumnnumber(lookinsheet, getcol, 1)
    ilastrow = Sheets(lookinsheet).UsedRange.Rows.Count
    Rng1 = Sheets(lookinsheet).Cells(1, idetclocde).Address
    
    rng2 = Sheets(lookinsheet).Cells(ilastrow, idetmo).Address
    search_table_range = Sheets(lookinsheet).Range(Rng1 & ":" & rng2).Address
    colmn = idetmo - idetclocde + 1
    'Sheets(lookupsheet).Cells(2, insemo).Value = "=IFERROR(VLOOKUP(" & colchar & ",'" & lookinsheet & "'!" & search_table_range & "," & colmn & ",0)," & err & ")"
    Workbooks(lookupbook).Activate
    If err = 0 Then
        Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & ",'[" & lookinbook & "]" & lookinsheet & "'!" & search_table_range & " ," & colmn & ",0)," & err & ")"
    Else

      Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & ",'[" & lookinbook & "]" & lookinsheet & "'!" & search_table_range & " ," & colmn & ",0)," & Chr(34) & err & Chr(34) & ")"
    End If
 'Sheets(lookupsheet).Cells(ivrow, ivcol).Value = "=IFERROR(VLOOKUP(" & colchar & "," & lookinsheet & "!" & search_table_range & " ," & colmn & ",0)," & Chr(34) & err & Chr(34) & ")"
    Sheets(lookupsheet).Cells(ivrow, ivcol).Copy
    Sheets(lookupsheet).Activate
    Sheets(lookupsheet).Range(pasterng1 & ":" & pasterng2).Select
    ActiveSheet.Paste
    Sheets(lookupsheet).Range(pasterng1 & ":" & pasterng2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets(lookupsheet).Range(pasterng1).Select

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''2
Public familyname As Variant
    
    
    Sub graph()
    ThisWorkbook.Worksheets("HOME").Activate
    startdt1 = Range("G2").Value
    
    enddt = Range("G3").Value
    shtadd ("Graph")
    familyname = Sheet1.ComboBox1.Value
    Range("A3").Select
    Selection.Value = familyname
    Call Macro1
    Range("A4").Value = "Date"
    Range("B4").Value = "Loan in(crs.)"
    Range("C4").Value = "JM approved in(crs.)"
    Range("D4").Value = "Group1 in(crs.)"
    Range("E4").Value = "Margin % on Approved stock"
    Range("F4").Value = "Margin% on group 1"
    
    Range("G4").Value = "Margin%"
    Call Macro4
    ''''''''''''''''''''''''''''''''''''''''''''''
    path1 = ThisWorkbook.Path & "\" & "report" & "\"
    
    yr = Format(Now, "yyyy")
    mon = Format(Now, "mmmm")
    dt = Format(Now, "d")
    fldrpt = path1 & "\" & yr & "\" & mon & "\" & dt
    'ActiveWorkbook.FollowHyperlink Address:=fldrpt
    
    
    Dim SelectedFiles As Object
    Set SelectedFiles = Application.FileDialog(msoFileDialogFilePicker)
    SelectedFiles.Show
    
    If SelectedFiles.SelectedItems.Count <> 0 Then
    'here is the code which will run for all files selected
    Dim fileOne
    Dim Wbk As Workbook
    For Each fileOne In SelectedFiles.SelectedItems
     Set Wbk = Workbooks.Open(fileOne)
     
    Next
    
    Else
    MsgBox "No file was selected...", vbOKOnly + vbCritical, "Error!"
    err.Clear
    End If
    
    
    ''''''''''''''''''''''''''''''''''''"Date"'''''''''''''''''''''
    
    Workbooks(Wbk.Name).Worksheets("Loan").Activate
    ccct = ActiveWorkbook.Worksheets("Loan").UsedRange.Columns.Count
    colchar1 = ColLtr(ccct)
    Range("B4:" & colchar1 & 4).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Graph").Activate
    Range("A5").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=True
    Application.CutCopyMode = False
    Columns("A:A").EntireColumn.AutoFit
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = 0
    .PatternTintAndShade = 0
    End With
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "[$-409]d-mmm-yy;@"
    '''''''''''''''''''''''''Loan''''''''''''''''''''''
    Workbooks(Wbk.Name).Worksheets("Loan").Activate
    rcct = ActiveWorkbook.Worksheets("Loan").UsedRange.Rows.Count
    ccct = ActiveWorkbook.Worksheets("Loan").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    ActiveSheet.Range("A5:" & celv).AutoFilter Field:=1, Criteria1:=familyname, Operator:=xlFilterValues
    Range("B6:" & celv).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Graph").Activate
    Range("B5").Select
    rcct = Sheets("Graph").UsedRange.Rows.Count
    ccct = Sheets("Graph").UsedRange.Columns.Count
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=True
    Application.CutCopyMode = False
    Workbooks(Wbk.Name).Worksheets("Loan").Activate
    Selection.AutoFilter
    '''''''''''''''''''''''''''''''jm approved''''''''''''''''''''''''''''
    
    Workbooks(Wbk.Name).Worksheets("JM approved").Activate
    rcct = ActiveWorkbook.Worksheets("JM approved").UsedRange.Rows.Count
    ccct = ActiveWorkbook.Worksheets("JM approved").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    ActiveSheet.Range("A5:" & celv).AutoFilter Field:=1, Criteria1:=familyname, Operator:=xlFilterValues
    Range("B6:" & celv).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Graph").Activate
    Range("C5").Select
    rcct = Sheets("Graph").UsedRange.Rows.Count
    ccct = Sheets("Graph").UsedRange.Columns.Count
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=True
    Application.CutCopyMode = False
    
    Workbooks(Wbk.Name).Worksheets("JM approved").Activate
    
    Selection.AutoFilter
    ''''''''''''''''''''Group1''''''''''''''''''''''''''''''''''
    
    Workbooks(Wbk.Name).Worksheets("Group1").Activate
    
    rcct = ActiveWorkbook.Worksheets("Group1").UsedRange.Rows.Count
    ccct = ActiveWorkbook.Worksheets("Group1").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    ActiveSheet.Range("A5:" & celv).AutoFilter Field:=1, Criteria1:=familyname, Operator:=xlFilterValues
    Range("B6:" & celv).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Graph").Activate
    Range("D5").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=True
    Application.CutCopyMode = False
    rcct = Sheets("Graph").UsedRange.Rows.Count
    ccct = Sheets("Graph").UsedRange.Columns.Count
    Workbooks(Wbk.Name).Worksheets("Group1").Activate
    Selection.AutoFilter
    '''''''''''''''''''''''''''''''Margin%''''''''''''''''''''''
    Workbooks(Wbk.Name).Worksheets("Margin%").Activate
    
    rcct = ActiveWorkbook.Worksheets("Margin%").UsedRange.Rows.Count
    ccct = ActiveWorkbook.Worksheets("Margin%").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    ActiveSheet.Range("A5:" & celv).AutoFilter Field:=1, Criteria1:=familyname, Operator:=xlFilterValues
    Range("B6:" & celv).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Graph").Activate
    Range("E5").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=True
    Application.CutCopyMode = False
    rcct = Sheets("Graph").UsedRange.Rows.Count
    ccct = Sheets("Graph").UsedRange.Columns.Count
    Workbooks(Wbk.Name).Worksheets("Margin%").Activate
    Selection.AutoFilter
    ''''''''''''''''''''''''''''''''% group1'''''''''''''''''''''''''''''''''''
    
    Workbooks(Wbk.Name).Worksheets("Margin% on group 1").Activate
    
    rcct = ActiveWorkbook.Worksheets("Margin% on group 1").UsedRange.Rows.Count
    ccct = ActiveWorkbook.Worksheets("Margin% on group 1").UsedRange.Columns.Count
    Cells(rcct, ccct).Select
    celv = Selection.Address(False, False)
    ActiveSheet.Range("A5:" & celv).AutoFilter Field:=1, Criteria1:=familyname, Operator:=xlFilterValues
    Range("B6:" & celv).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    ThisWorkbook.Worksheets("Graph").Activate
    Range("F5").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=True
    Application.CutCopyMode = False
    rcct = Sheets("Graph").UsedRange.Rows.Count
    ccct = Sheets("Graph").UsedRange.Columns.Count
    Workbooks(Wbk.Name).Worksheets("Margin% on group 1").Activate
    Selection.AutoFilter
    ThisWorkbook.Worksheets("Graph").Activate
    Range("G5").Select
    Selection.Value = "50"
    Range("G5").Select
    
    Selection.AutoFill Destination:=Range("G5:G" & rcct + 2), Type:=xlFillDefault
    
    Range("G5:G" & rcct + 2).Select
    
    Selection.Style = "Comma"
    '''''''''''''''''''''''''''''GRAPH''''''''''''''''''''''''''''''
    ThisWorkbook.Worksheets("Graph").Activate
    Cells(1, 1).Select
    rcct = Sheets("Graph").UsedRange.Rows.Count
    ccct = Sheets("Graph").UsedRange.Columns.Count
    Cells(rcct + 2, ccct).Select
    celv = Selection.Address(False, False)
    
    Range("A4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Call board
    
    
    Range("A4:D" & rcct + 2).Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlLine
    ActiveChart.SetSourceData Source:=Range("A4:D" & rcct + 2)
    ActiveChart.ChartArea.Select
    
    ActiveSheet.Shapes(1).ScaleWidth 1.2427084427, msoFalse, _
    msoScaleFromTopLeft
    ActiveSheet.Shapes(1).ScaleHeight 1.4756944444, msoFalse, _
    msoScaleFromTopLeft
    ActiveSheet.Shapes(1).IncrementLeft 13.5
    ActiveSheet.Shapes(1).IncrementTop 39.75
    ActiveSheet.Shapes(1).IncrementLeft -9
    ActiveSheet.Shapes(1).IncrementTop -13.5
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes(1).IncrementLeft 44.25
    ActiveSheet.Shapes(1).IncrementTop -181.5
    
    ActiveSheet.ChartObjects(1).Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).CategoryType = xlCategoryScale
    ActiveChart.ChartArea.Select
    ActiveChart.Legend.Select
    ActiveChart.ApplyLayout (3)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "in crs."
    Selection.Format.TextFrame2.TextRange.Characters.Text = "in crs."
    With Selection.Format.TextFrame2.TextRange.Characters(1, 7).ParagraphFormat
    .TextDirection = msoTextDirectionLeftToRight
    .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 2).Font
    .BaselineOffset = 0
    .Bold = msoTrue
    .NameComplexScript = "+mn-cs"
    .NameFarEast = "+mn-ea"
    .Fill.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(0, 0, 0)
    .Fill.Transparency = 0
    .Fill.Solid
    .Size = 10
    .Italic = msoFalse
    .Kerning = 12
    .Name = "+mn-lt"
    .UnderlineStyle = msoNoUnderline
    .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(3, 1).Font
    .BaselineOffset = 0
    .Bold = msoTrue
    .NameComplexScript = "+mn-cs"
    .NameFarEast = "+mn-ea"
    .Fill.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(0, 0, 0)
    .Fill.Transparency = 0
    .Fill.Solid
    .Size = 10
    .Italic = msoFalse
    .Kerning = 12
    .Name = "+mn-lt"
    .UnderlineStyle = msoNoUnderline
    .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(4, 4).Font
    .BaselineOffset = 0
    .Bold = msoTrue
    .NameComplexScript = "+mn-cs"
    .NameFarEast = "+mn-ea"
    .Fill.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(0, 0, 0)
    .Fill.Transparency = 0
    .Fill.Solid
    .Size = 10
    .Italic = msoFalse
    .Kerning = 12
    .Name = "+mn-lt"
    .UnderlineStyle = msoNoUnderline
    .Strike = msoNoStrike
    End With
    
    ActiveChart.ChartTitle.Select
    Selection.Caption = familyname
    '''''''''''''''''''''''''''''''''''graph2'''''''''''''''''
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes(1).ScaleWidth 0.9723385565, msoFalse, _
    msoScaleFromBottomRight
    ActiveSheet.Shapes(1).ScaleHeight 0.9458823529, msoFalse, _
    msoScaleFromBottomRight
    ActiveSheet.Shapes(1).ScaleWidth 1.0563126286, msoFalse, _
    msoScaleFromTopLeft
    Range("A4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A" & 4 & ":" & "A" & rcct + 2 & "," & "E" & 4 & ":" & "G" & rcct + 2).Select
    
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlLine
    ActiveChart.SetSourceData Source:=Range("A" & 4 & ":" & "A" & rcct + 2 & "," & "E" & 4 & ":" & "G" & rcct + 2)
    
    ActiveSheet.Shapes(2).IncrementLeft 85.5
    ActiveSheet.Shapes(2).IncrementTop 171
    ActiveWindow.ScrollRow = 2
    ActiveChart.Legend.Select
    ActiveChart.ApplyLayout (3)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleVertical)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "in %"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "in %"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).ParagraphFormat
    .TextDirection = msoTextDirectionLeftToRight
    .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).Font
    .BaselineOffset = 0
    .Bold = msoTrue
    .NameComplexScript = "+mn-cs"
    .NameFarEast = "+mn-ea"
    .Fill.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(0, 0, 0)
    .Fill.Transparency = 0
    .Fill.Solid
    .Size = 10
    .Italic = msoFalse
    .Kerning = 12
    .Name = "+mn-lt"
    .UnderlineStyle = msoNoUnderline
    .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes(2).IncrementLeft -1.5
    ActiveChart.ChartTitle.Select
    Application.CutCopyMode = False
    Selection.Caption = familyname & "(Margin%)"
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes(2).ScaleWidth 1.1468751094, msoFalse, _
    msoScaleFromTopLeft
    ActiveSheet.Shapes(2).IncrementLeft -23.25
    ActiveSheet.Shapes(2).IncrementTop -3
    ActiveSheet.Shapes(2).ScaleWidth 1.0563126286, msoFalse, _
    msoScaleFromTopLeft
    ActiveSheet.Shapes(2).IncrementTop 2.812519685
    ActiveSheet.ChartObjects(2).Activate
    ActiveSheet.Shapes(2).ScaleWidth 1.0599912132, msoFalse, _
    msoScaleFromTopLeft
    ActiveSheet.ChartObjects(2).Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).CategoryType = xlCategoryScale
    Cells.Select
    Range("S28").Activate
    ActiveWindow.Zoom = 80
    Cells.EntireColumn.AutoFit
    
    
    
    
    Workbooks(Wbk.Name).Close savechanges:=False
    
    
    
    ''''''''''''''''''''''''''SAVE''''''''''''''''''''''''''''''''''''''''''''''''''''
    Workbooks.Add
    pn = ActiveWorkbook.Name
    Workbooks(pn).Activate
    
    ThisWorkbook.Worksheets("Graph").Activate
    Sheets("Graph").Select
    
    Sheets("Graph").Copy Before:=Workbooks(pn).Sheets(1)
    For Each Sheet In ActiveWorkbook.Sheets
    If Sheet.Name = "Sheet1" Or Sheet.Name = "Sheet2" Or Sheet.Name = "Sheet3" Then
    Application.DisplayAlerts = False
     Sheet.Delete
     Application.DisplayAlerts = True
    Else
    
    End If
    Next Sheet
    Worksheets("Graph").Activate
    Cells(1, 1).Select
    Set fso = CreateObject("Scripting.FileSystemobject")
    spath = ThisWorkbook.Path & "\Graphreport"
    If Not fso.FolderExists(spath) Then fso.createfolder (spath)
    spath = ThisWorkbook.Path & "\Graphreport\" & Year(Now)
    If Not fso.FolderExists(spath) Then fso.createfolder (spath)
    spath = ThisWorkbook.Path & "\Graphreport\" & Year(Now) & "\" & MonthName(Month(Now))
    If Not fso.FolderExists(spath) Then fso.createfolder (spath)
    spath = ThisWorkbook.Path & "\Graphreport\" & Year(Now) & "\" & MonthName(Month(Now)) & "\" & Day(Now)
    If Not fso.FolderExists(spath) Then fso.createfolder (spath)
    savereport = spath
    
    Set fso = CreateObject("Scripting.filesystemobject")
      If fso.FolderExists(savereport) Then
          New_file_name = savereport & "\" & familyname & Format(Now, "ddmmyyyy_") & Format(startdt1, "dd-mmm-yyyyTo") & Format(enddt, "dd-mmm-yyyy") & ".xlsx"
          Workbooks(pn).SaveAs Filename:=New_file_name
          fname = familyname & Format(Now, "ddmmyyyy_") & Format(startdt1, "dd-mmm-yyTo") & Format(enddt, "dd-mmm-yyyy")
      Else
          fso.createfolder (savereport)
          
          New_file_name = savereport & "\" & familyname & Format(Now, "ddmmyyyy_") & Format(startdt1, "dd-mmm-yyyyTo") & Format(enddt, "dd-mmm-yyyy") & ".xlsx"
          Workbooks(pn).SaveAs Filename:=New_file_name
          fname = familyname & Format(Now, "ddmmyyyy_") & Format(startdt1, "dd-mmm-yyyyTo") & Format(enddt, "dd-mmm-yyyy")
     End If
    ThisWorkbook.Worksheets("Home").Activate
    End Sub
    
    
    Sub board()
    '
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
    End With
    End Sub
    '''''''''''''''''''''''''''''''ravirajmg13'''''''''''''''''''

Public clientcode As Variant
    
 Sub fltr()
    clientcode = Sheet1.ComboBox1.Value
    ThisWorkbook.Worksheets("Summary1").Activate
    rc1 = Sheets("Summary1").UsedRange.Rows.Count + 1
     cc1 = Sheets("Summary1").UsedRange.Columns.Count
    Cells(rc1, cc1).Select
    cv = Selection.Address(False, False)
ActiveSheet.Range("A2:" & cv).AutoFilter Field:=2, Criteria1:=clientcode, _
        Operator:=xlAnd
 Range("A2:" & cv).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
   shtadd (clientcode)
   Sheets(clientcode).Activate
   Range("A1").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
  ThisWorkbook.Worksheets("Summary1").Activate
  Rows(2).Select
  Selection.AutoFilter
  
''''''''''''''''''''''''''''''''''''''''''''''''save repart'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Workbooks.Add
pn = ActiveWorkbook.Name
Workbooks(pn).Activate

ThisWorkbook.Worksheets(clientcode).Activate
Sheets(clientcode).Select
Sheets(clientcode).Copy Before:=Workbooks(pn).Sheets(1)
 For Each Sheet In ActiveWorkbook.Sheets
        If Sheet.Name = "Sheet1" Or Sheet.Name = "Sheet2" Or Sheet.Name = "Sheet3" Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        Else
        
        End If
Next Sheet
Worksheets(clientcode).Activate
Cells(1, 1).Select
 
Set fso = CreateObject("Scripting.FileSystemobject")
 spath = ThisWorkbook.Path & "\ClientReport"
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\ClientReport\" & Year(Now)
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\ClientReport\" & Year(Now) & "\" & MonthName(Month(Now))
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 spath = ThisWorkbook.Path & "\ClientReport\" & Year(Now) & "\" & MonthName(Month(Now)) & "\" & Day(Now)
 If Not fso.FolderExists(spath) Then fso.createfolder (spath)
 savereport = spath
     Set fso = CreateObject("Scripting.filesystemobject")
             If fso.FolderExists(savereport) Then
                 New_file_name = savereport & "\" & clientcode & "_" & "Report_" & Format(Now, "ddmmyyyy") & ".xlsx"
                 Workbooks(pn).SaveAs Filename:=New_file_name
                 fname = clientcode & "_" & "Report_" & Format(Now, "ddmmyyyy")
             Else
                 fso.createfolder (savereport)
                 
                 New_file_name = savereport & "\" & clientcode & "_" & "Report_" & Format(Now, "ddmmyyyy") & ".xlsx"
                 Workbooks(pn).SaveAs Filename:=New_file_name
                 fname = clientcode & "_" & "Report_" & Format(Now, "ddmmyyyy")
            End If
    
    ActiveWorkbook.Close savechanges = False
    ThisWorkbook.Worksheets("Home").Activate
    
    
     ThisWorkbook.Worksheets("Home").Activate
    For Each sht In ThisWorkbook.Sheets
    If sht.Name = "Home" Or sht.Name = "Sheet1" Or sht.Name = "Sheet2" Or sht.Name = "Sheet3" Or sht.Name = "Working" Then
    Else
    sht.Activate
    Application.DisplayAlerts = False
    sht.Delete
    Application.DisplayAlerts = True
    End If
    Next
   ThisWorkbook.Worksheets("Home").Activate
    
         
End Sub
Private Sub Workbook_Open()
' Sheets(1).Activate
 Sheet1.ComboBox1.Clear

'On Error Resume Next
Set cn = CreateObject("ADODB.Connection")
With cn
 .Provider = "Microsoft.Jet.OLEDB.4.0"
  .ConnectionString = "Data Source=" & ActiveWorkbook.FullName & _
";Extended Properties=""Excel 8.0;HDR=Yes;"""
.Open
End With

strQuery = "SELECT distinct clientcode as Clientcode FROM [Working$]"
Set rs5 = cn.Execute(strQuery)
Client = 0

  Sheet1.ComboBox1.Clear
Do While Not rs5.EOF
    clientcode = rs5("Clientcode")
    If IsNull(clientcode) Then
    
    Else
       Sheet1.ComboBox1.AddItem (clientcode)
    End If
    rs5.movenext
Loop
'On Error GoTo 0
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''""omnysys"""""""""''''''''''''''''''''''''''''''''
Sub Auto_Open()
  
 Dim time1, time2

time1 = Now
time2 = Now + TimeValue("0:01:00")
    Do Until time1 >= time2
        DoEvents
        time1 = Now()
    Loop
  Columns("I:M").Select
   Selection.Clear
    
'    Worksheets("Sheet2").Activate
'    Range("C2").Activate
'    ActiveCell.FormulaR1C1 = J
'     J = J + 1
'     ActiveCell.FormulaR1C1 = J

 J = Cells(1, 15).Value
J = J + 1

Cells(1, 15).Select
ActiveCell.FormulaR1C1 = "=" & J
shtadd ("MRL")
shtadd ("NRML")

Sheets("Sheet1").Activate
    rc = Sheets("Sheet1").UsedRange.Rows.Count
    cc = Sheets("Sheet1").UsedRange.Columns.Count
    alk1 = 10
    colchar1 = ColLtr(9)
    colchar2 = ColLtr(alk1)
    colchar7 = ColLtr(alk1 + 1)
    alk3 = returncolumnnumber("Sheet1", "High", 1)
    alk4 = returncolumnnumber("Sheet1", "Low", 1)
    alk5 = returncolumnnumber("Sheet1", "LTP", 1)
     colchar4 = ColLtr(alk3) & 2
     colchar5 = ColLtr(alk4) & 2
     colchar6 = ColLtr(alk5) & 2
    colchar8 = ColLtr(alk1 + 2)
    colchar9 = ColLtr(alk1 + 3)

    Range(colchar1 & 1).Select
    ActiveCell.FormulaR1C1 = "Days Change"
    Range(colchar1 & 2).Select
    ActiveCell.Formula = "=(" & colchar4 & "-" & colchar5 & ")" & "/" & colchar5 & "*" & 100
    Selection.AutoFill Destination:=Range(colchar1 & "2:" & colchar1 & rc)
    Range(colchar1 & "2:" & colchar1 & rc).Select
    Selection.Copy
    Range(colchar1 & 2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
     Range(colchar1 & "2:" & colchar1 & rc).Select
    Selection.NumberFormat = "0.00"
     Range(colchar2 & 1).Select
    ActiveCell.FormulaR1C1 = "LTP High"
     Range(colchar2 & 2).Select
    ActiveCell.Formula = "=" & colchar4 & "-" & colchar6
    Selection.AutoFill Destination:=Range(colchar2 & "2:" & colchar2 & rc)
    Range(colchar2 & "2:" & colchar2 & rc).Select
    Selection.Copy
    Range(colchar2 & 2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range(colchar7 & 1).Select
    ActiveCell.Formula = "LTP Low"
    Range(colchar7 & 2).Select
    ActiveCell.Formula = "=" & colchar6 & "-" & colchar5
    Selection.AutoFill Destination:=Range(colchar7 & "2:" & colchar7 & rc)
    Range(colchar7 & "2:" & colchar7 & rc).Select
    Selection.Copy
    Range(colchar7 & 2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range(colchar8 & 1).Select
    ActiveCell.FormulaR1C1 = "Max Change"
    Range(colchar8 & 2).Select
    ActiveCell.FormulaR1C1 = "=MAX(RC[-2],RC[-1])"
    Selection.AutoFill Destination:=Range(colchar8 & "2:" & colchar8 & rc)
    Range(colchar8 & "2:" & colchar8 & rc).Select
    Selection.Copy
     Range(colchar8 & 2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    Range(colchar9 & 2).Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]>0,""Increase"",""Decrease"")"
    Range(colchar9 & 2).Select
    Selection.AutoFill Destination:=Range(colchar9 & "2:" & colchar9 & rc)
    
     
    '''''''''''''''''''''''''''''''''''''''''''''''''
'    Columns("H:K").Select
'   Selection.EntireColumn.Hidden = True
 Range(colchar2 & "1:" & colchar8 & rc).Select
 Selection.EntireColumn.Hidden = True
 ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range(colchar1 & 2), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Range("A1").Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range(colchar1 & "2:" & colchar1 & rc) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A2:N" & rc)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  alk2 = returncolumnnumber("Sheet1", "Symbol", 1)
    colchar3 = ColLtr(alk2)
  If J = 1 Then
    For i = 2 To rc
    If Range(colchar9 & i).Value = "Increase" Then
    Range(colchar1 & i).Select
     With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
   Else
    Range(colchar1 & i).Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    End If
    
 If Range(colchar1 & i).Text = "#DIV/0!" Then
    Range(colchar1 & i).Value = 0
End If
rng1 = Range(colchar1 & i).Value
Rng2 = Range("F" & i).Value
'If Range(colchar1 & i).Value > 10 Or Range("F " & i).Value > 10 Then
If rng1 > 10 Or Rng2 > 10 Then
    'MsgBox "Days Change increased by more than 10% for " + Range(colchar3 & i).Value + ""
     If Rng2 > 10 Then
    Sheets("NRML").Activate
    Range("A1").Value = "Symbol"

    rc = Sheets("NRML").UsedRange.Rows.Count
    Range("A" & rc + 1).Select
    ActiveCell.Formula = Worksheets("Sheet1").Range(colchar3 & i).Value
    Sheets("Sheet1").Activate
    Sheets("MRL").Activate
    Range("A1").Value = "Symbol"
    rc = Sheets("MRL").UsedRange.Rows.Count
    Range("A" & rc + 1).Select
    ActiveCell.Formula = Worksheets("Sheet1").Range(colchar3 & i).Value
  Sheets("Sheet1").Activate
    Else
    Sheets("MRL").Activate
    Range("A1").Value = "Symbol"
    rc = Sheets("MRL").UsedRange.Rows.Count
    Range("A" & rc + 1).Select
    ActiveCell.Formula = Worksheets("Sheet1").Range(colchar3 & i).Value
  Sheets("Sheet1").Activate
    End If
    Sheets("Sheet1").Activate
    Application.Speech.Speak ("Days Change" + Range(colchar9 & i).Value + "" + "by more than 10% for " + Range(colchar3 & i).Value + "")
    Application.Speech.Speak ("Days Change" + Range(colchar9 & i).Value + "" + "by more than 10% for " + Range(colchar3 & i).Value + "")
    Application.Speech.Speak ("Days Change" + Range(colchar9 & i).Value + "" + "by more than 10% for " + Range(colchar3 & i).Value + "")
    
    rc1 = Sheets("Sheet1").Cells(2, 12).CurrentRegion.Rows.Count
    Range("N" & i).Activate
    ActiveCell.FormulaR1C1 = Worksheets("Sheet1").Range(colchar3 & i).Value
  Sheets("Sheet1").Activate
    If Range(colchar9 & i).Value = "Increase" Then
     Range("N" & i).Select
    Cells(i, 14).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Else
    Range("N" & i).Select
    Cells(i, 14).Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
       
    End If
 End If

 Next
    Else
     For i = 2 To rc
    If Range(colchar9 & i).Value = "Increase" Then
    Range(colchar1 & i).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
   Else
    Range(colchar1 & i).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    End If
 
            If Range(colchar1 & i).Text = "#DIV/0!" Then
            Range(colchar1 & i).Value = 0
            End If
rng1 = Range(colchar1 & i).Value
Rng2 = Range("F" & i).Value

If rng1 > 10 Or Rng2 > 10 Then

 Dim Value As String
 Dim Val As Range
 ThisWorkbook.Worksheets("Sheet1").Activate
  Value = Range(colchar3 & i).Value
  Sheets("Sheet1").Activate
  Columns(12).Select
  
  ''''''''''''''''''''''''''''''''''''''
 
Range("N2:N" & rc).Select
  Set Val = Selection.Find(What:=Value, After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
 
Sheets("Sheet1").Activate
Range("N" & i).Activate
    If Val Is Nothing Then
 If Rng2 > 10 Then
    Sheets("NRML").Activate
    Range("A1").Value = "Symbol"

    rc = Sheets("NRML").UsedRange.Rows.Count
    Range("A" & rc + 1).Select
    ActiveCell.Formula = Worksheets("Sheet1").Range(colchar3 & i).Value
    Sheets("Sheet1").Activate
    Sheets("MRL").Activate
    Range("A1").Value = "Symbol"
    rc = Sheets("MRL").UsedRange.Rows.Count
    Range("A" & rc + 1).Select
    ActiveCell.Formula = Worksheets("Sheet1").Range(colchar3 & i).Value
  Sheets("Sheet1").Activate
    Else
    Sheets("MRL").Activate
    Range("A1").Value = "Symbol"
    rc = Sheets("MRL").UsedRange.Rows.Count
    Range("A" & rc + 1).Select
    ActiveCell.Formula = Worksheets("Sheet1").Range(colchar3 & i).Value
  Sheets("Sheet1").Activate
    End If
Sheets("Sheet1").Activate
Range("N" & i).Activate
    Application.Speech.Speak ("Days Change" + Range(colchar9 & i).Value + "" + "by more than 10% for " + Range(colchar3 & i).Value + "")
    Application.Speech.Speak ("Days Change" + Range(colchar9 & i).Value + "" + "by more than 10% for " + Range(colchar3 & i).Value + "")
    Application.Speech.Speak ("Days Change" + Range(colchar9 & i).Value + "" + "by more than 10% for " + Range(colchar3 & i).Value + "")
    Worksheets("Sheet1").Activate

    Range("N" & i).Activate
   ActiveCell.FormulaR1C1 = Worksheets("Sheet1").Range(colchar3 & i).Value
    If Range(colchar9 & i).Value = "Increase" Then
     Range("N" & i).Select
    Cells(i, 14).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Else
    Range("N" & i).Select
    Cells(i, 14).Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
      
    End With
    
  

    End If
   
   
   Else
   
   End If
  End If
     
    Next
    End If
   Call RefreshDataEachHour
   
            
'        Range(colchar2 & "1:" & colchar8 & RC).Select
'        Selection.ClearContents
'        Application.Wait DateAdd("s", 10, Now)
'        Range(colchar2 & "1:" & colchar8 & RC).Select
'        Selection.EntireColumn.Hidden = False
'        Worksheets("Sheet1").Activate
'        Range(colchar1 & "1:" & colchar9 & RC).Select
'        Selection.Delete Shift:=xlUp
        Range("A1").Select
        
End Sub


'''''After Certain Time Interval Data Will Get Refreshed Automatically'''''

Public Sub RefreshDataEachHour()
        RunWhen = Now + TimeValue("00:00:30")
        Application.OnTime EarliestTime:=RunWhen, Procedure:="Auto_Open", Schedule:=True
End Sub


 Function returncolumnnumber(parasheet, columnname, iRow)
    Sheets(parasheet).Activate
    tolst = Sheets(parasheet).UsedRange.Rows.Columns.Count
    checkstatus = ""
    For i = 1 To tolst
    findval = Sheets(parasheet).Cells(iRow, i).Value
    If Trim(UCase(columnname)) = Trim(UCase(findval)) Then
    checkstatus = "found"
    Exit For
    End If
    Next
    If checkstatus = "" Then
    returncolumnnumber = 0
    Exit Function
    End If
    returncolumnnumber = i
    End Function
    
Function ColLtr(icol)
    If icol > 0 And icol <= Columns.Count Then
    ColLtr = Evaluate("substitute(address(1, " & icol & ", 4), ""1"", """")")
    End If
    
    End Function




 Sub shtadd(shtnname)
       exists = False
       For i = 1 To ThisWorkbook.Worksheets.Count
               If ThisWorkbook.Worksheets(i).Name = shtnname Then
                   exists = True
               End If
       Next i
       If Not exists Then
           ThisWorkbook.Worksheets.Add(After:=Worksheets("Sheet1")).Name = shtnname
       Else
           cler (shtnname)
       End If
    End Sub
    Function cler(SheetName)
    ThisWorkbook.Worksheets(SheetName).Activate
       Cells.Select
       Cells.Clear
       Cells.Select
       Cells.Delete
       Cells(1, 1).Select
    End Function

""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
Sub Clientwise()
'''''''''''''''''''''''''''

ThisWorkbook.Worksheets("HOME").Activate
    Range("F2").Select
    Path = Selection.Value
    Path1 = Path
cv = GetFullFileName(Path1, "RBI Ageing Summary Report")
    
'    If cv <> "" Then
'    Set obj1 = Workbooks.Open(Filename:=Path1 & cv)
'    shtname = ActiveSheet.Name
'    rc = ActiveWorkbook.Worksheets(shtname).UsedRange.Rows.Count
'    cc = ActiveWorkbook.Worksheets(shtname).UsedRange.Columns.Count
'    Cells(rc, cc).Select
'    cv = Selection.Address(False, False)
'    Range("A1:" & cv).Select
'    Selection.Copy
'''    ThisWorkbook.Worksheets("Sheet1").Activate
''    Range("A1").Select

Filename = Path1 & cv
Module_function.Copydatas Filename, "Clientwise", "Clientwise"

'''''''''''''''''''''''''''''''''''''''''''
ThisWorkbook.Sheets("Clientwise").Activate

Columns(3).Select
 Selection.Insert Shift:=xlToRight
  Selection.Insert Shift:=xlToRight
     Count = 0
     Range("C2").Select
     Selection.Value = "No.of Days As per RBI"
     Range("D2").Select
     Selection.Value = "No.of Days As per JM"
rcount = Sheets("Clientwise").UsedRange.Rows.Count
ccty = Sheets("Clientwise").UsedRange.Columns.Count
     
For i = 3 To rcount
'countval = Cells(i, 7).Value
Count = 0

    For j = 10 To ccty - 2
   ' Count = 0
rbi_val1 = Cells(i, j).Value
rbi_val2 = Cells(i, j + 2).Value
         
      If rbi_val1 > 0.5 Then  'And rbi_val2 >= 0.5'
        Cells(i, 3).Value = Count
         Exit For
           Else
                If rbi_val1 < 0.5 And rbi_val2 >= 0.5 Then
                Count = Count + 1
                Cells(i, 3).Value = Count
                Exit For
                End If
                    If rbi_val1 < 0.5 Then
                    Count = Count + 1
                        For k = j + 2 To ccty
                        Cells(i, k).Select
                        rbi_val3 = Cells(i, k).Value
                        rbi_val4 = Cells(i, k + 2).Value
                            If rbi_val3 < 0.5 And rbi_val4 >= 0.5 Then
                            Count = Count + 1
                            Cells(i, 3).Value = Count
                            Exit For
                            Else
                        If rbi_val3 < 0.5 Then
                        Count = Count + 1
                    k = k + 1
                    If k = ccty Then
                    Cells(i, 3).Value = Count
                    j = k
                Exit For
                End If
                Cells(i, 3).Value = Count
                End If
              
             End If
            Next
             
        End If
      End If
  Exit For

Next
    Next
    ''''''''''''''''''''(jm)'''''''''''''''''''''''
    For i = 3 To rcount
'countval = Cells(i, 7).Value
Count = 0

    For j = 13 To ccty - 2
   ' Count = 0
rbi_val1 = Cells(i, j).Value
rbi_val2 = Cells(i, j + 2).Value
         
      If rbi_val1 > 0.5 Then  'And rbi_val2 >= 0.5'
        Cells(i, 4).Value = Count
         Exit For
           Else
                If rbi_val1 < 0.5 And rbi_val2 >= 0.5 Then
                Count = Count + 1
                Cells(i, 4).Value = Count
                Exit For
                End If
                    If rbi_val1 < 0.5 Then
                    Count = Count + 1
                        For k = j + 2 To ccty
                        Cells(i, k).Select
                        rbi_val3 = Cells(i, k).Value
                        rbi_val4 = Cells(i, k + 2).Value
                            If rbi_val3 < 0.5 And rbi_val4 >= 0.5 Then
                            Count = Count + 1
                            Cells(i, 4).Value = Count
                            Exit For
                            Else
                    If rbi_val3 < 0.5 Then
                    Count = Count + 1
                    k = k + 1
                    If k = ccty Then
                    Cells(i, 4).Value = Count
                    j = k
                    Exit For
                    End If
                    Cells(i, 4).Value = Count
                    End If
                    
               End If
              Next
               
          End If
        End If
    Exit For

  Next
    Next
'''''''''''''''''''''''''''''
  Range("A3").Select
'  For Each sht In ThisWorkbook.Sheets
'       If sht.Name = "Home" Or sht.Name = "Sheet2" Or sht.Name = "Sheet3" Then
'
'        Else
'         sht.Activate
'        Application.DisplayAlerts = False
'        sht.Delete
'        Application.DisplayAlerts = True
'
'    End If
'    Next
    ThisWorkbook.Sheets("Home").Activate

    
    
    
End Sub
Function GetFullFileName(strfilepath, strFileNamePartial)
    
    Dim objFS As Variant
    Dim objFolder As Variant
    Dim objFile As Variant
    Dim intLengthOfPartialName As Integer
    Dim strfilenamefull As String
    
    Set objFS = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFS.getfolder(strfilepath)
    
    'work out how long the partial file name is
    intLengthOfPartialName = Len(strFileNamePartial)
    
    For Each objFile In objFolder.Files
    
    'Test to see if the file matches the partial file name
    If Left(objFile.Name, intLengthOfPartialName) = strFileNamePartial Then
    
    'get the full file name
    strfilenamefull = objFile.Name
    Exit For
    
    Else
    
    End If
    
    Next objFile
    
    'Return the full file name as the function's value
    GetFullFileName = strfilenamefull
    
    End Function





    
    
    
    
    
    
    
    
    
    
    






