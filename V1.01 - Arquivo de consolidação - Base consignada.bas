Attribute VB_Name = "Módulo1"
Sub Carregar()

    'Tipo Var
    Dim valor As String
    valor = MsgBox("Processar todos os dados atualizados?", vbOKCancel, "VALIDAÇÃO DE ATIVAÇÃO DE MACROS")
    If valor = 1 Then

    Sheets("CARREGAR").Select
    Range("E2").Select
    Selection.Copy
    Range("H2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B4").Select

    Application.ScreenUpdating = False

    Sheets("DADOS CARREGADOS").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("B2").Select
    
' Carregamento 53
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\53.xlsx"
    
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[53.xlsx]53'!C4)"
    Range("C4").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Windows("53.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("DADOS CARREGADOS").Select
    Range("B2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("53.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close
    
' Carregamento 54
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\54.xlsx"
    
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[54.xlsx]54'!C4)"
    Range("C5").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("54.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("DADOS CARREGADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("54.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close

' Carregamento 55
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\55.xlsx"
    
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[55.xlsx]55'!C4)"
    Range("C6").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("55.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("DADOS CARREGADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("55.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close

' Carregamento 67
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\67.xlsx"
    
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[67.xlsx]67'!C4)"
    Range("C7").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("67.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("DADOS CARREGADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("67.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close

' Carregamento 81
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\81.xlsx"
    
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[81.xlsx]81'!C4)"
    Range("C8").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("81.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("DADOS CARREGADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("81.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close

' Carregamento 82
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\82.xlsx"
    
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C9").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[82.xlsx]82'!C4)"
    Range("C9").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("82.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("DADOS CARREGADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("82.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close
    
' Carregamento 87
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\87.xlsx"
    
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[87.xlsx]87'!C4)"
    Range("C10").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("87.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("DADOS CARREGADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("87.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close
    
    Windows("Arquivo de consolidação - Base consignada.xlsm").Activate
    Sheets("DADOS CARREGADOS").Select
    Range("B3").Select
    Sheets("CARREGAR").Select
    Range("B4").Select
    ActiveWorkbook.Save
    
    Else
    End If
    
    Application.ScreenUpdating = True
    
End Sub


