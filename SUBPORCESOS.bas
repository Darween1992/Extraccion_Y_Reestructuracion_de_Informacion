Attribute VB_Name = "SUBPORCESOS"
Sub DATOS_5_1()
Dim celda_activa As Range
Dim MODULO As String
Dim GALPON As String
Dim semana As String
Dim CORRAL As String
Dim i As Integer

Dim aumento1 As Integer
Dim aumento2 As Integer
aumento1 = 1
aumento2 = 6
aumento3 = 5
Dim j As Integer


Cells(6, 5).Select



For j = 1 To 8


' inicio corral1


    MODULO = Range("C6").Value
    GALPON = Range("C7").Value
    CORRAL = Cells(6, aumento3).Value
    semana = Range("C8").Value
    
    


Set celda_activa = Range("D4")
celda_activa.Offset(5, aumento1).Select  'aumenta 7  declarar variable fuera del ciclo que aumente de 7 en 7



Range(ActiveCell, celda_activa.Offset(11, aumento2)).Select 'aumenta 7 declarar variable fuera del ciclo que aumnte de 7 em 7
Selection.Copy

Sheets("BBDD").Select
Range("B3").End(xlDown).Offset(1, 4).Select


Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'semana
Range("B3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = semana
ActiveCell.Offset(1, 0).Select

    Next i
    
'modulo

Range("C3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = MODULO
ActiveCell.Offset(1, 0).Select

    Next i
'GALPON

Range("D3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = GALPON
ActiveCell.Offset(1, 0).Select

    Next i

'CORRAL
Range("E3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = CORRAL
ActiveCell.Offset(1, 0).Select

    Next i

    
     aumento1 = aumento1 + 7
     
     aumento2 = aumento2 + 7
    
     aumento3 = aumento3 + 7
    
    Sheets("MODULO 5-8").Select
    


 Next j
 

End Sub

Sub DATOS_5_2()
Dim celda_activa As Range
Dim MODULO As String
Dim GALPON As String
Dim semana As String
Dim CORRAL As String
Dim i As Integer

Dim aumento1 As Integer
Dim aumento2 As Integer
aumento1 = 1
aumento2 = 6
aumento3 = 5
Dim j As Integer


Cells(18, 5).Select



For j = 1 To 8


' inicio corral1


    MODULO = Range("C6").Value
    GALPON = Range("C19").Value
    CORRAL = Cells(6, aumento3).Value
    semana = Range("C8").Value
    
    


Set celda_activa = Range("D4")
celda_activa.Offset(17, aumento1).Select  'aumenta 7  declarar variable fuera del ciclo que aumente de 7 en 7



Range(ActiveCell, celda_activa.Offset(23, aumento2)).Select 'aumenta 7 declarar variable fuera del ciclo que aumnte de 7 em 7
Selection.Copy

Sheets("BBDD").Select
Range("B3").End(xlDown).Offset(1, 4).Select


Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'semana
Range("B3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = semana
ActiveCell.Offset(1, 0).Select

    Next i
    
'modulo

Range("C3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = MODULO
ActiveCell.Offset(1, 0).Select

    Next i
'GALPON

Range("D3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = GALPON
ActiveCell.Offset(1, 0).Select

    Next i

'CORRAL
Range("E3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = CORRAL
ActiveCell.Offset(1, 0).Select

    Next i

    
     aumento1 = aumento1 + 7
     
     aumento2 = aumento2 + 7
    
     aumento3 = aumento3 + 7
    
    Sheets("MODULO 5-8").Select
    


 Next j
 

End Sub

Sub DATOS_5_3()
Dim celda_activa As Range
Dim MODULO As String
Dim GALPON As String
Dim semana As String
Dim CORRAL As String
Dim i As Integer

Dim aumento1 As Integer
Dim aumento2 As Integer
aumento1 = 1
aumento2 = 6
aumento3 = 5
Dim j As Integer


Cells(30, 5).Select



For j = 1 To 12


' inicio corral1


    MODULO = Range("C6").Value
    GALPON = Range("C31").Value
    CORRAL = Cells(6, aumento3).Value
    semana = Range("C8").Value
    
    


Set celda_activa = Range("D4")
celda_activa.Offset(29, aumento1).Select  'aumenta 7  declarar variable fuera del ciclo que aumente de 7 en 7



Range(ActiveCell, celda_activa.Offset(35, aumento2)).Select 'aumenta 7 declarar variable fuera del ciclo que aumnte de 7 em 7
Selection.Copy

Sheets("BBDD").Select
Range("B3").End(xlDown).Offset(1, 4).Select


Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'semana
Range("B3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = semana
ActiveCell.Offset(1, 0).Select

    Next i
    
'modulo

Range("C3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = MODULO
ActiveCell.Offset(1, 0).Select

    Next i
'GALPON

Range("D3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = GALPON
ActiveCell.Offset(1, 0).Select

    Next i

'CORRAL
Range("E3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = CORRAL
ActiveCell.Offset(1, 0).Select

    Next i

    
     aumento1 = aumento1 + 7
     
     aumento2 = aumento2 + 7
    
     aumento3 = aumento3 + 7
    
    Sheets("MODULO 5-8").Select
    


 Next j
 

End Sub

Sub DATOS_8_1()
Dim celda_activa As Range
Dim MODULO As String
Dim GALPON As String
Dim semana As String
Dim CORRAL As String
Dim i As Integer

Dim aumento1 As Integer
Dim aumento2 As Integer
aumento1 = 1
aumento2 = 6
aumento3 = 5
Dim j As Integer


Cells(42, 5).Select



For j = 1 To 8


' inicio corral1


    MODULO = Range("C42").Value
    GALPON = Range("C43").Value
    CORRAL = Cells(42, aumento3).Value
    semana = Range("C44").Value
    
    


Set celda_activa = Range("D4")
celda_activa.Offset(41, aumento1).Select  'aumenta 7  declarar variable fuera del ciclo que aumente de 7 en 7



Range(ActiveCell, celda_activa.Offset(47, aumento2)).Select 'aumenta 7 declarar variable fuera del ciclo que aumnte de 7 em 7
Selection.Copy

Sheets("BBDD").Select
Range("B3").End(xlDown).Offset(1, 4).Select


Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'semana
Range("B3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = semana
ActiveCell.Offset(1, 0).Select

    Next i
    
'modulo

Range("C3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = MODULO
ActiveCell.Offset(1, 0).Select

    Next i
'GALPON

Range("D3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = GALPON
ActiveCell.Offset(1, 0).Select

    Next i

'CORRAL
Range("E3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = CORRAL
ActiveCell.Offset(1, 0).Select

    Next i

    
     aumento1 = aumento1 + 7
     
     aumento2 = aumento2 + 7
    
     aumento3 = aumento3 + 7
    
    Sheets("MODULO 5-8").Select
    


 Next j
 

End Sub

Sub DATOS_8_2()
Dim celda_activa As Range
Dim MODULO As String
Dim GALPON As String
Dim semana As String
Dim CORRAL As String
Dim i As Integer

Dim aumento1 As Integer
Dim aumento2 As Integer
aumento1 = 1
aumento2 = 6
aumento3 = 5
Dim j As Integer


Cells(54, 5).Select



For j = 1 To 9


' inicio corral1


    MODULO = Range("C54").Value
    GALPON = Range("C55").Value
    CORRAL = Cells(54, aumento3).Value
    semana = Range("C44").Value
    
    


Set celda_activa = Range("D4")
celda_activa.Offset(53, aumento1).Select  'aumenta 7  declarar variable fuera del ciclo que aumente de 7 en 7



Range(ActiveCell, celda_activa.Offset(59, aumento2)).Select 'aumenta 7 declarar variable fuera del ciclo que aumnte de 7 em 7
Selection.Copy

Sheets("BBDD").Select
Range("B3").End(xlDown).Offset(1, 4).Select


Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'semana
Range("B3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = semana
ActiveCell.Offset(1, 0).Select

    Next i
    
'modulo

Range("C3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = MODULO
ActiveCell.Offset(1, 0).Select

    Next i
'GALPON

Range("D3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = GALPON
ActiveCell.Offset(1, 0).Select

    Next i

'CORRAL
Range("E3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = CORRAL
ActiveCell.Offset(1, 0).Select

    Next i

    
     aumento1 = aumento1 + 7
     
     aumento2 = aumento2 + 7
    
     aumento3 = aumento3 + 7
    
    Sheets("MODULO 5-8").Select
    


 Next j
 

End Sub

Sub DATOS_8_3()
Dim celda_activa As Range
Dim MODULO As String
Dim GALPON As String
Dim semana As String
Dim CORRAL As String
Dim i As Integer

Dim aumento1 As Integer
Dim aumento2 As Integer
aumento1 = 1
aumento2 = 6
aumento3 = 5
Dim j As Integer


Cells(66, 5).Select



For j = 1 To 7


' inicio corral1


    MODULO = Range("C66").Value
    GALPON = Range("C67").Value
    CORRAL = Cells(66, aumento3).Value
    semana = Range("C44").Value
    
    


Set celda_activa = Range("D4")
celda_activa.Offset(65, aumento1).Select  'aumenta 7  declarar variable fuera del ciclo que aumente de 7 en 7



Range(ActiveCell, celda_activa.Offset(71, aumento2)).Select 'aumenta 7 declarar variable fuera del ciclo que aumnte de 7 em 7
Selection.Copy

Sheets("BBDD").Select
Range("B3").End(xlDown).Offset(1, 4).Select


Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'semana
Range("B3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = semana
ActiveCell.Offset(1, 0).Select

    Next i
    
'modulo

Range("C3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = MODULO
ActiveCell.Offset(1, 0).Select

    Next i
'GALPON

Range("D3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = GALPON
ActiveCell.Offset(1, 0).Select

    Next i

'CORRAL
Range("E3").End(xlDown).Offset(1, 0).Select

For i = 1 To 7

ActiveCell.Value = CORRAL
ActiveCell.Offset(1, 0).Select

    Next i

    
     aumento1 = aumento1 + 7
     
     aumento2 = aumento2 + 7
    
     aumento3 = aumento3 + 7
    
    Sheets("MODULO 5-8").Select
    


 Next j
 

End Sub







