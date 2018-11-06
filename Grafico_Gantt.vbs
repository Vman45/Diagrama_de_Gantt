'Created by Mauricio Mani
'VBScript Excel
'La idea es crear un Grafico de Gantt.
'Utilizaremos dos hojas principales, la primer hoja contendra los valores de un grafico de gantt, la segunda sera el grafico.
'La primer hoja contiene: Nivel, actividad (task), responsable (assigment), dia de inicio (start-end), dia de fin, numero de dias.
'Sin embargo, puede contener costos, si es critica la actividad, tipo de actividad, etc.

Sub Grafico_Gantt()
'La plantilla de la hoja 1 debera contener: actividad, responsable, dia de inicio del proyecto, fin del proyecto y numeo de dias por actividad.
'Primero crearemos la nueva hoja y la seleccionaremos
Dim Hojagant As Worksheet, datos As Worksheet, numact As Integer, numday As Integer, r1 As Range, valor As Integer, coin As Integer
Set datos = ActiveSheet
Set Hojagant = Worksheets.Add(After:=Sheets(Worksheets.Count))
'Utilizamos un TextBox para ingresar el titulo
Dim titulo As String
'Creamos variables de las hojas de nuestro archivo
Hojagant.Select
ActiveSheet.Name = "Gráfico de Gantt"
titulo = InputBox("Titulo del gráfico de Gantt", "Gráfico de Gantt", "Grafico de Gantt")
Range("B1").Value = titulo
Range("B1").Select
With Selection.Font
    .Name = "Corbel"
    .Size = 42
    .ThemeColor = xlThemeColorAccent3
    .TintAndShade = -0.249977111117893
End With
With Selection
    .HorizontalAlignment = xlLeft
End With
Selection.Font.Bold = True
'Estamos seleccionando los titulos y modificando de ser necesario la fuente.
Range("B3").Value = "Nivel"
Range("C3").Value = "Actividad"
Range("D3").Value = "Responsable"
Range("E3").Value = "Inicio"
Range("F3").Value = "Fin"
Range("G3").Value = "Dias"
Range("B3:G3").Select
With Selection.Font
    .Name = "Calibri"
    .Size = 11
    .Color = -14515628
    .Bold = True
End With
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
Columns("A:A").ColumnWidth = 1
datos.Select
Range("A2:D2").Select
Range(Selection, Selection.End(xlDown)).Select
numact = ((Selection.Count) / 4)
Selection.Copy
Hojagant.Select
Range("C4").Select
Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Columns("C:F").EntireColumn.AutoFit
nivel = 1
For i = 4 To (numact + 3)
    Cells(i, 2).Value = nivel
    nivel = nivel + 1
Next i
For i = 4 To (numact + 3)
    Cells(i, 7).Value = (DateDiff("d", Cells(i, 5), Cells(i, 6))) + 1
Next i
Cells(3, 9).Value = Cells(4, 5).Value
Range("I3").Select
Selection.DataSeries Rowcol:=xlRows, Type:=xlChronological, Date:=xlDay, _
        Step:=1, Stop:=Cells((numact + 3), "F").Value
Range(Selection, Selection.End(xlToRight)).Select
With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .Orientation = 90
        .IndentLevel = 0
        .ReadingOrder = xlContext
End With
Columns("H:ZZ").Select
Selection.ColumnWidth = 2
Rows(3).Select
Selection.RowHeight = 60
numday = (DateDiff("d", Cells(4, 5), Cells((numact + 3), 6))) + 1
i = 5
Do While i < ((2 * numact) + 3)
    Rows(i).Select
    Selection.Insert Shift:=xlDown
    i = i + 2
    Loop
i = 4
Do While i < ((2 * numact) + 3)
    For e = 2 To 7
        Range(Cells(i, e), Cells(i + 1, e)).Select
        With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
    End With
    Selection.Merge
    Next e
    i = i + 2
    Loop
e = 4
Do While e < (2 * (numact + 2))
    Cells(e, 8).Value = "E"
    e = e + 1
    Cells(e, 8).Value = "R"
    e = e + 1
Loop
Set r1 = Range(Cells(4, 9), Cells(4, (Cells(4, "G") + 8)))
    r1.Interior.Color = vbBlue
For i = 6 To (2 * numact + 3) Step 2
    coin = Application.Match(Cells(i, 5), Range(Cells(3, 9), Cells(3, (numday + 8))))
    valor = Cells(i, "G")
    Set r1 = Range(Cells(i, (8 + coin)), Cells(i, ((valor + 7) + coin)))
    r1.Interior.Color = vbBlue
Next i
Range(Cells(3, "B"), Cells(3, numday + 8)).Select
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
Range(Cells((2 * numact + 3), "B"), Cells((2 * numact + 3), numday + 8)).Select
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
ActiveWindow.DisplayGridlines = False
ActiveWindow.SmallScroll Down:=0
End Sub