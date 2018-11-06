'Created by Mauricio Mani
'VBScript Excel
Sub Insertar()
'
' Insertar Macro
' Creado para insertar valores para diagrama de Gantt
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Actividades"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Responsable"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Inicio"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Fin"
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$D$2"), , xlYes).Name = _
        "Tabla1"
    Range("Tabla1[#All]").Select
    ActiveSheet.ListObjects("Tabla1").TableStyle = "TableStyleLight1"
    ActiveSheet.ShowDataForm
End Sub