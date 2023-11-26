Attribute VB_Name = "Módulo2"
Sub MiPrimeraMacro()
Attribute MiPrimeraMacro.VB_Description = "Mi primera macro con la gravadora"
Attribute MiPrimeraMacro.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' MiPrimeraMacro Macro
' Mi primera macro con la gravadora
'
' Acceso directo: Ctrl+Mayús+P
'
    ActiveCell.FormulaR1C1 = "Excel - Curso de Macros"
    Range("A1").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Range("D4").Select
End Sub
