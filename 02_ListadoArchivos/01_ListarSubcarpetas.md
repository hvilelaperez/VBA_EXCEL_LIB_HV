# Listar Subcarpetas

Lista todas las subcarpetas de una carpeta seleccionada recursivamente.

## Código VBA

```vba
Option Explicit

' ============================================================
' LISTAR SUBCARPETAS EN HOJA
' ============================================================
Public Sub ListarSubcarpetas()
    Dim rutaCarpeta As String
    Dim fso As Object
    Dim carpeta As Object
    Dim subCarpeta As Object
    Dim fila As Long
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar carpeta"
        If .Show <> -1 Then Exit Sub
        rutaCarpeta = .SelectedItems(1)
    End With
    
    Application.ScreenUpdating = False
    
    ' Limpiar y preparar hoja
    ActiveSheet.Cells.Clear
    ActiveSheet.Cells(1, 1).Value = "Nombre"
    ActiveSheet.Cells(1, 2).Value = "Ruta"
    ActiveSheet.Cells(1, 3).Value = "Fecha"
    ActiveSheet.Range("A1:C1").Font.Bold = True
    
    fila = 2
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpeta = fso.GetFolder(rutaCarpeta)
    
    Call RecorrerCarpetas(carpeta, fila)
    
    ActiveSheet.Columns("A:C").AutoFit
    Application.ScreenUpdating = True
    
    MsgBox "Subcarpetas encontradas: " & (fila - 2), vbInformation
End Sub


Private Sub RecorrerCarpetas(carpeta As Object, ByRef fila As Long)
    Dim subCarpeta As Object
    
    For Each subCarpeta In carpeta.SubFolders
        ActiveSheet.Cells(fila, 1).Value = subCarpeta.Name
        ActiveSheet.Cells(fila, 2).Value = subCarpeta.Path
        ActiveSheet.Cells(fila, 3).Value = subCarpeta.DateLastModified
        fila = fila + 1
        
        Call RecorrerCarpetas(subCarpeta, fila)
    Next subCarpeta
End Sub
```

---

*Módulo: 02_ListadoArchivos | Archivo: 01_ListarSubcarpetas.md*
