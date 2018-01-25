# Buscar Archivos Recursivo

Búsqueda profunda de archivos en todas las subcarpetas.

## Código VBA

```vba
Option Explicit

' ============================================================
' BÚSQUEDA RECURSIVA DE ARCHIVOS
' ============================================================
Public Sub BuscarArchivosRecursivo()
    Dim rutaCarpeta As String
    Dim fso As Object
    Dim carpeta As Object
    Dim contador As Long
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Carpeta de búsqueda"
        If .Show <> -1 Then Exit Sub
        rutaCarpeta = .SelectedItems(1)
    End With
    
    Application.ScreenUpdating = False
    
    ActiveSheet.Cells.Clear
    ActiveSheet.Cells(1, 1).Value = "Nombre"
    ActiveSheet.Cells(1, 2).Value = "Ruta Completa"
    ActiveSheet.Cells(1, 3).Value = "Tamaño (KB)"
    ActiveSheet.Range("A1:C1").Font.Bold = True
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpeta = fso.GetFolder(rutaCarpeta)
    contador = 0
    
    Call BuscarEnCarpeta(carpeta, contador)
    
    ActiveSheet.Columns("A:C").AutoFit
    Application.ScreenUpdating = True
    
    MsgBox "Archivos encontrados: " & contador, vbInformation
End Sub


Private Sub BuscarEnCarpeta(carpeta As Object, ByRef contador As Long)
    Dim archivo As Object
    Dim subCarpeta As Object
    
    For Each archivo In carpeta.Files
        contador = contador + 1
        ActiveSheet.Cells(contador + 1, 1).Value = archivo.Name
        ActiveSheet.Cells(contador + 1, 2).Value = archivo.Path
        ActiveSheet.Cells(contador + 1, 3).Value = archivo.Size \ 1024
    Next archivo
    
    For Each subCarpeta In carpeta.SubFolders
        Call BuscarEnCarpeta(subCarpeta, contador)
    Next subCarpeta
End Sub
```

---

*Módulo: 02_ListadoArchivos | Archivo: 04_BuscarArchivosRecursivo.md*
