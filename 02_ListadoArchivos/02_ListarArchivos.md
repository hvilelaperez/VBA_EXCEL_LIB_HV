# Listar Archivos

Lista archivos de una carpeta con diferentes opciones de filtrado.

## Código VBA

```vba
Option Explicit

' ============================================================
' LISTAR ARCHIVOS DE UNA CARPETA
' ============================================================
Public Sub ListarArchivos()
    Dim rutaCarpeta As String
    Dim nombreArchivo As String
    Dim fila As Long
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar carpeta"
        If .Show <> -1 Then Exit Sub
        rutaCarpeta = .SelectedItems(1)
    End With
    
    If Right(rutaCarpeta, 1) <> "\" Then rutaCarpeta = rutaCarpeta & "\"
    
    Application.ScreenUpdating = False
    
    ActiveSheet.Cells.Clear
    ActiveSheet.Cells(1, 1).Value = "Nombre"
    ActiveSheet.Cells(1, 2).Value = "Tamaño (KB)"
    ActiveSheet.Cells(1, 3).Value = "Fecha"
    ActiveSheet.Range("A1:C1").Font.Bold = True
    
    fila = 2
    nombreArchivo = Dir(rutaCarpeta & "*.*")
    
    Do While nombreArchivo <> ""
        ActiveSheet.Cells(fila, 1).Value = nombreArchivo
        ActiveSheet.Cells(fila, 2).Value = FileLen(rutaCarpeta & nombreArchivo) \ 1024
        ActiveSheet.Cells(fila, 3).Value = FileDateTime(rutaCarpeta & nombreArchivo)
        fila = fila + 1
        nombreArchivo = Dir()
    Loop
    
    ActiveSheet.Columns("A:C").AutoFit
    Application.ScreenUpdating = True
    
    MsgBox "Archivos encontrados: " & (fila - 2), vbInformation
End Sub


' ============================================================
' LISTAR ARCHIVOS POR EXTENSIÓN
' ============================================================
Public Sub ListarArchivosPorExtension()
    Dim rutaCarpeta As String
    Dim extension As String
    Dim nombreArchivo As String
    Dim fila As Long
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar carpeta"
        If .Show <> -1 Then Exit Sub
        rutaCarpeta = .SelectedItems(1)
    End With
    
    extension = InputBox("Extensión a buscar:", "Extensión", "xlsx")
    If extension = "" Then Exit Sub
    
    If Right(rutaCarpeta, 1) <> "\" Then rutaCarpeta = rutaCarpeta & "\"
    
    ActiveSheet.Cells.Clear
    fila = 2
    
    nombreArchivo = Dir(rutaCarpeta & "*." & extension)
    
    Do While nombreArchivo <> ""
        ActiveSheet.Cells(fila, 1).Value = nombreArchivo
        ActiveSheet.Cells(fila, 2).Value = FileLen(rutaCarpeta & nombreArchivo) \ 1024
        fila = fila + 1
        nombreArchivo = Dir()
    Loop
    
    MsgBox "Archivos encontrados: " & (fila - 2), vbInformation
End Sub
```

---

*Módulo: 02_ListadoArchivos | Archivo: 02_ListarArchivos.md*
