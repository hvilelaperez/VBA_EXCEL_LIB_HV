# Contar Archivos

Cuenta archivos por extensión o tipo en una carpeta.

## Código VBA

```vba
Option Explicit

' ============================================================
' CONTAR ARCHIVOS EN CARPETA
' ============================================================
Public Sub ContarArchivos()
    Dim rutaCarpeta As String
    Dim nombreArchivo As String
    Dim contador As Long
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar carpeta"
        If .Show <> -1 Then Exit Sub
        rutaCarpeta = .SelectedItems(1)
    End With
    
    If Right(rutaCarpeta, 1) <> "\" Then rutaCarpeta = rutaCarpeta & "\"
    
    contador = 0
    nombreArchivo = Dir(rutaCarpeta & "*.*")
    
    Do While nombreArchivo <> ""
        contador = contador + 1
        nombreArchivo = Dir()
    Loop
    
    MsgBox "Total de archivos: " & contador, vbInformation
End Sub


' ============================================================
' CONTAR POR MÚLTIPLES EXTENSIONES
' ============================================================
Public Sub ContarPorExtensiones()
    Dim rutaCarpeta As String
    Dim extensiones As Variant
    Dim ext As Variant
    Dim contador As Long
    Dim resultado As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar carpeta"
        If .Show <> -1 Then Exit Sub
        rutaCarpeta = .SelectedItems(1)
    End With
    
    If Right(rutaCarpeta, 1) <> "\" Then rutaCarpeta = rutaCarpeta & "\"
    
    extensiones = Array("xlsx", "xlsm", "pdf", "docx", "txt", "csv")
    resultado = ""
    
    For Each ext In extensiones
        contador = 0
        Dim nombreArchivo As String
        nombreArchivo = Dir(rutaCarpeta & "*." & ext)
        
        Do While nombreArchivo <> ""
            contador = contador + 1
            nombreArchivo = Dir()
        Loop
        
        resultado = resultado & UCase(ext) & ": " & contador & vbCrLf
    Next ext
    
    MsgBox resultado, vbInformation, "Conteo por Extensión"
End Sub
```

---

*Módulo: 02_ListadoArchivos | Archivo: 03_ContarArchivos.md*
