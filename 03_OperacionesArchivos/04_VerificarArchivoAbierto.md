# Verificar si Archivo está Abierto

Detecta si un archivo Excel específico está abierto en la sesión actual.

## Código VBA

```vba
Option Explicit

' ============================================================
' VERIFICAR SI ARCHIVO ESTÁ ABIERTO
' ============================================================
Public Function ArchivoAbierto(nombreArchivo As String) As Boolean
    On Error Resume Next
    
    Dim wb As Workbook
    Set wb = Application.Workbooks(nombreArchivo)
    
    ArchivoAbierto = Not wb Is Nothing
    
    On Error GoTo 0
End Function


' ============================================================
' LISTAR ARCHIVOS ABIERTOS
' ============================================================
Public Sub ListarArchivosAbiertos()
    Dim wb As Workbook
    Dim mensaje As String
    
    mensaje = "Libros abiertos:" & vbCrLf & vbCrLf
    
    For Each wb In Application.Workbooks
        mensaje = mensaje & "- " & wb.Name
        If wb.Path <> "" Then
            mensaje = mensaje & vbCrLf & "  " & wb.Path
        End If
        mensaje = mensaje & vbCrLf
    Next wb
    
    MsgBox mensaje, vbInformation
End Sub


' ============================================================
' CERRAR TODOS EXCEPTO EL ACTUAL
' ============================================================
Public Sub CerrarTodosExceptoActual()
    Dim wb As Workbook
    
    For Each wb In Application.Workbooks
        If wb.Name <> ActiveWorkbook.Name Then
            wb.Close saveChanges:=False
        End If
    Next wb
    
    MsgBox "Archivos cerrados.", vbInformation
End Sub
```

---

*Módulo: 03_OperacionesArchivos | Archivo: 04_VerificarArchivoAbierto.md*
