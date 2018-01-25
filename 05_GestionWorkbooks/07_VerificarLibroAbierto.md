# Verificar si Libro Está Abierto

Funciones para detectar si un archivo Excel está abierto en la sesión actual.

## Código VBA

```vba
Option Explicit

' ============================================================
' VERIFICAR SI LIBRO ESTÁ ABIERTO (Método 1)
' Usa FreeFile y bloqueo de archivo
' ============================================================
Public Function LibroAbierto_FreeFile(rutaArchivo As String) As Boolean
    Dim archivoLibre As Long
    Dim errorNumero As Long
    
    On Error Resume Next
    
    archivoLibre = FreeFile()
    Open rutaArchivo For Input Lock Read As #archivoLibre
    Close archivoLibre
    errorNumero = Err
    
    On Error GoTo 0
    
    Select Case errorNumero
        Case 0:     LibroAbierto_FreeFile = False
        Case 70:    LibroAbierto_FreeFile = True
        Case Else:  Err.Raise errorNumero
    End Select
End Function


' ============================================================
' VERIFICAR SI LIBRO ESTÁ ABIERTO (Método 2)
' Usa colección Workbooks
' ============================================================
Public Function LibroAbierto_Workbooks(nombreArchivo As String) As Boolean
    Dim wb As Workbook
    
    On Error Resume Next
    Set wb = Application.Workbooks(nombreArchivo)
    
    LibroAbierto_Workbooks = Not wb Is Nothing
    
    On Error GoTo 0
End Function


' ============================================================
' VERIFICAR POR RUTA COMPLETA
' ============================================================
Public Function LibroAbiertoPorRuta(rutaCompleta As String) As Boolean
    Dim nombreArchivo As String
    nombreArchivo = Mid(rutaCompleta, InStrRev(rutaCompleta, "\") + 1)
    
    LibroAbiertoPorRuta = LibroAbierto_Workbooks(nombreArchivo)
End Function


' ============================================================
' CERRAR SI ESTÁ ABIERTO
' ============================================================
Public Sub CerrarSiAbierto(nombreArchivo As String, Optional guardar As Boolean = False)
    If LibroAbierto_Workbooks(nombreArchivo) Then
        Application.Workbooks(nombreArchivo).Close saveChanges:=guardar
        MsgBox "Libro cerrado: " & nombreArchivo, vbInformation
    Else
        MsgBox "El libro no está abierto.", vbExclamation
    End If
End Sub


' ============================================================
' LISTAR TODOS LOS LIBROS ABIERTOS
' ============================================================
Public Sub ListarLibrosAbiertos()
    Dim wb As Workbook
    Dim mensaje As String
    Dim contador As Long
    
    mensaje = "Libros abiertos:" & vbCrLf & vbCrLf
    contador = 0
    
    For Each wb In Application.Workbooks
        contador = contador + 1
        mensaje = mensaje & contador & ". " & wb.Name
        
        If wb.Path <> "" Then
            mensaje = mensaje & vbCrLf & "   Ruta: " & wb.Path
        Else
            mensaje = mensaje & vbCrLf & "   (Sin guardar)"
        End If
        
        mensaje = mensaje & vbCrLf & "   Hojas: " & wb.Sheets.Count & vbCrLf
    Next wb
    
    mensaje = mensaje & "Total: " & contador & " libro(s)"
    
    MsgBox mensaje, vbInformation, "Libros Abiertos"
End Sub


' ============================================================
' CERRAR TODOS EXCEPTO EL ACTUAL
' ============================================================
Public Sub CerrarTodosExceptoActual()
    Dim wb As Workbook
    Dim contador As Long
    
    contador = 0
    
    For Each wb In Application.Workbooks
        If wb.Name <> ActiveWorkbook.Name Then
            wb.Close saveChanges:=False
            contador = contador + 1
        End If
    Next wb
    
    MsgBox contador & " libro(s) cerrado(s).", vbInformation
End Sub
```

## Comparación de Métodos

| Método | Ventaja | Desventaja |
|--------|---------|------------|
| `FreeFile` | Detecta archivos no Excel | Más complejo |
| `Workbooks` | Simple y directo | Solo para libros Excel |

## Descripción de Funciones

| Función | Descripción |
|---------|-------------|
| `LibroAbierto_FreeFile` | Verifica con bloqueo de archivo |
| `LibroAbierto_Workbooks` | Verifica en colección Workbooks |
| `LibroAbiertoPorRuta` | Verifica por ruta completa |
| `CerrarSiAbierto` | Cierra si está abierto |
| `ListarLibrosAbiertos` | Muestra todos los libros abiertos |
| `CerrarTodosExceptoActual` | Cierra todos menos el activo |

---

*Módulo: 05_GestionWorkbooks | Archivo: 07_VerificarLibroAbierto.md*
