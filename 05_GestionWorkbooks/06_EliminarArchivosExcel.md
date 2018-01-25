# Eliminar Archivos Excel

Diferentes métodos para eliminar archivos Excel de forma segura con confirmación.

## Código VBA

```vba
Option Explicit

' ============================================================
' ELIMINAR ARCHIVO CON FileSystemObject
' Método seguro que verifica existencia primero
' ============================================================
Public Sub EliminarConFSO()
    Dim fso As Object
    Dim rutaArchivo As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    rutaArchivo = ThisWorkbook.Path & "\ArchivoAEliminar.xlsx"
    
    If fso.FileExists(rutaArchivo) Then
        fso.DeleteFile rutaArchivo, True
        MsgBox "Archivo eliminado.", vbInformation
    Else
        MsgBox "El archivo no existe.", vbExclamation
    End If
    
    Set fso = Nothing
End Sub


' ============================================================
' ELIMINAR ARCHIVO CON Kill
' Método clásico de VBA
' ============================================================
Public Sub EliminarConKill()
    Dim rutaArchivo As String
    
    rutaArchivo = ThisWorkbook.Path & "\ArchivoAEliminar.xlsx"
    
    If Dir(rutaArchivo) <> "" Then
        Kill rutaArchivo
        MsgBox "Archivo eliminado.", vbInformation
    Else
        MsgBox "El archivo no existe.", vbExclamation
    End If
End Sub


' ============================================================
' ELIMINAR ARCHIVOS DE UNA CARPETA
' Elimina todos los archivos que coincidan con patrón
' ============================================================
Public Sub EliminarArchivosCarpeta()
    Dim rutaCarpeta As String
    Dim patron As String
    Dim nombreArchivo As String
    Dim contador As Long
    
    rutaCarpeta = ThisWorkbook.Path & "\Temporal\"
    patron = "*.tmp"
    
    nombreArchivo = Dir(rutaCarpeta & patron)
    contador = 0
    
    Do While nombreArchivo <> ""
        Kill rutaCarpeta & nombreArchivo
        contador = contador + 1
        nombreArchivo = Dir()
    Loop
    
    MsgBox contador & " archivos eliminados.", vbInformation
End Sub


' ============================================================
' ELIMINAR CON CONFIRMACIÓN DEL USUARIO
' Pide confirmación antes de eliminar
' ============================================================
Public Sub EliminarConConfirmacion()
    Dim rutaArchivo As String
    Dim respuesta As VbMsgBoxResult
    
    rutaArchivo = ThisWorkbook.Path & "\ArchivoImportante.xlsx"
    
    If Dir(rutaArchivo) = "" Then
        MsgBox "El archivo no existe.", vbExclamation
        Exit Sub
    End If
    
    respuesta = MsgBox("¿Está seguro de eliminar este archivo?" & vbCrLf & _
                       rutaArchivo, vbQuestion + vbYesNo, "Confirmar Eliminación")
    
    If respuesta = vbYes Then
        Kill rutaArchivo
        MsgBox "Archivo eliminado.", vbInformation
    Else
        MsgBox "Eliminación cancelada.", vbExclamation
    End If
End Sub


' ============================================================
' ELIMINAR TODOS LOS ARCHivos EXCEPTO ACTUAL
' ============================================================
Public Sub EliminarArchivosExceptoActual()
    Dim rutaCarpeta As String
    Dim nombreArchivo As String
    Dim nombreActual As String
    Dim contador As Long
    
    rutaCarpeta = ThisWorkbook.Path & "\"
    nombreActual = ThisWorkbook.Name
    contador = 0
    
    nombreArchivo = Dir(rutaCarpeta & "*.xlsx")
    
    Do While nombreArchivo <> ""
        If nombreArchivo <> nombreActual Then
            Kill rutaCarpeta & nombreArchivo
            contador = contador + 1
        End If
        nombreArchivo = Dir()
    Loop
    
    MsgBox contador & " archivos eliminados (se conservó el actual).", vbInformation
End Sub


' ============================================================
' ELIMINAR CARPETA TEMPORAL COMPLETA
' ============================================================
Public Sub EliminarCarpetaTemporal()
    Dim fso As Object
    Dim rutaCarpeta As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    rutaCarpeta = ThisWorkbook.Path & "\Temporal"
    
    If fso.FolderExists(rutaCarpeta) Then
        fso.DeleteFolder rutaCarpeta, True
        MsgBox "Carpeta temporal eliminada.", vbInformation
    Else
        MsgBox "La carpeta temporal no existe.", vbExclamation
    End If
    
    Set fso = Nothing
End Sub
```

## Comparación de Métodos

| Método | Ventaja | Desventaja |
|--------|---------|------------|
| `Kill` | Nativo VBA, rápido | No maneja errores bien |
| `FileSystemObject` | Robusta, verifica existencia | Requiere creación de objeto |

## Notas de Seguridad

- **Siempre** verificar que el archivo existe antes de eliminar
- Usar `MsgBox` con `vbYesNo` para confirmación
- Evitar eliminar archivos del sistema
- Hacer copia de seguridad antes de operaciones masivas

---

*Módulo: 05_GestionWorkbooks | Archivo: 06_EliminarArchivosExcel.md*
