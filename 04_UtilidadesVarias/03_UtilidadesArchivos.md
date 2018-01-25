# Utilidades de Archivos - Funciones Auxiliares

Funciones de utilidad para trabajar con archivos.

## Código VBA

```vba
Option Explicit

' ============================================================
' VERIFICAR SI ARCHIVO EXISTE
' ============================================================
Public Function ArchivoExiste(rutaArchivo As String) As Boolean
    On Error GoNoExiste
    If Len(Dir(rutaArchivo)) > 0 Then
        ArchivoExiste = True
    Else
        ArchivoExiste = False
    End If
    Exit Function
NoExiste:
    ArchivoExiste = False
End Function


' ============================================================
' OBTENER EXTENSIÓN DEL ARCHIVO
' ============================================================
Public Function ObtenerExtension(rutaArchivo As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ObtenerExtension = fso.GetExtensionName(rutaArchivo)
    Set fso = Nothing
End Function


' ============================================================
' OBTENER NOMBRE BASE SIN EXTENSIÓN
' ============================================================
Public Function ObtenerNombreBase(rutaArchivo As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ObtenerNombreBase = fso.GetBaseName(rutaArchivo)
    Set fso = Nothing
End Function


' ============================================================
' OBTENER CARPETA PADRE
' ============================================================
Public Function ObtenerCarpetaPadre(rutaArchivo As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ObtenerCarpetaPadre = fso.GetParentFolderName(rutaArchivo)
    Set fso = Nothing
End Function


' ============================================================
' OBTENER TAMAÑO EN KB
' ============================================================
Public Function ObtenerTamanoKB(rutaArchivo As String) As Long
    If ArchivoExiste(rutaArchivo) Then
        ObtenerTamanoKB = FileLen(rutaArchivo) \ 1024
    Else
        ObtenerTamanoKB = 0
    End If
End Function


' ============================================================
' LIMPIAR NOMBRE DE ARCHIVO
' ============================================================
Public Function LimpiarNombreArchivo(nombre As String) As String
    Dim caracteres As Variant
    Dim i As Long
    
    caracteres = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    LimpiarNombreArchivo = nombre
    
    For i = LBound(caracteres) To UBound(caracteres)
        LimpiarNombreArchivo = Replace(LimpiarNombreArchivo, CStr(caracteres(i)), "_")
    Next i
End Function


' ============================================================
' RENOMBRAR ARCHIVO
' ============================================================
Public Sub RenombrarArchivo(rutaOriginal As String, nuevoNombre As String)
    Dim fso As Object
    Dim rutaNueva As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    rutaNueva = fso.GetParentFolderName(rutaOriginal) & "\" & nuevoNombre
    
    If ArchivoExiste(rutaOriginal) And Not ArchivoExiste(rutaNueva) Then
        Name rutaOriginal As rutaNueva
        MsgBox "Archivo renombrado.", vbInformation
    Else
        MsgBox "Error al renombrar.", vbExclamation
    End If
    
    Set fso = Nothing
End Sub


' ============================================================
' ELIMINAR ARCHIVO CON CONFIRMACIÓN
' ============================================================
Public Sub EliminarArchivo(rutaArchivo As String)
    If Not ArchivoExiste(rutaArchivo) Then
        MsgBox "El archivo no existe.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("¿Eliminar archivo?", vbQuestion + vbYesNo) = vbYes Then
        Kill rutaArchivo
        MsgBox "Archivo eliminado.", vbInformation
    End If
End Sub
```

## Resumen de Funciones

| Función | Descripción |
|---------|-------------|
| `ArchivoExiste` | Verifica existencia |
| `ObtenerExtension` | Retorna extensión |
| `ObtenerNombreBase` | Nombre sin extensión |
| `ObtenerCarpetaPadre` | Carpeta contenedora |
| `ObtenerTamanoKB` | Tamaño en kilobytes |
| `LimpiarNombreArchivo` | Remueve caracteres inválidos |

---

*Módulo: 04_UtilidadesVarias | Archivo: 03_UtilidadesArchivos.md*
