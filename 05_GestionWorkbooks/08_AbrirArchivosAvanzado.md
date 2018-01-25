# Abrir Archivos Excel Avanzado

Métodos avanzados para abrir archivos Excel con manejo de errores y opciones especiales.

## Código VBA

```vba
Option Explicit

' ============================================================
' ABRIR ARCHIVO CON VERIFICACIÓN COMPLETA
' Verifica existencia, si está abierto y permisos
' ============================================================
Public Function AbrirArchivoSeguro(rutaArchivo As String, _
                                    Optional soloLectura As Boolean = False, _
                                    Optional mostrarError As Boolean = True) As Workbook
    Dim nombreArchivo As String
    
    ' Verificar si el archivo existe
    If Dir(rutaArchivo) = "" Then
        If mostrarError Then
            MsgBox "El archivo no existe:" & vbCrLf & rutaArchivo, vbExclamation
        End If
        Set AbrirArchivoSeguro = Nothing
        Exit Function
    End If
    
    ' Obtener nombre del archivo
    nombreArchivo = Mid(rutaArchivo, InStrRev(rutaArchivo, "\") + 1)
    
    ' Verificar si ya está abierto
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Application.Workbooks(nombreArchivo)
    On Error GoTo 0
    
    If Not wb Is Nothing Then
        ' Ya está abierto, retornar referencia
        Set AbrirArchivoSeguro = wb
        Exit Function
    End If
    
    ' Abrir el archivo
    On Error GoTo ManejarError
    
    Set wb = Workbooks.Open( _
        fileName:=rutaArchivo, _
        ReadOnly:=soloLectura, _
        UpdateLinks:=0) ' No actualizar vínculos
    
    Set AbrirArchivoSeguro = wb
    Exit Function
    
ManejarError:
    If mostrarError Then
        MsgBox "Error al abrir:" & vbCrLf & Err.Description, vbCritical
    End If
    Set AbrirArchivoSeguro = Nothing
End Function


' ============================================================
' ABRIR ARCHIVO Y CERRAR AL TERMINAR
' Abre, ejecuta macro y cierra automáticamente
' ============================================================
Public Sub AbrirEjecutarCerrar()
    Dim rutaArchivo As String
    Dim wb As Workbook
    
    rutaArchivo = ThisWorkbook.Path & "\Proceso.xlsx"
    
    Set wb = AbrirArchivoSeguro(rutaArchivo, soloLectura:=False)
    
    If wb Is Nothing Then
        MsgBox "No se pudo abrir el archivo.", vbExclamation
        Exit Sub
    End If
    
    ' Ejecutar macro del archivo
    Application.Run "'" & wb.Name & "'!MiMacroProceso"
    
    ' Guardar y cerrar
    wb.Close saveChanges:=True
    
    MsgBox "Proceso completado.", vbInformation
End Sub


' ============================================================
' ABRIR MÚLTIPLES ARCHIVOS EN SOLO LECTURA
' ============================================================
Public Sub AbrirMultiplesSoloLectura()
    Dim dialogo As FileDialog
    Dim rutaArchivo As Variant
    Dim contador As Long
    
    Set dialogo = Application.FileDialog(msoFileDialogOpen)
    dialogo.Title = "Seleccionar archivos"
    dialogo.AllowMultiSelect = True
    dialogo.Filters.Clear
    dialogo.Filters.Add "Excel", "*.xls; *.xlsx; *.xlsm"
    
    If dialogo.Show <> -1 Then Exit Sub
    
    Application.ScreenUpdating = False
    contador = 0
    
    For Each rutaArchivo In dialogo.SelectedItems
        Dim wb As Workbook
        Set wb = AbrirArchivoSeguro(CStr(rutaArchivo), soloLectura:=True)
        If Not wb Is Nothing Then contador = contador + 1
    Next rutaArchivo
    
    Application.ScreenUpdating = True
    
    MsgBox contador & " archivos abiertos.", vbInformation
End Sub


' ============================================================
' ABRIR ARCHIVO CON CONTRASEÑA
' ============================================================
Public Function AbrirArchivoConContrasena(rutaArchivo As String, _
                                           contrasena As String) As Workbook
    Dim wb As Workbook
    
    On Error GoTo ManejarError
    
    Set wb = Workbooks.Open( _
        fileName:=rutaArchivo, _
        Password:=contrasena)
    
    Set AbrirArchivoConContrasena = wb
    Exit Function
    
ManejarError:
    MsgBox "Error al abrir (contraseña incorrecta?):" & vbCrLf & Err.Description, vbCritical
    Set AbrirArchivoConContrasena = Nothing
End Function


' ============================================================
' CERRAR TODOS LOS ARCHIVOS EXCEPTO ESTE
' ============================================================
Public Sub CerrarTodosExceptoEste()
    Dim wb As Workbook
    Dim contador As Long
    
    contador = 0
    
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            wb.Close saveChanges:=False
            contador = contador + 1
        End If
    Next wb
    
    MsgBox contador & " archivos cerrados.", vbInformation
End Sub
```

## Descripción de Funciones

| Función/Macro | Descripción |
|---------------|-------------|
| `AbrirArchivoSeguro` | Abre con verificación completa |
| `AbrirEjecutarCerrar` | Abre, ejecuta y cierra |
| `AbrirMultiplesSoloLectura` | Abre varios en solo lectura |
| `AbrirArchivoConContrasena` | Abre con contraseña |
| `CerrarTodosExceptoEste` | Cierra todos menos el actual |

## Parámetros de Workbooks.Open

| Parámetro | Descripción | Valor |
|-----------|-------------|-------|
| `ReadOnly` | Solo lectura | True/False |
| `Password` | Contraseña | String |
| `UpdateLinks` | Actualizar vínculos | 0=No, 1=Externos, 2=Preguntar |
| `Format` | Formato del archivo | xlExcel8, xlOpenXMLWorkbook, etc. |

---

*Módulo: 05_GestionWorkbooks | Archivo: 08_AbrirArchivosAvanzado.md*
