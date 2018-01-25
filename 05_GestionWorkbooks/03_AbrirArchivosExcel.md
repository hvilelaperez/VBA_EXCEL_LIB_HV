# Abrir Archivos Excel

Diferentes métodos para abrir archivos Excel con opciones de solo lectura y verificación.

## Código VBA

```vba
Option Explicit

' ============================================================
' ABRIR ARCHIVO CON FUNCIÓN REUTILIZABLE
' Función que abre un libro y retorna el objeto Workbook
' ============================================================
Public Function AbrirArchivoExcel(ByVal rutaArchivo As String, _
                                   Optional ByVal soloLectura As Boolean = False) As Workbook
    Application.DisplayAlerts = False
    
    Workbooks.Open fileName:=rutaArchivo, ReadOnly:=soloLectura
    
    Set AbrirArchivoExcel = ActiveWorkbook
    
    Application.DisplayAlerts = True
End Function


' ============================================================
' EJEMPLO 1: Abrir y cerrar sin guardar
' ============================================================
Public Sub EjemploAbrirCerrar()
    Dim wb As Workbook
    
    ' Abrir archivo
    Set wb = Workbooks.Open("C:\Datos\Muestra.xlsx")
    
    ' Realizar operaciones...
    
    ' Cerrar sin guardar cambios
    wb.Close SaveChanges:=False
End Sub


' ============================================================
' EJEMPLO 2: Abrir, modificar y guardar
' ============================================================
Public Sub EjemploAbrirModificarGuardar()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' Abrir archivo
    Set wb = Workbooks.Open("C:\Datos\Muestra.xlsx")
    
    ' Establecer referencia a hoja
    Set ws = wb.Worksheets(1)
    
    ' Escribir datos
    ws.Cells(1, 1).Value = "Hola VBA!"
    
    ' Guardar cambios y cerrar
    wb.Close SaveChanges:=True
End Sub


' ============================================================
' EJEMPLO 3: Abrir solo lectura
' ============================================================
Public Sub EjemploAbrirSoloLectura()
    Dim wb As Workbook
    
    Set wb = Workbooks.Open( _
        fileName:="C:\Datos\Muestra.xlsx", _
        ReadOnly:=True)
    
    ' Leer datos (no se puede escribir en modo lectura)
    MsgBox "Dato en A1: " & wb.Sheets(1).Cells(1, 1).Value
    
    ' Cerrar sin guardar
    wb.Close SaveChanges:=False
End Sub


' ============================================================
' ABRIR ARCHIVO Y VERIFICAR SI YA ESTÁ ABIERTO
' ============================================================
Public Sub AbrirSiNoEstaAbierto()
    Dim nombreArchivo As String
    Dim rutaCompleta As String
    
    rutaCompleta = "C:\Datos\BaseDatos.xlsx"
    nombreArchivo = Mid(rutaCompleta, InStrRev(rutaCompleta, "\") + 1)
    
    ' Verificar si ya está abierto
    If Not ArchivoExcelAbierto(nombreArchivo) Then
        Workbooks.Open rutaCompleta
        MsgBox "Archivo abierto correctamente.", vbInformation
    Else
        MsgBox "El archivo ya está abierto.", vbExclamation
    End If
End Sub


' ============================================================
' FUNCIÓN: Verificar si archivo Excel está abierto
' ============================================================
Public Function ArchivoExcelAbierto(nombreArchivo As String) As Boolean
    On Error Resume Next
    
    Dim wb As Workbook
    Set wb = Application.Workbooks(nombreArchivo)
    
    ArchivoExcelAbierto = Not wb Is Nothing
    
    On Error GoTo 0
End Function


' ============================================================
' ABRIR MÚLTIPLES ARCHIVOS
' ============================================================
Public Sub AbrirMultiplesArchivos()
    Dim dialogo As FileDialog
    Dim rutaArchivo As Variant
    
    Set dialogo = Application.FileDialog(msoFileDialogOpen)
    dialogo.Title = "Seleccionar archivos Excel"
    dialogo.AllowMultiSelect = True
    dialogo.Filters.Clear
    dialogo.Filters.Add "Archivos Excel", "*.xls; *.xlsx; *.xlsm"
    
    If dialogo.Show <> -1 Then Exit Sub
    
    For Each rutaArchivo In dialogo.SelectedItems
        Workbooks.Open CStr(rutaArchivo)
    Next rutaArchivo
    
    MsgBox dialogo.SelectedItems.Count & " archivos abiertos.", vbInformation
End Sub
```

## Descripción de Funciones

| Función/Macro | Descripción |
|---------------|-------------|
| `AbrirArchivoExcel` | Función reutilizable para abrir archivos |
| `EjemploAbrirCerrar` | Abre y cierra sin guardar |
| `EjemploAbrirModificarGuardar` | Abre, modifica y guarda |
| `EjemploAbrirSoloLectura` | Abre en modo solo lectura |
| `AbrirSiNoEstaAbierto` | Verifica antes de abrir |
| `ArchivoExcelAbierto` | Verifica si un libro está abierto |
| `AbrirMultiplesArchivos` | Abre varios archivos seleccionados |

## Parámetros de Workbooks.Open

| Parámetro | Descripción |
|-----------|-------------|
| `fileName` | Ruta completa del archivo |
| `ReadOnly` | True = solo lectura |
| `Password` | Contraseña del archivo |
| `UpdateLinks` | Actualizar vínculos externos |

---

*Módulo: 05_GestionWorkbooks | Archivo: 03_AbrirArchivosExcel.md*
