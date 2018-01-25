# Separar Hojas a Archivos Distintos

Crea archivos individuales para cada hoja del libro actual.

## Código VBA

```vba
Option Explicit

' ============================================================
' SEPARAR HOJAS A ARCHIVOS INDIVIDUALES
' Crea un archivo por cada hoja del libro
' ============================================================
Public Sub SepararHojasAArchivos()
    Dim ws As Worksheet
    Dim nombreArchivo As String
    Dim rutaGuardado As String
    Dim contador As Long
    Dim formato As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Preguntar formato
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Guardar como .xlsx (No) o .xlsm (Sí con macros)?", _
                       vbYesNo + vbQuestion, "Formato")
    
    If respuesta = vbYes Then
        formato = xlOpenXMLWorkbookMacroEnabled
    Else
        formato = xlOpenXMLWorkbook
    End If
    
    contador = 0
    
    For Each ws In ThisWorkbook.Worksheets
        ' Crear nombre de archivo
        nombreArchivo = ws.Name
        ' Limpiar caracteres no válidos
        nombreArchivo = LimpiarNombre(nombreArchivo)
        
        rutaGuardado = ThisWorkbook.Path & "\" & nombreArchivo & ".xlsx"
        
        ' Copiar hoja a libro nuevo
        ws.Copy
        
        ' Guardar como nuevo archivo
        ActiveWorkbook.SaveAs fileName:=rutaGuardado, FileFormat:=formato
        ActiveWorkbook.Close SaveChanges:=False
        
        contador = contador + 1
    Next ws
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox contador & " archivos creados.", vbInformation
End Sub


' ============================================================
' SEPARAR HOJAS SELECCIONADAS
' Solo separa las hojas que estén seleccionadas
' ============================================================
Public Sub SepararHojasSeleccionadas()
    Dim ws As Worksheet
    Dim nombreArchivo As String
    Dim rutaGuardado As String
    Dim contador As Long
    
    If ActiveWindow.SelectedSheets.Count < 2 Then
        MsgBox "Seleccione las hojas a separar (Ctrl + clic).", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    contador = 0
    
    For Each ws In ActiveWindow.SelectedSheets
        nombreArchivo = LimpiarNombre(ws.Name)
        rutaGuardado = ThisWorkbook.Path & "\" & nombreArchivo & ".xlsx"
        
        ws.Copy
        ActiveWorkbook.SaveAs fileName:=rutaGuardado, FileFormat:=xlOpenXMLWorkbook
        ActiveWorkbook.Close SaveChanges:=False
        
        contador = contador + 1
    Next ws
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox contador & " archivos creados.", vbInformation
End Sub


' ============================================================
' SEPARAR CON PREFIJO PERSONALIZADO
' ============================================================
Public Sub SepararConPrefijo()
    Dim ws As Worksheet
    Dim prefijo As String
    Dim rutaGuardado As String
    
    prefijo = InputBox("Ingrese prefijo para los archivos:", "Prefijo", "Copia_")
    If prefijo = "" Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        rutaGuardado = ThisWorkbook.Path & "\" & prefijo & LimpiarNombre(ws.Name) & ".xlsx"
        
        ws.Copy
        ActiveWorkbook.SaveAs fileName:=rutaGuardado, FileFormat:=xlOpenXMLWorkbook
        ActiveWorkbook.Close SaveChanges:=False
    Next ws
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Archivos creados con prefijo: " & prefijo, vbInformation
End Sub


' ============================================================
' FUNCIÓN AUXILIAR: Limpiar nombre de archivo
' ============================================================
Private Function LimpiarNombre(nombre As String) As String
    Dim resultado As String
    Dim caracteres As Variant
    Dim i As Long
    
    caracteres = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    resultado = nombre
    
    For i = LBound(caracteres) To UBound(caracteres)
        resultado = Replace(resultado, CStr(caracteres(i)), "_")
    Next i
    
    ' Truncar si es muy largo
    If Len(resultado) > 31 Then resultado = Left(resultado, 31)
    
    LimpiarNombre = resultado
End Function
```

## Descripción de Macros

| Macro | Descripción |
|-------|-------------|
| `SepararHojasAArchivos` | Separa todas las hojas a archivos |
| `SepararHojasSeleccionadas` | Separa solo hojas seleccionadas |
| `SepararConPrefijo` | Agrega prefijo a los nombres |

## Ejemplo

```
Libro Actual (3 hojas):
├── Ventas
├── Inventario
└── Reportes

Resultado:
├── Ventas.xlsx
├── Inventario.xlsx
└── Reportes.xlsx
```

---

*Módulo: 06_OperacionesHojas | Archivo: 05_SepararHojasArchivos.md*
