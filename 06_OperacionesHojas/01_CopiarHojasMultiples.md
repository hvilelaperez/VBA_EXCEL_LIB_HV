# Copiar Hojas de Múltiples Archivos

Copia todas las hojas de varios archivos Excel en un solo libro nuevo.

## Código VBA

```vba
Option Explicit

' ============================================================
' COPIAR HOJAS DE MÚLTIPLES ARCHIVOS
' Abre varios archivos y copia sus hojas al libro actual
' ============================================================
Public Sub CopiarHojasMultiplesArchivos()
    Dim rutaCarpeta As String
    Dim nombreArchivo As String
    Dim wbOrigen As Workbook
    Dim sht As Worksheet
    Dim fso As Object
    Dim contador As Long
    
    ' Seleccionar carpeta
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar carpeta con archivos Excel"
        If .Show <> -1 Then Exit Sub
        rutaCarpeta = .SelectedItems(1)
    End With
    
    If Right(rutaCarpeta, 1) <> "\" Then rutaCarpeta = rutaCarpeta & "\"
    
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    contador = 0
    
    ' Recorrer archivos Excel en la carpeta
    nombreArchivo = Dir(rutaCarpeta & "*.xl*")
    
    Do While nombreArchivo <> ""
        ' Abrir archivo en modo lectura
        Set wbOrigen = Workbooks.Open( _
            fileName:=rutaCarpeta & nombreArchivo, _
            ReadOnly:=True)
        
        ' Copiar cada hoja al libro actual
        For Each sht In wbOrigen.Worksheets
            sht.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            
            ' Renombrar hoja con nombre del archivo
            ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = _
                fso.GetBaseName(nombreArchivo) & "-" & sht.Name
            
            contador = contador + 1
        Next sht
        
        wbOrigen.Close SaveChanges:=False
        nombreArchivo = Dir()
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Hojas copiadas: " & contador, vbInformation
End Sub


' ============================================================
' ELIMINAR TODAS LAS HOJAS EXCEPTO UNA
' Limpia el libro dejando solo una hoja
' ============================================================
Public Sub EliminarHojasExcepto()
    Dim sht As Worksheet
    Dim nombreConservar As String
    
    nombreConservar = InputBox("Nombre de la hoja a conservar:", "Conservar Hoja", "Hoja1")
    If nombreConservar = "" Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name <> nombreConservar Then
            sht.Delete
        End If
    Next sht
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Se eliminaron todas las hojas excepto: " & nombreConservar, vbInformation
End Sub


' ============================================================
' OBTENER NOMBRE BASE SIN EXTENSIÓN
' Función auxiliar
' ============================================================
Public Function ObtenerNombreBase(nombreArchivo As String) As String
    If InStr(nombreArchivo, ".") > 0 Then
        ObtenerNombreBase = Left(nombreArchivo, InStr(nombreArchivo, ".") - 1)
    Else
        ObtenerNombreBase = nombreArchivo
    End If
End Function
```

## Flujo de Ejecución

```
CarpetaSeleccionada/
├── Ventas.xlsx      → Hoja1 → Ventas-Hoja1
├── Inventario.xlsx  → Hoja1 → Inventario-Hoja1
└── Reportes.xlsx    → Hoja1 → Reportes-Hoja1
                      Hoja2 → Reportes-Hoja2
```

## Descripción de Macros

| Macro | Descripción |
|-------|-------------|
| `CopiarHojasMultiplesArchivos` | Copia todas las hojas de archivos en una carpeta |
| `EliminarHojasExcepto` | Elimina hojas dejando solo una específica |
| `ObtenerNombreBase` | Obtiene nombre sin extensión |

---

*Módulo: 06_OperacionesHojas | Archivo: 01_CopiarHojasMultiples.md*
