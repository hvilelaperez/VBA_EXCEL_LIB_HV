# Diálogo para Seleccionar Archivo Excel

Muestra un diálogo de selección de archivo con filtros para diferentes formatos de Excel.

## Código VBA

```vba
Option Explicit

' ============================================================
' SELECCIONAR ARCHIVO EXCEL
' Diálogo con filtro para archivos Excel
' ============================================================
Public Sub SeleccionarArchivoExcel()
    Dim fd As Object
    Dim archivoSeleccionado As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .InitialFileName = ActiveWorkbook.Path & "\"
        .Filters.Clear
        .Filters.Add "Excel 2007+", "*.xlsx"
        .Filters.Add "Excel con macros", "*.xlsm"
        .Filters.Add "Excel 97-2003", "*.xls"
        .Filters.Add "Todos los archivos", "*.*"
        .Title = "Seleccionar archivo Excel"
        
        If .Show = True Then
            archivoSeleccionado = .SelectedItems(1)
            MsgBox "Archivo seleccionado: " & archivoSeleccionado, vbInformation
        Else
            MsgBox "No se seleccionó ningún archivo.", vbExclamation
        End If
    End With
    
    Set fd = Nothing
End Sub


' ============================================================
' SELECCIONAR Y ABRIR ARCHIVO
' Selecciona y abre el archivo directamente
' ============================================================
Public Sub SeleccionarYAbrir()
    Dim fd As Object
    Dim rutaArchivo As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path & "\"
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xls; *.xlsx; *.xlsm; *.xlsb"
        .Title = "Seleccionar archivo para abrir"
        
        If .Show = True Then
            rutaArchivo = .SelectedItems(1)
            
            ' Abrir el archivo
            Dim wb As Workbook
            Set wb = Workbooks.Open(rutaArchivo, ReadOnly:=True)
            
            MsgBox "Archivo abierto: " & wb.Name, vbInformation
            
            ' Aquí puedes trabajar con el archivo...
            
            wb.Close SaveChanges:=False
        End If
    End With
    
    Set fd = Nothing
End Sub


' ============================================================
' SELECCIONAR ARCHIVO Y GUARDAR RUTA EN CELDA
' ============================================================
Public Sub SeleccionarYGuardarRuta()
    Dim fd As Object
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path & "\"
        .Filters.Clear
        .Filters.Add "Excel", "*.xls; *.xlsx; *.xlsm"
        .Title = "Seleccionar archivo"
        
        If .Show = True Then
            ' Guardar ruta en celda activa
            ActiveCell.Value = .SelectedItems(1)
            MsgBox "Ruta guardada en celda.", vbInformation
        End If
    End With
    
    Set fd = Nothing
End Sub


' ============================================================
' SELECCIONAR MÚLTIPLES ARCHIVOS
' ============================================================
Public Sub SeleccionarMultiples()
    Dim fd As Object
    Dim i As Long
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = True
        .InitialFileName = ThisWorkbook.Path & "\"
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xls; *.xlsx; *.xlsm"
        .Title = "Seleccionar múltiples archivos"
        
        If .Show = True Then
            For i = 1 To .SelectedItems.Count
                Debug.Print "Archivo " & i & ": " & .SelectedItems(i)
            Next i
            
            MsgBox .SelectedItems.Count & " archivos seleccionados.", vbInformation
        End If
    End With
    
    Set fd = Nothing
End Sub
```

## Filtros Comunes

| Filtro | Extensión |
|--------|-----------|
| Excel 2007+ | `*.xlsx` |
| Excel con macros | `*.xlsm` |
| Excel binario | `*.xlsb` |
| Excel 97-2003 | `*.xls` |
| Todos | `*.*` |

## Descripción de Macros

| Macro | Descripción |
|-------|-------------|
| `SeleccionarArchivoExcel` | Muestra diálogo y retorna ruta |
| `SeleccionarYAbrir` | Selecciona y abre el archivo |
| `SeleccionarYGuardarRuta` | Guarda la ruta en la celda activa |
| `SeleccionarMultiples` | Permite elegir varios archivos |

---

*Módulo: 05_GestionWorkbooks | Archivo: 04_DialogoSeleccionArchivo.md*
