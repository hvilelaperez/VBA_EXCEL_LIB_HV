# Fusionar Archivos Excel

Combina múltiples archivos Excel en un solo libro, copiando sus hojas.

## Código VBA

```vba
Option Explicit

' ============================================================
' FUSIONAR ARCHIVOS SELECCIONADOS
' Permite elegir varios archivos y combinar sus hojas
' ============================================================
Public Sub FusionarArchivosSeleccionados()
    Dim dialogo As FileDialog
    Dim archivos() As String
    Dim i As Long
    Dim wbOrigen As Workbook
    Dim wbDestino As Workbook
    
    On Error GoTo ManejarError
    
    ' Seleccionar archivos
    Set dialogo = Application.FileDialog(msoFileDialogOpen)
    dialogo.Title = "Seleccionar archivos a fusionar"
    dialogo.AllowMultiSelect = True
    dialogo.Filters.Clear
    dialogo.Filters.Add "Archivos Excel", "*.xlsx; *.xlsm; *.xls"
    
    If dialogo.Show <> -1 Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' Crear libro destino
    Set wbDestino = Workbooks.Add(xlWorksheet)
    
    ' Abrir cada archivo y copiar hojas
    For i = 1 To dialogo.SelectedItems.Count
        Set wbOrigen = Workbooks.Open( _
            fileName:=dialogo.SelectedItems(i), _
            ReadOnly:=True)
        
        ' Copiar todas las hojas
        Dim sht As Worksheet
        For Each sht In wbOrigen.Sheets
            sht.Copy After:=wbDestino.Sheets(wbDestino.Sheets.Count)
            
            ' Renombrar para evitar conflictos
            Dim nuevoNombre As String
            nuevoNombre = ObtenerNombreLimpio(wbOrigen.Name) & "_" & sht.Name
            
            ' Truncar si es muy largo
            If Len(nuevoNombre) > 31 Then nuevoNombre = Left(nuevoNombre, 31)
            
            wbDestino.Sheets(wbDestino.Sheets.Count).Name = nuevoNombre
        Next sht
        
        wbOrigen.Close SaveChanges:=False
    Next i
    
    ' Eliminar hoja inicial en blanco
    Application.DisplayAlerts = False
    wbDestino.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
    
    MsgBox "Archivos fusionados correctamente.", vbInformation
    Exit Sub
    
ManejarError:
    MsgBox "Error: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
End Sub


' ============================================================
' FUSIONAR ARCHIVOS DE UNA CARPETA
' Procesa todos los Excel de una carpeta
' ============================================================
Public Sub FusionarArchivosCarpeta()
    Dim rutaCarpeta As String
    Dim nombreArchivo As String
    Dim wbOrigen As Workbook
    Dim wbDestino As Workbook
    Dim contador As Long
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar carpeta con archivos"
        If .Show <> -1 Then Exit Sub
        rutaCarpeta = .SelectedItems(1)
    End With
    
    If Right(rutaCarpeta, 1) <> "\" Then rutaCarpeta = rutaCarpeta & "\"
    
    Application.ScreenUpdating = False
    
    ' Crear libro destino
    Set wbDestino = Workbooks.Add(xlWorksheet)
    contador = 0
    
    ' Procesar cada archivo
    nombreArchivo = Dir(rutaCarpeta & "*.xl*")
    
    Do While nombreArchivo <> ""
        Set wbOrigen = Workbooks.Open( _
            fileName:=rutaCarpeta & nombreArchivo, _
            ReadOnly:=True)
        
        ' Copiar primera hoja de cada archivo
        wbOrigen.Sheets(1).Copy After:=wbDestino.Sheets(wbDestino.Sheets.Count)
        wbDestino.Sheets(wbDestino.Sheets.Count).Name = _
            Left(nombreArchivo, InStrRev(nombreArchivo, ".") - 1)
        
        wbOrigen.Close SaveChanges:=False
        contador = contador + 1
        nombreArchivo = Dir()
    Loop
    
    ' Eliminar hoja inicial
    Application.DisplayAlerts = False
    wbDestino.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
    
    MsgBox contador & " archivos fusionados.", vbInformation
End Sub


' ============================================================
' FUNCIÓN AUXILIAR: Limpiar nombre de archivo
' ============================================================
Private Function ObtenerNombreLimpio(nombre As String) As String
    Dim resultado As String
    Dim caracteres As Variant
    Dim i As Long
    
    caracteres = Array("\", "/", ":", "*", "?", """", "<", ">", "|", ".")
    
    resultado = Left(nombre, InStrRev(nombre, ".") - 1)
    
    For i = LBound(caracteres) To UBound(caracteres)
        resultado = Replace(resultado, CStr(caracteres(i)), "_")
    Next i
    
    ObtenerNombreLimpio = resultado
End Function
```

## Flujo de Ejecución

```
Archivo1.xlsx          Archivo2.xlsx
┌──────────────┐       ┌──────────────┐
│ Hoja1        │       │ Hoja1        │
│ Hoja2        │       │ Hoja2        │
└──────────────┘       └──────────────┘
         │                     │
         └──────────┬──────────┘
                    ▼
         ┌──────────────────┐
         │ Libro Fusionado  │
         ├──────────────────┤
         │ Archivo1_Hoja1   │
         │ Archivo1_Hoja2   │
         │ Archivo2_Hoja1   │
         │ Archivo2_Hoja2   │
         └──────────────────┘
```

## Descripción de Macros

| Macro | Descripción |
|-------|-------------|
| `FusionarArchivosSeleccionados` | Fusiona archivos elegidos por el usuario |
| `FusionarArchivosCarpeta` | Fusiona todos los Excel de una carpeta |

---

*Módulo: 06_OperacionesHojas | Archivo: 02_FusionarArchivosExcel.md*
