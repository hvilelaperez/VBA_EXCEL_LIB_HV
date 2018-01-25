# Combinar Datos de 2 Workbooks

Fusiona datos de dos libros diferentes en uno solo.

## Código VBA

```vba
Option Explicit

' ============================================================
' COMBINAR DOS ARCHIVOS DE CARPETA
' Fusiona todos los Excel de una carpeta en uno nuevo
' ============================================================
Public Sub CombinarArchivosCarpeta()
    Dim rutaCarpeta As String
    Dim nombreArchivo As String
    Dim wbDestino As Workbook
    Dim wbOrigen As Workbook
    Dim contador As Long
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar carpeta con archivos"
        If .Show <> -1 Then Exit Sub
        rutaCarpeta = .SelectedItems(1)
    End With
    
    If Right(rutaCarpeta, 1) <> "\" Then rutaCarpeta = rutaCarpeta & "\"
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Crear libro destino
    Set wbDestino = Workbooks.Add(xlWorksheet)
    contador = 0
    
    ' Recorrer archivos
    nombreArchivo = Dir(rutaCarpeta & "*.xls*")
    
    Do While nombreArchivo <> ""
        ' Abrir archivo origen
        Set wbOrigen = Workbooks.Open( _
            fileName:=rutaCarpeta & nombreArchivo, _
            ReadOnly:=True)
        
        ' Copiar hojas al destino
        Dim sht As Worksheet
        For Each sht In wbOrigen.Sheets
            sht.Copy After:=wbDestino.Sheets(wbDestino.Sheets.Count)
            
            ' Renombrar (máx 29 chars para dejar espacio)
            Dim nuevoNombre As String
            nuevoNombre = Left(nombreArchivo, InStrRev(nombreArchivo, ".") - 1)
            If wbOrigen.Sheets.Count > 1 Then
                nuevoNombre = nuevoNombre & "_" & sht.Index
            End If
            wbDestino.Sheets(wbDestino.Sheets.Count).Name = nuevoNombre
        Next sht
        
        wbOrigen.Close SaveChanges:=False
        contador = contador + 1
        nombreArchivo = Dir()
    Loop
    
    ' Eliminar hoja inicial vacía
    Application.DisplayAlerts = False
    wbDestino.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox contador & " archivos combinados.", vbInformation
End Sub


' ============================================================
' COMBINAR DOS ARCHIVOS ESPECÍFICOS
' Fusiona datos de dos archivos en una hoja
' ============================================================
Public Sub CombinarDosArchivos()
    Dim rutaArchivo1 As String
    Dim rutaArchivo2 As String
    Dim wb1 As Workbook
    Dim wb2 As Workbook
    Dim wsDestino As Worksheet
    
    ' Seleccionar primer archivo
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Seleccionar primer archivo"
        .Filters.Clear
        .Filters.Add "Excel", "*.xls; *.xlsx; *.xlsm"
        If .Show <> -1 Then Exit Sub
        rutaArchivo1 = .SelectedItems(1)
    End With
    
    ' Seleccionar segundo archivo
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Seleccionar segundo archivo"
        .Filters.Clear
        .Filters.Add "Excel", "*.xls; *.xlsx; *.xlsm"
        If .Show <> -1 Then Exit Sub
        rutaArchivo2 = .SelectedItems(1)
    End With
    
    Application.ScreenUpdating = False
    
    ' Abrir archivos
    Set wb1 = Workbooks.Open(rutaArchivo1, ReadOnly:=True)
    Set wb2 = Workbooks.Open(rutaArchivo2, ReadOnly:=True)
    
    ' Crear hoja destino en libro actual
    Set wsDestino = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsDestino.Name = "Combinado_" & Format(Now, "HHMMSS")
    
    ' Copiar datos de primer archivo
    wb1.Sheets(1).UsedRange.Copy wsDestino.Range("A1")
    
    ' Copiar datos de segundo archivo debajo
    Dim ultimaFila As Long
    ultimaFila = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).Row + 1
    wb2.Sheets(1).UsedRange.Copy wsDestino.Cells(ultimaFila, 1)
    
    ' Cerrar archivos originales
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    
    MsgBox "Archivos combinados correctamente.", vbInformation
End Sub


' ============================================================
' COMBINAR CON NOMBRE DE ARCHIVO COLUMNA
' Agrega columna con origen de cada registro
' ============================================================
Public Sub CombinaConOrigen()
    Dim rutaCarpeta As String
    Dim nombreArchivo As String
    Dim wbOrigen As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar carpeta"
        If .Show <> -1 Then Exit Sub
        rutaCarpeta = .SelectedItems(1)
    End With
    
    If Right(rutaCarpeta, 1) <> "\" Then rutaCarpeta = rutaCarpeta & "\"
    
    Application.ScreenUpdating = False
    
    Set wsDestino = ThisWorkbook.Sheets.Add
    wsDestino.Name = "Consolidado"
    
    ' Encabezados
    wsDestino.Range("A1").Value = "Archivo Origen"
    wsDestino.Range("B1").Value = "Dato1"
    wsDestino.Range("C1").Value = "Dato2"
    wsDestino.Range("A1:C1").Font.Bold = True
    
    ultimaFila = 2
    
    nombreArchivo = Dir(rutaCarpeta & "*.xl*")
    
    Do While nombreArchivo <> ""
        Set wbOrigen = Workbooks.Open(rutaCarpeta & nombreArchivo, ReadOnly:=True)
        Set wsOrigen = wbOrigen.Sheets(1)
        
        ' Copiar datos con origen
        Dim i As Long
        For i = 2 To wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
            wsDestino.Cells(ultimaFila, 1).Value = nombreArchivo
            wsDestino.Cells(ultimaFila, 2).Value = wsOrigen.Cells(i, 1).Value
            wsDestino.Cells(ultimaFila, 3).Value = wsOrigen.Cells(i, 2).Value
            ultimaFila = ultimaFila + 1
        Next i
        
        wbOrigen.Close SaveChanges:=False
        nombreArchivo = Dir()
    Loop
    
    wsDestino.Columns("A:C").AutoFit
    Application.ScreenUpdating = True
    
    MsgBox "Consolidado completado.", vbInformation
End Sub
```

## Descripción de Macros

| Macro | Descripción |
|-------|-------------|
| `CombinarArchivosCarpeta` | Combina todos los Excel de una carpeta |
| `CombinarDosArchivos` | Fusiona dos archivos específicos |
| `CombinaConOrigen` | Agrega columna con nombre de archivo origen |

---

*Módulo: 06_OperacionesHojas | Archivo: 06_CombinarDosWorkbooks.md*
