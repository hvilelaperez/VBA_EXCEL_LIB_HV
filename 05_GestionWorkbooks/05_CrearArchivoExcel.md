# Crear Archivo Excel

Crea nuevos libros Excel con datos iniciales y los guarda en ubicación específica.

## Código VBA

```vba
Option Explicit

' ============================================================
' CREAR LIBRO CON DATOS INICIALES
' Crea un libro, agrega datos y lo guarda
' ============================================================
Public Sub CrearLibroConDatos()
    Dim wbNuevo As Workbook
    Dim ws As Worksheet
    Dim rutaGuardado As String
    
    rutaGuardado = ThisWorkbook.Path & "\NuevoLibro.xlsx"
    
    ' Crear nuevo libro
    Set wbNuevo = Workbooks.Add
    Set ws = wbNuevo.Worksheets(1)
    
    ' Agregar datos de ejemplo
    ws.Cells(1, 1).Value = "Nombre"
    ws.Cells(1, 2).Value = "Edad"
    ws.Cells(1, 3).Value = "Ciudad"
    
    ws.Cells(2, 1).Value = "Juan"
    ws.Cells(2, 2).Value = 25
    ws.Cells(2, 3).Value = "Lima"
    
    ' Guardar y cerrar
    With wbNuevo
        .SaveAs fileName:=rutaGuardado
        .Close
    End With
    
    MsgBox "Libro creado: " & rutaGuardado, vbInformation
End Sub


' ============================================================
' CREAR LIBRO CON FORMATO ESPECÍFICO
' Usa xlOpenXMLWorkbook para formato moderno
' ============================================================
Public Sub CrearLibroConFormato()
    Dim wbNuevo As Workbook
    Dim ws As Worksheet
    Dim rutaGuardado As String
    
    rutaGuardado = ThisWorkbook.Path & "\LibroFormato.xlsx"
    
    Set wbNuevo = Workbooks.Add
    Set ws = wbNuevo.Worksheets(1)
    
    ' Encabezados con formato
    ws.Range("A1:C1").Value = Array("Producto", "Precio", "Cantidad")
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(200, 200, 255)
    
    ' Datos
    ws.Range("A2:C4").Value = Array( _
        Array("Lápiz", 0.5, 100), _
        Array("Cuaderno", 2.5, 50), _
        Array("Bolígrafo", 1.0, 75))
    
    ' Ajustar columnas
    ws.Columns("A:C").AutoFit
    
    ' Guardar
    With wbNuevo
        .SaveAs fileName:=rutaGuardado, _
                FileFormat:=xlOpenXMLWorkbook
        .Close
    End With
    
    MsgBox "Libro con formato creado.", vbInformation
End Sub


' ============================================================
' CREAR MÚLTIPLES LIBROS
' Crea varios libros basados en lista
' ============================================================
Public Sub CrearMultiplesLibros()
    Dim wsLista As Worksheet
    Dim i As Long
    Dim nombreLibro As String
    Dim rutaBase As String
    Dim wbNuevo As Workbook
    Dim contador As Long
    
    Set wsLista = ThisWorkbook.Sheets("Lista")
    rutaBase = ThisWorkbook.Path & "\"
    contador = 0
    
    ' Recorrer lista de nombres
    For i = 2 To wsLista.Cells(wsLista.Rows.Count, 1).End(xlUp).Row
        nombreLibro = Trim(wsLista.Cells(i, 1).Value)
        
        If nombreLibro <> "" Then
            Set wbNuevo = Workbooks.Add
            wbNuevo.Sheets(1).Name = "Datos"
            wbNuevo.Sheets(1).Cells(1, 1).Value = "Contenido de " & nombreLibro
            
            wbNuevo.SaveAs fileName:=rutaBase & nombreLibro & ".xlsx"
            wbNuevo.Close
            
            contador = contador + 1
        End If
    Next i
    
    MsgBox contador & " libros creados.", vbInformation
End Sub


' ============================================================
' CREAR LIBRO CON HOJA SEGÚN PLANTILLA
' ============================================================
Public Sub CrearConPlantilla()
    Dim wbNuevo As Workbook
    Dim ws As Worksheet
    Dim rutaPlantilla As String
    
    ' Crear libro con hoja en blanco
    Set wbNuevo = Workbooks.Add(xlWorksheet)
    Set ws = wbNuevo.Worksheets(1)
    
    ' Configurar hoja
    ws.Name = "Reporte"
    ws.Cells(1, 1).Value = "REPORTE MENSUAL"
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    
    ws.Cells(3, 1).Value = "Fecha:"
    ws.Cells(3, 2).Value = Date
    ws.Cells(4, 1).Value = "Generado por:"
    ws.Cells(4, 2).Value = Application.UserName
    
    ' Guardar
    wbNuevo.SaveAs ThisWorkbook.Path & "\Reporte_" & Format(Date, "YYYYMMDD") & ".xlsx"
    wbNuevo.Close
    
    MsgBox "Reporte creado.", vbInformation
End Sub
```

## Descripción de Macros

| Macro | Descripción |
|-------|-------------|
| `CrearLibroConDatos` | Crea libro con datos de ejemplo |
| `CrearLibroConFormato` | Crea libro con formato y colores |
| `CrearMultiplesLibros` | Crea varios libros desde una lista |
| `CrearConPlantilla` | Crea libro con estructura predefinida |

## Constantes de FileFormat

| Constante | Formato |
|-----------|---------|
| `xlOpenXMLWorkbook` | .xlsx |
| `xlOpenXMLWorkbookMacroEnabled` | .xlsm |
| `xlExcel12` | .xlsb |

---

*Módulo: 05_GestionWorkbooks | Archivo: 05_CrearArchivoExcel.md*
