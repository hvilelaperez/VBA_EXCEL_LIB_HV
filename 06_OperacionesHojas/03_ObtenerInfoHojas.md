# Obtener Información de Hojas

Obtiene lista de todas las hojas de un libro y sus propiedades.

## Código VBA

```vba
Option Explicit

' ============================================================
' OBTENER NOMBRES DE HOJAS
' Lista todos los nombres de hojas en el libro actual
' ============================================================
Public Sub ObtenerNombresHojas()
    Dim ws As Worksheet
    Dim fila As Long
    
    ' Limpiar hoja de resultados
    ActiveSheet.Cells.Clear
    
    ' Encabezados
    ActiveSheet.Cells(1, 1).Value = "Nombre"
    ActiveSheet.Cells(1, 2).Value = "Índice"
    ActiveSheet.Cells(1, 3).Value = "Visibilidad"
    ActiveSheet.Range("A1:C1").Font.Bold = True
    
    fila = 2
    
    For Each ws In ThisWorkbook.Worksheets
        ActiveSheet.Cells(fila, 1).Value = ws.Name
        ActiveSheet.Cells(fila, 2).Value = ws.Index
        
        Select Case ws.Visible
            Case xlSheetVisible
                ActiveSheet.Cells(fila, 3).Value = "Visible"
            Case xlSheetHidden
                ActiveSheet.Cells(fila, 3).Value = "Oculta"
            Case xlSheetVeryHidden
                ActiveSheet.Cells(fila, 3).Value = "Muy Oculta"
        End Select
        
        fila = fila + 1
    Next ws
    
    MsgBox "Se encontraron " & (fila - 2) & " hojas.", vbInformation
End Sub


' ============================================================
' OBTENER HOJAS DE OTRO LIBRO
' Lista las hojas de un archivo específico
' ============================================================
Public Sub ObtenerHojasOtroLibro()
    Dim rutaArchivo As String
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' Seleccionar archivo
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Seleccionar archivo Excel"
        .Filters.Clear
        .Filters.Add "Excel", "*.xls; *.xlsx; *.xlsm"
        
        If .Show <> -1 Then Exit Sub
        rutaArchivo = .SelectedItems(1)
    End With
    
    ' Abrir en modo lectura
    Set wb = Workbooks.Open(rutaArchivo, ReadOnly:=True)
    
    ' Listar hojas
    Dim mensaje As String
    mensaje = "Hojas en " & wb.Name & ":" & vbCrLf & vbCrLf
    
    For Each ws In wb.Worksheets
        mensaje = mensaje & "- " & ws.Name & vbCrLf
    Next ws
    
    wb.Close SaveChanges:=False
    
    MsgBox mensaje, vbInformation
End Sub


' ============================================================
' CONTAR HOJAS POR TIPO
' Cuenta hojas visibles, ocultas, etc.
' ============================================================
Public Sub ContarHojasPorTipo()
    Dim ws As Worksheet
    Dim visibles As Long
    Dim ocultas As Long
    Dim muyOcultas As Long
    
    visibles = 0
    ocultas = 0
    muyOcultas = 0
    
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Visible
            Case xlSheetVisible: visibles = visibles + 1
            Case xlSheetHidden: ocultas = ocultas + 1
            Case xlSheetVeryHidden: muyOcultas = muyOcultas + 1
        End Select
    Next ws
    
    MsgBox "Resumen de hojas:" & vbCrLf & _
           "Visibles: " & visibles & vbCrLf & _
           "Ocultas: " & ocultas & vbCrLf & _
           "Muy ocultas: " & muyOcultas & vbCrLf & _
           "Total: " & (visibles + ocultas + muyOcultas), vbInformation
End Sub


' ============================================================
' RENOMBRAR HOJAS
' Renombra hojas basándose en contenido
' ============================================================
Public Sub RenombrarHojas()
    Dim ws As Worksheet
    Dim nuevoNombre As String
    
    For Each ws In ThisWorkbook.Worksheets
        ' Usar valor de celda A1 como nuevo nombre
        nuevoNombre = Trim(CStr(ws.Range("A1").Value))
        
        If nuevoNombre <> "" And nuevoNombre <> ws.Name Then
            On Error Resume Next
            ws.Name = nuevoNombre
            On Error GoTo 0
        End If
    Next ws
    
    MsgBox "Proceso de renombrado completado.", vbInformation
End Sub
```

## Descripción de Macros

| Macro | Descripción |
|-------|-------------|
| `ObtenerNombresHojas` | Lista nombres e índice de hojas |
| `ObtenerHojasOtroLibro` | Lista hojas de archivo externo |
| `ContarHojasPorTipo` | Cuenta por visibilidad |
| `RenombrarHojas` | Renombra según contenido |

## Constantes de Visibilidad

| Constante | Valor | Descripción |
|-----------|-------|-------------|
| `xlSheetVisible` | -1 | Hoja visible |
| `xlSheetHidden` | 0 | Hoja oculta |
| `xlSheetVeryHidden` | 2 | Hoja muy oculta (solo VBA) |

---

*Módulo: 06_OperacionesHojas | Archivo: 03_ObtenerInfoHojas.md*
