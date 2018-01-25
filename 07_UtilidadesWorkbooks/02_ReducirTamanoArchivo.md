# Reducir Tamaño de Archivo Excel

Limpia archivos Excel eliminando espacio no utilizado y reduciendo su tamaño.

## Código VBA

```vba
Option Explicit

' ============================================================
' REDUCIR TAMAÑO DE ARCHIVO (ExcelDiet)
' Elimina filas y columnas no utilizadas
' ============================================================
Public Sub ReducirTamanoArchivo()
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim ultimaColumna As Long
    Dim rangoEliminar As Range
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        
        ' Encontrar última fila con fórmulas
        Dim colFormula As Range
        Set colFormula = ws.Cells.Find(What:="*", After:=ws.Range("A1"), _
            LookIn:=xlFormulas, LookAt:=xlPart, _
            SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
        
        ' Encontrar última fila con valores
        Dim colValor As Range
        Set colValor = ws.Cells.Find(What:="*", After:=ws.Range("A1"), _
            LookIn:=xlValues, LookAt:=xlPart, _
            SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
        
        ' Encontrar última fila con fórmulas
        Dim filaFormula As Range
        Set filaFormula = ws.Cells.Find(What:="*", After:=ws.Range("A1"), _
            LookIn:=xlFormulas, LookAt:=xlPart, _
            SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        
        ' Encontrar última fila con valores
        Dim filaValor As Range
        Set filaValor = ws.Cells.Find(What:="*", After:=ws.Range("A1"), _
            LookIn:=xlValues, LookAt:=xlPart, _
            SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        
        On Error GoTo 0
        
        ' Determinar última columna
        ultimaColumna = 0
        If Not colFormula Is Nothing Then ultimaColumna = colFormula.Column
        If Not colValor Is Nothing Then
            ultimaColumna = Application.WorksheetFunction.Max(ultimaColumna, colValor.Column)
        End If
        
        ' Determinar última fila
        ultimaFila = 0
        If Not filaFormula Is Nothing Then ultimaFila = filaFormula.Row
        If Not filaValor Is Nothing Then
            ultimaFila = Application.WorksheetFunction.Max(ultimaFila, filaValor.Row)
        End If
        
        ' Verificar shapes (objetos gráficos)
        Dim shp As Shape
        For Each shp In ws.Shapes
            Dim j As Long, k As Long
            j = 0: k = 0
            On Error Resume Next
            j = shp.TopLeftCell.Row
            k = shp.TopLeftCell.Column
            On Error GoTo 0
            
            If j > 0 And k > 0 Then
                Do Until ws.Cells(j, k).Top > shp.Top + shp.Height
                    j = j + 1
                Loop
                If j > ultimaFila Then ultimaFila = j
                
                Do Until ws.Cells(j, k).Left > shp.Left + shp.Width
                    k = k + 1
                Loop
                If k > ultimaColumna Then ultimaColumna = k
            End If
        Next shp
        
        ' Eliminar filas y columnas excedentes
        If ultimaColumna < ws.Columns.Count Then
            ws.Range(ws.Cells(1, ultimaColumna + 1).Address & ":IV65536").Delete
        End If
        
        If ultimaFila < ws.Rows.Count Then
            ws.Range(ws.Cells(ultimaFila + 1, 1).Address & ":IV65536").Delete
        End If
    Next ws
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Espacio no utilizado eliminado.", vbInformation
End Sub


' ============================================================
' LIMPIAR CELDAS VACÍAS EXCEDENTES
' Elimina formato de celdas fuera del rango utilizado
' ============================================================
Public Sub LimpiarCeldasExcedentes()
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim ultimaColumna As Long
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
        ultimaFila = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
        ultimaColumna = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
        
        ' Eliminar formato de celdas excedentes
        If ultimaFila < 1000 Then
            ws.Range(ws.Cells(ultimaFila + 1, 1), ws.Cells(1000, ws.Columns.Count)).ClearFormats
        End If
    Next ws
    
    Application.ScreenUpdating = True
    
    MsgBox "Celdas excedentes limpiadas.", vbInformation
End Sub


' ============================================================
' ELIMINAR OBJETOS INNECESARIOS
' Elimina formas, imágenes, etc.
' ============================================================
Public Sub EliminarObjetosInnecesarios()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim contador As Long
    
    contador = 0
    
    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            ' Eliminar todas las formas (cuidado: irreversible)
            shp.Delete
            contador = contador + 1
        Next shp
    Next ws
    
    MsgBox contador & " objetos eliminados.", vbInformation
End Sub


' ============================================================
' MOSTRAR TAMAÑO ACTUAL DEL ARCHIVO
' ============================================================
Public Sub MostrarTamanoArchivo()
    Dim tamano As Double
    
    tamano = FileLen(ThisWorkbook.FullName) / 1024 / 1024 ' En MB
    
    MsgBox "Tamaño actual del archivo: " & Format(tamano, "0.00") & " MB", vbInformation
End Sub
```

## Descripción de Macros

| Macro | Descripción |
|-------|-------------|
| `ReducirTamanoArchivo` | Elimina filas/columnas no utilizadas |
| `LimpiarCeldasExcedentes` | Limpia formato de celdas vacías |
| `EliminarObjetosInnecesarios` | Elimina shapes, imágenes, etc. |
| `MostrarTamanoArchivo` | Muestra tamaño en MB |

## Consejos Adicionales

1. **Guardar como**: Archivo > Guardar como > Herramientas > Opciones generales
2. **Comprimir imágenes**: Seleccionar imagen > Formato > Comprimir
3. **Eliminar hojas ocultas**: Desocultar y eliminar si no son necesarias
4. **Convertir fórmulas a valores**: Reduce dependencias

## Nota

- `ReducirTamanoArchivo` es similar a la macro ExcelDiet original
- Siempre hacer copia de seguridad antes de ejecutar
- El resultado puede variar según el tipo de contenido

---

*Módulo: 07_UtilidadesWorkbooks | Archivo: 02_ReducirTamanoArchivo.md*
