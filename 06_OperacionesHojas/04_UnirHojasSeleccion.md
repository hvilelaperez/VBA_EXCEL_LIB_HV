# Unir Hojas de Selección

Combina datos de múltiples hojas seleccionadas en la hoja activa.

## Código VBA

```vba
Option Explicit

' ============================================================
' UNIR HOJAS SELECCIONADAS
' Copia datos de hojas seleccionadas a la hoja activa
' ============================================================
Public Sub UnirHojasSeleccionadas()
    Dim wsActiva As Worksheet
    Dim wsOrigen As Worksheet
    Dim ultimaFilaOrigen As Long
    Dim primeraFilaVacia As Long
    Dim hojaNombre As String
    
    Set wsActiva = ActiveSheet
    
    ' Verificar que hay más de una hoja seleccionada
    If ActiveWindow.SelectedSheets.Count < 2 Then
        MsgBox "Seleccione al menos 2 hojas antes de ejecutar.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    For Each wsOrigen In ActiveWindow.SelectedSheets
        ' Saltar la hoja activa
        If Not wsOrigen Is wsActiva Then
            ' Encontrar última fila con datos en hoja origen
            ultimaFilaOrigen = wsOrigen.UsedRange.Cells(wsOrigen.UsedRange.Cells.Count).Row
            
            ' Encontrar primera fila vacía en hoja activa
            primeraFilaVacia = wsActiva.UsedRange.Cells(wsActiva.UsedRange.Cells.Count).Row + 1
            
            ' Copiar datos (excluyendo encabezado si es necesario)
            wsOrigen.Range(wsOrigen.Rows(2), wsOrigen.Rows(ultimaFilaOrigen)).Copy _
                wsActiva.Rows(primeraFilaVacia)
        End If
    Next wsOrigen
    
    Application.ScreenUpdating = True
    
    MsgBox "Hojas unidas correctamente.", vbInformation
End Sub


' ============================================================
' UNIR HOJAS CON SELECCIÓN MÚLTIPLE DE ARCHIVOS
' Selecciona varias hojas de diferentes libros
' ============================================================
Public Sub UnirHojasDeMultiplesLibros()
    Dim dialogo As FileDialog
    Dim rutaArchivo As Variant
    Dim wbOrigen As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim primeraFilaVacia As Long
    Dim contador As Long
    
    Set dialogo = Application.FileDialog(msoFileDialogOpen)
    dialogo.Title = "Seleccionar archivos a unir"
    dialogo.AllowMultiSelect = True
    dialogo.Filters.Clear
    dialogo.Filters.Add "Excel", "*.xls; *.xlsx; *.xlsm"
    
    If dialogo.Show <> -1 Then Exit Sub
    
    Application.ScreenUpdating = False
    Set wsDestino = ActiveSheet
    contador = 0
    
    For Each rutaArchivo In dialogo.SelectedItems
        Set wbOrigen = Workbooks.Open(CStr(rutaArchivo), ReadOnly:=True)
        
        For Each wsOrigen In wbOrigen.Sheets
            ' Encontrar última fila con datos
            Dim ultimaFila As Long
            ultimaFila = wsOrigen.UsedRange.Cells(wsOrigen.UsedRange.Cells.Count).Row
            
            ' Encontrar espacio disponible
            primeraFilaVacia = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).Row + 1
            
            ' Copiar datos
            If ultimaFila > 1 Then
                wsOrigen.Range("A2:A" & ultimaFila).Copy wsDestino.Cells(primeraFilaVacia, 1)
                contador = contador + (ultimaFila - 1)
            End If
        Next wsOrigen
        
        wbOrigen.Close SaveChanges:=False
    Next rutaArchivo
    
    Application.ScreenUpdating = True
    
    MsgBox "Registros unidos: " & contador, vbInformation
End Sub


' ============================================================
' FUNCIÓN AUXILIAR: Obtener última fila con datos
' ============================================================
Private Function ObtenerUltimaFila(ws As Worksheet, Optional columna As Long = 1) As Long
    ObtenerUltimaFila = ws.Cells(ws.Rows.Count, columna).End(xlUp).Row
End Function


' ============================================================
' FUNCIÓN AUXILIAR: Obtener última columna con datos
' ============================================================
Private Function ObtenerUltimaColumna(ws As Worksheet, Optional fila As Long = 1) As Long
    ObtenerUltimaColumna = ws.Cells(fila, ws.Columns.Count).End(xlToLeft).Column
End Function
```

## Flujo de Ejecución

```
Hoja1 (origen)      Hoja2 (origen)
┌──────────────┐    ┌──────────────┐
│ Encabezado   │    │ Encabezado   │
│ Dato1        │    │ Dato4        │
│ Dato2        │    │ Dato5        │
│ Dato3        │    │ Dato6        │
└──────────────┘    └──────────────┘
         │                  │
         └────────┬─────────┘
                  ▼
         ┌──────────────────┐
         │ Hoja Activa      │
         ├──────────────────┤
         │ Encabezado       │
         │ Dato1            │
         │ Dato2            │
         │ Dato3            │
         │ Dato4            │
         │ Dato5            │
         │ Dato6            │
         └──────────────────┘
```

## Notas

- Selecciona las hojas con `Ctrl + clic` antes de ejecutar
- Omite el encabezado de cada hoja origen (fila 1)
- Copia solo la primera columna (A)

---

*Módulo: 06_OperacionesHojas | Archivo: 04_UnirHojasSeleccion.md*
