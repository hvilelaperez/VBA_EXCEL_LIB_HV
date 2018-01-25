# Eliminar Millón de Filas

Métodos eficientes para eliminar grandes volúmenes de filas (hasta 1 millón) en Excel usando VBA, con optimización de rendimiento.

## Código VBA

```vba
Sub DeleteRowsWithValuesStrings()
    Const MAX_SZ As Byte = 240

    Dim i As Long, j As Long, t As Double, ws As Worksheet
    Dim memArr As Variant, max As Long, tmp As String

    Set ws = Worksheets(1)
    max = GetMaxCell(ws.UsedRange).Row
    FastWB True:    t = Timer

    With ws
        If max > 1 Then
            If IndexOfValInRowOrCol("Test String", , ws.UsedRange) > 0 Then
                memArr = .Range(.Cells(1, 1), .Cells(max, 1)).Value2
                For i = max To 1 Step -1

                    If memArr(i, 1) = "Test String" Then
                        tmp = tmp & "A" & i & ","
                        If Len(tmp) > MAX_SZ Then
                           .Range(Left(tmp, Len(tmp) - 1)).EntireRow.Delete Shift:=xlUp
                           tmp = vbNullString

                        End If
                    End If

                Next
                If Len(tmp) > 0 Then
                    .Range(Left(tmp, Len(tmp) - 1)).EntireRow.Delete Shift:=xlUp
                End If
                .Calculate
            End If
        End If
    End With
    FastWB False:   InputBox "Duración: ", "Duración", Timer - t
End Sub


Public Sub FastWB(Optional ByVal opt As Boolean = True)
    With Application
        .Calculation = IIf(opt, xlCalculationManual, xlCalculationAutomatic)
        .DisplayAlerts = Not opt
        .DisplayStatusBar = Not opt
        .EnableAnimations = Not opt
        .EnableEvents = Not opt
        .ScreenUpdating = Not opt
    End With
    FastWS , opt
End Sub

Public Sub FastWS(Optional ByVal ws As Worksheet = Nothing, _
                  Optional ByVal opt As Boolean = True)
    If ws Is Nothing Then
        For Each ws In Application.ActiveWorkbook.Sheets
            EnableWS ws, opt
        Next
    Else
        EnableWS ws, opt
    End If
End Sub

Private Sub EnableWS(ByVal ws As Worksheet, ByVal opt As Boolean)
    With ws
        .DisplayPageBreaks = False
        .EnableCalculation = Not opt
        .EnableFormatConditionsCalculation = Not opt
        .EnablePivotTable = Not opt
    End With
End Sub



Public Function GetMaxCell(Optional ByRef rng As Range = Nothing) As Range

    'Devuelve la última celda con valor, o A1 si la hoja está vacía

    Const NONEMPTY As String = "*"
    Dim lRow As Range, lCol As Range

    If rng Is Nothing Then Set rng = Application.ActiveWorkbook.ActiveSheet.UsedRange
    If WorksheetFunction.CountA(rng) = 0 Then
        Set GetMaxCell = rng.Parent.Cells(1, 1)
    Else
        With rng
            Set lRow = .Cells.Find(What:=NONEMPTY, LookIn:=xlFormulas, _
                                        After:=.Cells(1, 1), _
                                        SearchDirection:=xlPrevious, _
                                        SearchOrder:=xlByRows)
            If Not lRow Is Nothing Then
                Set lCol = .Cells.Find(What:=NONEMPTY, LookIn:=xlFormulas, _
                                            After:=.Cells(1, 1), _
                                            SearchDirection:=xlPrevious, _
                                            SearchOrder:=xlByColumns)

                Set GetMaxCell = .Parent.Cells(lRow.Row, lCol.Column)
            End If
        End With
    End If
End Function


Public Function IndexOfValInRowOrCol( _
                                    ByVal searchVal As String, _
                                    Optional ByRef ws As Worksheet = Nothing, _
                                    Optional ByRef rng As Range = Nothing, _
                                    Optional ByRef vertical As Boolean = True, _
                                    Optional ByRef rowOrColNum As Long = 1 _
                                    ) As Long

    'Devuelve la posición en fila o columna, o 0 si no se encuentran coincidencias

    Dim usedRng As Range, result As Variant, searchRow As Long, searchCol As Long

    result = CVErr(9999) '- generar error personalizado

    Set usedRng = GetUsedRng(ws, rng)
    If Not usedRng Is Nothing Then
        If rowOrColNum < 1 Then rowOrColNum = 1
        With Application
            If vertical Then
                result = .Match(searchVal, rng.Columns(rowOrColNum), 0)
            Else
                result = .Match(searchVal, rng.Rows(rowOrColNum), 0)
            End If
        End With
    End If
    If IsError(result) Then IndexOfValInRowOrCol = 0 Else IndexOfValInRowOrCol = result
End Function

Sub DeleteRowsWithValuesNewSheet()  '100K registros   10K a eliminar
                                    'Prueba 1:        2.40234375 seg
                                    'Prueba 2:        2.41796875 seg
                                    'Prueba 3:        2.40234375 seg
                                    '1M registros     100K a eliminar
                                    'Prueba 1:        32.9140625 seg
                                    'Prueba 2:        33.1484375 seg
                                    'Prueba 3:        32.90625   seg
    Dim oldWs As Worksheet, newWs As Worksheet, rowHeights() As Long
    Dim wsName As String, t As Double, oldUsedRng As Range

    FastWB True:    t = Timer

    Set oldWs = Worksheets(1)
    wsName = oldWs.Name

    Set oldUsedRng = oldWs.Range("A1", GetMaxCell(oldWs.UsedRange))

    If oldUsedRng.Rows.Count > 1 Then                           'Si la hoja no está vacía
        Set newWs = Sheets.Add(After:=oldWs)                    'Agregar nueva hoja
        With oldUsedRng
            .AutoFilter Field:=1, Criteria1:="<>Test String"
            .Copy                                               'Copiar datos visibles
        End With
        With newWs.Cells
            .PasteSpecial xlPasteColumnWidths
            .PasteSpecial xlPasteAll                            'Pegar datos en nueva hoja
            .Cells(1, 1).Select                                 'Deseleccionar área de pegado
            .Cells(1, 1).Copy                                   'Limpiar portapapeles
        End With
        oldWs.Delete                                            'Eliminar hoja antigua
        newWs.Name = wsName
    End If
    FastWB False:   InputBox "Duración: ", "Duración", Timer - t
End Sub




Public Sub ExcelHero()
    Dim t#, crit As Range, data As Range, ws As Worksheet
    Dim r&, fc As Range, lc As Range, fr1 As Range, fr2 As Range
    FastWB True
    t = Timer

        Set fc = ActiveSheet.UsedRange.Item(1)
        Set lc = GetMaxCell
        Set data = ActiveSheet.Range(fc, lc)
        Set ws = Sheets.Add
        With data
            Set fr1 = data.Worksheet.Range(fc, fc.Offset(, lc.Column))
            Set fr2 = ws.Range(ws.Cells(fc.Row, fc.Column), ws.Cells(fc.Row, lc.Column))
            With fr2
                fr1.Copy
                .PasteSpecial xlPasteColumnWidths: .PasteSpecial xlPasteAll
                .Item(1).Select
            End With
            Set crit = .Resize(2, 1).Offset(, lc.Column + 1)
            crit = [{"Column 1";"<>Test String"}]
            .AdvancedFilter xlFilterCopy, crit, fr2
            .Worksheet.Delete
        End With

    FastWB False
    r = ws.UsedRange.Rows.Count
    Debug.Print "Filas: " & r & ", Duración: " & Timer - t & " segundos"
End Sub
```

## Descripción

| Método | Descripción | Rendimiento |
|--------|-------------|-------------|
| **DeleteRowsWithValuesStrings** | Elimina filas agrupando rangos de dirección | ~2.4 seg (100K registros) |
| **DeleteRowsWithValuesNewSheet** | Crea nueva hoja con datos filtrados y elimina la original | ~33 seg (1M registros) |
| **ExcelHero** | Usa AdvancedFilter para copiar datos a nueva hoja | Rápido para grandes volúmenes |

### Funciones auxiliares

| Función | Descripción |
|---------|-------------|
| `FastWB` | Optimiza el rendimiento desactivando cálculos, eventos y pantalla |
| `FastWS` | Optimiza hojas de trabajo individuales |
| `GetMaxCell` | Encuentra la última celda con contenido en un rango |
| `IndexOfValInRowOrCol` | Busca un valor y devuelve su posición |

### Optimización de rendimiento

Las funciones `FastWB` y `FastWS` desactivan:
- Cálculo automático
- Alertas de pantalla
- Barra de estado
- Animaciones
- Eventos
- Actualización de pantalla

---

*Módulo: 14 - VBA Autofilter | Archivo: 01 - Delete 1 millions rows.txt*
