# Array Bidimensional para Filtrar Datos

Código VBA para procesar y filtrar datos usando un array bidimensional, concatenando valores y transfiriendo resultados a otra hoja.

## Código VBA

```vba
Sub FILTER_DATA()
Dim T1 As Long, T2 As Long, last_Column As Integer
Dim Data As Variant, i As Long, j As Long, TMP As Variant, MA As String, T As Double
Data = SHEET1.Range("A2:C" & SHEET1.Cells(Rows.Count, "C").End(xlUp).Row).Value2

T1 = 0
T1 = GetTickCount
j = 0
ReDim TMP(1 To 3, 1 To 1)
For i = 1 To UBound(Data, 1)
    MA = CStr(Data(i, 1) & Data(i, 2))
    If MA <> vbNullString Then
        j = j + 1
        ReDim Preserve TMP(1 To 3, 1 To j)
        TMP(1, j) = Data(i, 1)
        TMP(2, j) = Data(i, 2)
        TMP(3, j) = Data(i, 3)
    Else
        TMP(3, j) = TMP(3, j) & " ; " & Data(i, 3)
    End If
    
Next
Sheet2.Range("A2:C1048576").ClearContents
'On Error Resume Next
Sheet2.Activate
ReDim Preserve TMP(1 To 3, 1 To j)
'Debug.Print j
Sheet2.Range("A2").Resize(3, UBound(TMP, 2)).Value = TMP
Sheet2.Range("A2:" & Sheet2.Cells(4, Columns.Count).End(xlToLeft).Address).Copy
Sheet2.Range("D10").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

'Debug.Print last_Column
'Sheet3.Range("A2").Resize(UBound(TMP, 2), 3).Value = Application.WorksheetFunction.Transpose(TMP)
T2 = GetTickCount
MsgBox ((T2 - T1) / 1000), vbInformation, "El tiempo de ejecución en segundos es "
 'On Error Resume Next
End Sub
```

## Descripción

| Variable | Función |
|----------|---------|
| `Data` | Array bidimensional con datos origen (columnas A:C) |
| `TMP` | Array bidimensional de resultados filtrados |
| `MA` | Cadena concatenada de columnas 1 y 2 para detectar vacíos |
| `T1/T2` | Tiempos de inicio y fin para medir rendimiento |

### Lógica del filtro

1. Lee todos los datos desde A2 hasta la última fila con datos en columna C
2. Para cada fila, concatena columnas A y B
3. Si la concatenación no está vacía → crea nueva entrada en TMP
4. Si está vacía → concatena el valor de columna C al registro anterior con `; `
5. Voltea los resultados y los pega en Sheet2 comenzando en D10

### Características

- **Rendimiento**: Mide tiempo de ejecución con `GetTickCount`
- **Limpieza**: Borra contenido previo en Sheet2 antes de pegar
- **Transposición**: Convierte filas a columnas al pegar

---

*Módulo: 05 - VBA Array | Archivo: 01 - Noi du lieu bang cach su dung mang 2 chieu*