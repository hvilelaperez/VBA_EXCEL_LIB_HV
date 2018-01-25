# Verificar Existencia de Hoja

Múltiples formas de verificar si una hoja de trabajo existe en un libro.

## Código VBA

```vba
Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
  Dim sht As Worksheet

  If wb Is Nothing Then Set wb = ThisWorkbook
  On Error Resume Next
  Set sht = wb.Sheets(shtName)
  On Error GoTo 0
  SheetExists = Not sht Is Nothing
End Function



Function CheckIfSheetExists(SheetName As String) As Boolean
  CheckIfSheetExists = False
  For Each WS In Worksheets
    If SheetName = WS.name Then
      CheckIfSheetExists = True
      Exit Function
    End If
  Next WS
End Function


Function isSheetExist(shName As String) As Boolean
  isSheetExist = Not Evaluate("IsError(" & shName & "!1:1)")
End Function


Private Function SheetExists(sname) As Boolean
'   Devuelve VERDADERO si la hoja existe en el libro activo
    Dim x As Object
    On Error Resume Next
    Set x = ActiveWorkbook.Sheets(sname)
    If Err = 0 Then SheetExists = True _
        Else SheetExists = False
End Function
```

## Descripción

| Función | Método | Retorno |
|---------|--------|---------|
| `SheetExists(shtName, wb)` | Usa `On Error Resume Next` para intentar asignar la hoja | Boolean - `True` si existe |
| `CheckIfSheetExists(SheetName)` | Itera por todas las hojas comparando nombres | Boolean - `True` si encuentra coincidencia |
| `isSheetExist(shName)` | Usa `Evaluate` con `IsError` para verificar la hoja | Boolean - `True` si no hay error |
| `SheetExists(sname)` (Private) | Intenta acceder a la hoja y verifica el estado de `Err` | Boolean - `True` si `Err = 0` |

---

*Módulo: 06 - VBA Worksheets | Archivo: 02 - SheetExiste.txt*
