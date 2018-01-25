# Renombrar Hojas de Trabajo

Procedimiento para renombrar todas las hojas de un libro añadiendo un sufijo `_v1` al nombre actual, evitando duplicar el sufijo si ya existe.

## Código VBA

```vba
Sub RENAME_SHEET()

Dim ws As Worksheet
Dim ws1 As Worksheet
Dim strErr As String

On Error Resume Next
For Each ws In ActiveWorkbook.Sheets
Set ws1 = Sheets(ws.Name)
    If Right(ws1.Name, 3) <> "_v1" Then
        ws.Name = ws.Name & "_v1"
    Else
         strErr = strErr & ws.Name & " ya tiene el sufijo " & "_v1" & vbNewLine
    End If
Set ws1 = Nothing
Next
On Error GoTo 0

If Len(strErr) > 0 Then MsgBox strErr, vbOKOnly, "Estas hojas ya existían"

End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `RENAME_SHEET` | Procedimiento principal que recorre todas las hojas del libro activo |
| `ws` / `ws1` | Variables de tipo Worksheet para iterar y validar |
| `strErr` | Cadena que acumula los nombres de hojas que ya tienen el sufijo |
| `On Error Resume Next` | Manejo de errores para evitar interrupciones al acceder a hojas |
| Sufijo `_v1` | Se añade al final del nombre de cada hoja |
| `MsgBox` | Muestra un mensaje con las hojas que ya tenían el sufijo |

---

*Módulo: 06 - VBA Worksheets | Archivo: 00 - Rename sheets.txt*
