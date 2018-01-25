# Listar Comentarios del Libro

Macro que lista todos los comentarios de la hoja activa en un nuevo libro de Excel, mostrando autor, libro, hoja, rango y texto del comentario.

## Código VBA

```vba
' Autor: Henry Vilela - hvilelaperez@gmail.com
Sub ListarComentarios()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cmt As Comment
    Dim cmtCount As Long
    cmtCount = 2
    On Error Resume Next
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0
    Application.ScreenUpdating = False
    Set wb = Workbooks.Add(xlWorksheet)
    With wb.Sheets(1)
        .Range("$A$1") = "Autor"
        .Range("$B$1") = "Libro"
        .Range("$C$1") = "Hoja"
        .Range("$D$1") = "Rango"
        .Range("$E$1") = "Comentario"
    End With
    For Each cmt In ws.Comments
        With wb.Sheets(1)
            .Cells(cmtCount, 1) = cmt.author
            .Cells(cmtCount, 2) = cmt.Parent.Parent.Parent.Name
            .Cells(cmtCount, 3) = cmt.Parent.Parent.Name
            .Cells(cmtCount, 4) = cmt.Parent.Address
            .Cells(cmtCount, 5) = LimpiarComentario(cmt.author, cmt.Text)
        End With
        cmtCount = cmtCount + 1
    Next
    wb.Sheets(1).UsedRange.WrapText = False
    Application.ScreenUpdating = True
    Set ws = Nothing
    Set wb = Nothing
End Sub

Private Function LimpiarComentario(author As String, cmt As String) As String
    Dim tmp As String
    tmp = Application.WorksheetFunction.Substitute(cmt, author & ":", "")
    tmp = Application.WorksheetFunction.Substitute(tmp, Chr(10), "")
    LimpiarComentario = tmp
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Procedimiento** | `ListarComentarios` - Extrae todos los comentarios de la hoja activa |
| **Función auxiliar** | `LimpiarComentario` - Elimina el nombre del autor y saltos de línea del texto |
| **Encabezados** | Autor, Libro, Hoja, Rango, Comentario |
| **Ubicación** | Crea un nuevo libro de trabajo con los resultados |
| **Optimización** | Desactiva actualización de pantalla para mejorar rendimiento |
| **Uso** | Ejecutar macro desde la hoja que contiene los comentarios |

---

*Módulo: 31_CodigosGenerales | Archivo: 39 - Let ke cac comments trong workbook.txt*
