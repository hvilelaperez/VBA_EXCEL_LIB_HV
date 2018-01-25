# Transferir Datos de Excel a Word - Múltiples Métodos

Macros VBA con diferentes enfoques para transferir datos de Excel a Word: reemplazo simple, manejo de textos largos superiores a 255 caracteres, y funciones auxiliares para búsqueda y reemplazo.

> **Autor:** Henry Vilela - hvilelaperez@gmail.com

## Código VBA

```vba
Option Explicit
Public t As Object

Sub test()
    Dim num_of_cust As Long
    Dim num_of_column As Long
    Dim i As Long, j As Long, arr()
    Dim template As Object
    Application.ScreenUpdating = False
    num_of_column = 6
    
    num_of_cust = Sheet2.Cells(Rows.Count, "A").End(xlUp).Row - 1
    arr = Sheet2.Range("A2:F5").Value
    ReDim Preserve arr(1 To num_of_cust, 1 To num_of_column)
    
    With CreateObject("word.application") ' Late binding
        .Visible = True
        
        For i = 1 To num_of_cust
            Set template = .Documents.Open("C:\Users\tien-duy.nguyen\Desktop\VBA Word\template.docx")
            Set t = template.Content
            For j = 1 To num_of_column
                t.Find.Execute _
                    FindText:=Sheet2.Cells(1, j).Value, _
                    ReplaceWith:=arr(i, j), _
                    Replace:=wdReplaceAll
            Next
            template.SaveAs Filename:=ThisWorkbook.Path & Application.PathSeparator & i & "-documento.docx"
        Next
        .Quit
    End With
    Set t = Nothing
    Set template = Nothing
    Application.ScreenUpdating = True
End Sub


Sub EXCEL_TO_WORD_MORE255()
    Const wdReplaceAll As Long = 2
    Const wdFindContinue As Long = 1
    Dim longString As String
    Dim wd As Object, doc As Object, sel As Object
    Dim dataObj As New dataObject  '## Requiere referencia a MSForms
    
    Set wd = GetObject(, "Word.Application")
    Set doc = wd.ActiveDocument
    
    longString = Worksheets("3A- Multi-Family Housing").Range("A4").Value
    
    dataObj.SetText longString
    dataObj.PutinClipboard
    
    Set sel = doc.Range
    sel.Select
    
    With doc.Range.Find
        .ClearFormatting
        .Text = "[[3A Description]]"
        .Replacement.ClearFormatting
        .Replacement.Text = "^c"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
    End With
End Sub


Sub EXCEL_TO_WORD_ADVANCED()
    Dim oSourceDoc As Object, dataObj As Object, selectRng As Object
    Dim srchTxt As String, sValue As String
    Dim replaceRng As Range
    Dim i&, j&, arr(), num_of_cust&, num_of_column&, t As Object, lTextAfter&
    
    Application.ScreenUpdating = False
    'Set dataObj = CreateObject("MSForms.DataObject")
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    'Número de columnas que contienen el código a buscar
    num_of_column = 6
    
    'Número de columnas que contienen datos, sin contar la fila de encabezados
    num_of_cust = Sheet1.Cells(Rows.Count, "A").End(xlUp).Row - 1
    
    arr = Sheet1.Range("A2:F5").Value
    ReDim Preserve arr(1 To num_of_cust, 1 To num_of_column)
    
    With CreateObject("word.application") ' Late binding
        .Visible = True
        
        For i = 1 To num_of_cust
            Set oSourceDoc = .Documents.Open("C:\Users\tien-duy.nguyen\Desktop\VBA Word\template.docx")
            Set t = oSourceDoc.Content
            For j = 1 To num_of_column
                sValue = Sheet1.Cells(i + 1, j).Value
                If Len(sValue) <= 250 Then
                    t.Find.Execute _
                    FindText:=Sheet1.Cells(1, j).Value, _
                    ReplaceWith:=sValue, _
                    Replace:=wdReplaceAll
                Else
                    'Para textos largos: copiar al portapapeles y pegar
                    dataObj.SetText sValue
                    dataObj.PutinClipboard
                    Set selectRng = oSourceDoc.Range
                    selectRng.Select
                    lTextAfter = Len(sValue) - 250
                    
                    With oSourceDoc.Range.Find
                        .Text = Sheet1.Cells(1, j).Value
                        .Forward = True
                        .Replacement.Text = "^c"
                        .Wrap = wdFindContinue
                        .Execute Replace:=wdReplaceAll
                    End With
                End If
            Next j
            oSourceDoc.SaveAs Filename:=ThisWorkbook.Path & Application.PathSeparator & i & "-documento.docx"
        Next i
        .Quit
    End With
    
    Set t = Nothing
    Set oSourceDoc = Nothing
    Application.ScreenUpdating = True
End Sub


Sub LongStringFindReplace()
    Dim oSourceDoc As Object
    Set oSourceDoc = CreateObject("Word.Application")
    Dim srchTxt As String
    Dim replaceRng As Range
    Dim i As Long
    
    Set oSourceDoc = CreateObject("Word.Application").Documents.Open("C:\Users\tien-duy.nguyen\Desktop\VBA Word\template.docx")
    
    With oSourceDoc
        'Establecer cadena de búsqueda
        srchTxt = .Paragraphs(1).Range.Text
        srchTxt = Left(srchTxt, Len(srchTxt) - 1) 'Eliminar marca de párrafo
        'Establecer texto de reemplazo y copiar al portapapeles
        Set replaceRng = .Paragraphs(2).Range
        replaceRng.MoveEnd Unit:=wdCharacter, Count:=-1 'Eliminar marca de párrafo
        replaceRng.Copy
        .Close
    End With
    
    ActiveDocument.Range(0, 0).Select
    If Len(srchTxt) > 250 Then
        i = Len(srchTxt) - 250
        With Selection.Find
            .Text = Left(srchTxt, 250)
            .Forward = True
            .Wrap = wdFindContinue
            Do While .Execute
                'Mover el final de la selección para igualar la longitud de srchTxt
                Selection.MoveEnd Unit:=wdCharacter, Count:=i
                'Comparar la selección con la cadena de búsqueda
                If Selection.Text = srchTxt Then
                    'Reemplazar la selección con el contenido del portapapeles
                    Selection.Paste
                Else
                    Selection.MoveRight
                End If
            Loop
        End With
    Else
        ResetFRParameters
        With Selection.Find
            .Text = srchTxt
            .Replacement.Text = "^c"
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End If
lbl_Exit:
    Exit Sub
End Sub


Sub ResetFRParameters()
    With t.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
lbl_Exit:
    Exit Sub
End Sub


Public Function FIND_REPLACEMULTI_EXCEL_TO_WORD(srchTxt As String, replaceRng As Range)
    Dim oSourceDoc As Object
    Set oSourceDoc = CreateObject("Word.Application")
    Dim i As Long
    
    Set oSourceDoc = CreateObject("Word.Application").Documents.Open("C:\Users\tien-duy.nguyen\Desktop\VBA Word\template.docx")
    
    With oSourceDoc
        'Establecer cadena de búsqueda
        srchTxt = .Paragraphs(1).Range.Text
        srchTxt = Left(srchTxt, Len(srchTxt) - 1) 'Eliminar marca de párrafo
        'Establecer texto de reemplazo y copiar al portapapeles
        Set replaceRng = .Paragraphs(2).Range
        replaceRng.MoveEnd Unit:=wdCharacter, Count:=-1 'Eliminar marca de párrafo
        replaceRng.Copy
        .Close
    End With
    
    ActiveDocument.Range(0, 0).Select
    If Len(srchTxt) > 250 Then
        i = Len(srchTxt) - 250
        With Selection.Find
            .Text = Left(srchTxt, 250)
            .Forward = True
            .Wrap = wdFindContinue
            Do While .Execute
                'Mover el final de la selección para igualar la longitud de srchTxt
                Selection.MoveEnd Unit:=wdCharacter, Count:=i
                'Comparar la selección con la cadena de búsqueda
                If Selection.Text = srchTxt Then
                    'Reemplazar la selección con el contenido del portapapeles
                    Selection.Paste
                Else
                    Selection.MoveRight
                End If
            Loop
        End With
    Else
        ResetFRParameters
        With Selection.Find
            .Text = srchTxt
            .Replacement.Text = "^c"
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End If
lbl_Exit:
    Exit Function
End Function
```

## Descripción

| Sub/Función | Descripción |
|---|---|
| `test` | Versión simplificada: reemplaza marcadores con valores del arreglo directamente |
| `EXCEL_TO_WORD_MORE255` | Reemplaza un texto largo específico usando portapapeles como intermediario |
| `EXCEL_TO_WORD_ADVANCED` | Versión completa con manejo condicional de textos <= 250 y > 250 caracteres |
| `LongStringFindReplace` | Reemplaza textos largos comparando fragmentos de 250 caracteres |
| `ResetFRParameters` | Restaura los parámetros de búsqueda a valores predeterminados |
| `FIND_REPLACEMULTI_EXCEL_TO_WORD` | Función que encapsula la lógica de búsqueda y reemplazo múltiple |

### Comparativa de Métodos

| Método | Ventaja | Limitación |
|---|---|---|
| `Find.Execute` directo | Rápido y simple | Limitado a 250 caracteres |
| Portapapeles (`^c`) | Maneja cualquier longitud | Requiere `MSForms.DataObject` |
| Comparación por fragmentos | Preciso para textos largos | Más complejo de implementar |

### Puntos Clave

- `wdReplaceAll = 2`: reemplaza todas las ocurrencias
- `wdFindContinue = 1`: continúa la búsqueda desde el final del documento
- `^c` en `Replacement.Text` significa "contenido del portapapeles"
- El parámetro `MatchWholeWord = True` en `ResetFRParameters` asegura coincidencia de palabras completas

---

*Módulo: 18 - VBA Excel Word | Archivo: 01 - ExcelToWord - MultiTest*
