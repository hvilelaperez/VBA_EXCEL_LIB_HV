# Transferir Datos de Excel a Word

Macro VBA para transferir datos de una hoja de cálculo de Excel a documentos Word mediante búsqueda y reemplazo de marcadores de posición, incluyendo manejo de textos largos superiores a 250 caracteres.

> **Autor:** Henry Vilela - hvilelaperez@gmail.com

## Código VBA

```vba
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
```

## Descripción

| Componente | Descripción |
|---|---|
| **Fuente de datos** | Hoja `Sheet1` con encabezados en fila 1 y datos desde fila 2 |
| **Documento destino** | Plantilla Word (`template.docx`) con marcadores de posición |
| **Número de columnas** | 6 columnas de datos (A a F) |

### Flujo de Ejecución

1. Lee los datos de Excel en un arreglo bidimensional
2. Abre Word mediante late binding (sin referencia directa)
3. Para cada cliente/fila, abre la plantilla Word
4. Recorre cada columna y reemplaza el marcador de posición con el valor correspondiente
5. Si el texto tiene **250 caracteres o menos**: reemplazo directo con `Find.Execute`
6. Si el texto tiene **más de 250 caracteres**: copia al portapapeles y reemplaza con `^c`
7. Guarda el documento con nombre secuencial

### Puntos Clave

- La función `Find.Execute` tiene un límite de 250 caracteres en el texto de búsqueda
- Para textos largos se utiliza `MSForms.DataObject` como puente al portapapeles
- Se utiliza `^c` (reemplazo por contenido del portapapeles) para superar la limitación
- El nombre del archivo genera secuencialmente: `1-documento.docx`, `2-documento.docx`, etc.

---

*Módulo: 18 - VBA Excel Word | Archivo: 00 - ExcelToWord*
