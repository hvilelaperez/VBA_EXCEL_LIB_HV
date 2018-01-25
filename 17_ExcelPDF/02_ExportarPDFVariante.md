# Exportar PDF con Nombre Dinámico

Exporta la hoja a PDF utilizando un nombre de archivo dinámico basado en una fecha y un valor de celda, recorriendo múltiples registros.

## Código VBA

```vba
Option Explicit
Sub ExportarA_PDF()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim cRng As Range, DATA As Variant, I As Long
    Dim sPath As String, sName As String, sDate As Variant
    DATA = Sheet1.Range("A2:I" & Sheet1.Cells(Sheet1.Rows.Count, "A").End(xlUp).Row).Value
    On Error GoTo ManejoError
    For Each cRng In Sheet2.Range("G1:G" & Sheet2.Cells(Sheet2.Rows.Count, "G").End(xlUp).Row)
        Sheet2.Range("E4").Value = cRng.Value
        sPath = vbNullString: sName = vbNullString
        sPath = ThisWorkbook.Path
        If sPath = vbNullString Then sPath = Application.DefaultFilePath
        sPath = sPath & "\"
        sDate = Application.WorksheetFunction.VLookup(cRng.Value, DATA, 2, 0)
        On Error Resume Next
        If CDate(sDate) Then
            If Err.Number <> 0 Then GoTo ManejoError
            sName = CStr(Format(CDate(sDate), "dd-mm-yyyy")) & cRng.Value & ".Pdf"
        End If
        sPath = sPath & sName
        Sheet2.ExportAsFixedFormat _
                     xlTypePDF, _
                     sPath, _
                     xlQualityStandard, _
                     True, _
                     False, , , False
    Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
ManejoError:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "¡Error encontrado! No se creó ningún archivo PDF"
    Exit Sub
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Origen de datos** | Rango A2:I en Sheet1 contiene la información para nombres y fechas |
| **Recorrido** | Itera sobre cada valor en la columna G de Sheet2 |
| **Nombre del archivo** | Se genera con formato `dd-mm-yyyy` + valor de celda + ".Pdf" |
| **Búsqueda** | Usa `VLookup` para obtener la fecha asociada a cada registro |
| **Ruta** | Usa la ruta del libro actual o la ruta predeterminada de Excel |

| Parámetro de ExportAsFixedFormat | Valor |
|--------------------------------|-------|
| Type | xlTypePDF (formato PDF) |
| Quality | xlQualityStandard (calidad estándar) |
| IncludeDocProperties | True |
| IgnorePrintAreas | False |

---

*Módulo: 11 - VBA Excel PDF | Archivo: 01 - Export PDF*
