# Guardar como PDF

Exporta el libro activo o una hoja específica a formato PDF.

## Código VBA

```vba
Option Explicit

' ============================================================
' GUARDAR COMO PDF CON DIÁLOGO
' ============================================================
Public Sub GuardarComoPDF()
    Dim nombrePDF As String
    
    nombrePDF = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & ".pdf"
    
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "Guardar como PDF"
        .InitialFileName = ThisWorkbook.Path & "\" & nombrePDF
        
        If .Show Then
            ActiveWorkbook.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                fileName:=.SelectedItems(1), _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
            
            MsgBox "PDF guardado.", vbInformation
        End If
    End With
End Sub


' ============================================================
' GUARDAR PDF AUTOMÁTICO
' ============================================================
Public Sub GuardarPDFAutomatico()
    Dim rutaPDF As String
    
    rutaPDF = ThisWorkbook.Path & "\" & Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & ".pdf"
    
    ActiveWorkbook.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=rutaPDF, _
        Quality:=xlQualityStandard, _
        OpenAfterPublish:=True
    
    MsgBox "PDF generado.", vbInformation
End Sub


' ============================================================
' GUARDAR SOLO HOJA ACTIVA
' ============================================================
Public Sub GuardarHojaActivaComoPDF()
    Dim rutaPDF As String
    
    rutaPDF = ThisWorkbook.Path & "\" & ActiveSheet.Name & ".pdf"
    
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=rutaPDF, _
        Quality:=xlQualityStandard
    
    MsgBox "Hoja exportada.", vbInformation
End Sub
```

---

*Módulo: 03_OperacionesArchivos | Archivo: 03_GuardarComoPDF.md*
