# Importar CSV Grande

Importa archivos CSV grandes de forma eficiente.

## Código VBA

```vba
Option Explicit

' ============================================================
' IMPORTAR CSV - MÉTODO RÁPIDO
' ============================================================
Public Sub ImportarCSVRapido()
    Dim rutaArchivo As String
    Dim buf(1 To 16384) As Variant
    Dim i As Long
    
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Seleccionar CSV"
        .Filters.Clear
        .Filters.Add "CSV", "*.csv"
        
        If .Show <> -1 Then Exit Sub
        rutaArchivo = .SelectedItems(1)
    End With
    
    ' Configurar buffer
    For i = 1 To 16384
        buf(i) = Array(i, 2) ' 2 = xlTextFormat
    Next i
    
    Application.ScreenUpdating = False
    
    Workbooks.OpenText _
        fileName:=rutaArchivo, _
        DataType:=xlDelimited, _
        Comma:=True, _
        FieldInfo:=buf
    
    Erase buf
    
    ' Copiar a libro actual
    ActiveSheet.UsedRange.Copy ThisWorkbook.Sheets(1).Range("A1")
    ActiveWorkbook.Close False
    
    Application.ScreenUpdating = True
    
    MsgBox "CSV importado.", vbInformation
End Sub


' ============================================================
' IMPORTAR CON SEPARADOR PERSONALIZADO
' ============================================================
Public Sub ImportarCSVConSeparador()
    Dim rutaArchivo As String
    Dim separador As String
    Dim buf(1 To 16384) As Variant
    Dim i As Long
    
    separador = InputBox("Separador (ej: ; o ,):", "Separador", ";")
    If separador = "" Then Exit Sub
    
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Seleccionar archivo"
        .Filters.Clear
        .Filters.Add "Archivos", "*.csv; *.txt"
        
        If .Show <> -1 Then Exit Sub
        rutaArchivo = .SelectedItems(1)
    End With
    
    For i = 1 To 16384
        buf(i) = Array(i, 2)
    Next i
    
    Workbooks.OpenText _
        fileName:=rutaArchivo, _
        DataType:=xlDelimited, _
        Semicolon:=IIf(separador = ";", True, False), _
        Comma:=IIf(separador = ",", True, False), _
        FieldInfo:=buf
    
    Erase buf
    
    ActiveSheet.UsedRange.Copy ThisWorkbook.Sheets(1).Range("A1")
    ActiveWorkbook.Close False
    
    MsgBox "CSV importado con separador: " & separador, vbInformation
End Sub
```

---

*Módulo: 04_UtilidadesVarias | Archivo: 02_ImportarCSVGrande.md*
