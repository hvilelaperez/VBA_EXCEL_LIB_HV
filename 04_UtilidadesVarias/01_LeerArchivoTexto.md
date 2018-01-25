# Leer Archivo de Texto

Lee el contenido de archivos de texto y lo carga en hojas de cálculo.

## Código VBA

```vba
Option Explicit

' ============================================================
' LEER ARCHIVO LÍNEA POR LÍNEA
' ============================================================
Public Sub LeerArchivoLineaPorLinea()
    Dim fso As Object
    Dim archivo As Object
    Dim textoStream As Object
    Dim rutaArchivo As String
    Dim linea As String
    Dim fila As Long
    
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Seleccionar archivo"
        .Filters.Clear
        .Filters.Add "Archivos de texto", "*.txt; *.log"
        
        If .Show <> -1 Then Exit Sub
        rutaArchivo = .SelectedItems(1)
    End With
    
    Application.ScreenUpdating = False
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set archivo = fso.GetFile(rutaArchivo)
    Set textoStream = archivo.OpenAsTextStream(1)
    
    fila = 1
    
    Do While Not textoStream.AtEndOfStream
        linea = textoStream.ReadLine
        ActiveSheet.Cells(fila, 1).Value = linea
        fila = fila + 1
    Loop
    
    textoStream.Close
    Application.ScreenUpdating = True
    
    MsgBox "Líneas leídas: " & (fila - 1), vbInformation
    
    Set textoStream = Nothing
    Set archivo = Nothing
    Set fso = Nothing
End Sub


' ============================================================
' LEER CON DELIMITADOR
' ============================================================
Public Sub LeerConDelimitador()
    Dim fso As Object
    Dim archivo As Object
    Dim textoStream As Object
    Dim rutaArchivo As String
    Dim linea As String
    Dim partes As Variant
    Dim fila As Long
    Dim columna As Long
    
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Seleccionar archivo"
        .Filters.Clear
        .Filters.Add "Archivos", "*.txt; *.csv"
        
        If .Show <> -1 Then Exit Sub
        rutaArchivo = .SelectedItems(1)
    End With
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set archivo = fso.GetFile(rutaArchivo)
    Set textoStream = archivo.OpenAsTextStream(1)
    
    fila = 1
    
    Do While Not textoStream.AtEndOfStream
        linea = textoStream.ReadLine
        partes = Split(linea, ",")
        
        For columna = 0 To UBound(partes)
            ActiveSheet.Cells(fila, columna + 1).Value = partes(columna)
        Next columna
        
        fila = fila + 1
    Loop
    
    textoStream.Close
    
    MsgBox "Filas importadas: " & (fila - 1), vbInformation
    
    Set textoStream = Nothing
    Set archivo = Nothing
    Set fso = Nothing
End Sub
```

---

*Módulo: 04_UtilidadesVarias | Archivo: 01_LeerArchivoTexto.md*
