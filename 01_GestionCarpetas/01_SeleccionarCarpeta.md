# Seleccionar Carpeta con FileDialog

Abre un diálogo para seleccionar una o múltiples carpetas/archivos y muestra la ruta seleccionada.

## Código VBA

```vba
Option Explicit

' ============================================================
' SELECCIONAR CARPETA - Diálogo de selección de carpeta
' ============================================================
Public Sub SeleccionarCarpeta()
    Dim dialogo As FileDialog
    Dim itemSeleccionado As Variant
    
    Set dialogo = Application.FileDialog(msoFileDialogFolderPicker)
    
    With dialogo
        .Title = "Seleccionar Carpeta de Trabajo"
        .InitialFileName = ThisWorkbook.Path & "\"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            For Each itemSeleccionado In .SelectedItems
                MsgBox "Carpeta seleccionada: " & vbCrLf & itemSeleccionado, _
                       vbInformation, "Carpeta Seleccionada"
            Next itemSeleccionado
        Else
            MsgBox "No se seleccionó ninguna carpeta.", vbExclamation, "Cancelado"
        End If
    End With
    
    Set dialogo = Nothing
End Sub


' ============================================================
' SELECCIONAR ARCHIVOS - Diálogo para abrir archivos
' ============================================================
Public Sub SeleccionarArchivos()
    Dim dialogo As FileDialog
    Dim itemSeleccionado As Variant
    Dim contador As Long
    
    Set dialogo = Application.FileDialog(msoFileDialogOpen)
    
    With dialogo
        .Title = "Seleccionar Archivos Excel"
        .InitialFileName = ThisWorkbook.Path & "\"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xls; *.xlsx; *.xlsm; *.xlsb"
        .Filters.Add "Todos los archivos", "*.*"
        
        If .Show = -1 Then
            contador = 0
            For Each itemSeleccionado In .SelectedItems
                contador = contador + 1
                Debug.Print "Archivo " & contador & ": " & itemSeleccionado
            Next itemSeleccionado
            
            MsgBox "Se seleccionaron " & contador & " archivo(s).", vbInformation
        End If
    End With
    
    Set dialogo = Nothing
End Sub


' ============================================================
' OBTENER RUTA CARPETA COMO FUNCIÓN
' ============================================================
Public Function ObtenerRutaCarpeta(Optional titulo As String = "Seleccionar carpeta", _
                                    Optional rutaInicial As String = "") As String
    Dim dialogo As FileDialog
    
    If rutaInicial = "" Then rutaInicial = ThisWorkbook.Path & "\"
    
    Set dialogo = Application.FileDialog(msoFileDialogFolderPicker)
    
    With dialogo
        .Title = titulo
        .InitialFileName = rutaInicial
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            ObtenerRutaCarpeta = .SelectedItems(1)
        Else
            ObtenerRutaCarpeta = ""
        End If
    End With
    
    Set dialogo = Nothing
End Function
```

## Uso

1. Ejecuta `SeleccionarCarpeta` para elegir una carpeta
2. Ejecuta `SeleccionarArchivos` para elegir múltiples archivos Excel
3. Usa `ObtenerRutaCarpeta()` como función para obtener la ruta como String

---

*Módulo: 01_GestionCarpetas | Archivo: 01_SeleccionarCarpeta.md*
