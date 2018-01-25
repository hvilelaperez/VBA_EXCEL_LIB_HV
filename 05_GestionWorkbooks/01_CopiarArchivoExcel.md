# Copiar Archivo Excel

Copia archivos Excel de una ubicación a otra, creando múltiples copias con diferentes nombres.

## Código VBA

```vba
Option Explicit

' ============================================================
' SELECCIONAR ARCHIVO PLANTILLA
' Abre diálogo para elegir el archivo origen
' ============================================================
Public Sub SeleccionarPlantilla()
    Dim rutaArchivo As String
    
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Seleccionar archivo plantilla"
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xls; *.xlsx; *.xlsm"
        
        If .Show = -1 Then
            ActiveSheet.Range("A1").Value = .SelectedItems(1)
        End If
    End With
End Sub


' ============================================================
' SELECCIONAR CARPETA DESTINO
' Abre diálogo para elegir dónde copiar
' ============================================================
Public Sub SeleccionarCarpetaDestino()
    Dim dialogo As FileDialog
    
    Set dialogo = Application.FileDialog(msoFileDialogFolderPicker)
    dialogo.AllowMultiSelect = False
    
    If dialogo.Show = -1 Then
        ActiveSheet.Range("A2").Value = dialogo.SelectedItems(1)
    End If
End Sub


' ============================================================
' COPIAR ARCHIVO CON NUEVOS NOMBRES
' Crea múltiples copias con nombres de la columna B
' ============================================================
Public Sub CopiarArchivoConNombres()
    Dim rutaOrigen As String
    Dim carpetaDestino As String
    Dim extension As String
    Dim celda As Range
    Dim contador As Long
    
    ' Leer rutas desde la hoja
    rutaOrigen = ActiveSheet.Range("A1").Value
    carpetaDestino = ActiveSheet.Range("A2").Value
    
    If rutaOrigen = "" Or carpetaDestino = "" Then
        MsgBox "Debe seleccionar archivo origen y carpeta destino.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener extensión del archivo original
    extension = Mid(rutaOrigen, InStrRev(rutaOrigen, "."))
    
    ' Asegurar backslash al final
    If Right(carpetaDestino, 1) <> "\" Then carpetaDestino = carpetaDestino & "\"
    
    contador = 0
    
    ' Recorrer columna B para nuevos nombres
    For Each celda In ActiveSheet.Range("B2:B100")
        If celda.Value <> "" Then
            FileCopy rutaOrigen, carpetaDestino & celda.Value & "." & extension
            contador = contador + 1
        End If
    Next celda
    
    MsgBox "Proceso completado." & vbCrLf & _
           "Archivos copiados: " & contador, vbInformation
End Sub


' ============================================================
' COPIAR ARCHIVO ESPECÍFICO
' Copia un archivo a ubicación específica
' ============================================================
Public Sub CopiarArchivoEspecifico()
    Dim rutaOrigen As String
    Dim rutaDestino As String
    
    rutaOrigen = ThisWorkbook.Path & "\Original.xlsx"
    rutaDestino = ThisWorkbook.Path & "\Copia_" & Format(Now, "YYYYMMDD_HHmmss") & ".xlsx"
    
    If Dir(rutaOrigen) = "" Then
        MsgBox "El archivo origen no existe.", vbExclamation
        Exit Sub
    End If
    
    FileCopy rutaOrigen, rutaDestino
    
    MsgBox "Archivo copiado a:" & vbCrLf & rutaDestino, vbInformation
End Sub
```

## Estructura de Hoja

| A (Ruta Origen) | B (Nuevos Nombres) |
|-----------------|-------------------|
| C:\Archivo.xlsx | Reporte_Enero |
| | Reporte_Febrero |
| | Reporte_Marzo |

## Funciones Utilizadas

| Función | Descripción |
|---------|-------------|
| `FileCopy` | Copia un archivo de origen a destino |
| `Dir` | Verifica si existe un archivo |

---

*Módulo: 05_GestionWorkbooks | Archivo: 01_CopiarArchivoExcel.md*
