# Obtener Nombre Disponible

Genera nombres de archivo únicos cuando ya existe un archivo con ese nombre.

## Código VBA

```vba
Option Explicit

' ============================================================
' OBTENER NOMBRE DE ARCHIVO DISPONIBLE
' Agrega número incremental si el archivo ya existe
' ============================================================
Public Function ObtenerNombreDisponible(rutaCompleta As String) As String
    Dim fso As Object
    Dim carpeta As String
    Dim nombreBase As String
    Dim extension As String
    Dim contador As Long
    Dim rutaTemporal As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    carpeta = fso.GetParentFolderName(rutaCompleta)
    nombreBase = fso.GetBaseName(rutaCompleta)
    extension = fso.GetExtensionName(rutaCompleta)
    
    If Not fso.FileExists(rutaCompleta) Then
        ObtenerNombreDisponible = rutaCompleta
        Exit Function
    End If
    
    contador = 1
    rutaTemporal = fso.BuildPath(carpeta, nombreBase & " (" & contador & ")." & extension)
    
    Do While fso.FileExists(rutaTemporal)
        contador = contador + 1
        rutaTemporal = fso.BuildPath(carpeta, nombreBase & " (" & contador & ")." & extension)
    Loop
    
    ObtenerNombreDisponible = rutaTemporal
    Set fso = Nothing
End Function


' ============================================================
' GUARDAR CON NOMBRE DISPONABLE
' ============================================================
Public Sub GuardarConNombreDisponible()
    Dim rutaFinal As String
    
    rutaFinal = ObtenerNombreDisponible(ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    ThisWorkbook.SaveAs rutaFinal
    
    MsgBox "Guardado como: " & rutaFinal, vbInformation
End Sub
```

## Ejemplo

```
Ruta original:    C:\Reportes\Ventas.xlsx  (ya existe)
Ruta resultante:  C:\Reportes\Ventas (1).xlsx
```

---

*Módulo: 03_OperacionesArchivos | Archivo: 01_ObtenerNombreDisponible.md*
