# Crear Carpetas desde Rango de Excel

Crea carpetas automáticamente basándose en los valores de un rango de celdas.

## Código VBA

```vba
Option Explicit

' ============================================================
' CREAR CARPETAS DESDE RANGO
' Lee valores de la columna A y crea carpetas
' ============================================================
Public Sub CrearCarpetasDesdeRango()
    Dim hojaTrabajo As Worksheet
    Dim celda As Range
    Dim rutaBase As String
    
    Set hojaTrabajo = ActiveSheet
    rutaBase = ThisWorkbook.Path & "\"
    
    For Each celda In hojaTrabajo.Range("A2:A" & hojaTrabajo.Cells(hojaTrabajo.Rows.Count, 1).End(xlUp).Row)
        If celda.Value <> "" Then
            If Dir(rutaBase & Trim(CStr(celda.Value)), vbDirectory) = "" Then
                MkDir rutaBase & Trim(CStr(celda.Value))
                celda.Interior.Color = RGB(144, 238, 144) ' Verde
            Else
                celda.Interior.Color = RGB(255, 255, 200) ' Amarillo
            End If
        End If
    Next celda
    
    MsgBox "Proceso completado.", vbInformation
End Sub


' ============================================================
' CREAR ESTRUCTURA DE CARPETAS ANIDADAS
' ============================================================
Public Sub CrearEstructuraAnidada()
    Dim fso As Object
    Dim rutaBase As String
    Dim carpetas As Variant
    Dim subCarpetas As Variant
    Dim i As Long, j As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    rutaBase = ThisWorkbook.Path & "\Proyecto"
    
    carpetas = Array("Documentos", "Recursos", "Resultados")
    subCarpetas = Array("Internos", "Externos", "Temporal")
    
    If Not fso.FolderExists(rutaBase) Then fso.CreateFolder rutaBase
    
    For i = LBound(carpetas) To UBound(carpetas)
        Dim rutaCarpeta As String
        rutaCarpeta = rutaBase & "\" & carpetas(i)
        If Not fso.FolderExists(rutaCarpeta) Then fso.CreateFolder rutaCarpeta
        
        For j = LBound(subCarpetas) To UBound(subCarpetas)
            If Not fso.FolderExists(rutaCarpeta & "\" & subCarpetas(j)) Then
                fso.CreateFolder rutaCarpeta & "\" & subCarpetas(j)
            End If
        Next j
    Next i
    
    MsgBox "Estructura creada.", vbInformation
    Set fso = Nothing
End Sub
```

## Estructura de Hoja

| A (Nombre Carpeta) | B (Estado) |
|-------------------|------------|
| Reporte01 | *(se llena)* |
| Reporte02 | *(se llena)* |

---

*Módulo: 01_GestionCarpetas | Archivo: 03_CrearCarpetasDesdeRango.md*
