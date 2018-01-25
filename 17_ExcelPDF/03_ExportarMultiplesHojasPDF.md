# Exportar Múltiples Hojas a un Solo Archivo PDF

Exporta todas las hojas de un libro de Excel a un único archivo PDF, seleccionando todas las hojas previamente.

## Código VBA

```vba
Sub EXPORTAR_PDF()
    Dim wb As Workbook, sh As Worksheet, i As Integer
    Dim sFolder As String, sFile As Variant, sFilePDF As String
    Dim arrHojas()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Exportar PDF
    'expression .ExportAsFixedFormat(Type, Filename, Quality, _
    'IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish)
    ReDim arrHojas(1 To Worksheets.Count)
    For Each shi In Worksheets
        i = i + 1
        sFile = ThisWorkbook.Path & "\" & shi.Name & ".pdf"
        sFilePDF = CStr(sFile)
        arrHojas(i) = shi.Name
    Next shi
    ReDim Preserve arrHojas(1 To i)
    Sheets(arrHojas).Select
    ActiveSheet.ExportAsFixedFormat _
                    xlTypePDF, _
                    "test", _
                    xlQualityStandard, _
                    True, _
                    False, , , False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Function CrearNombreArchivoGuardado(ByVal FileName As String) As String
    Dim n As Long
    Dim sFolder As String, sFile As String, sExt As String, sTmpFile As String
    sTmpFile = FileName
    With CreateObject("Scripting.FileSystemObject")
        sExt = .GetExtensionName(FileName)
        sFolder = .GetParentFolderName(FileName)
        sFile = .GetBaseName(FileName)
        Do While .FileExists(sTmpFile) = True
            n = n + 1
            sTmpFile = .BuildPath(sFolder, sFile & "(" & n & ")." & sExt)
        Loop
        CrearNombreArchivoGuardado = sTmpFile
    End With
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Sub `EXPORTAR_PDF`** | Selecciona todas las hojas y las exporta a un solo PDF |
| **Función `CrearNombreArchivoGuardado`** | Genera un nombre único si el archivo ya existe (agrega número incremental) |
| **Array de hojas** | Almacena los nombres de todas las hojas para seleccionarlas juntas |
| **Nombre del archivo** | Se usa "test" como nombre base del PDF |

| Flujo del procedimiento |
|------------------------|
| 1. Recopilar nombres de todas las hojas en un array |
| 2. Seleccionar todas las hojas simultáneamente |
| 3. Exportar como PDF con `ExportAsFixedFormat` |
| 4. La función auxiliar evita sobrescribir archivos existentes |

| Parámetro | Valor |
|-----------|-------|
| Type | xlTypePDF |
| Quality | xlQualityStandard |
| IncludeDocProperties | True |
| IgnorePrintAreas | False |
| OpenAfterPublish | False |

---

*Módulo: 11 - VBA Excel PDF | Archivo: 02 - Export MultiSheets to 1 file PDF*
