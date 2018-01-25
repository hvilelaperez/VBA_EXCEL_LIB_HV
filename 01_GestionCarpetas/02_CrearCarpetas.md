# Crear Carpetas

Crea una o múltiples carpetas en disco con verificación de existencia previa.

## Código VBA

```vba
Option Explicit

' ============================================================
' CREAR CARPETA INDIVIDUAL
' ============================================================
Public Sub CrearCarpeta()
    Dim rutaCompleta As String
    
    rutaCompleta = ThisWorkbook.Path & "\MisReportes"
    
    If Dir(rutaCompleta, vbDirectory) = "" Then
        MkDir rutaCompleta
        MsgBox "Carpeta creada: " & rutaCompleta, vbInformation
    Else
        MsgBox "La carpeta ya existe.", vbExclamation
    End If
End Sub


' ============================================================
' CREAR MÚLTIPLES CARPETAS
' ============================================================
Public Sub CrearMultiplesCarpetas()
    Dim rutaBase As String
    Dim carpetas As Variant
    Dim i As Long
    
    rutaBase = ThisWorkbook.Path & "\"
    carpetas = Array("Datos", "Reportes", "Backups", "Temporal", "Configuracion")
    
    For i = LBound(carpetas) To UBound(carpetas)
        If Dir(rutaBase & CStr(carpetas(i)), vbDirectory) = "" Then
            MkDir rutaBase & CStr(carpetas(i))
        End If
    Next i
    
    MsgBox "Carpetas creadas.", vbInformation
End Sub


' ============================================================
' CREAR CARPETA CON FileSystemObject
' ============================================================
Public Sub CrearCarpetaConFSO()
    Dim fso As Object
    Dim rutaCompleta As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    rutaCompleta = ThisWorkbook.Path & "\NuevaCarpeta"
    
    If Not fso.FolderExists(rutaCompleta) Then
        fso.CreateFolder rutaCompleta
        MsgBox "Carpeta creada.", vbInformation
    Else
        MsgBox "La carpeta ya existe.", vbExclamation
    End If
    
    Set fso = Nothing
End Sub
```

## Descripción de Macros

| Macro | Descripción |
|-------|-------------|
| `CrearCarpeta` | Crea una carpeta individual |
| `CrearMultiplesCarpetas` | Crea 5 carpetas estándar |
| `CrearCarpetaConFSO` | Usa FileSystemObject |

---

*Módulo: 01_GestionCarpetas | Archivo: 02_CrearCarpetas.md*
