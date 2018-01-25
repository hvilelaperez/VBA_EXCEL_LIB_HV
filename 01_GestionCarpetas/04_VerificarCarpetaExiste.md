# Verificar si Carpeta o Archivo Existe

Funciones para comprobar la existencia de carpetas y archivos antes de operar con ellos.

## Código VBA

```vba
Option Explicit

' ============================================================
' VERIFICAR SI UNA CARPETA EXISTE
' ============================================================
Public Function CarpetaExiste(rutaCarpeta As String) As Boolean
    On Error GoTo NoExiste
    
    Dim atributos As Long
    atributos = GetAttr(rutaCarpeta)
    CarpetaExiste = (atributos And vbDirectory) = vbDirectory
    Exit Function
    
NoExiste:
    CarpetaExiste = False
End Function


' ============================================================
' VERIFICAR SI UN ARCHIVO EXISTE
' ============================================================
Public Function ArchivoExiste(rutaArchivo As String) As Boolean
    On Error GoTo NoExiste
    
    Dim atributos As Long
    atributos = GetAttr(rutaArchivo)
    ArchivoExiste = ((atributos And vbDirectory) <> vbDirectory)
    Exit Function
    
NoExiste:
    ArchivoExiste = False
End Function


' ============================================================
' VERIFICAR CON Dir()
' ============================================================
Public Function CarpetaExisteDir(rutaCarpeta As String) As Boolean
    CarpetaExisteDir = (Dir(rutaCarpeta, vbDirectory) <> "")
End Function


' ============================================================
' VERIFICAR CON FileSystemObject
' ============================================================
Public Function CarpetaExisteFSO(rutaCarpeta As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    CarpetaExisteFSO = fso.FolderExists(rutaCarpeta)
    Set fso = Nothing
End Function
```

## Comparación de Métodos

| Método | Ventaja | Desventaja |
|--------|---------|------------|
| `GetAttr` | Nativo VBA | Requiere manejo de errores |
| `Dir()` | Simple | Puede fallar con rutas largas |
| `FileSystemObject` | Robusto | Requiere creación de objeto |

---

*Módulo: 01_GestionCarpetas | Archivo: 04_VerificarCarpetaExiste.md*
