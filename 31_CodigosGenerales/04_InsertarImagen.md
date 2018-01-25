# Insertar Imagen en Hoja de Cálculo

Macros para insertar imágenes (logos, capturas de pantalla) en posiciones específicas de una hoja de cálculo, manteniendo la proporción original.

## Código VBA

```vba
Option Explicit

Sub Inserer_Logo()
    Dim fNameAndPath As String, sh As Worksheet, rng As Range, rng1 As Range
    Dim img As Object, zoomF As Integer
    Dim shp As Shape, countLine As Integer, nameShape As String
    Dim calcul As Double
    Set sh = ActiveSheet
    On Error Resume Next
    Application.ScreenUpdating = False
    nameShape = Application.Caller
    Set shp = sh.Shapes(nameShape)
    countLine = shp.TopLeftCell.Row

    '========================================================================'
    Set rng = sh.Range("B" & (countLine + 3))
    Set rng1 = sh.Range("B1", "F1")
    '========================================================================'

    rng.Select

    fNameAndPath = Application.GetOpenFilename(Title:="Seleccionar una imagen")
    Set img = LoadPicture(fNameAndPath)
    calcul = img.Width / img.Height
    Set shp = sh.Shapes.AddPicture(fNameAndPath, msoFalse, msoCTrue, rng.Left, rng.Top, rng1.Width, rng1.Width / calcul)

    Application.ScreenUpdating = True
End Sub


Sub Inserer_Capture_Criteres_Selection()
    Dim fNameAndPath As String, sh As Worksheet, rng As Range, rng1 As Range
    Dim img As Object, zoomF As Integer
    Dim shp As Shape, countLine As Integer, nameShape As String
    Dim calcul As Double
    Set sh = ActiveSheet
    On Error Resume Next
    Application.ScreenUpdating = False
    nameShape = Application.Caller
    Set shp = sh.Shapes(nameShape)
    countLine = shp.TopLeftCell.Row
    Debug.Print countLine

    '========================================================================'
    Set rng = sh.Range("C" & (countLine + 2))
    Set rng1 = sh.Range("C1", "S1")
    '========================================================================'

    rng.Select

    fNameAndPath = Application.GetOpenFilename(Title:="Seleccionar una imagen")
    Set img = LoadPicture(fNameAndPath)
    calcul = img.Width / img.Height
    Set shp = sh.Shapes.AddPicture(fNameAndPath, msoFalse, msoCTrue, rng.Left, rng.Top, rng1.Width, rng1.Width / calcul)

    Application.ScreenUpdating = True
End Sub


Sub Inserer_Capture_Documents_Demandes()
    Dim fNameAndPath As String, sh As Worksheet, rng As Range, rng1 As Range
    Dim img As Object, zoomF As Integer
    Dim shp As Shape, countLine As Integer, nameShape As String
    Dim calcul As Double
    Set sh = ActiveSheet
    On Error Resume Next
    Application.ScreenUpdating = False
    nameShape = Application.Caller
    Set shp = sh.Shapes(nameShape)
    countLine = shp.TopLeftCell.Row

    '========================================================================'
    Set rng = sh.Range("W" & (countLine + 2))
    Set rng1 = sh.Range("W1", "AM1")
    '========================================================================'

    rng.Select

    fNameAndPath = Application.GetOpenFilename(Title:="Seleccionar una imagen")
    Set img = LoadPicture(fNameAndPath)
    calcul = img.Width / img.Height
    Set shp = sh.Shapes.AddPicture(fNameAndPath, msoFalse, msoCTrue, rng.Left, rng.Top, rng1.Width, rng1.Width / calcul)

    Application.ScreenUpdating = True
End Sub
```

## Descripción

| Macro | Descripción |
|-------|-------------|
| `Inserer_Logo` | Inserta un logo en la posición calculada a partir del botón que la llama (columnas B-F) |
| `Inserer_Capture_Criteres_Selection` | Inserta una captura de pantalla en la posición de selección (columnas C-S) |
| `Inserer_Capture_Documents_Demandes` | Inserta una captura de documentos solicitados (columnas W-AM) |

| Elemento técnico | Descripción |
|------------------|-------------|
| `Application.Caller` | Identifica qué botón/shape ejecutó la macro |
| `AddPicture` | Método que inserta la imagen manteniendo proporción |
| `calcul = img.Width / img.Height` | Calcula la proporción de la imagen original |

---

*Módulo: 31_CodigosGenerales | Archivo: 04_InsertarImagen.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
