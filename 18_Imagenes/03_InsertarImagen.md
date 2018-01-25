# Insertar Imagen con Botón

Permite insertar una imagen seleccionada desde el explorador de archivos, ajustándola proporcionalmente al ancho de un rango de referencia.

## Código VBA

```vba
Option Explicit


Sub InsertarLogo()
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
    'zoomF = 100
    '========================================================================'

    rng.Select

    fNameAndPath = Application.GetOpenFilename(Title:="Seleccionar una imagen")
    Set img = LoadPicture(fNameAndPath)
    calcul = img.Width / img.Height
    'Set img = sh.Pictures.Insert(fNameAndPath)
    Set shp = sh.Shapes.AddPicture(fNameAndPath, msoFalse, msoCTrue, rng.Left, rng.Top, rng1.Width, rng1.Width / calcul)
    
    Application.ScreenUpdating = True
End Sub


Sub InsertarCapturaCriteriosSeleccion()
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
    'zoomF = 100
    '========================================================================'

    rng.Select

    fNameAndPath = Application.GetOpenFilename(Title:="Seleccionar una imagen")
    Set img = LoadPicture(fNameAndPath)
    calcul = img.Width / img.Height
    'Set img = sh.Pictures.Insert(fNameAndPath)
    Set shp = sh.Shapes.AddPicture(fNameAndPath, msoFalse, msoCTrue, rng.Left, rng.Top, rng1.Width, rng1.Width / calcul)
    
    Application.ScreenUpdating = True
End Sub


Sub InsertarCapturaDocumentosSolicitados()
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
    'zoomF = 100
    '========================================================================'

    rng.Select

    fNameAndPath = Application.GetOpenFilename(Title:="Seleccionar una imagen")
    Set img = LoadPicture(fNameAndPath)
    calcul = img.Width / img.Height
    'Set img = sh.Pictures.Insert(fNameAndPath)
    Set shp = sh.Shapes.AddPicture(fNameAndPath, msoFalse, msoCTrue, rng.Left, rng.Top, rng1.Width, rng1.Width / calcul)
    
    Application.ScreenUpdating = True
End Sub
```

## Descripción

| Sub | Descripción |
|-----|-------------|
| **InsertarLogo** | Inserta imagen en celda B (fila del botón + 3), con ancho del rango B1:F1 |
| **InsertarCapturaCriteriosSeleccion** | Inserta imagen en celda C (fila del botón + 2), con ancho del rango C1:S1 |
| **InsertarCapturaDocumentosSolicitados** | Inserta imagen en celda W (fila del botón + 2), con ancho del rango W1:AM1 |

| Característica | Detalle |
|----------------|---------|
| **Activación** | Se ejecuta desde un botón de formulario (Application.Caller) |
| **Proporción** | Calcula la relación ancho/alto para mantener la proporción |
| **Diálogo** | Utiliza `GetOpenFilename` con título en español |
| **Actualización** | Desactiva la actualización de pantalla para mayor velocidad |

---

*Módulo: 12 - VBA Images | Archivo: 02 - InsertImage*
