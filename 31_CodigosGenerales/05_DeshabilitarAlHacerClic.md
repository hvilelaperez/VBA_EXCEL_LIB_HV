# Deshabilitar Filas al Hacer Clic en Botón

Macros que ocultan o muestran filas específicas al hacer clic en botones u opciones de selección (OptionButton).

## Código VBA

```vba
Option Explicit

Sub hide_Conception()
    Dim shp As Shape, rowTxtBox As Integer, rngToHide As Range
    Dim sh As Worksheet
    Set sh = ActiveSheet

    Set shp = sh.Shapes("TextBox 164")
    rowTxtBox = shp.TopLeftCell.Row
    Set rngToHide = sh.Range("A" & (rowTxtBox - 2))

    If sh.OLEObjects("OptionButton1").Object.Value = True Then
        rngToHide.EntireRow.Hidden = False
    Else
        rngToHide.EntireRow.Hidden = True
    End If
End Sub

Sub hide_Clause()
    Dim shp As Shape, shp2 As Shape, rowTxtBox As Integer
    Dim sh As Worksheet, rngToHide As Range, rngToHide2 As Range
    Set sh = ActiveSheet

    Set shp = sh.Shapes("Graphic 247")
    rowTxtBox = shp.TopLeftCell.Row
    Set rngToHide = sh.Range("A" & (rowTxtBox - 2))

    Select Case sh.Range("L" & (rowTxtBox - 3))
        Case "": rngToHide.EntireRow.Hidden = True
        Case "No": rngToHide.EntireRow.Hidden = True
        Case "Si": rngToHide.EntireRow.Hidden = False
    End Select

    Set rngToHide2 = sh.Range("A" & (rowTxtBox - 8))
    Select Case sh.Range("X" & (rowTxtBox - 9))
        Case "": rngToHide.EntireRow.Hidden = True
        Case "No": rngToHide.EntireRow.Hidden = True
        Case "Si": rngToHide.EntireRow.Hidden = False
    End Select
End Sub
```

## Descripción

| Macro | Descripción |
|-------|-------------|
| `hide_Conception` | Muestra/oculta filas según el valor del OptionButton1 (verdadero = mostrar, falso = ocultar) |
| `hide_Clause` | Muestra/oculta filas basándose en el valor de celdas de texto ("Si"/"No"/vacío) en columnas L y X |

| Elemento técnico | Descripción |
|------------------|-------------|
| `Shapes("TextBox 164")` | Referencia al cuadro de texto que define la posición |
| `OLEObjects("OptionButton1")` | Referencia al botón de opción para verificar su estado |
| `.EntireRow.Hidden` | Propiedad que oculta/muestra la fila completa |

---

*Módulo: 31_CodigosGenerales | Archivo: 05_DeshabilitarAlHacerClic.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
