# Propiedades de Checkbox

Código VBA para crear controles CheckBox con formato personalizado y superponer cuadros de texto.

## Código VBA

```vba
Sub CreateFormsCheckBoxes()

    Dim objCtrl As Object
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim dblWidth As Double
    Dim dblHeight As Double
    Dim shp As Shape
    Dim strCaption As String
    Dim lngCapLen As Long
        
    strCaption = "Mi CheckBox" 'Editar con el texto requerido para el CheckBox.
    lngCapLen = Len(strCaption)
    
    With Sheets("Hoja1")
        With ActiveCell
            'Posición y tamaño respecto a la ubicación de la celda
            'Tamaño y posición ajustados ligeramente para evitar cubrir las líneas de cuadrícula
            dblLeft = .Left + 1
            dblWidth = .Width * 3 - 2
            dblHeight = .Height - 2
            dblTop = .Top + 1
        End With
        Set objCtrl = .CheckBoxes.Add(dblLeft, dblTop, dblWidth, dblHeight)
        objCtrl.Name = "chkBox1"
        objCtrl.Characters.Text = ""    'Sin texto dentro del CheckBox
        
        'Superponer un TextBox sobre el CheckBox
        dblLeft = dblLeft + 15      'Mover el TextBox a la derecha para que el checkbox sea visible
        dblWidth = dblWidth - 15    'Reducir ancho del TextBox para alinear con el lado derecho del CheckBox
        Set shp = .Shapes.AddTextbox(msoTextOrientationHorizontal, dblLeft, dblTop, dblWidth, dblHeight)
    End With
    
    shp.TextFrame2.TextRange.Characters.Text = strCaption       'Insertar el texto
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle     'Centrar fuente verticalmente
    shp.Line.Visible = msoFalse              'Sin borde visible
    
    With shp.TextFrame2.TextRange.Characters(1, lngCapLen).Font
        .NameComplexScript = "Times New Roman"
        .Size = 12
        .Bold = msoTrue
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 0, 0)
        
        '**************************************************************
        'Lo siguiente se refiere a la Fuente; no al relleno de la forma.
        'Todo sin probar.
        'Si se hace clic derecho en la forma en la hoja de trabajo
        'y se selecciona la cinta de herramientas de dibujo que aparece,
        'ver el bloque de Estilos de Word Art al que aplican estas opciones.
        '.Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
        '.Fill.ForeColor.TintAndShade = 0
        '.Fill.ForeColor.Brightness = 0
        '.Fill.Transparency = 0
        '.Fill.Solid
        '***************************************************************
    End With
    
    ActiveCell.Offset(2, 0).Select  'Mover la selección fuera de las formas
    
End Sub


Sub ClearShapes()
'Usar esta sub para eliminar todos los CheckBox y TextBox durante las pruebas.
Dim shp As Shape

For Each shp In ActiveSheet.Shapes
    shp.Delete
Next shp

End Sub
```

## Descripción

| Subrutina | Función |
|-----------|---------|
| `CreateFormsCheckBoxes` | Crea un CheckBox con texto superpuesto y formato personalizado |
| `ClearShapes` | Elimina todas las formas de la hoja activa |

### Pasos de CreateFormsCheckBoxes

1. Define el texto del CheckBox (`strCaption`)
2. Calcula posición y tamaño basado en la celda activa
3. Crea el control CheckBox en la hoja
4. Superpone un TextBox con el texto formateado
5. Aplica formato de fuente (Times New Roman, 12pt, negrita, rojo)

---

*Módulo: 02 - VBA Button, checkbox, canvas | Archivo: 01 - Property Checkbox*