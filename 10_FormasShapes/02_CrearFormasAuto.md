# Crear Formas AutoShape

Código VBA para crear todas las formas AutoShape disponibles en Excel con estilos aleatorios.

## Código VBA

```vba
Option Explicit


Private Sub cmdCreateAutoshapes_Click()
    Dim i As Integer
    Dim t As Integer
    Dim shp As Shape
    
    If ActiveSheet.Shapes.Count > 2 Then
        DeleteAllAutoShapes
    End If
    
    Randomize
    t = 15
    For i = 1 To 137
        Set shp = ActiveSheet.Shapes.AddShape(i, 48, t, 96, 60)
        shp.TextFrame.Characters.Text = i
        If CInt(Application.Version) >= 12 Then
            If i = 25 Then
                'Tratar como línea
                shp.ShapeStyle = msoLineStylePreset1
            Else
                'Seleccionar estilo aleatoriamente
                shp.ShapeStyle = Int(Rnd() * 42 + 1)
            End If
        End If
        t = t + 75
    Next
    ' saltar 138 - no soportado
    If CInt(Application.Version) >= 12 Then
        For i = 139 To 179
            Set shp = ActiveSheet.Shapes.AddShape(i, 48, t, 96, 60)
            shp.TextFrame.Characters.Text = i
            shp.ShapeStyle = Int(Rnd() * 42 + 1)
            t = t + 75
        Next
        ' Estas formas no tienen TextFrame
        For i = 180 To 183
            Set shp = ActiveSheet.Shapes.AddShape(i, 48, t, 96, 60)
            t = t + 75
        Next
    End If
    
End Sub

Private Sub cmdDeleteShapes_Click()
    DeleteAllAutoShapes
End Sub
```

## Descripción

| Subrutina | Función |
|-----------|---------|
| `cmdCreateAutoshapes_Click` | Crea todas las formas AutoShape disponibles (1-183) |
| `cmdDeleteShapes_Click` | Elimina todas las formas AutoShape creadas |

### Características

- **Formas creadas**: 1 a 183 (excepto 138 que no es soportada)
- **Estilo**: Seleccionado aleatoriamente entre 42 estilos preestablecidos
- **Excepción**: La forma 25 se trata como línea con estilo preestablecido 1
- **Formas 180-183**: No incluyen cuadro de texto

### Requisitos

- Excel 2007 o superior (versión >= 12) para estilos de forma
- Sub `DeleteAllAutoShapes` debe estar disponible en el módulo

---

*Módulo: 04 - VBA Shapes | Archivo: 01 - Create auto shapes*