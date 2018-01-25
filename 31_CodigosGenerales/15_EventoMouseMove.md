# Evento MouseMove

Evento que se ejecuta al mover el mouse sobre una etiqueta, escribiendo texto y cambiando colores.

## Código VBA

```vba
Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sh As Worksheet, ultimaFila&
    Dim shp As Shape
    Set sh = ActiveSheet
    If sh.Range("A1") = "" Then
    ultimaFila = 0
    Else
        ultimaFila = sh.Range("A" & Rows.Count).End(xlUp).Row
    End If

        sh.Cells(ultimaFila + 1, 1) = "I    have    a    great    day!"
    If ultimaFila > 20 Then ActiveWindow.SmallScroll (1)
    Set shp = sh.Shapes("Rectangle 1")
    With shp.Fill.ForeColor
    If .RGB = RGB(0, 0, 255) Then
        .RGB = RGB(197, 240, 255)
    Else
        .RGB = RGB(0, 0, 255)
    End If
     End With
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| Evento | `MouseMove` - Se ejecuta al mover el mouse sobre el control |
| Control | Label1 (etiqueta) |
| Acción 1 | Escribe texto en la siguiente fila vacía |
| Acción 2 | Desplaza la ventana si hay más de 20 filas |
| Acción 3 | Alterna el color de un rectángulo entre azul y celeste |

---

*Módulo: 31_CodigosGenerales | Archivo: 15_EventoMouseMove.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
