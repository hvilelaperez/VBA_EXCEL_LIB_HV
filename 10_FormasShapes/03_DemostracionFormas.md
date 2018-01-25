# Demostración de Formas

Código VBA completo para crear, formatear y conectar formas de flujo con formato compatible con Excel 2003 y 2007+.

## Código VBA

```vba
Option Explicit

Private m_Worksheet As Worksheet

Private Sub cmdClearShapes_Click()
    DeleteAllAutoShapes 'en módulo MCommon
End Sub

Private Sub cmdRunDemo_Click()

    Dim oShape1 As Shape
    Dim oShape2 As Shape
    Dim oConnector As Shape
    
    Set m_Worksheet = ActiveSheet
    
    If m_Worksheet.Shapes.Count > 2 Then
        DeleteAllAutoShapes
    End If
    
    Set oShape1 = AddShapeToRange(msoShapeFlowchartProcess, "B3:D5")
    AddFormattedTextToShape oShape1, "Primera Forma"
    
    Set oShape2 = AddShapeToRange(msoShapeFlowchartAlternateProcess, "B9:D11")
    AddFormattedTextToShape oShape2, "Segunda Forma"
    
    Set oConnector = AddConnectorBetweenShapes(msoConnectorStraight, oShape1, oShape2)
    
    If CInt(Application.Version) < 12 Then
        FormatShape2003 oShape1
        FormatShape2003 oShape2
        FormatConnector2003 oConnector
        'IlluminateShapeText2003
    Else
        FormatShape2007 oShape1
        FormatShape2007 oShape2
        FormatConnector2007 oConnector
        'IlluminateShapeText2007
    End If
    
    Set oShape1 = Nothing
    Set oShape2 = Nothing
    Set oConnector = Nothing
    
    Set m_Worksheet = Nothing

End Sub

Private Function AddShapeToRange(ShapeType As MsoAutoShapeType, _
                                 sAddress As String) As Shape
    With m_Worksheet.Range(sAddress)
        Set AddShapeToRange = m_Worksheet.Shapes.AddShape(ShapeType, .Left, .Top, .Width, .Height)
    End With
End Function

Private Sub AddFormattedTextToShape(oShape As Shape, _
                                    sText As String)
    If Len(sText) > 0 Then
        With oShape.TextFrame
            .Characters.Text = sText
            .Characters.Font.Name = "Garamond"
            .Characters.Font.Size = 12
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
    End If
End Sub

Private Function AddConnectorBetweenShapes(ConnectorType As MsoConnectorType, _
                                          oBeginShape As Shape, _
                                          oEndShape As Shape) As Shape

    ' NOTA: Estas constantes de sitio de conexión solo funcionan para formas rectangulares con
    '       4 puntos de conexión. La llamada a RerouteConnections a continuación
    '       redirigirá automáticamente el conector al camino más corto entre las formas.
    Const LADO_SUPERIOR As Integer = 1
    Const LADO_IZQUIERDO As Integer = 2
    Const LADO_INFERIOR As Integer = 3
    Const LADO_DERECHO As Integer = 4
    
    Dim oConnector As Shape
    Dim x1 As Single
    Dim x2 As Single
    Dim y1 As Single
    Dim y2 As Single
    
    With oBeginShape
        x1 = .Left + .Width / 2
        y1 = .Top + .Height
    End With
    
    With oEndShape
        x2 = .Left + .Width / 2
        y2 = .Top
    End With
    
    ' Excel 2007 usa coordenadas absolutas para el segundo punto,
    ' de la función AddConnector. Versiones anteriores de Excel
    ' usan coordenadas relativas. Pero, ... (continúa abajo)
    If CInt(Application.Version) < 12 Then
        x2 = x2 - x1
        y2 = y2 - y1
    End If
    
    Set oConnector = m_Worksheet.Shapes.AddConnector(ConnectorType, x1, y1, x2, y2)
    
    ' ... se pueden usar valores Single positivos si se conectan
    ' los puntos finales con BeginConnect y EndConnect:
    oConnector.ConnectorFormat.BeginConnect oBeginShape, LADO_INFERIOR
    oConnector.ConnectorFormat.EndConnect oEndShape, LADO_SUPERIOR
    oConnector.RerouteConnections
    
    Set AddConnectorBetweenShapes = oConnector
    
    Set oConnector = Nothing

End Function

Private Sub FormatConnector2003(oConnector As Shape)
    If oConnector.Connector Or oConnector.Type = msoLine Then
        'aproximación del estilo de línea preestablecido #17 de Excel 2007
        With oConnector
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.Weight = 2
            .Line.ForeColor.RGB = RGB(192, 0, 0)
        End With
    End If
End Sub

Private Sub FormatConnector2007(oConnector As Shape)
    With oConnector
        If .Connector Or .Type = msoLine Then
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .ShapeStyle = msoLineStylePreset17
        End If
    End With
End Sub

Private Sub FormatShape2003(oShape As Shape)
    'aproximación del estilo de forma preestablecido #2 de Excel 2007
    With oShape
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(79, 129, 189)
        .Line.Weight = 2
        .Shadow.OffsetX = 0.8
        .Shadow.OffsetY = 0.8
        .Shadow.ForeColor.RGB = RGB(192, 192, 192)
        .Shadow.Transparency = 0.5
        .Shadow.Visible = msoTrue
    End With
End Sub

Private Sub FormatShape2007(oShape As Shape)
    oShape.ShapeStyle = msoShapeStylePreset2
End Sub

Private Sub IlluminateShapeText2003()
    On Error Resume Next

    Dim oShape As Shape
    Dim numChars As Integer
    
    For Each oShape In m_Worksheet.Shapes
        If oShape.Type = msoAutoShape And oShape.AutoShapeType > 0 Then
            'Genera error cuando la forma no contiene texto:
            numChars = oShape.TextFrame.Characters.Count
            If Err.Number <> 0 Then
                Debug.Print oShape.Name
                Err.Clear
            ElseIf numChars > 0 Then
                With oShape.TextFrame.Characters(1, 1).Font
                    .Name = "Garamond"
                    'Lo siguiente no funciona en Excel 2003 y anteriores.
                    'Redondea el color al preestablecido más cercano.
                    .Color = RGB(192, 80, 77)
                    .Size = 24
                    .Bold = True
                End With
            End If
        End If
    Next
End Sub

Private Sub IlluminateShapeText2007()
    Dim oShape As Shape
    
    For Each oShape In m_Worksheet.Shapes
        If oShape.Type = msoAutoShape And oShape.AutoShapeType > 0 Then
            'No genera error cuando la forma no contiene texto:
            If oShape.TextFrame2.HasText Then
                With oShape.TextFrame2.TextRange.Characters(1, 1).Font
                    .Name = "Garamond"
                    'Establece el color correctamente - sin redondeo como
                    'en versiones anteriores de Excel.
                    .Fill.ForeColor.RGB = RGB(192, 80, 77)
                    .Size = 24
                    .Bold = True
                    .Reflection.Type = msoReflectionType1
                    .Shadow.Type = msoShadow14
                    .Glow.Color.RGB = RGB(192, 137, 45)
                    .Glow.radius = 5
                End With
            End If
            
        End If
    Next
End Sub
```

## Descripción

### Subrutinas principales

| Subrutina | Función |
|-----------|---------|
| `cmdRunDemo_Click` | Crea dos formas de flujo y las conecta con un conector |
| `cmdClearShapes_Click` | Elimina todas las formas de la hoja |
| `AddShapeToRange` | Agrega una forma en el rango especificado |
| `AddFormattedTextToShape` | Inserta texto formateado en una forma |
| `AddConnectorBetweenShapes` | Crea un conector entre dos formas |

### Formato por versión de Excel

| Subrutina | Excel | Descripción |
|-----------|-------|-------------|
| `FormatShape2003` | 2003 | Formato manual con colores y sombras |
| `FormatShape2007` | 2007+ | Usa estilo preestablecido #2 |
| `FormatConnector2003` | 2003 | Línea roja con punta de flecha triangular |
| `FormatConnector2007` | 2007+ | Usa estilo preestablecido #17 |
| `IlluminateShapeText2003` | 2003 | Ilumina texto con color rojo |
| `IlluminateShapeText2007` | 2007+ | Agrega reflexión, sombra y resplandor |

### Constantes de conexión

| Constante | Valor | Posición |
|-----------|-------|----------|
| `LADO_SUPERIOR` | 1 | Parte superior de la forma |
| `LADO_IZQUIERDO` | 2 | Lado izquierdo |
| `LADO_INFERIOR` | 3 | Parte inferior |
| `LADO_DERECHO` | 4 | Lado derecho |

---

*Módulo: 04 - VBA Shapes | Archivo: 03 - Shapes demo*