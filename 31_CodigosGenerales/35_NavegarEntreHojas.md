# Navegar Entre Hojas de Cálculo

Macros para navegar rápidamente entre diferentes posiciones de hojas de cálculo en Excel usando accesos de teclado.

## Código VBA

```vba
Sub Navegacion_1()
'
' Macro Navegacion_1
'
' Tecla de acceso rápido del teclado: Ctrl+l

'
    ActiveSheet.Range("A11").Select
End Sub
Sub Imagen8_Clic()

End Sub
Sub Navegacion_2()
'
' Macro Navegacion_2
'
' Tecla de acceso rápido del teclado: Ctrl+p
    Dim shp As Shape, filaCuadroTexto As Integer, rngOcultar As Range
    Dim sh As Worksheet
    Set sh = ActiveSheet
    
    Set shp = sh.Shapes("Image 217")
    filaCuadroTexto = shp.TopLeftCell.Row + 15
    
'
    Range("B" & filaCuadroTexto).Select
End Sub

Sub Navegacion_F()
'
' Macro Navegacion_F
'
' Tecla de acceso rápido del teclado: Ctrl+f
    Dim shp As Shape, filaCuadroTexto As Integer, rngOcultar As Range
    Dim sh As Worksheet
    Set sh = ActiveSheet
    
    Set shp = sh.Shapes("Image 328")
    filaCuadroTexto = shp.TopLeftCell.Row + 10
'
    Range("B" & filaCuadroTexto).Select
End Sub
```

## Descripción

| Sub | Tecla | Descripción |
|-----|-------|-------------|
| `Navegacion_1` | Ctrl+l | Selecciona la celda A11 de la hoja activa |
| `Navegacion_2` | Ctrl+p | Navega a la celda B calculada a partir de la posición de una imagen específica |
| `Navegacion_F` | Ctrl+f | Navega a la celda B calculada a partir de la posición de otra imagen |

### Variables utilizadas

| Variable | Tipo | Descripción |
|----------|------|-------------|
| `shp` | Shape | Referencia a la imagen/objeto gráfico en la hoja |
| `filaCuadroTexto` | Integer | Fila calculada sumando un offset a la fila de la imagen |
| `rngOcultar` | Range | Rango para operaciones de ocultamiento (declarada pero no utilizada) |
| `sh` | Worksheet | Referencia a la hoja activa |

### Lógica

1. Se obtiene la referencia de un objeto gráfico (Shape) por nombre
2. Se calcula la fila destino sumando un offset a la fila donde se ubica la esquina superior izquierda de la imagen
3. Se selecciona la celda en la columna B de esa fila calculada

---

*Módulo: 96 - VBA Codes general | Archivo: 14 - Naviration entre les feuilles.txt*
