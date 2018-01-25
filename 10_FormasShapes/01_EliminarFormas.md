# Eliminar Todas las Formas

Código VBA para eliminar todas las formas AutoShape de una hoja de cálculo, preservando los botones de comando.

## Código VBA

```vba
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Sub DeleteAllAutoShapes()
    
    'Normalmente, se puede hacer lo siguiente:
    '
    '   Worksheet.Shapes.SelectAll
    '   Selection.Delete
    '
    ' Sin embargo, eso también eliminaría los botones de comando.
    ' Por lo tanto, el código siguiente se escribió para demostrar una forma de
    ' usar el objeto ShapeRange
    
    Dim arrShapeNames()     As Variant 'debe ser Variant para funcionar con Shapes.Range()
    Dim shp                 As Excel.Shape
    Dim sr                  As Excel.ShapeRange
    Dim ws                  As Worksheet
    Dim i                   As Integer
    
    Set ws = ActiveSheet
    
    ReDim arrShapeNames(ws.Shapes.Count - 1)
    i = 0
    For Each shp In ws.Shapes
        If shp.Type = msoAutoShape Or shp.Type = msoCallout Then '¡sin botones!
            arrShapeNames(i) = shp.Name
            i = i + 1
        End If
    Next
    
    If i = 0 Then Exit Sub  ' no había formas válidas
    
    ' El primer ReDim pudo haber incluido espacio para tipos de forma
    ' inválidos - en este caso los botones. Este ReDim
    ' ajusta el tamaño del array en consecuencia
    ReDim Preserve arrShapeNames(i - 1)
    Set sr = ws.Shapes.Range(arrShapeNames)
    sr.Select
    sr.Delete
    
    Set sr = Nothing
    Set ws = Nothing
    
End Sub

Public Sub SleepAndDo(NumberOfMilliseconds As Long)
  Sleep NumberOfMilliseconds
  DoEvents
End Sub
```

## Descripción

| Subrutina | Función |
|-----------|---------|
| `DeleteAllAutoShapes` | Elimina todas las formas AutoShape y Callout preservando botones |
| `SleepAndDo` | Pausa la ejecución por milisegundos especificados |

### Proceso de eliminación

1. Recorre todas las formas de la hoja activa
2. Filtra solo formas `msoAutoShape` y `msoCallout` (excluye botones)
3. Almacena nombres de formas válidas en array
4. Crea ShapeRange con las formas filtradas
5. Selecciona y elimina el grupo

### Tipos de forma excluidos

| Tipo | Descripción |
|------|-------------|
| `msoFormControl` | Botones de comando y otros controles de formulario |
| Cualquier otro tipo | No se elimina |

---

*Módulo: 04 - VBA Shapes | Archivo: 00 - Delelte all shapes*