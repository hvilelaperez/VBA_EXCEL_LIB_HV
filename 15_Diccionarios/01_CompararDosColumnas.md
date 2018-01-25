# Comparar Valores de Dos Columnas

Compara los valores de dos columnas y extrae los valores que existen en la segunda columna pero no en la primera, utilizando un diccionario para la verificación eficiente.

## Código VBA

```vba
Sub CompararDosColumnas()
    Dim Dic As Object, KQ As Range
    Dim TmpEntrada As Variant, TmpSalida As Variant, Tmp As Variant, I As Long, J As Long
    Set Dic = CreateObject("Scripting.Dictionary") 'Inicializar diccionario
    With Sheet1
        TmpSalida = .Range("A3:A" & .Cells(.Rows.Count, "A").End(xlUp).Row).Value 'Datos existentes en columna A
        TmpEntrada = .Range("G3:G" & .Cells(.Rows.Count, "G").End(xlUp).Row).Value 'Datos en columna G
        Set KQ = .Range("A" & .Cells(.Rows.Count, "A").End(xlUp).Row).Offset(1) 'Ubicar resultado
    End With
    For I = LBound(TmpSalida, 1) To UBound(TmpSalida, 1) 'Recorrer columna A
        If Not Dic.Exists(CStr(TmpSalida(I, 1))) Then Dic.Add CStr(TmpSalida(I, 1)), I 'Agregar elementos no repetidos al Dic
    Next
    ReDim Tmp(0) 'Crear array unidimensional nuevo para guardar resultado tras comparar
    J = -1 'El array unidimensional empieza desde el elemento 0 por eso J = -1
    For I = LBound(TmpEntrada, 1) To UBound(TmpEntrada, 1) 'Recorrer columna G
        If Not Dic.Exists(CStr(TmpEntrada(I, 1))) Then 'Si el código no existe en el Dic
            J = J + 1 'Por eso J = -1 arriba
            ReDim Preserve Tmp(J) 'Expandir array sin perder datos antiguos
            Tmp(J) = TmpEntrada(I, 1) 'Colocar código que está en columna G pero no en columna A en el array de resultados
        End If
    Next
    KQ.Resize(UBound(Tmp, 1) + 1, 1) = Application.Transpose(Tmp) 'Pegar resultado en la hoja de cálculo
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Objetivo** | Encontrar valores de la columna G que no existan en la columna A |
| **Método** | Utiliza un diccionario para almacenar valores de la columna A y verificar existencia |
| **Salida** | Los valores únicos de columna G que no están en columna A se escriben debajo de la última fila de columna A |
| **Ventaja** | El diccionario permite búsquedas en tiempo O(1), mucho más rápido que buscar en rangos |

---

*Módulo: 09 - VBA Dictionary | Archivo: 00 - So sanh gia tri 2 cot*
