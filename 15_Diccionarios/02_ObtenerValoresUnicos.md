# Obtener Valores Únicos de una Columna

Extrae los valores únicos (sin duplicados) de un rango de una columna utilizando un diccionario de VBA.

## Código VBA

```vba
'//Localizar y eliminar duplicados en una columna
Function ObtenerValoresUnicos1D(ByVal rng As Range) As Variant
    If rng.Count = 1 Then ObtenerValoresUnicos1D = rng.Value: Exit Function
    Dim Dic As Object, i As Long, arr()
    arr = rng.Value
    Set Dic = CreateObject("Scripting.Dictionary")
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) <> "" And Dic.Exists(arr(i, 1)) = False Then
            Dic.Add arr(i, 1), ""
        End If
    Next i
    ObtenerValoresUnicos1D = Dic.Keys
End Function

Sub FILTRAR_ELEMENTOS_UNICOS()
    Dim sh As Worksheet, arr As Variant, i As Integer
    Dim rng As Range
    Dim Dic As Object
    
    Set sh = ActiveSheet
    Set rng = sh.Range("I3:I27")
    arr = rng.Value
    Set Dic = CreateObject("Scripting.Dictionary")
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) <> "" And Dic.Exists(arr(i, 1)) = False Then
            Dic.Add arr(i, 1), ""
        End If
    Next i
    sh.Range("J3").Resize(5, 1) = Application.Transpose(Dic.Keys)
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Función `ObtenerValoresUnicos1D`** | Recibe un rango y retorna un array con los valores únicos |
| **Sub `FILTRAR_ELEMENTOS_UNICOS`** | Filtra valores únicos del rango I3:I27 y los coloca en J3 |
| **Método** | Utiliza un diccionario donde las claves son los valores únicos |
| **Filtro** | Excluye celdas vacías y valores duplicados |
| **Resultado** | Los valores únicos se escriben en el rango destino mediante `Transpose` |

---

*Módulo: 09 - VBA Dictionary | Archivo: 01 - Lay gia tri duy nhat*
