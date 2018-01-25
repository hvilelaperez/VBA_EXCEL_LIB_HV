# Uso de RefersTo

Muestra cómo trabajar con nombres definidos (Named Ranges) y sus propiedades RefersTo.

## Código VBA

```vba
Option Explicit
Sub test()
Dim p As Variant
    Debug.Print ActiveSheet.Name
'    p = Sheets(ActiveSheet.Name).Names("Print_Area").RefersToRange.Value
'    MsgBox "Print_Area: " & UBound(p, 1) & " filas, " & _
'    UBound(p, 2) & " columnas"
'    Debug.Print ActiveWorkbook.Names("MyRange").Name
    ActiveSheet.Cells.Interior.Pattern = xlNone
    ThisWorkbook.Names.Add Name:="MyRange", RefersTo:="=PrintArea!$M$1:$D$12"
    Debug.Print ActiveWorkbook.Names("MyRange").RefersTo
    Debug.Print ActiveWorkbook.Names("MyRange").Name
    ActiveSheet.Range("MyRange").Interior.Color = vbYellow
    
    Debug.Print ActiveWorkbook.Names("MyRange").RefersTo Like "*!*"
    Debug.Print ActiveWorkbook.Names("MyRange").RefersTo Like "*[[]*"
    Debug.Print ActiveWorkbook.Names("MyRange").RefersTo Like "*REF!*"
    
    ActiveSheet.PageSetup.PrintArea = ActiveSheet.Range("MyRange").Address
End Sub



Sub test_refer()
    Dim countRow&, countCol&
    countRow = 20
    countCol = 10
    ActiveSheet.Range("rngTest").Interior.Color = xlColorIndexNone
    
 ThisWorkbook.Names.Item("rngTest").RefersTo = ThisWorkbook.Names.Item("rngTest").RefersToRange.Resize(countRow, countCol)
 ActiveSheet.Range("rngTest").Select
    ActiveSheet.Range("rngTest").Interior.Color = vbYellow
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `Names.Add` | Crea un nuevo nombre definido |
| `RefersTo` | Referencia del nombre al rango |
| `RefersToRange` | Obtiene el rango referenciado |
| `Resize` | Redimensiona el rango referenciado |
| Patrones Like | Detectan referencias rotas (#REF!) |

---

*Módulo: 31_CodigosGenerales | Archivo: 17_UsoRefersTo.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
