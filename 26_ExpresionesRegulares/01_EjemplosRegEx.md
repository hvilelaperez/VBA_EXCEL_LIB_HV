# Ejemplos de Expresiones Regulares en VBA

Colección de ejemplos prácticos que demuestran el uso de expresiones regulares (RegExp) en VBA para validación de datos, búsqueda de patrones y formato de texto.

## Código VBA

```vba
Option Explicit

Function RegExample(inputString As String, _
                    Optional globalRe As Boolean = True, _
                    Optional ignoreCase As Boolean = True) As Boolean
    
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")

    With re
        .Global = globalRe
        .Pattern = "^[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})$"
        .ignoreCase = ignoreCase
        RegExample = .test(inputString)
    End With

Sub timtheo_vitri()
    Dim str As String, kqua
    str = "Lua nep la lua nep lang,lua len lop lop ,long nang lang lang "
    With CreateObject("VBScript.RegExp")
        .Global = True
        .IgnoreCase = True
        .Pattern = "Lua"
            Set kqua = .Execute(str) ' El resultado sera un array con 3 cadenas: Lua lua lua
        .Pattern = "^Lua"
            Set kqua = .Execute(str) ' El resultado sera un solo valor: la primera cadena Lua
    End With
' De manera similar se puede aplicar para otros casos ^^
End Sub

Sub Regx1()
    Dim str As String, kqua
    str = " Messi dang co bong, anh sut bong, vaooooo)"
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "[vao]{5}"
        Set kqua = .Execute(str) ' Retorna el resultado como la cadena vaooo
        .IgnoreCase = True
        .Pattern = "[a-z]* +"
        Set kqua = .Execute(str) ' Separa todas las letras de a-z seguidas de un espacio
    End With
End sub


Dim rRange As Range
Dim rCell As Range

Dim re As Object
Set re = CreateObject("vbscript.regexp")
With re
  .Pattern = "^\d\d?:\d\d[aApP][mM] - \d\d?:\d\d[aApP][mM]$"
  .Global = False
  .IgnoreCase = False
End With

Set rRange = Range("A2", "G225")

For Each rCell In rRange.Cells
    If re.Test(rCell) Then
        rCell.Interior.Color = RGB(0, 250, 0)
    Else
        rCell.Interior.Color = RGB(250, 0, 0)
    End If
Next rCell
```

## Descripción

| Función/Sub | Descripción |
|---|---|
| `RegExample` | Valida si un string cumple con el formato de correo electrónico usando una expresión regular |
| `timtheo_vitri` | Busca la palabra "Lua" en una cadena de texto, mostrando diferencias entre búsqueda global y con ancla de inicio |
| `Regx1` | Ejemplos de patrones: búsqueda de caracteres repetidos `[vao]{5}` y separación de palabras por espacios `[a-z]* +` |
| Código general | Resalta celdas en verde si el contenido coincide con un formato de hora (HH:MM AM/PM), de lo contrario las marca en rojo |

---

*Módulo: 25 - VBA Regular expression | Archivo: 00 - Codes samples Regular expression.txt*
