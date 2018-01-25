# Eliminar Acentos

Macro para eliminar acentos de caracteres franceses en celdas de Excel, reemplazándolos por sus equivalentes sin tilde.

## Código VBA

```vba
Const accent As String = "ÀÁÂÃÄÅàáâãäåÈÉÊËèéêëÌÍÎÏìíîïÒÓÔÕÖòóôõöÙÚÛÜùúûüÝÿñçÇ"
Const noAccent As String = "AAAAAAaaaaaaOOOOOOooooooEEEEeeeeIIIIiiiiUUUUuuuuyNnCc"

'========================================================================'
'Aquí para ejecutar la macro - no es necesario usar "Option Explicit"
' Esta macro solo elimina acentos franceses, no acentos vietnamitas

Public Sub Remove_Accent()
Dim c As Range
For Each c In Cells.SpecialCells(xlCellTypeConstants, 23)
If TypeName(c.Value) = "String" Then
    If InStr(1, c.Text, "@") > 0 Then
        c = noneAccents(c.Text)
    Else
        c = noneAccents(c.Text)
    End If
End If
Next
End Sub

'Fin de la macro
'========================================================================'

'========================================================================'
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
'Línea de código para la función - No tocar aquí

Private Function noneAccents(ByRef s As String) As String
Dim i As Integer
Dim lettre As String * 1
noneAccents = s
For i = 1 To Len(accent)
    lettre = Mid$(accent, i, 1)
    If InStr(noneAccents, lettre) > 0 Then
        noneAccents = Replace(noneAccents, lettre, Mid$(noAccent, i, 1))
    End If
Next i
End Function

'Fin de la función
'========================================================================'
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `accent` | Constante con caracteres acentuados franceses (incluye ñ y ç) |
| `noAccent` | Constante con los caracteres equivalentes sin acento |
| `Remove_Accent` | Sub principal que recorre todas las celdas con constantes y elimina acentos |
| `noneAccents` | Función auxiliar que reemplaza cada acento por su versión sin tilde |
| `SpecialCells(xlCellTypeConstants, 23)` | Selecciona solo celdas que contienen constantes de texto |

> **Nota:** Esta macro está diseñada específicamente para caracteres franceses. Para otros idiomas se deben ajustar las tablas de caracteres.

---

*Módulo: 31_CodigosGenerales | Archivo: 02_EliminarAcentos.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
