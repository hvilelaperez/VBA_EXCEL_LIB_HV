# Aplicación de Búsqueda (Find)

Busca y elimina usuarios de una hoja con confirmación.

## Código VBA

```vba
Sub EliminarUsuarioConConfirmacion()
Dim oFound As Range
Dim oLookin As Range
Dim sLookFor As String
SearchUserName = InputBox("Por favor ingrese un usuario")
sLookFor = SearchUserName
Set oLookin = Worksheets("Sheet13").UsedRange
Set oFound = oLookin.Find(what:=sLookFor, _
LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
If Not oFound Is Nothing Then
uChoice = MsgBox("¿Está seguro de que desea eliminar el usuario: " & _
vbNewLine & sLookFor, vbQuestion + vbYesNo, "Confirmación de Eliminación")
If uChoice = vbYes Then
oLookin.Range("A" & oFound.Row).EntireRow.Delete
MsgBox "El usuario: " & vbNewLine & sLookFor & _
vbNewLine & "ha sido eliminado exitosamente."
End If
Else
MsgBox "No se encontró nada para: " & vbNewLine & _
sLookFor
End If
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| Función | `Find` - Busca un valor en un rango |
| Parámetro `LookIn` | xlValues - Busca en valores de celda |
| Parámetro `LookAt` | xlPart - Busca coincidencia parcial |
| Parámetro `MatchCase` | False - No distingue mayúsculas/minúsculas |
| InputBox | Solicita el nombre del usuario a buscar |
| MsgBox | Confirma la eliminación antes de ejecutarla |

---

*Módulo: 31_CodigosGenerales | Archivo: 16_AplicacionBuscar.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
