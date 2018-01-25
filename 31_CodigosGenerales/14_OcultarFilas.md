# Ocultar Filas

Módulo para ocultar/mostrar filas basado en controles de formulario y condiciones específicas.

## Código VBA

```vba
Option Explicit
Sub ocultar_Concepcion()
    Dim shp As Shape, filaCuadroTexto As Integer, rangoOcultar As Range
    Dim sh As Worksheet
    Set sh = ActiveSheet
    
    Set shp = sh.Shapes("TextBox 164")
    filaCuadroTexto = shp.TopLeftCell.Row
    Set rangoOcultar = sh.Range("A" & (filaCuadroTexto - 2))

    If sh.OLEObjects("OptionButton5").Object.Value = True Then
       rangoOcultar.EntireRow.Hidden = False
    Else
        rangoOcultar.EntireRow.Hidden = True
    End If
    
End Sub

Sub ocultar_Clausula()
    Dim shp As Shape, shp2 As Shape, filaCuadroTexto As Integer
    Dim sh As Worksheet, rangoOcultar As Range, rangoOcultar2 As Range
    Set sh = ActiveSheet
    
    Set shp = sh.Shapes("Graphic 247")
    filaCuadroTexto = shp.TopLeftCell.Row
    Set rangoOcultar = sh.Range("A" & (filaCuadroTexto - 2))
    Select Case sh.Range("L" & (filaCuadroTexto - 4))
        Case "": rangoOcultar.EntireRow.Hidden = True
        Case "Non": rangoOcultar.EntireRow.Hidden = True
        Case "Oui": rangoOcultar.EntireRow.Hidden = False
        End Select
    Set rangoOcultar2 = sh.Range("A" & (filaCuadroTexto - 8))
    Select Case sh.Range("X" & (filaCuadroTexto - 10))
        Case "": rangoOcultar2.EntireRow.Hidden = True
        Case "Non": rangoOcultar2.EntireRow.Hidden = True
        Case "Oui": rangoOcultar2.EntireRow.Hidden = False
        End Select
End Sub


Sub ocultar_CES()
    Dim shp As Shape, shp2 As Shape, filaCuadroTexto As Integer
    Dim sh As Worksheet, rangoOcultar As Range, rangoOcultar2 As Range, rangoOcultar3 As Range
    Dim numFilasOcultar As Integer
    Call OptimzeCode_Begin
    
    Set sh = ActiveSheet
    Set shp = sh.Shapes("Rectangle 35")
    filaCuadroTexto = shp.TopLeftCell.Row
    Set rangoOcultar = sh.Range("A" & (filaCuadroTexto + 1), "A" & (filaCuadroTexto + 34))
    Debug.Print filaCuadroTexto
    numFilasOcultar = sh.Range("G" & (filaCuadroTexto + 2)).Value
    Set rangoOcultar2 = sh.Range("A" & (filaCuadroTexto + 3 + numFilasOcultar * 2), "A" & (filaCuadroTexto + 34))
    Set rangoOcultar3 = sh.Range("A" & (filaCuadroTexto + 5 + numFilasOcultar * 2), "A" & (filaCuadroTexto + 34))
    If sh.Range("G" & filaCuadroTexto).Value = "Corps d'�tat s�par�s" Then
        rangoOcultar.EntireRow.Hidden = True
    ElseIf (sh.Range("G" & (filaCuadroTexto + 2)) = "") Or (numFilasOcultar = 0) Then
        MsgBox "Nb de Lots doit �tre sup�rieur � 0", vbCritical
        rangoOcultar.EntireRow.Hidden = False
        rangoOcultar2.EntireRow.Hidden = True
    ElseIf numFilasOcultar > 0 Then
        rangoOcultar.EntireRow.Hidden = False
        rangoOcultar3.EntireRow.Hidden = True
    Else:
        rangoOcultar2.EntireRow.Hidden = False
    End If
    
    Call OptimizeCode_End
End Sub

Sub ocultar_Decoupage()
    Dim shp As Shape, shp2 As Shape, filaCuadroTexto As Integer
    Dim sh As Worksheet, rangoOcultar As Range, rangoOcultar2 As Range, rangoOcultar3 As Range
    Dim numFilasOcultar As Integer
    Call OptimzeCode_Begin
    Set sh = ActiveSheet
    Set shp = sh.Shapes("Rectangle 16")
    filaCuadroTexto = shp.TopLeftCell.Row
    Set rangoOcultar = sh.Range("A" & (filaCuadroTexto + 1), "A" & (filaCuadroTexto + 18))
    
    numFilasOcultar = sh.Range("G" & (filaCuadroTexto + 2)).Value
    
    Set rangoOcultar3 = sh.Range("A" & (filaCuadroTexto + 1), "A" & (filaCuadroTexto + 3 + numFilasOcultar * 2))
    If sh.Range("G" & filaCuadroTexto).Value = "Non" Then
        rangoOcultar.EntireRow.Hidden = True
    ElseIf (numFilasOcultar <= 0) Then
        MsgBox "Nb de tranches doit �tre sup�rieur � 0", vbCritical
        rangoOcultar.EntireRow.Hidden = False
        rangoOcultar2.EntireRow.Hidden = True
    Else:
        rangoOcultar.EntireRow.Hidden = True
        rangoOcultar3.EntireRow.Hidden = False
    End If
    Call OptimizeCode_End
End Sub
```

## Descripción

| Sub | Función |
|-----|---------|
| `ocultar_Concepcion` | Oculta/muestra filas según el estado de un botón de opción |
| `ocultar_Clausula` | Oculta filas basado en valores "Oui" o "Non" en celdas específicas |
| `ocultar_CES` | Oculta filas según la cantidad de lotes especificada |
| `ocultar_Decoupage` | Oculta filas basado en tranches (secciones) |

---

*Módulo: 31_CodigosGenerales | Archivo: 14_OcultarFilas.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
