# Manejo Básico de Errores

Ejemplos de manejo básico de errores con validación de entrada de datos y uso de `On Error GoTo`.

## Código VBA

```vba
Sub CAN_BAC_HAI()

Dim number As Variant
TINHCAN:
number = InputBox("Ingrese un número para calcular la raíz cuadrada", "Raíz cuadrada")

If number = "" Then Exit Sub
If number < 0 Then
    MsgBox "Debe ingresar un número positivo > 0", vbOKOnly
    GoTo TINHCAN
End If
If Not IsNumeric(number) Then
    MsgBox "Los datos ingresados deben ser numéricos", vbOKOnly
    GoTo TINHCAN
End If

MsgBox Round(Sqr(number), 3)

End Sub



Sub CAN_BAC_HAI2()

    Dim number As Variant, intentar As Variant
    
TINHCAN:
    
    number = InputBox("Ingrese datos para calcular la raíz")
    
    On Error GoTo ERROR
    If number = "" Then Exit Sub
    MsgBox "El resultado es: " & Round(Sqr(number), 3), vbOKOnly
    Exit Sub
ERROR:
    MsgBox "Ocurrió un error con los datos ingresados", vbOKOnly
    intentar = MsgBox("¿Desea continuar?", vbYesNo)
    If intentar = vbYes Then Resume TINHCAN
    
End Sub
```

## Descripción

| Procedimiento | Descripción |
|---------------|-------------|
| `CAN_BAC_HAI` | Calcula raíz cuadrada con validación manual de errores |
| `CAN_BAC_HAI2` | Calcula raíz cuadrada con manejo de errores usando `On Error GoTo` |
| `TINHCAN` | Etiqueta de retorno para reintentar la operación |
| `ERROR` | Etiqueta de manejo de errores que muestra mensaje y ofrece continuar |
| `Resume TINHCAN` | Retoma la ejecución desde la etiqueta especificada |
| Validaciones | Verifica que el valor sea numérico y positivo |

---

*Módulo: 07 - VBA Error Handler | Archivo: 00 - Bay loi.txt*
