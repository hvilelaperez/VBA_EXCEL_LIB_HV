# Verificar si un Variant es Array

Código VBA para verificar si una variable Variant contiene un array o una instancia de objeto.

## Código VBA

```vba
' Archivo fuente vacío - contenido no disponible
' Utilizar la función IsArray() de VBA para verificar:
'
'   If IsArray(miVariable) Then
'       ' La variable contiene un array
'   Else
'       ' La variable contiene un valor simple o instancia
'   End If
```

## Descripción

### Funciones de verificación en VBA

| Función | Retorna True cuando |
|---------|---------------------|
| `IsArray()` | La variable contiene un array |
| `IsObject()` | La variable contiene un objeto |
| `IsEmpty()` | La variable no ha sido inicializada |
| `IsNull()` | La variable contiene Null |
| `IsNumeric()` | La variable contiene un valor numérico |

### Ejemplo de uso

```vba
Sub VerificarTipoVariable()
    Dim miVar As Variant
    
    miVar = Array(1, 2, 3)
    If IsArray(miVar) Then
        Debug.Print "miVar es un array con " & UBound(miVar) & " elementos"
    End If
    
    miVar = "Hola"
    If Not IsArray(miVar) Then
        Debug.Print "miVar es un valor simple: " & miVar
    End If
End Sub
```

### Nota

> Archivo fuente original vacío. Se incluye referencia a la función `IsArray()` como guía para verificar el tipo de variable.

---

*Módulo: 05 - VBA Array | Archivo: 03 - Check a variant is array or is instance*