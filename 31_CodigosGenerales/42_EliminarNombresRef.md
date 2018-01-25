# Eliminar Nombres de Rango Referencia (#REF)

Macros para eliminar nombres de rangos con errores de referencia (#REF) en libros de Excel.

## Código VBA

```vba
Sub ELIMINAR_NOMBRE_RANGO()
On Error Resume Next
    Dim nombreRng As Name
    For Each nombreRng In ThisWorkbook.Sheets("FAC.11x").Names
        nombreRng.Delete
    Next nombreRng
    
'    For Each nombre In ActiveWorkbook.Names
'        nombre.Delete
'    Next nombre

' Elimina todos los nombres en el libro activo
 
' Con un error #REF - confirma antes de ejecutar
'
'Dim N As Name
'
'If MsgBox("¿Está seguro?", vbYesNo + vbDefaultButton2, "Confirmar macro") = vbNo Then Exit Sub
'
'For Each N In ActiveWorkbook.Names
'
'If InStr(N.Value, "#REF") Then N.Delete
'
'Next N
End Sub


Sub EliminarNombresErroneos()
Dim nm As Excel.Name
Dim vPrueba As Variant
Dim i As Long
Dim ListaHojas As Worksheet
Set ListaHojas = Sheets("ELIMINADOS")
For Each nm In ActiveWorkbook.Names
    vPrueba = Empty
    On Error Resume Next
    vPrueba = Application.Evaluate(nm.RefersTo)
    On Error GoTo 0
    If TypeName(vPrueba) = "Error" Then
        i = i + 1
        ListaHojas.Cells(i, 1).Value = nm.Name
        If IsError(Application.Match(nm.Name, Sheets("NO ELIMINAR").Columns("A"), 0)) Then nm.Delete
    End If
Next nm
End Sub
```

## Descripción

### Primer procedimiento: ELIMINAR_NOMBRE_RANGO

| Elemento | Descripción |
|----------|-------------|
| `On Error Resume Next` | Ignora errores y continúa ejecutando |
| `ThisWorkbook.Sheets("FAC.11x").Names` | Accede a los nombres definidos en la hoja específica |
| `nombreRng.Delete` | Elimina cada nombre de rango |

### Segundo procedimiento: EliminarNombresErroneos

| Elemento | Descripción |
|----------|-------------|
| `nm.RefersTo` | Obtiene la referencia del nombre |
| `Application.Evaluate` | Evalúa la referencia para verificar si es válida |
| `TypeName(vPrueba) = "Error"` | Detecta si la referencia es un error (#REF) |
| `ListaHojas.Cells(i, 1)` | Registra el nombre en la hoja "ELIMINADOS" |
| `Application.Match` | Verifica si el nombre está en la lista "NO ELIMINAR" |

### Flujo de ejecución (EliminarNombresErroneos)

1. Recorre todos los nombres definidos en el libro activo
2. Intenta evaluar la referencia de cada nombre
3. Si la evaluación produce un error, el nombre tiene una referencia inválida
4. Registra el nombre en la hoja "ELIMINADOS"
5. Verifica que el nombre no esté en la lista protegida "NO ELIMINAR"
6. Si no está protegido, lo elimina

### Hojas requeridas

| Hoja | Propósito |
|------|-----------|
| "ELIMINADOS" | Registro de nombres eliminados |
| "NO ELIMINAR" | Lista de nombres que no deben eliminarse (columna A) |

---

*Módulo: 96 - VBA Codes general | Archivo: 35 -Delete all named range #ref.txt*
