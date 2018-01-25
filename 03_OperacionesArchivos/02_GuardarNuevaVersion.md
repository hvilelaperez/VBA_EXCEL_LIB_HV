# Guardar Nueva Versión

Guarda el libro actual con versión incremental (_v1, _v2, etc.) si ya existe.

## Código VBA

```vba
Option Explicit

' ============================================================
' GUARDAR NUEVA VERSIÓN
' ============================================================
Public Sub GuardarNuevaVersion()
    Dim rutaCarpeta As String
    Dim nombreBase As String
    Dim extension As String
    Dim contador As Long
    Dim guardarComo As String
    
    On Error GoTo NoGuardado
    
    rutaCarpeta = ThisWorkbook.Path & "\"
    nombreBase = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    extension = Mid(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "."))
    On Error GoTo 0
    
    ' Verificar si ya tiene versión
    If InStr(nombreBase, "_v") > 0 Then
        Dim partes As Variant
        partes = Split(nombreBase, "_v")
        nombreBase = partes(0)
    End If
    
    contador = 1
    guardarComo = rutaCarpeta & nombreBase & "_v" & contador & extension
    
    Do While Dir(guardarComo) <> ""
        contador = contador + 1
        guardarComo = rutaCarpeta & nombreBase & "_v" & contador & extension
    Loop
    
    ThisWorkbook.SaveAs guardarComo
    MsgBox "Versión " & contador & " guardada.", vbInformation
    Exit Sub
    
NoGuardado:
    MsgBox "Error al guardar.", vbCritical
End Sub
```

## Ejemplo

| Nombre Original | Resultado |
|-----------------|-----------|
| Reporte.xlsx | Reporte_v1.xlsx |
| Reporte.xlsx (ya existe _v1) | Reporte_v2.xlsx |

---

*Módulo: 03_OperacionesArchivos | Archivo: 02_GuardarNuevaVersion.md*
