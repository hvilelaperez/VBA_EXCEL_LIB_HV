# Evento KeyDown para Formato Automático de Fecha

Macro VBA que usa el evento `KeyDown` en un cuadro de texto para insertar automáticamente barras diagonales (`/`) al escribir una fecha, forzando el formato `dd/mm/yyyy` y manejando la tecla de retroceso.

> **Autor:** Henry Vilela - hvilelaperez@gmail.com

## Código VBA

```vba
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyBack Then
        If Len(Me.TextBox1) = 4 Then
            Me.TextBox1 = Left(Me.TextBox1, 2)
            KeyCode = False
        ElseIf Len(Me.TextBox1) = 7 Then
            Me.TextBox1 = Left(Me.TextBox1, 5)
            KeyCode = False
        End If
    Else
        If Len(Me.TextBox1) = 2 Or Len(Me.TextBox1) = 5 Then
            'Agregar barra diagonal
            Me.TextBox1 = Me.TextBox1 & "/"
        End If
    End If
End Sub
```

## Descripción

| Condición | Acción |
|---|---|
| Se presiona **Retroceso** con 4 caracteres | Elimina la barra: `dd/mm` → `dd` |
| Se presiona **Retroceso** con 7 caracteres | Elimina la barra: `dd/mm/yy` → `dd/mm` |
| Se presiona **cualquier tecla** con 2 caracteres | Agrega barra: `dd` → `dd/` |
| Se presiona **cualquier tecla** con 5 caracteres | Agrega barra: `dd/mm` → `dd/mm/` |

### Flujo de Ejecución

```
Entrada del usuario → Evento KeyDown → Verificar longitud actual

  Longitud = 2 y tecla ≠ Retroceso → Agregar "/"
  Longitud = 5 y tecla ≠ Retroceso → Agregar "/"
  Longitud = 4 y tecla = Retroceso → Eliminar "/"
  Longitud = 7 y tecla = Retroceso → Eliminar "/"
```

### Ejemplo de Uso

| Paso | Entrada | Tecla | Resultado |
|---|---|---|---|
| 1 | `""` | `1` | `1` |
| 2 | `1` | `5` | `15` |
| 3 | `15` | `cualquier` | `15/` |
| 4 | `15/` | `0` | `15/0` |
| 5 | `15/0` | `6` | `15/06` |
| 6 | `15/06` | `cualquier` | `15/06/` |
| 7 | `15/06/` | `2` | `15/06/2` |
| 8 | `15/06/2` | `0` | `15/06/20` |

### Puntos Clave

- `vbKeyBack` (valor 8) corresponde a la tecla de retroceso
- `KeyCode = False` cancela la pulsación de la tecla
- El control automático de barras mejora la experiencia del usuario
- No valida caracteres numéricos, solo controla la inserción de separadores

---

*Módulo: 20 - VBA Worksheets Events | Archivo: 02 - Event Keydown de k phai nhap slash voi ngay thang*
