# Remover Password de Proyecto VBA

Técnica avanzada que utiliza Hook de API para desproteger proyectos VBA bloqueados con contraseña. Intercepts la función `DialogBoxParamA` de user32.dll para bypass del diálogo de contraseña.

## Código VBA

```vba
Option Explicit

Private Const PAGE_EXECUTE_READWRITE = &H40

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Long, Source As Long, ByVal Length As Long)

Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Long, _
        ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
        ByVal lpProcName As String) As Long

Private Declare Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, _
        ByVal pTemplateName As Long, ByVal hWndParent As Long, _
        ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer

Dim HookBytes(0 To 5) As Byte
Dim OriginBytes(0 To 5) As Byte
Dim pFunc As Long
Dim Flag As Boolean

Private Function GetPtr(ByVal Value As Long) As Long
    GetPtr = Value
End Function

Public Sub RecoverBytes()
    If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 6
End Sub

Public Function Hook() As Boolean
    Dim TmpBytes(0 To 5) As Byte
    Dim p As Long
    Dim OriginProtect As Long

    Hook = False

    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")


    If VirtualProtect(ByVal pFunc, 6, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then

        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, 6
        If TmpBytes(0) <> &H68 Then

            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 6

            p = GetPtr(AddressOf MyDialogBoxParam)

            HookBytes(0) = &H68
            MoveMemory ByVal VarPtr(HookBytes(1)), ByVal VarPtr(p), 4
            HookBytes(5) = &HC3

            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 6
            Flag = True
            Hook = True
        End If
    End If
End Function

Private Function MyDialogBoxParam(ByVal hInstance As Long, _
        ByVal pTemplateName As Long, ByVal hWndParent As Long, _
        ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
    If pTemplateName = 4070 Then
        MyDialogBoxParam = 1
    Else
        RecoverBytes
        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
                           hWndParent, lpDialogFunc, dwInitParam)
        Hook
    End If
End Function


Sub unprotected()
    If Hook Then
        MsgBox "¡El proyecto VBA está desprotegido!", vbInformation, "*****"
    End If
End Sub
```

## Descripción

| Elemento | Descripción |
|---|---|
| `MoveMemory` | API para copiar bloques de memoria entre direcciones |
| `VirtualProtect` | API para cambiar los permisos de protección de memoria de una región |
| `GetProcAddress` | API para obtener la dirección de una función en un módulo DLL |
| `DialogBoxParamA` | API de user32 para crear diálogos modales; es la función que se intercepta |
| `HookBytes` | Array que almacena los bytes del hook ( PUSH + dirección + RET ) |
| `OriginBytes` | Bytes originales de la función para restaurar el hook |
| `MyDialogBoxParam` | Función reemplazo que intercepta el diálogo de contraseña (ID 4070) |
| `Hook` | Instala el hook modificando los primeros 6 bytes de `DialogBoxParamA` |
| `RecoverBytes` | Restaura los bytes originales de la función |
| `unprotected` | Sub principal que activa el hook y muestra confirmación |

**Advertencia:** Este código modifica memoria de procesos en tiempo de ejecución. Úselo solo para fines educativos o sobre archivos propios.

---

*Módulo: 94 - VBA Crack - Remove password | Archivo: 00 - Remove passWord.txt*
