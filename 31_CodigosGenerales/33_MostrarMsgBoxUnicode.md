# Mostrar MsgBox con Unicode

Función que muestra cuadros de mensaje (MsgBox) con caracteres Unicode (vietnamitas, etc.) usando la API de Windows.

## Código VBA

```vba
' Autor: Henry Vilela - hvilelaperez@gmail.com
' Autor original: Nguyen Duy Tuan
' Fuente: www.bluesofts.net

' Declaración de funciones API de User32.DLL
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, _
    ByVal lpText As String, ByVal lpCaption As String, _
    ByVal wType As Long) As Long

Function MsgBoxUnicode(ByVal PromptUnicode As Variant, _
    Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
    Optional ByVal TitleUnicode As Variant = vbNullString) As VbMsgBoxResult

    ' BStrMsg, BStrTitle son cadenas Unicode
    Dim BStrMsg, BStrTitle

    ' StrConv convierte la cadena a formato Unicode
    BStrMsg = StrConv(PromptUnicode, vbUnicode)
    BStrTitle = StrConv(TitleUnicode, vbUnicode)

    ' Mostrar cuadro de mensaje con soporte Unicode
    MsgBoxUnicode = MessageBoxW(GetActiveWindow, BStrMsg, BStrTitle, Buttons)
End Function

'==================================================================
' Funciones de conversión de codificación TCVN3 y VNI a Unicode
' Autor: Binh - OverAC (www.giaiphapexcel.com)

Function TCVN3aUnicode(vnstr As String)
    Dim c As String, i As Integer
    For i = 1 To Len(vnstr)
        c = Mid(vnstr, i, 1)
        Select Case c
        Case "a": c = ChrW$(97)
        Case ChrW$(225): c = ChrW$(225)
        Case ChrW$(224): c = ChrW$(224)
        Case ChrW$(7843): c = ChrW$(7843)
        Case ChrW$(227): c = ChrW$(227)
        Case ChrW$(7841): c = ChrW$(7841)
        Case ChrW$(259): c = ChrW$(259)
        Case ChrW$(7855): c = ChrW$(7855)
        Case ChrW$(7857): c = ChrW$(7857)
        Case ChrW$(7859): c = ChrW$(7859)
        Case ChrW$(7861): c = ChrW$(7861)
        Case ChrW$(7863): c = ChrW$(7863)
        Case ChrW$(226): c = ChrW$(226)
        Case ChrW$(7845): c = ChrW$(7845)
        Case ChrW$(7847): c = ChrW$(7847)
        Case ChrW$(7849): c = ChrW$(7849)
        Case ChrW$(7851): c = ChrW$(7851)
        Case ChrW$(7853): c = ChrW$(7853)
        Case "e": c = ChrW$(101)
        Case ChrW$(233): c = ChrW$(233)
        Case ChrW$(232): c = ChrW$(232)
        Case ChrW$(7867): c = ChrW$(7867)
        Case ChrW$(7869): c = ChrW$(7869)
        Case ChrW$(7865): c = ChrW$(7865)
        Case ChrW$(234): c = ChrW$(234)
        Case ChrW$(7871): c = ChrW$(7871)
        Case ChrW$(7873): c = ChrW$(7873)
        Case ChrW$(7875): c = ChrW$(7875)
        Case ChrW$(7877): c = ChrW$(7877)
        Case ChrW$(7879): c = ChrW$(7879)
        Case "o": c = ChrW$(111)
        Case ChrW$(243): c = ChrW$(243)
        Case ChrW$(242): c = ChrW$(242)
        Case ChrW$(7887): c = ChrW$(7887)
        Case ChrW$(245): c = ChrW$(245)
        Case ChrW$(7885): c = ChrW$(7885)
        Case ChrW$(244): c = ChrW$(244)
        Case ChrW$(7889): c = ChrW$(7889)
        Case ChrW$(7891): c = ChrW$(7891)
        Case ChrW$(7893): c = ChrW$(7893)
        Case ChrW$(7895): c = ChrW$(7895)
        Case ChrW$(7897): c = ChrW$(7897)
        Case ChrW$(417): c = ChrW$(417)
        Case ChrW$(7899): c = ChrW$(7899)
        Case ChrW$(7901): c = ChrW$(7901)
        Case ChrW$(7903): c = ChrW$(7903)
        Case ChrW$(7905): c = ChrW$(7905)
        Case ChrW$(7907): c = ChrW$(7907)
        Case "i": c = ChrW$(105)
        Case ChrW$(237): c = ChrW$(237)
        Case ChrW$(236): c = ChrW$(236)
        Case ChrW$(7881): c = ChrW$(7881)
        Case ChrW$(297): c = ChrW$(297)
        Case ChrW$(7883): c = ChrW$(7883)
        Case "u": c = ChrW$(117)
        Case ChrW$(250): c = ChrW$(250)
        Case ChrW$(249): c = ChrW$(249)
        Case ChrW$(7911): c = ChrW$(7911)
        Case ChrW$(361): c = ChrW$(361)
        Case ChrW$(7909): c = ChrW$(7909)
        Case ChrW$(432): c = ChrW$(432)
        Case ChrW$(7913): c = ChrW$(7913)
        Case ChrW$(7915): c = ChrW$(7915)
        Case ChrW$(7917): c = ChrW$(7917)
        Case ChrW$(7919): c = ChrW$(7919)
        Case ChrW$(7921): c = ChrW$(7921)
        Case "y": c = ChrW$(121)
        Case ChrW$(253): c = ChrW$(253)
        Case ChrW$(7923): c = ChrW$(7923)
        Case ChrW$(7927): c = ChrW$(7927)
        Case ChrW$(7929): c = ChrW$(7929)
        Case ChrW$(7925): c = ChrW$(7925)
        Case ChrW$(273): c = ChrW$(273)
        Case "A": c = ChrW$(65)
        Case ChrW$(258): c = ChrW$(258)
        Case ChrW$(194): c = ChrW$(194)
        Case "E": c = ChrW$(69)
        Case ChrW$(202): c = ChrW$(202)
        Case "O": c = ChrW$(79)
        Case ChrW$(212): c = ChrW$(212)
        Case ChrW$(416): c = ChrW$(416)
        Case "I": c = ChrW$(73)
        Case "U": c = ChrW$(85)
        Case ChrW$(431): c = ChrW$(431)
        Case "Y": c = ChrW$(89)
        Case ChrW$(272): c = ChrW$(272)
        End Select
        TCVN3aUnicode = TCVN3aUnicode + c
    Next i
End Function

Function VNIaUnicode(vnstr As String)
    Dim c As String, i As Integer
    Dim db As Boolean
    For i = 1 To Len(vnstr)
        db = False
        If i < Len(vnstr) Then
            c = Mid(vnstr, i + 1, 1)
            ' Verificar si el siguiente carácter es un signo diacrítico VNI
            If c = ChrW$(777) Or c = ChrW$(769) Or c = ChrW$(768) Or _
               c = ChrW$(777) Or c = ChrW$(771) Or c = ChrW$(803) Or _
               c = ChrW$(7843) Or c = ChrW$(7855) Or c = ChrW$(7857) Or _
               c = ChrW$(7859) Or c = ChrW$(7861) Or c = ChrW$(7863) Or _
               c = ChrW$(7867) Or c = ChrW$(7869) Or c = ChrW$(7865) Or _
               c = ChrW$(7887) Or c = ChrW$(7889) Or c = ChrW$(7891) Or _
               c = ChrW$(7893) Or c = ChrW$(7895) Or c = ChrW$(7897) Or _
               c = ChrW$(7899) Or c = ChrW$(7901) Or c = ChrW$(7903) Or _
               c = ChrW$(7905) Or c = ChrW$(7907) Or c = ChrW$(7911) Or _
               c = ChrW$(7913) Or c = ChrW$(7915) Or c = ChrW$(7917) Or _
               c = ChrW$(7919) Or c = ChrW$(7921) Or c = ChrW$(7923) Or _
               c = ChrW$(7927) Or c = ChrW$(7929) Then db = True
        End If
        If db Then
            c = Mid(vnstr, i, 2)
            Select Case c
            Case "a" & ChrW$(777): c = ChrW$(225)
            Case "a" & ChrW$(769): c = ChrW$(224)
            Case "a" & ChrW$(768): c = ChrW$(7843)
            Case "a" & ChrW$(771): c = ChrW$(227)
            Case "a" & ChrW$(803): c = ChrW$(7841)
            Case "a" & ChrW$(7855): c = ChrW$(259)
            Case "a" & ChrW$(7857): c = ChrW$(7855)
            Case "a" & ChrW$(7859): c = ChrW$(7857)
            Case "a" & ChrW$(7861): c = ChrW$(7859)
            Case "a" & ChrW$(7863): c = ChrW$(7861)
            Case "a" & ChrW$(7845): c = ChrW$(7863)
            Case "e" & ChrW$(777): c = ChrW$(233)
            Case "e" & ChrW$(769): c = ChrW$(232)
            Case "e" & ChrW$(768): c = ChrW$(7867)
            Case "e" & ChrW$(771): c = ChrW$(7869)
            Case "e" & ChrW$(803): c = ChrW$(7865)
            Case "e" & ChrW$(7871): c = ChrW$(234)
            Case "o" & ChrW$(777): c = ChrW$(243)
            Case "o" & ChrW$(769): c = ChrW$(242)
            Case "o" & ChrW$(768): c = ChrW$(7887)
            Case "o" & ChrW$(771): c = ChrW$(245)
            Case "o" & ChrW$(803): c = ChrW$(7885)
            Case "o" & ChrW$(7889): c = ChrW$(244)
            Case "u" & ChrW$(777): c = ChrW$(250)
            Case "u" & ChrW$(769): c = ChrW$(249)
            Case "u" & ChrW$(768): c = ChrW$(7911)
            Case "u" & ChrW$(771): c = ChrW$(361)
            Case "u" & ChrW$(803): c = ChrW$(7909)
            Case "y" & ChrW$(777): c = ChrW$(253)
            Case "y" & ChrW$(769): c = ChrW$(7923)
            Case "y" & ChrW$(768): c = ChrW$(7927)
            Case "y" & ChrW$(803): c = ChrW$(7929)
            Case "A" & ChrW$(777): c = ChrW$(193)
            Case "A" & ChrW$(769): c = ChrW$(192)
            Case "A" & ChrW$(768): c = ChrW$(7842)
            Case "A" & ChrW$(771): c = ChrW$(195)
            Case "A" & ChrW$(803): c = ChrW$(7840)
            Case "E" & ChrW$(777): c = ChrW$(201)
            Case "E" & ChrW$(769): c = ChrW$(200)
            Case "E" & ChrW$(768): c = ChrW$(7866)
            Case "E" & ChrW$(771): c = ChrW$(7868)
            Case "E" & ChrW$(803): c = ChrW$(7864)
            Case "E" & ChrW$(7870): c = ChrW$(202)
            Case "O" & ChrW$(777): c = ChrW$(211)
            Case "O" & ChrW$(769): c = ChrW$(210)
            Case "O" & ChrW$(768): c = ChrW$(7886)
            Case "O" & ChrW$(771): c = ChrW$(213)
            Case "O" & ChrW$(803): c = ChrW$(7884)
            Case "O" & ChrW$(7888): c = ChrW$(212)
            Case Else
            End Select
        Else
            c = Mid(vnstr, i, 1)
            Select Case c
            Case ChrW$(417): c = ChrW$(417)
            Case ChrW$(237): c = ChrW$(237)
            Case ChrW$(236): c = ChrW$(236)
            Case ChrW$(7881): c = ChrW$(7881)
            Case ChrW$(297): c = ChrW$(297)
            Case ChrW$(7883): c = ChrW$(7883)
            Case ChrW$(432): c = ChrW$(432)
            Case ChrW$(7925): c = ChrW$(7925)
            Case ChrW$(273): c = ChrW$(273)
            Case ChrW$(416): c = ChrW$(416)
            Case Else
            End Select
        End If
        VNIaUnicode = VNIaUnicode + c
        If db Then i = i + 1
    Next i
End Function

Function UNICODEaVNI(ByVal vnstr As String)
    Dim c As String, i As Integer
    For i = 1 To Len(vnstr)
        c = Mid(vnstr, i, 1)
        Select Case c
        Case ChrW$(97): c = "a"
        Case ChrW$(225): c = "a" & ChrW$(777)
        Case ChrW$(224): c = "a" & ChrW$(769)
        Case ChrW$(7843): c = "a" & ChrW$(768)
        Case ChrW$(227): c = "a" & ChrW$(771)
        Case ChrW$(7841): c = "a" & ChrW$(803)
        Case ChrW$(259): c = "a" & ChrW$(7855)
        Case ChrW$(7855): c = "a" & ChrW$(7857)
        Case ChrW$(7857): c = "a" & ChrW$(7859)
        Case ChrW$(7859): c = "a" & ChrW$(7861)
        Case ChrW$(7861): c = "a" & ChrW$(7863)
        Case ChrW$(7863): c = "a" & ChrW$(7845)
        Case ChrW$(226): c = "a" & ChrW$(7845)
        Case ChrW$(7845): c = "a" & ChrW$(7847)
        Case ChrW$(7847): c = "a" & ChrW$(7849)
        Case ChrW$(7849): c = "a" & ChrW$(7851)
        Case ChrW$(7851): c = "a" & ChrW$(7853)
        Case ChrW$(7853): c = "a" & ChrW$(7845)
        Case ChrW$(101): c = "e"
        Case ChrW$(233): c = "e" & ChrW$(777)
        Case ChrW$(232): c = "e" & ChrW$(769)
        Case ChrW$(7867): c = "e" & ChrW$(768)
        Case ChrW$(7869): c = "e" & ChrW$(771)
        Case ChrW$(7865): c = "e" & ChrW$(803)
        Case ChrW$(234): c = "e" & ChrW$(7871)
        Case ChrW$(7871): c = "e" & ChrW$(7873)
        Case ChrW$(7873): c = "e" & ChrW$(7875)
        Case ChrW$(7875): c = "e" & ChrW$(7877)
        Case ChrW$(7877): c = "e" & ChrW$(7879)
        Case ChrW$(7879): c = "e" & ChrW$(7871)
        Case ChrW$(111): c = "o"
        Case ChrW$(243): c = "o" & ChrW$(777)
        Case ChrW$(242): c = "o" & ChrW$(769)
        Case ChrW$(7887): c = "o" & ChrW$(768)
        Case ChrW$(245): c = "o" & ChrW$(771)
        Case ChrW$(7885): c = "o" & ChrW$(803)
        Case ChrW$(244): c = "o" & ChrW$(7889)
        Case ChrW$(7889): c = "o" & ChrW$(7891)
        Case ChrW$(7891): c = "o" & ChrW$(7893)
        Case ChrW$(7893): c = "o" & ChrW$(7895)
        Case ChrW$(7895): c = "o" & ChrW$(7897)
        Case ChrW$(7897): c = "o" & ChrW$(7889)
        Case ChrW$(417): c = ChrW$(417)
        Case ChrW$(7899): c = ChrW$(417) & ChrW$(7901)
        Case ChrW$(7901): c = ChrW$(417) & ChrW$(7903)
        Case ChrW$(7903): c = ChrW$(417) & ChrW$(7905)
        Case ChrW$(7905): c = ChrW$(417) & ChrW$(7907)
        Case ChrW$(7907): c = ChrW$(417) & ChrW$(7899)
        Case ChrW$(105): c = "i"
        Case ChrW$(237): c = "i" & ChrW$(777)
        Case ChrW$(236): c = "i" & ChrW$(769)
        Case ChrW$(7881): c = "i" & ChrW$(768)
        Case ChrW$(297): c = "i" & ChrW$(771)
        Case ChrW$(7883): c = "i" & ChrW$(803)
        Case ChrW$(117): c = "u"
        Case ChrW$(250): c = "u" & ChrW$(777)
        Case ChrW$(249): c = "u" & ChrW$(769)
        Case ChrW$(7911): c = "u" & ChrW$(768)
        Case ChrW$(361): c = "u" & ChrW$(771)
        Case ChrW$(7909): c = "u" & ChrW$(803)
        Case ChrW$(432): c = ChrW$(432)
        Case ChrW$(7913): c = ChrW$(432) & ChrW$(7915)
        Case ChrW$(7915): c = "u" & ChrW$(7917)
        Case ChrW$(7917): c = ChrW$(432) & ChrW$(7919)
        Case ChrW$(7919): c = ChrW$(432) & ChrW$(7921)
        Case ChrW$(7921): c = ChrW$(432) & ChrW$(7913)
        Case ChrW$(121): c = "y"
        Case ChrW$(253): c = "y" & ChrW$(777)
        Case ChrW$(7923): c = "y" & ChrW$(769)
        Case ChrW$(7927): c = "y" & ChrW$(768)
        Case ChrW$(7929): c = "y" & ChrW$(803)
        Case ChrW$(7925): c = "y" & ChrW$(7929)
        Case ChrW$(273): c = "d" & ChrW$(9)
        Case ChrW$(65): c = "A"
        Case ChrW$(193): c = "A" & ChrW$(777)
        Case ChrW$(192): c = "A" & ChrW$(769)
        Case ChrW$(7842): c = "A" & ChrW$(768)
        Case ChrW$(195): c = "A" & ChrW$(771)
        Case ChrW$(7840): c = "A" & ChrW$(803)
        Case ChrW$(258): c = "A" & ChrW$(7855)
        Case ChrW$(7854): c = "A" & ChrW$(7857)
        Case ChrW$(7856): c = "A" & ChrW$(7859)
        Case ChrW$(7858): c = "A" & ChrW$(7861)
        Case ChrW$(7860): c = "A" & ChrW$(7863)
        Case ChrW$(7862): c = "A" & ChrW$(7845)
        Case ChrW$(194): c = "A" & ChrW$(7845)
        Case ChrW$(7844): c = "A" & ChrW$(7847)
        Case ChrW$(7846): c = "A" & ChrW$(7849)
        Case ChrW$(7848): c = "A" & ChrW$(7851)
        Case ChrW$(7850): c = "A" & ChrW$(7853)
        Case ChrW$(7852): c = "A" & ChrW$(7845)
        Case ChrW$(69): c = "E"
        Case ChrW$(201): c = "E" & ChrW$(777)
        Case ChrW$(200): c = "E" & ChrW$(769)
        Case ChrW$(7866): c = "E" & ChrW$(768)
        Case ChrW$(7868): c = "E" & ChrW$(771)
        Case ChrW$(7864): c = "E" & ChrW$(803)
        Case ChrW$(202): c = "E" & ChrW$(7871)
        Case ChrW$(7870): c = "E" & ChrW$(7873)
        Case ChrW$(7872): c = "E" & ChrW$(7875)
        Case ChrW$(7874): c = "E" & ChrW$(7877)
        Case ChrW$(7876): c = "E" & ChrW$(7879)
        Case ChrW$(7878): c = "E" & ChrW$(7871)
        Case ChrW$(79): c = "O"
        Case ChrW$(211): c = "O" & ChrW$(777)
        Case ChrW$(210): c = "O" & ChrW$(769)
        Case ChrW$(7886): c = "O" & ChrW$(768)
        Case ChrW$(213): c = "O" & ChrW$(771)
        Case ChrW$(7884): c = "O" & ChrW$(803)
        Case ChrW$(212): c = "O" & ChrW$(7889)
        Case ChrW$(7888): c = "O" & ChrW$(7891)
        Case ChrW$(7890): c = "O" & ChrW$(7893)
        Case ChrW$(7892): c = "O" & ChrW$(7895)
        Case ChrW$(7894): c = "O" & ChrW$(7897)
        Case ChrW$(7896): c = "O" & ChrW$(7889)
        Case ChrW$(416): c = ChrW$(416)
        Case ChrW$(7898): c = ChrW$(416) & ChrW$(7900)
        Case ChrW$(7900): c = ChrW$(416) & ChrW$(7902)
        Case ChrW$(7902): c = ChrW$(416) & ChrW$(7904)
        Case ChrW$(7904): c = ChrW$(416) & ChrW$(7906)
        Case ChrW$(7906): c = ChrW$(416) & ChrW$(7898)
        Case ChrW$(73): c = "I"
        Case ChrW$(205): c = "I" & ChrW$(777)
        Case ChrW$(204): c = "I" & ChrW$(769)
        Case ChrW$(7880): c = "I" & ChrW$(768)
        Case ChrW$(296): c = "I" & ChrW$(771)
        Case ChrW$(7882): c = "I" & ChrW$(803)
        Case ChrW$(85): c = "U"
        Case ChrW$(218): c = "U" & ChrW$(777)
        Case ChrW$(217): c = "U" & ChrW$(769)
        Case ChrW$(7910): c = "U" & ChrW$(768)
        Case ChrW$(360): c = "U" & ChrW$(771)
        Case ChrW$(7908): c = "U" & ChrW$(803)
        Case ChrW$(431): c = ChrW$(431)
        Case ChrW$(7912): c = ChrW$(431) & ChrW$(7914)
        Case ChrW$(7914): c = ChrW$(431) & ChrW$(7916)
        Case ChrW$(7916): c = ChrW$(431) & ChrW$(7918)
        Case ChrW$(7918): c = ChrW$(431) & ChrW$(7920)
        Case ChrW$(7920): c = ChrW$(431) & ChrW$(7912)
        Case ChrW$(89): c = "Y"
        Case ChrW$(221): c = "Y" & ChrW$(777)
        Case ChrW$(7922): c = "Y" & ChrW$(769)
        Case ChrW$(7926): c = "Y" & ChrW$(768)
        Case ChrW$(7928): c = "Y" & ChrW$(803)
        Case ChrW$(7924): c = "Y" & ChrW$(7929)
        Case ChrW$(272): c = "D" & ChrW$(9)
        End Select
        UNICODEaVNI = UNICODEaVNI + c
    Next i
End Function

' Función abreviada de conversión TCVN3 a Unicode
Function TCVN3(strTCVN3 As String)
    TCVN3 = TCVN3aUnicode(strTCVN3)
End Function

' Función abreviada de conversión VNI a Unicode
Function VNI(strVNI As String)
    VNI = VNIaUnicode(strVNI)
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Función principal** | `MsgBoxUnicode` - Muestra MsgBox con caracteres Unicode usando API de Windows |
| **API Windows** | `MessageBoxW` - Versión Unicode de MessageBox, `GetActiveWindow` - Obtiene ventana activa |
| **Conversión TCVN3** | `TCVN3aUnicode` - Convierte codificación TCVN3 a Unicode |
| **Conversión VNI** | `VNIaUnicode` - Convierte codificación VNI a Unicode |
| **Conversión inversa** | `UNICODEaVNI` - Convierte Unicode a codificación VNI |
| **Funciones abreviadas** | `TCVN3()` y `VNI()` como atajos de conversión |
| **Uso** | `MsgBoxUnicode "Texto con caracteres especiales"` para mostrar mensajes Unicode |

---

*Módulo: 31_CodigosGenerales | Archivo: 44 - MsgBox hien thi tieng viet Unicode.txt*
