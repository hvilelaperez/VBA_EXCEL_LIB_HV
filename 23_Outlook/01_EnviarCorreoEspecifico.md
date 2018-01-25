# Enviar Correo Electrónico con Cuenta Específica en Outlook

Macros VBA para enviar correos electrónicos desde Excel utilizando Outlook, incluyendo la capacidad de enviar con una cuenta específica y obtener información de cuentas configuradas.

> **Autor:** Henry Vilela - hvilelaperez@gmail.com

## Código VBA

```vba
Option Explicit

Sub Mail_workbook_Outlook_1()

    Dim outApp As Object
    Dim outMail As Object

    Set outApp = CreateObject("Outlook.Application")
    Set outMail = outApp.CreateItem(0)

    On Error Resume Next
    With outMail
        .To = "tienduy.nguyen2@estp.fr"
        .CC = ""
        .BCC = ""
        .Subject = "Este es el asunto del correo"
        .Body = "Hola"
        .Attachments.Add ActiveWorkbook.FullName
        .Send
    End With
    On Error GoTo 0

    Set outMail = Nothing
    Set outApp = Nothing
End Sub

Sub Which_Account_Number()
    'No olvides establecer una referencia a Outlook en el editor VBA
    Dim outApp As Object
    Set outApp = CreateObject("Outlook.Application")
    Dim i As Long

    For i = 1 To outApp.session.Accounts.Count
        MsgBox outApp.session.Accounts.Item(i) & " : Esta es la cuenta número " & i
    Next i
    End Sub

Sub MAIL_WITH_SPECIFIC_ACCOUNT()

    Dim outApp As Object
    Dim outMail As Object
    Dim outAccount As Object
    Dim strbody As String, StrFrom As String, i As Integer, itemAccount As Integer, signature As String
    Dim foundAccount As Boolean
    
    Set outApp = CreateObject("Outlook.Application")
    Set outMail = outApp.CreateItem(0)

    StrFrom = "tienduy.nguyen2@estp.fr"
    strbody = "Línea 1 con 2 saltos de línea detrás pero en realidad hace una interlínea, en todo caso, sin problema <br>" _
          & "<br>" _
          & "<br>" _
          & "Línea 2 : OK" _
          & "Línea 3 : OK<br>" _
          & "Línea 4 : OK<br>" _
          & "Línea 5 : OK<br>" _
          & "Línea 6 con 2 saltos de línea detrás pero en realidad hace una interlínea, en todo caso, sin problema <br>" _
          & "<br>" _
          & "<br>" _
          & "Línea 7 : OK" & "<br>" _
          & "Línea 8 con UN SOLO salto de línea detrás pero en realidad hace una interlínea. Creo que se debe a que hay muchas cosas en esta línea pero ¿cómo solucionarlo? <br>" _
          & "<br>" _
          & "Línea 9 : OK" & "<br>" _
          & "Línea 10 : OK" & "<br>" _
          & "<br>"

    strbody = "estimado" & " " & "nombre" & ",<BR><BR>" & _
              "como indicado por teléfono, le confirmo su reserva en el restaurante " & _
              "<B>" & "restaurante" & "</B>"
               
    foundAccount = False
    'Dar el índice del elemento de correo específico
    For i = 1 To outApp.session.Accounts.Count
        If outApp.session.Accounts.Item(i).smtpAddress = StrFrom Then
            itemAccount = i
            foundAccount = True
            Exit For
        End If
    Next i
    
    If foundAccount Then
        On Error Resume Next
        With outMail
            .Display
            .To = "tien-duy.nguyen@vinci-construction.fr"
            .CC = ""
            .BCC = ""
            .Subject = "Este es el asunto del correo"
            .HTMLBody = strbody & "<br>" & .HTMLBody

    '        Set .SendOnBehaftOfName = "tienduy9x@outlook.com" --> no funciona
    '       'Establecer el remitente con el correo específico
            Set .SendUsingAccount = outApp.session.Accounts.Item(itemAccount)
    '        .Send   'o usar .Display
            
    '        'Eliminar después de enviar
    '        .DeleteAfterSubmit = True 'Establecer en True si no se guarda una copia del mensaje al enviar
            
        End With
        On Error GoTo 0
    Else
        MsgBox "No se encontró una cuenta con este correo" & vbNewLine & "Debe verificar el correo del remitente"
    End If
    
    Set outMail = Nothing
    Set outApp = Nothing
    
End Sub
```

## Descripción

| Sub/Función | Descripción |
|---|---|
| `Mail_workbook_Outlook_1` | Envía el libro de trabajo activo como adjunto mediante Outlook |
| `Which_Account_Number` | Muestra todas las cuentas de Outlook configuradas con su número de índice |
| `MAIL_WITH_SPECIFIC_ACCOUNT` | Envía un correo utilizando una cuenta específica de Outlook, con soporte para formato HTML |

### Puntos Clave

- Utiliza `CreateObject("Outlook.Application")` para late binding (sin referencia directa)
- `MAIL_WITH_SPECIFIC_ACCOUNT` busca la cuenta por dirección SMTP y la configura con `SendUsingAccount`
- Soporta formato HTML en el cuerpo del mensaje mediante `<br>` y etiquetas `<B>`
- Los archivos adjuntos se agregan con `.Attachments.Add`

---

*Módulo: 17 - VBA Outlook | Archivo: 00 - VBA Outlook with specific mail*
