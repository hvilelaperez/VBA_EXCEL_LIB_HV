# Enviar Hoja Activa por Outlook

Macro VBA para copiar la hoja activa de Excel a un nuevo libro temporal, guardarlo en formato compatible y enviarlo por correo electrónico mediante Outlook.

> **Autor:** Henry Vilela - hvilelaperez@gmail.com

## Código VBA

```vba
Sub Mail_ActiveSheet()
    'Guía: http://www.hocexcel.online/gui-email-tu-excel-bang-vba
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set Sourcewb = ActiveWorkbook

    ActiveSheet.Copy
    Set Destwb = ActiveWorkbook

    With Destwb
        If Val(Application.Version) < 12 Then
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
            Select Case Sourcewb.FileFormat
            Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If .HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56: FileExtStr = ".xls": FileFormatNum = 56
            Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
            End Select
        End If
    End With

    TempFilePath = Environ$("temp") & "\"
    TempFileName = "Parte de " & Sourcewb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .to = "hocexcel@hocexcel.online"
            .CC = ""
            .BCC = ""
            .Subject = "Cambiar el asunto del correo aquí"
            .Body = "Cambiar el contenido del correo aquí"
            .Attachments.Add Destwb.FullName
            .Send
        End With
        On Error GoTo 0
        .Close savechanges:=False
    End With

    'Eliminar el archivo temporal enviado
    Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub
```

## Descripción

| Paso | Descripción |
|---|---|
| 1 | Desactiva actualización de pantalla y eventos para evitar parpadeos |
| 2 | Copia la hoja activa a un nuevo libro temporal |
| 3 | Detecta el formato del libro origen (.xlsx, .xlsm, .xls, .xlsb) |
| 4 | Guarda el libro temporal en la carpeta de archivos temporales del sistema |
| 5 | Crea un correo en Outlook, adjunta el archivo y lo envía |
| 6 | Elimina el archivo temporal con `Kill` |
| 7 | Restaura la configuración de la aplicación |

### Puntos Clave

- Detecta automáticamente la versión de Excel y el formato del libro
- Utiliza la variable de entorno `temp` del sistema para archivos temporales
- El archivo temporal se elimina después del envío para no dejar residuos
- Incluye formato de fecha y hora en el nombre del archivo para evitar conflictos

---

*Módulo: 17 - VBA Outlook | Archivo: 01 - Gui sheet qua outlook*
