# Aceptar Reuniones de Outlook Automáticamente

Macros VBA para aceptar automáticamente solicitudes de reunión en Outlook, tanto de forma reactiva (al recibir) como de forma programada (buscando en la bandeja de entrada).

> **Autor:** Henry Vilela - hvilelaperez@gmail.com

## Código VBA

```vba
Public WithEvents GItems As Outlook.Items
'Actualizado por ExtendOffice 20180814

Private Sub Application_Startup()
    Set GItems = Outlook.Application.Session.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub GItems_ItemAdd(ByVal Item As Object)
    Dim xMtRequest As MeetingItem
    Dim xAppointmentItem As AppointmentItem
    Dim xMtResponse As MeetingItem
    
    If Item.Class = olMeetingRequest Then
        Set xMtRequest = Item
        Set xAppointmentItem = xMtRequest.GetAssociatedAppointment(True)
        If xAppointmentItem.GetOrganizer.Name = "Nombre del Remitente" Then
            With xAppointmentItem
                .ReminderMinutesBeforeStart = 45
                .Categories = "Categoría Naranja"
                .Save
            End With
            Set xMtResponse = xAppointmentItem.Respond(olMeetingAccepted)
            xMtResponse.Send
            xMtRequest.Delete
        End If
    End If
End Sub


Sub AcceptMeeting()
    Dim myNameSpace As Outlook.NameSpace
    Dim myFolder As Outlook.Folder
    Dim myMtgReq As Outlook.MeetingItem
    Dim myAppt As Outlook.AppointmentItem
    Dim myMtg As Outlook.MeetingItem
    
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set myMtgReq = myFolder.Items.Find("[MessageClass] = 'IPM.Schedule.Meeting.Request'")
    
    If TypeName(myMtgReq) <> "Nothing" Then
        Set myAppt = myMtgReq.GetAssociatedAppointment(True)
        Set myMtg = myAppt.Respond(olResponseAccepted, True)
        myMtg.Send
    End If
End Sub
```

## Descripción

| Sub/Función | Descripción |
|---|---|
| `Application_Startup` | Se ejecuta al iniciar Outlook, conecta el evento de bandeja de entrada |
| `GItems_ItemAdd` | Evento reactivo: se dispara al recibir un nuevo elemento en la bandeja |
| `AcceptMeeting` | Macro programada: busca y acepta la primera solicitud de reunión pendiente |

### Flujo del Evento Reactivo (`GItems_ItemAdd`)

1. Verifica si el elemento recibido es una solicitud de reunión (`olMeetingRequest`)
2. Obtiene la cita asociada a la solicitud
3. Compara el nombre del organizador con el nombre configurado
4. Configura recordatorio (45 min) y categoría
5. Acepta la reunión y envía la respuesta
6. Elimina la solicitud original

### Flujo de la Macro Programada (`AcceptMeeting`)

1. Accede a la bandeja de entrada de Outlook mediante MAPI
2. Busca la primera solicitud de reunión pendiente
3. Si la encuentra, obtiene la cita asociada y acepta
4. Envía la respuesta de aceptación

### Puntos Clave

- Requiere que el código se ejecute desde **ThisOutlookSession** (evento reactivo)
- Cambiar `"Nombre del Remitente"` por el nombre real del organizador a filtrar
- `olMeetingAccepted` acepta la reunión; usar `olMeetingTentative` para aceptación tentativa

---

*Módulo: 17 - VBA Outlook | Archivo: 02 - Auto accept meeting*
