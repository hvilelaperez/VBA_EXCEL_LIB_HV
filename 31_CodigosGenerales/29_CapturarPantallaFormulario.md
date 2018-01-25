# Capturar Pantalla de Formulario

Macro que captura la pantalla actual (ventana activa) y la exporta como archivo PDF.

## Código VBA

```vba
' Autor: Henry Vilela - hvilelaperez@gmail.com
' Adaptado de Kenneth Hobs
' Fuente: http://www.ozgrid.com/forum/showthread.php?t=157677
Option Explicit

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
    ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Const VK_SNAPSHOT = 44
Const VK_LMENU = 164
Const KEYEVENTF_KEYUP = 2
Const KEYEVENTF_EXTENDEDKEY = 1

Private Sub CommandButton1_Click()
    ' Cambiar al nombre de tu botón
    Dim pdfName As String
    Dim newWS As Worksheet

    ' Simular pulsación de Alt+PrtScn para capturar ventana activa
    keybd_event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY + KEYEVENTF_KEYUP, 0
    keybd_event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY + KEYEVENTF_KEYUP, 0

    ' Permitir que el sistema procese las pulsaciones de teclas
    DoEvents

    ' Crear nueva hoja y pegar la imagen capturada
    Set newWS = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count))
    newWS.PasteSpecial Format:="Bitmap", Link:=False, DisplayAsIcon:=False

    ' Generar nombre del archivo PDF con fecha
    pdfName = ActiveWorkbook.Path & "\" & Me.Name & " " & Format(Now, "yyyy-mmm-dd") & ".pdf"

    ' Exportar como PDF
    newWS.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=pdfName, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' Limpiar: eliminar hoja temporal y cerrar formulario
    newWS.Delete
    Unload Me
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Función** | Captura la ventana activa y la exporta como PDF |
| **API Windows** | `keybd_event` - Simula pulsaciones de teclado para captura de pantalla |
| **Constantes** | `VK_SNAPSHOT` (PrtScn), `VK_LMENU` (Alt izquierdo) |
| **Proceso** | Simula Alt+PrtScn → pega imagen → exporta a PDF → limpia |
| **Nombre PDF** | Se genera automáticamente con el nombre del formulario y la fecha |
| **Uso** | Asignar a un botón de formulario UserForm |

---

*Módulo: 31_CodigosGenerales | Archivo: 38 - Capture screen userform.txt*
