# Variante para Remover Protección de Archivos Excel

Herramienta completa que manipula archivos Excel (.xlsx/.xlsm/.xlsb) como archivos ZIP para remover protección de hojas, libros, objetos y nombres definidos. Incluye tres subrutinas principales.

## Código VBA

```vba
Option Explicit
'Desarrollado por d.nguyen @ HocExcel.Online
'Todos los derechos reservados.

Sub RemoveObject()

Dim dialogBox As Object
Dim sourceFullName As String
Dim sourceFilePath As String
Dim sourceFileName As String
Dim sourceFileType As String
Dim newFileName As Variant
Dim tempFileName As String
Dim zipFilePath As Variant
Dim oApp As Object
Dim fso As Object
Dim xmlSheetFile As String
Dim xmlFile As Integer
Dim xmlFileContent As String
Dim xmlStartProtectionCode As Double
Dim xmlEndProtectionCode As Double
Dim xmlProtectionString As String

'Abrir cuadro de diálogo para seleccionar archivo
Set dialogBox = Application.FileDialog(msoFileDialogFilePicker)
dialogBox.AllowMultiSelect = False
dialogBox.Title = "Seleccionar archivo para remover objetos"

If dialogBox.Show = -1 Then
    sourceFullName = dialogBox.SelectedItems(1)
Else
    MsgBox "¡Por favor seleccione un archivo!"
    Exit Sub
End If

TurnOnOff
On Error GoTo xError

sourceFullName = xlChangeType(sourceFullName)
'Obtener ruta del carpeta, tipo y nombre del archivo
sourceFilePath = Left(sourceFullName, InStrRev(sourceFullName, "\"))
sourceFileType = Mid(sourceFullName, InStrRev(sourceFullName, ".") + 1)
sourceFileName = Mid(sourceFullName, Len(sourceFilePath) + 1)
sourceFileName = Left(sourceFileName, InStrRev(sourceFileName, ".") - 1)

'Usar fecha y hora para crear nombre de archivo único
tempFileName = "Temp" & Format(Now, " dd-mmm-yy h-mm-ss")

'Copiar y renombrar archivo original a archivo ZIP
newFileName = sourceFilePath & tempFileName & ".zip"
On Error Resume Next
FileCopy sourceFullName, newFileName

If Err.Number <> 0 Then
    MsgBox "No se pudo copiar " & sourceFullName & vbNewLine _
        & "Verifique que el archivo esté cerrado e intente de nuevo"
    Exit Sub
End If
On Error GoTo 0

'Crear carpeta para descomprimir
zipFilePath = sourceFilePath & tempFileName & "\"
MkDir zipFilePath

'Extraer archivos a la carpeta creada
Set oApp = CreateObject("Shell.Application")
oApp.Namespace(zipFilePath).CopyHere oApp.Namespace(newFileName).items

'Recorrer cada archivo en la carpeta \xl\worksheets del archivo descomprimido
xmlSheetFile = Dir(zipFilePath & "\xl\worksheets\*.xml*")
Do While xmlSheetFile <> ""

    'Leer texto del archivo a una variable
    xmlFile = FreeFile
    Open zipFilePath & "xl\worksheets\" & xmlSheetFile For Input As xmlFile
    xmlFileContent = Input(LOF(xmlFile), xmlFile)
    Close xmlFile

'Manipular el texto del archivo para remover la protección
xmlStartProtectionCode = 0
xmlStartProtectionCode = InStr(1, xmlFileContent, "<drawing")
If xmlStartProtectionCode > 0 Then

    xmlEndProtectionCode = InStr(xmlStartProtectionCode, _
        xmlFileContent, "/>") + 2 '"/>" tiene 2 caracteres
    xmlProtectionString = Mid(xmlFileContent, xmlStartProtectionCode, _
        xmlEndProtectionCode - xmlStartProtectionCode)
    xmlFileContent = Replace(xmlFileContent, xmlProtectionString, "")

End If

    'Escribir el texto modificado al archivo
    xmlFile = FreeFile
    Open zipFilePath & "xl\worksheets\" & xmlSheetFile For Output As xmlFile
    Print #xmlFile, xmlFileContent
    Close xmlFile

    'Ir al siguiente archivo XML en el directorio
    xmlSheetFile = Dir

Loop


'Crear archivo ZIP vacío
Open sourceFilePath & tempFileName & ".zip" For Output As #1
Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
Close #1

'Mover archivos al archivo ZIP
oApp.Namespace(sourceFilePath & tempFileName & ".zip").CopyHere _
oApp.Namespace(zipFilePath).items
'Mantener script esperando hasta que termine la compresión
On Error Resume Next
Do Until oApp.Namespace(sourceFilePath & tempFileName & ".zip").items.Count = _
    oApp.Namespace(zipFilePath).items.Count
    Application.Wait (Now + TimeValue("0:00:01"))
Loop
On Error GoTo 0

'Eliminar archivos y carpetas creados durante la ejecución
Set fso = CreateObject("scripting.filesystemobject")
fso.deletefolder sourceFilePath & tempFileName

'Renombrar el archivo final a extensión original
Name sourceFilePath & tempFileName & ".zip" As sourceFilePath & sourceFileName & Format(Now, " ddmmmyyhmmss") & "." & sourceFileType

'Mostrar mensaje de confirmación
MsgBox "Todos los objetos del libro han sido removidos.", _
vbInformation + vbOKOnly, Title:="Limpiar Objetos"
TurnOnOff True
Exit Sub

xError:
MsgBox "¡Error!", vbExclamation, Title:="Limpiar Objetos"
TurnOnOff True
End Sub

Sub RemoveAllName()

Dim dialogBox As Object
Dim sourceFullName As String
Dim sourceFilePath As String
Dim sourceFileName As String
Dim sourceFileType As String
Dim newFileName As Variant
Dim tempFileName As String
Dim zipFilePath As Variant
Dim oApp As Object
Dim fso As Object
Dim xmlSheetFile As String
Dim xmlFile As Integer
Dim xmlFileContent As String
Dim xmlStartProtectionCode As Double
Dim xmlEndProtectionCode As Double
Dim xmlProtectionString As String
 
'Abrir cuadro de diálogo para seleccionar archivo
Set dialogBox = Application.FileDialog(msoFileDialogFilePicker)
dialogBox.AllowMultiSelect = False
dialogBox.Title = "Seleccionar archivo para remover todos los nombres"

If dialogBox.Show = -1 Then
    sourceFullName = dialogBox.SelectedItems(1)
Else
    MsgBox "¡Por favor seleccione un archivo!"
    Exit Sub
End If

TurnOnOff
On Error GoTo xError

Dim regex As Object, str As String
Set regex = CreateObject("VBScript.RegExp")

sourceFullName = xlChangeType(sourceFullName)
'Obtener ruta del carpeta, tipo y nombre del archivo
sourceFilePath = Left(sourceFullName, InStrRev(sourceFullName, "\"))
sourceFileType = Mid(sourceFullName, InStrRev(sourceFullName, ".") + 1)
sourceFileName = Mid(sourceFullName, Len(sourceFilePath) + 1)
sourceFileName = Left(sourceFileName, InStrRev(sourceFileName, ".") - 1)

'Usar fecha y hora para crear nombre de archivo único
tempFileName = "Temp" & Format(Now, " dd-mmm-yy h-mm-ss")

'Copiar y renombrar archivo original a archivo ZIP
newFileName = sourceFilePath & tempFileName & ".zip"
On Error Resume Next
FileCopy sourceFullName, newFileName

If Err.Number <> 0 Then
    MsgBox "No se pudo copiar " & sourceFullName & vbNewLine _
        & "Verifique que el archivo esté cerrado e intente de nuevo"
    Exit Sub
End If
On Error GoTo 0

'Crear carpeta para descomprimir
zipFilePath = sourceFilePath & tempFileName & "\"
MkDir zipFilePath

'Extraer archivos a la carpeta creada
Set oApp = CreateObject("Shell.Application")
oApp.Namespace(zipFilePath).CopyHere oApp.Namespace(newFileName).items

'Leer texto del archivo xl\workbook.xml a una variable
xmlFile = FreeFile
Open zipFilePath & "xl\workbook.xml" For Input As xmlFile
xmlFileContent = Input(LOF(xmlFile), xmlFile)
Close xmlFile

With regex
  .Pattern = "(<definedName\s.*>#REF!<\/definedName>)"
  .Global = True 'Si es False, solo reemplazaría el primero
End With
xmlFileContent = regex.Replace(xmlFileContent, "")

'Escribir el texto modificado al archivo
xmlFile = FreeFile
Open zipFilePath & "xl\workbook.xml" & xmlSheetFile For Output As xmlFile
Print #xmlFile, xmlFileContent
Close xmlFile

'Crear archivo ZIP vacío
Open sourceFilePath & tempFileName & ".zip" For Output As #1
Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
Close #1

'Mover archivos al archivo ZIP
oApp.Namespace(sourceFilePath & tempFileName & ".zip").CopyHere _
oApp.Namespace(zipFilePath).items
'Mantener script esperando hasta que termine la compresión
On Error Resume Next
Do Until oApp.Namespace(sourceFilePath & tempFileName & ".zip").items.Count = _
    oApp.Namespace(zipFilePath).items.Count
    Application.Wait (Now + TimeValue("0:00:01"))
Loop
On Error GoTo 0

'Eliminar archivos y carpetas creados durante la ejecución
Set fso = CreateObject("scripting.filesystemobject")
fso.deletefolder sourceFilePath & tempFileName

'Renombrar el archivo final a extensión original
Name sourceFilePath & tempFileName & ".zip" As sourceFilePath & sourceFileName & Format(Now, " ddmmmyyhmmss") & "." & sourceFileType

'Mostrar mensaje de confirmación
MsgBox "Todos los nombres definidos han sido removidos.", _
vbInformation + vbOKOnly, Title:="Protección de Contraseña"
TurnOnOff True
Exit Sub

xError:
MsgBox "¡Error!", vbExclamation, Title:="Limpiar Objetos"
TurnOnOff True
End Sub

Sub RemoveProtection()

Dim dialogBox As Object
Dim sourceFullName As String
Dim sourceFilePath As String
Dim sourceFileName As String
Dim sourceFileType As String
Dim newFileName As Variant
Dim tempFileName As String
Dim zipFilePath As Variant
Dim oApp As Object
Dim fso As Object
Dim xmlSheetFile As String
Dim xmlFile As Integer
Dim xmlFileContent As String
Dim xmlStartProtectionCode As Double
Dim xmlEndProtectionCode As Double
Dim xmlProtectionString As String

'Abrir cuadro de diálogo para seleccionar archivo
Set dialogBox = Application.FileDialog(msoFileDialogFilePicker)
dialogBox.AllowMultiSelect = False
dialogBox.Title = "Seleccionar archivo para remover protección"

If dialogBox.Show = -1 Then
    sourceFullName = dialogBox.SelectedItems(1)
Else
    MsgBox "¡Por favor seleccione un archivo!"
    Exit Sub
End If

TurnOnOff
On Error GoTo xError

sourceFullName = xlChangeType(sourceFullName)
'Obtener ruta del carpeta, tipo y nombre del archivo
sourceFilePath = Left(sourceFullName, InStrRev(sourceFullName, "\"))
sourceFileType = Mid(sourceFullName, InStrRev(sourceFullName, ".") + 1)
sourceFileName = Mid(sourceFullName, Len(sourceFilePath) + 1)
sourceFileName = Left(sourceFileName, InStrRev(sourceFileName, ".") - 1)

'Usar fecha y hora para crear nombre de archivo único
tempFileName = "Temp" & Format(Now, " dd-mmm-yy h-mm-ss")

'Copiar y renombrar archivo original a archivo ZIP
newFileName = sourceFilePath & tempFileName & ".zip"
On Error Resume Next
FileCopy sourceFullName, newFileName

If Err.Number <> 0 Then
    MsgBox "No se pudo copiar " & sourceFullName & vbNewLine _
        & "Verifique que el archivo esté cerrado e intente de nuevo"
    Exit Sub
End If
On Error GoTo 0

'Crear carpeta para descomprimir
zipFilePath = sourceFilePath & tempFileName & "\"
MkDir zipFilePath

'Extraer archivos a la carpeta creada
Set oApp = CreateObject("Shell.Application")
oApp.Namespace(zipFilePath).CopyHere oApp.Namespace(newFileName).items

'Recorrer cada archivo en la carpeta \xl\worksheets del archivo descomprimido
xmlSheetFile = Dir(zipFilePath & "\xl\worksheets\*.xml*")
Do While xmlSheetFile <> ""

    'Leer texto del archivo a una variable
    xmlFile = FreeFile
    Open zipFilePath & "xl\worksheets\" & xmlSheetFile For Input As xmlFile
    xmlFileContent = Input(LOF(xmlFile), xmlFile)
    Close xmlFile

    'Manipular el texto del archivo
    xmlStartProtectionCode = 0
    xmlStartProtectionCode = InStr(1, xmlFileContent, "<sheetProtection")

    If xmlStartProtectionCode > 0 Then

        xmlEndProtectionCode = InStr(xmlStartProtectionCode, _
            xmlFileContent, "/>") + 2 '"/>" tiene 2 caracteres
        xmlProtectionString = Mid(xmlFileContent, xmlStartProtectionCode, _
            xmlEndProtectionCode - xmlStartProtectionCode)
        xmlFileContent = Replace(xmlFileContent, xmlProtectionString, "")

    End If

    'Escribir el texto modificado al archivo
    xmlFile = FreeFile
    Open zipFilePath & "xl\worksheets\" & xmlSheetFile For Output As xmlFile
    Print #xmlFile, xmlFileContent
    Close xmlFile

    'Ir al siguiente archivo XML en el directorio
    xmlSheetFile = Dir

Loop

'Leer texto del archivo xl\workbook.xml a una variable
xmlFile = FreeFile
Open zipFilePath & "xl\workbook.xml" For Input As xmlFile
xmlFileContent = Input(LOF(xmlFile), xmlFile)
Close xmlFile

'Manipular el texto del archivo para remover la protección del libro
xmlStartProtectionCode = 0
xmlStartProtectionCode = InStr(1, xmlFileContent, "<workbookProtection")
If xmlStartProtectionCode > 0 Then

    xmlEndProtectionCode = InStr(xmlStartProtectionCode, _
        xmlFileContent, "/>") + 2 '"/>" tiene 2 caracteres
    xmlProtectionString = Mid(xmlFileContent, xmlStartProtectionCode, _
        xmlEndProtectionCode - xmlStartProtectionCode)
    xmlFileContent = Replace(xmlFileContent, xmlProtectionString, "")

End If



'Manipular el texto del archivo para remover la contraseña de modificación
xmlStartProtectionCode = 0
xmlStartProtectionCode = InStr(1, xmlFileContent, "<fileSharing")
If xmlStartProtectionCode > 0 Then

    xmlEndProtectionCode = InStr(xmlStartProtectionCode, xmlFileContent, _
        "/>") + 2 '"/>" tiene 2 caracteres
    xmlProtectionString = Mid(xmlFileContent, xmlStartProtectionCode, _
        xmlEndProtectionCode - xmlStartProtectionCode)
    xmlFileContent = Replace(xmlFileContent, xmlProtectionString, "")

End If

'Escribir el texto modificado al archivo
xmlFile = FreeFile
Open zipFilePath & "xl\workbook.xml" & xmlSheetFile For Output As xmlFile
Print #xmlFile, xmlFileContent
Close xmlFile

'Crear archivo ZIP vacío
Open sourceFilePath & tempFileName & ".zip" For Output As #1
Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
Close #1

'Mover archivos al archivo ZIP
oApp.Namespace(sourceFilePath & tempFileName & ".zip").CopyHere _
oApp.Namespace(zipFilePath).items
'Mantener script esperando hasta que termine la compresión
On Error Resume Next
Do Until oApp.Namespace(sourceFilePath & tempFileName & ".zip").items.Count = _
    oApp.Namespace(zipFilePath).items.Count
    Application.Wait (Now + TimeValue("0:00:01"))
Loop
On Error GoTo 0

'Eliminar archivos y carpetas creados durante la ejecución
Set fso = CreateObject("scripting.filesystemobject")
fso.deletefolder sourceFilePath & tempFileName

'Renombrar el archivo final a extensión original
Name sourceFilePath & tempFileName & ".zip" As sourceFilePath & sourceFileName & Format(Now, " ddmmmyyhmmss") & "." & sourceFileType

'Mostrar mensaje de confirmación
MsgBox "Las contraseñas de protección del libro y las hojas han sido removidas.", _
vbInformation + vbOKOnly, Title:="Protección de Contraseña"
TurnOnOff True
Exit Sub

xError:
MsgBox "¡Error!", vbExclamation, Title:="Limpiar Objetos"
TurnOnOff True
End Sub

Function xlChangeType(ByVal sourceFullName As String) As String

Dim sourceFilePath As String, sourceFileType As String, sourceFileName As String
Dim strNewPathFix As String
'Obtener ruta del carpeta, tipo y nombre del archivo
sourceFilePath = Left(sourceFullName, InStrRev(sourceFullName, "\"))
sourceFileType = Mid(sourceFullName, InStrRev(sourceFullName, ".") + 1)
sourceFileName = Mid(sourceFullName, Len(sourceFilePath) + 1)
sourceFileName = Left(sourceFileName, InStrRev(sourceFileName, ".") - 1)
If Not sourceFileType = "xlsb" Then
    xlChangeType = sourceFullName
    Exit Function
End If
    Dim strCurrentFileExt   As String
    Dim strNewFileExt       As String
    Dim xlFile              As Workbook
    Dim strNewName          As String
    Dim strFolderPath       As String
    Dim strTimeNow As String
    strCurrentFileExt = ".xlsb"
    strNewFileExt = ".xlsm"

    strFolderPath = sourceFilePath
    If Right(strFolderPath, 1) <> "\" Then
        strFolderPath = strFolderPath & "\"
    End If
            strTimeNow = Format(Now, "ddmmmyyhmmss")
            Application.DisplayAlerts = False
            Set xlFile = Workbooks.Open(sourceFullName, , True)
            strNewName = Replace(sourceFileName, strCurrentFileExt, strNewFileExt)
            Select Case strNewFileExt
            Case ".xlsx"
                xlFile.SaveAs strFolderPath & strNewName & strTimeNow, XlFileFormat.xlOpenXMLWorkbook
            Case ".xlsm"
                xlFile.SaveAs strFolderPath & strNewName & strTimeNow, XlFileFormat.xlOpenXMLWorkbookMacroEnabled
            End Select
            xlFile.Close
            Application.DisplayAlerts = True

ClearMemory:
    strNewPathFix = strFolderPath & strNewName & strTimeNow & strNewFileExt
    strCurrentFileExt = vbNullString
    strNewFileExt = vbNullString
    Set xlFile = Nothing
    strNewName = vbNullString
    strFolderPath = vbNullString
    xlChangeType = strNewPathFix
End Function
```

## Descripción

| Sub/Función | Descripción |
|---|---|
| `RemoveObject` | Remueve todos los objetos (gráficos, formas, etc.) de las hojas del libro manipulando el XML interno |
| `RemoveAllName` | Remueve todos los nombres definidos (`definedNames`) que referencian `#REF!` usando RegExp |
| `RemoveProtection` | Remueve la protección de hojas (`<sheetProtection`), del libro (`<workbookProtection`) y la contraseña de modificación (`<fileSharing`) |
| `xlChangeType` | Convierte archivos `.xlsb` a `.xlsm` antes de procesar, ya que el formato binario no se puede manipular como ZIP |

**Proceso común:** Todos los subs siguen el patrón: copiar archivo → renombrar a .zip → descomprimir → modificar XML → recomprimir → renombrar extensión original.

**Advertencia:** Cierre el archivo Excel antes de ejecutar cualquier sub. Se crean archivos temporales que se eliminan automáticamente.

---

*Módulo: 94 - VBA Crack - Remove password | Archivo: 01 - Remove.txt*
