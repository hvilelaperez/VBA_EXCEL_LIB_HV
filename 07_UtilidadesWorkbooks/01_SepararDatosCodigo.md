# Separar Datos y Código en Workbooks

Técnicas para separar el código VBA de los datos en libros diferentes.

## Código VBA

```vba
Option Explicit

' ============================================================
' SEPARAR LIBRO EN DATOS Y CÓDIGO
' Exporta módulos VBA y crea copia sin macros
' ============================================================
Public Sub SepararDatosYC()
    Dim rutaDestino As String
    Dim nombreBase As String
    Dim ws As Worksheet
    
    nombreBase = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    rutaDestino = ThisWorkbook.Path & "\"
    
    Application.ScreenUpdating = False
    
    ' 1. Guardar copia como .xlsx (sin macros)
    ThisWorkbook.SaveCopyAs rutaDestino & nombreBase & "_Datos.xlsx"
    
    ' 2. Exportar módulos VBA
    Dim vbComp As Object
    Dim contador As Long
    contador = 0
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 1 Then ' vbext_ct_StdModule
            vbComp.Export rutaDestino & "Modulo_" & vbComp.Name & ".bas"
            contador = contador + 1
        ElseIf vbComp.Type = 2 Then ' vbext_ct_ClassModule
            vbComp.Export rutaDestino & "Clase_" & vbComp.Name & ".cls"
            contador = contador + 1
        ElseIf vbComp.Type = 3 Then ' vbext_ct_MSForm
            vbComp.Export rutaDestino & "Form_" & vbComp.Name & ".frm"
            contador = contador + 1
        End If
    Next vbComp
    
    Application.ScreenUpdating = True
    
    MsgBox "Proceso completado." & vbCrLf & _
           "Copia sin macros: " & nombreBase & "_Datos.xlsx" & vbCrLf & _
           "Módulos exportados: " & contador, vbInformation
End Sub


' ============================================================
' ABRIR LIBRO CON CÓDIGO SEPARADO
' Abre el libro de código junto con el de datos
' ============================================================
Public Sub AbrirConCodigoSeparado()
    Dim rutaLibroDatos As String
    Dim rutaLibroCodigo As String
    
    rutaLibroDatos = ThisWorkbook.Path & "\Datos.xlsx"
    rutaLibroCodigo = ThisWorkbook.Path & "\Codigo.xlsm"
    
    ' Verificar si el libro de código ya está abierto
    If Not LibroAbierto("Codigo.xlsm") Then
        Workbooks.Open rutaLibroCodigo
    End If
    
    ' Ejecutar macro del libro de código
    Application.Run "Codigo.xlsm!MiMacro"
End Sub


' ============================================================
' FUNCIÓN: Verificar si libro está abierto
' ============================================================
Private Function LibroAbierto(nombreArchivo As String) As Boolean
    Dim wb As Workbook
    
    On Error Resume Next
    Set wb = Application.Workbooks(nombreArchivo)
    LibroAbierto = Not wb Is Nothing
    On Error GoTo 0
End Function


' ============================================================
' EVENTO: Al abrir libro, buscar libro de código
' ============================================================
' Colocar en módulo del libro de datos
Private Sub Workbook_Open()
    Dim rutaCodigo As String
    
    On Error Resume Next
    
    ' Verificar si libro de código ya está abierto
    If Workbooks("Codigo.xlsm").Name = "" Then
    
        On Error GoTo 0
        
        ' Abrir libro de código
        rutaCodigo = ThisWorkbook.Path & "\Codigo.xlsm"
        
        If Dir(rutaCodigo) <> "" Then
            Workbooks.Open rutaCodigo
        End If
    End If
    
    On Error GoTo 0
    
    ' Ejecutar macro de inicialización
    Application.Run "Codigo.xlsm!Inicializar"
End Sub


' ============================================================
' EXPORTAR TODOS LOS MÓDULOS
' ============================================================
Public Sub ExportarModulos()
    Dim vbComp As Object
    Dim rutaExportacion As String
    
    rutaExportacion = ThisWorkbook.Path & "\Modulos_Exportados\"
    
    ' Crear carpeta si no existe
    If Dir(rutaExportacion, vbDirectory) = "" Then
        MkDir rutaExportacion
    End If
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' Módulo estándar
                vbComp.Export rutaExportacion & vbComp.Name & ".bas"
            Case 2 ' Clase
                vbComp.Export rutaExportacion & vbComp.Name & ".cls"
            Case 3 ' Formulario
                vbComp.Export rutaExportacion & vbComp.Name & ".frm"
        End Select
    Next vbComp
    
    MsgBox "Módulos exportados a:" & vbCrLf & rutaExportacion, vbInformation
End Sub


' ============================================================
' IMPORTAR MÓDULOS DESDE CARPETA
' ============================================================
Public Sub ImportarModulos()
    Dim rutaImportacion As String
    Dim nombreArchivo As String
    Dim contador As Long
    
    rutaImportacion = ThisWorkbook.Path & "\Modulos_Exportados\"
    
    If Dir(rutaImportacion) = "" Then
        MsgBox "Carpeta no encontrada.", vbExclamation
        Exit Sub
    End If
    
    nombreArchivo = Dir(rutaImportacion & "*.bas")
    contador = 0
    
    Do While nombreArchivo <> ""
        ThisWorkbook.VBProject.Import rutaImportacion & nombreArchivo
        contador = contador + 1
        nombreArchivo = Dir()
    Loop
    
    MsgBox contador & " módulos importados.", vbInformation
End Sub
```

## Estructura de Archivos

```
Proyecto/
├── Datos.xlsx          (solo datos, sin macros)
├── Codigo.xlsm         (solo código VBA)
└── Modulos_Exportados/
    ├── Modulo1.bas
    ├── Clase1.cls
    └── Formulario1.frm
```

## Notas

- Para exportar/importar módulos, habilitar acceso al objeto VBProject
- Archivo > Opciones > Centro de confianza > Configuración del centro de confianza
- Marcar "Confiar en el acceso a proyectos de VBA"

---

*Módulo: 07_UtilidadesWorkbooks | Archivo: 01_SepararDatosCodigo.md*
