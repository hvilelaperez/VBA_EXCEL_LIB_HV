# Exportar Hoja a Archivo PDF

Exporta una hoja de cálculo de Excel a un archivo PDF, con selección de carpeta destino y manejo de archivos existentes.

## Código VBA

```vba
Sub ObtenerDatosSpinner()
    Dim ws As Worksheet 'Establecer variable ws como la hoja activa
    Set ws = ActiveSheet
    Dim rngData As Range 'Establecer rango de la tabla de datos comenzando desde D3
    Set rngData = Range(ws.Range("D3").End(xlDown), ws.Range("D3").End(xlToRight))
    Dim Data As Variant 'Revisar contenido de la tabla de Datos
    Data = rngData.Value
    Dim iCell As Integer 'Posición del resultado del Spinner
    iCell = ws.Range("M2").Value

    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TableTest")

    tbl.DataBodyRange.Delete 'Eliminar datos en la Tabla
    For x = LBound(Data) To UBound(Data) 'Recorrer todo el Data desde la primera hasta la última fila
        If iCell = Data(x, 3) Then 'Si el resultado del Spinner = resultado de la columna F del Data
            With tbl.ListRows.Add 'Agregar datos a la Tabla
                .Range(, 2).Value = Data(x, 1) 'Obtener contenido de la columna 2
                .Range(, 3).Value = Data(x, 2) 'Obtener contenido de la columna 3
            End With
        End If
    Next
    tbl.DataBodyRange.AutoFit 'Ajustar automáticamente el ancho en la Tabla
End Sub


Sub ExportarPDF()
    'Buscar la última fila de la tabla
    Dim maxR As Integer
    maxR = Sheet1.Range("F" & Rows.Count).End(xlUp).Value 'Guardar en la columna que contiene la última fila, columna F

    'Determinar la ruta hacia la carpeta para guardar el resultado
    Set xSht = ActiveSheet
    Set xFileDlg = Application.FileDialog(msoFileDialogFolderPicker)
    
    If xFileDlg.Show = True Then
        xFolder = xFileDlg.SelectedItems(1)
    Else
        MsgBox "Debe especificar una carpeta para guardar el PDF." & vbCrLf & vbCrLf & "Presione Aceptar para salir de esta macro.", vbCritical, "Debe Especificar Carpeta de Destino"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False 'Omitir la actualización de la pantalla

    For x = 1 To maxR
        With ActiveSheet.Range("M2") '<== Posición del resultado del Spin Button
            .Value = x
            Call ObtenerDatosSpinner '<== Llamar al comando para obtener datos en PXK después de cada cambio del resultado del Spin
            
            xFile = xFolder + "\" + xSht.Range("K4").Value + ".pdf" 'Determinar el nombre del archivo a guardar, el nombre se toma de la posición K4
            
            'Verificar si el archivo ya existe, ignorar el nombre
            If Len(Dir(xFile)) > 0 Then
                xYesorNo = MsgBox(xFile & " ya existe." & vbCrLf & vbCrLf & "¿Desea sobrescribirlo?", _
                                  vbYesNo + vbQuestion, "Archivo Existe")
                On Error Resume Next
                If xYesorNo = vbYes Then
                    Kill xFile
                Else
                    MsgBox "Si no sobrescribe el PDF existente, no puedo continuar." _
                                & vbCrLf & vbCrLf & "Presione Aceptar para salir de esta macro.", vbCritical, "Saliendo de la Macro"
                    Exit Sub
                End If
                If Err.Number <> 0 Then
                    MsgBox "No se puede eliminar el archivo existente. Por favor asegúrese de que el archivo no esté abierto o protegido contra escritura." _
                                & vbCrLf & vbCrLf & "Presione Aceptar para salir de esta macro.", vbCritical, "No se Puede Eliminar el Archivo"
                    Exit Sub
                End If
            End If
            
            Set xUsedRng = xSht.UsedRange
            
            If Application.WorksheetFunction.CountA(xUsedRng.Cells) <> 0 Then
                'Guardar con formato de archivo PDF
                xSht.ExportAsFixedFormat Type:=xlTypePDF, Filename:=xFile, Quality:=xlQualityStandard
            Else
                MsgBox "La hoja activa no puede estar en blanco" 'Error cuando la tabla no tiene datos
                Exit Sub
            End If
        End With
    Next
    Application.ScreenUpdating = True 'Restaurar el modo de actualización de pantalla después de completar el bucle

    MsgBox "¡Completado!" 'Notificación de trabajo finalizado
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Sub `ObtenerDatosSpinner`** | Carga datos en una tabla según el valor seleccionado por un Spin Button |
| **Sub `ExportarPDF`** | Exporta cada vista del Spinner a un archivo PDF independiente |
| **Selección de carpeta** | Utiliza `FileDialog` para que el usuario elija dónde guardar |
| **Nombre del archivo** | Se toma del valor en la celda K4 de la hoja |
| **Manejo de duplicados** | Verifica si el archivo existe y ofrece sobrescribir |

| Flujo del procedimiento |
|------------------------|
| 1. Determinar la última fila de datos |
| 2. Solicitar carpeta de destino al usuario |
| 3. Recorrer cada valor del Spinner |
| 4. Cargar datos correspondientes en la tabla |
| 5. Exportar a PDF con el nombre basado en K4 |
| 6. Verificar si el archivo ya existe |
| 7. Guardar en formato PDF de calidad estándar |

---

*Módulo: 11 - VBA Excel PDF | Archivo: 00 - Xuat file PDF*
