# Conectar Excel con Access

Macros para establecer conexión entre Excel y bases de datos Access, permitiendo transferir datos en ambas direcciones.

## Código VBA

```vba
Function TaoChuoiKetNoi() As String
    Dim sAppPath As String
    ' Puedes modificar según tus necesidades de conexión
    sAppPath = ThisWorkbook.Path
    TaoChuoiKetNoi = "Driver={Microsoft Access Driver (*.mdb)}; Dbq=" & sAppPath & "\" & DBName & "; UID=Admin; PWD=;"

End Function




Function KetNoi() As Long

    On Error GoTo ErrorHandle

    Set gcnObj = CreateObject("ADODB.Connection")
    With gcnObj
        .Mode = 3                                     'es decir, adModeReadWrite
        'Si después de este tiempo no conecta, mostrará error
        .ConnectionTimeout = 30
        'CursorTypeEnum
        '--------------
        'adOpenDynamic     = 2
        'adOpenForwardOnly = 0
        'adOpenKeySet      = 1
        'adOpenStatic      = 3

        'The CursorLocationEnum:
        '-------------------------
        'adUseClient = 3 | Usa cursores del lado del cliente proporcionados por una biblioteca local.
        'Los servicios de cursor local a menudo permiten muchas funciones
        'que los cursores del controlador no tienen,
        'por lo que usar esta configuración puede proporcionar ventajas
        'con respecto a las funciones que se habilitarán.
        'Para compatibilidad con versiones anteriores, el sinónimo adUseClientBatch también es compatible.
        'adUseServer = 2 | Predeterminado. Usa cursores proporcionados por el proveedor de datos o controlador.
        'Estos cursores son a veces muy flexibles y permiten
        'mayor sensibilidad a los cambios que otros realizan en la fuente de datos.
        'Sin embargo, algunas funciones del Microsoft Cursor Service para OLE DB,
        'como los objetos Recordset desasociados,
        'no se pueden simular con cursores del lado del servidor
        'y estas funciones no estarán disponibles con esta configuración.
        'adUseNone = 1 | No usa servicios de cursor.
        '(Esta constante está obsoleta y aparece solo por compatibilidad con versiones anteriores.)
        .CursorLocation = 3
        .ConnectionString = TaoChuoiKetNoi
        .Open
    End With
    ' Si es exitoso, establecer el valor de la función en 1
    KetNoi = 1
    ' Cerrar la conexión
    gcnObj.Close

ErrorExit:

    Exit Function

ErrorHandle:
    ' Si hay error, establecer el código de error para la función KetNoi
    KetNoi = Err.Number
    Err.Clear
    Resume ErrorExit

End Function




Sub ExcelToAccess()

    Dim lLastRow As Long, i As Long
    Dim wsObj As Worksheet
    Dim arrFieldNames As Variant, arrValues As Variant, arrRecordvals As Variant
    Dim sTableName As String
    Dim lCon  As Long
    Dim oRs   As Object

    On Error GoTo ErrorHandle


    sTableName = "TB_NhanVien"

    Set wsObj = Application.ThisWorkbook.Worksheets("CSDL")
    lLastRow = FindLastRow(wsObj)

    If lLastRow = 1 Then
        ' Si la fila es 1
        ' Significa que no hay datos
        MsgBox "No hay datos para exportar." & vbCrLf & _
               "Por favor, verifique.", vbOKOnly + vbInformation, gcsAppName
        GoTo ErrorExit
    End If

    lCon = KetNoi
    If lCon = 1 Then
        gcnObj.Open
        arrFieldNames = Array("MaSo", "HoTen", _
                              "NgaySinh", "GioiTinh", _
                              "NgayNhap")             'Puedes cambiar según tus campos

        ' Acelerar
        SpeedOn
        ' Crear objeto recordset
        Set oRs = CreateObject("ADODB.Recordset")
        oRs.CursorLocation = 3                        'adUseClient
        oRs.Open sTableName, gcnObj, 3, 4             'adOpenStatic, adLockBatchOptimistic
        arrValues = wsObj.Range("A2:E" & lLastRow)
        For i = 1 To UBound(arrValues, 1)
            If Len(arrValues(i, 1)) > 0 Then
                arrRecordvals = Array(arrValues(i, 1), arrValues(i, 2), _
                                      arrValues(i, 3), arrValues(i, 4), _
                                      arrValues(i, 5))
                oRs.AddNew arrFieldNames, arrRecordvals
                Application.StatusBar = "Estás ingresando registro " & i & "/" & (lLastRow - 1)
            End If
        Next i

        Application.StatusBar = "Actualizando... Por favor espera un momento..."
        oRs.UpdateBatch
        MsgBox "Datos exportados exitosamente.", vbOKOnly + vbInformation, gcsAppName

    Else
        MsgBox "No se pudo conectar a la base de datos." & vbCrLf & _
               "Exportación fallida.", vbOKOnly + vbInformation, gcsAppName
    End If

ErrorExit:

    ' Liberar variables
    Set wsObj = Nothing
    Set oRs = Nothing
    SpeedOff
    If Not gcnObj Is Nothing Then
        If (gcnObj.State And adStateOpen) = adStateOpen Then
            gcnObj.Close
        End If
    End If
    Exit Sub

ErrorHandle:
    ' Manejar error aquí
    MsgBox "Ocurrió un error. Por favor, verifique.", vbOKOnly + vbInformation, gcsAppName
    Debug.Print Err.Number & ", " & Err.Description
    Resume ErrorExit

End Sub




Sub AccessToExcel()
    Dim lCon  As Long
    Dim sSQL  As String
    Dim adoCommand As Object                          'ADODB.Command
    Dim oRs   As Object                               'ADODB.Recordset
    Dim wsObj As Worksheet
    Dim lLastRow As Long


    Set wsObj = Application.ThisWorkbook.Worksheets("FilterFromAccess")
    lLastRow = FindLastRow(wsObj)
    lCon = KetNoi
    If lCon <> 1 Then
        MsgBox "No se pudo conectar a la base de datos." & vbCrLf & _
               "Exportación fallida.", vbOKOnly + vbInformation, gcsAppName
    Else
        ' Abrir conexión
        gcnObj.Open
        sSQL = "SELECT MaSo, HoTen, NgaySinh, GioiTinh, NgayNhap " & _
               "FROM TB_NhanVien " & _
               "WHERE MaSo='" & "HTN005" & "';"
        Set adoCommand = CreateObject("ADODB.Command")
        With adoCommand
            .CommandType = 1                          '1: adCmdText, 2: adCmdTable, 4: adCmdStoredProc
            .ActiveConnection = gcnObj
            .CommandText = sSQL
        End With
        Set oRs = CreateObject("ADODB.Recordset")
        oRs.Open adoCommand, , 3, 4
        If oRs.EOF Then
            MsgBox "No hay registros que cumplan la condición.", vbOKOnly + vbInformation, gcsAppName
        Else
            wsObj.Cells(lLastRow + 1, 1).CopyFromRecordset oRs
        End If

    End If

ErrorExit:
    Set adoCommand = Nothing
    Set oRs = Nothing
    Set wsObj = Nothing
    If Not gcnObj Is Nothing Then
        If (gcnObj.State And adStateOpen) = adStateOpen Then
            gcnObj.Close
        End If
    End If
    Exit Sub

ErrorHandle:
    ' Manejar error aquí
    MsgBox "Ocurrió un error. Por favor, verifique.", vbOKOnly + vbInformation, gcsAppName
    Debug.Print Err.Number & ", " & Err.Description
    Resume ErrorExit
End Sub
```

## Descripción

### Funciones

| Función | Descripción |
|---------|-------------|
| `TaoChuoiKetNoi` | Genera la cadena de conexión a Access |
| `KetNoi` | Establece y verifica la conexión a la base de datos |
| `ExcelToAccess` | Transfiere datos de Excel a Access |
| `AccessToExcel` | Transfiere datos de Access a Excel |

### Configuración de conexión

| Propiedad | Valor | Descripción |
|-----------|-------|-------------|
| Driver | Microsoft Access Driver (*.mdb) | Controlador ODBC para Access |
| Mode | 3 (adModeReadWrite) | Acceso de lectura y escritura |
| ConnectionTimeout | 30 segundos | Tiempo máximo de espera |
| CursorLocation | 3 (adUseClient) | Cursores del lado del cliente |

### Configuración de Recordset

| Propiedad | Valor | Descripción |
|-----------|-------|-------------|
| CursorType | 3 (adOpenStatic) | Cursor estático |
| LockType | 4 (adLockBatchOptimistic) | Bloqueo por lotes optimista |

### Proceso ExcelToAccess

1. Verifica si hay datos en la hoja
2. Establece conexión con Access
3. Crea recordset y lo abre
4. Recorre los datos y agrega registros
5. Actualiza en lote y muestra confirmación

### Proceso AccessToExcel

1. Establece conexión con Access
2. Ejecuta consulta SQL SELECT
3. Copia resultados a la hoja de Excel
4. Muestra mensaje de éxito o error

### Manejo de errores

- `ErrorHandle`: Captura errores y muestra información
- `ErrorExit`: Punto de salida para liberar recursos
- `Debug.Print`: Registra errores en la ventana Inmediato

---

*Módulo: 16 - VBA SQL | Archivo: 05 - Ket noi giua Excel va Access.txt*
