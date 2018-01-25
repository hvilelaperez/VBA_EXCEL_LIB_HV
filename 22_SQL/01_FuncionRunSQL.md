# Función RunSQL para Excel

Función para ejecutar consultas SQL sobre archivos Excel usando ADODB, permitiendo leer, insertar y actualizar datos directamente con SQL.

## Código VBA

```vba
Function RunSQL(Target As Range, sql As String, Optional FilePath As String = Empty, _
                Optional IncludeHeader As Boolean = True)
    If FilePath = Empty Then
        FilePath = ThisWorkbook.Path & "\" & ThisWorkbook.Name
    End If
    'Obtener extensión del archivo
    Dim Ext As String
    Ext = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "."))
    
    Dim cn As Object, rs As Object
    
    '---Conectar a la fuente de datos---
    'Estas configuraciones solo aplican para archivos Excel como xlsx, xlsm, xlsb
    Set cn = CreateObject("ADODB.Connection")
    With cn
        If Ext = "xlsx" Or Ext = "xlsb" Or Ext = "xlsm" Then
            .provider = "Microsoft.ACE.OLEDB.12.0"
            .connectionstring = "Data Source=" & FilePath & ";" & _
                                "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
        ElseIf Ext = "xls" Then
            .provider = "Microsoft.Jet.OLEDB.4.0"
            .connectionstring = "Data Source=" & FilePath & ";" & _
                                "Extended Properties=""Excel 8.0;HDR=YES"";"
        End If
        On Error GoTo ERR_Handler
        .Open
    End With
    
    Dim blankPos As Integer
    blankPos = InStr(sql, " ")
    Dim SQLtype As String
    SQLtype = LCase(Left(sql, blankPos - 1))
    '---Ejecutar la consulta SQL SELECT---
    If SQLtype = "update" Or SQLtype = "insert" Then
        Set rs = cn.Execute(sql, , adExecuteNoRecords)
    Else
        Set rs = cn.Execute(sql)
    End If
    
     ' Verificar si hay datos.
    If Not rs.EOF Then
        'Transferir resultado.
        'Worksheets("Result").Range("A1").CopyFromRecordset rs
        If IncludeHeader = True Then
            'Copiar los títulos del recordset
            Dim iCol As Integer
            For iCols = 0 To rs.Fields.count - 1
                Target.Cells(1, iCols + 1).value = rs.Fields(iCols).Name
            Next
            Target.Cells(2, 1).CopyFromRecordset rs
        Else
            Target.CopyFromRecordset rs
        End If
    ' Cerrar el recordset
        rs.Close
    Else
        MsgBox "Error: No se devolvieron registros.", vbCritical
    End If
    
    '---Limpiar recursos---
    cn.Close
    Set cn = Nothing
    Set rs = Nothing
END_Handler:

    Exit Function
    
ERR_Handler:
    If Err.Number = 3704 Then Exit Function
    If Err.Number = 91 Then Exit Function
    MsgBox "Error " & Err.Number & "," & Err.Description
    Set RunSQL = Nothing
    Resume END_Handler:
End Function
```

## Descripción

### Parámetros

| Parámetro | Tipo | Requerido | Descripción |
|-----------|------|-----------|-------------|
| `Target` | Range | Sí | Celda destino donde se pegará el resultado |
| `sql` | String | Sí | Consulta SQL a ejecutar |
| `FilePath` | String | No | Ruta del archivo (default: libro actual) |
| `IncludeHeader` | Boolean | No | Incluir nombres de columna (default: True) |

### Proveedores de conexión

| Extensión | Proveedor | Versión |
|-----------|-----------|---------|
| .xlsx, .xlsb, .xlsm | Microsoft.ACE.OLEDB.12.0 | Excel 12.0 Xml |
| .xls | Microsoft.Jet.OLEDB.4.0 | Excel 8.0 |

### Tipos de consulta soportados

| Tipo | Descripción |
|------|-------------|
| `SELECT` | Lee datos del archivo Excel |
| `INSERT` | Inserta nuevos registros |
| `UPDATE` | Actualiza registros existentes |

### Ejemplo de uso

```vba
Sub EjemploRunSQL()
    Dim resultado As Range
    Set resultado = Hoja1.Range("A1")
    
    ' Consulta SELECT
    RunSQL resultado, "SELECT * FROM [Hoja1$]", , True
    
    ' Consulta con filtro
    RunSQL resultado, "SELECT * FROM [Hoja1$] WHERE [Edad] > 25"
End Sub
```

### Manejo de errores

- Error 3704: Conexión ya cerrada
- Error 91: Objeto no inicializado
- Otros errores: Muestra mensaje con número y descripción

---

*Módulo: 16 - VBA SQL | Archivo: 00 - Function RunSQL GPE.txt*
