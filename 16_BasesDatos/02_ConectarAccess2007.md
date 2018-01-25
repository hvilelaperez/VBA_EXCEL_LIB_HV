# Conectar a Base de Datos Access 2007-2010

Establece una conexión ADODB con una base de datos Access 2007/2010 (.accdb) y ejecuta consultas para leer registros.

## Código VBA

```vba
Option Explicit
 
Public Const dbName = "student.accdb"
 
Sub ADODB_Conectar()
    Dim dbPath As String
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim strConn As String
    Dim query As String
    
    On Error GoTo ManejoError
    
    ' Crear dbPath desde la carpeta actual y dbName
    dbPath = Application.ActiveWorkbook.Path & "\" & dbName
    ' Información para conectar a Access 2007/2010
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & dbPath & ";" & _
        "User Id=admin;Password="
    ' Abrir conexión
    conn.Open (strConn)
    ' Definir consulta
    query = "SELECT * FROM student"
    ' Ejecutar la consulta
    rs.Open query, conn, adOpenKeyset
    ' Mostrar cantidad de registros
    MsgBox rs.RecordCount
    ' Mostrar datos de la base de datos
    Do Until rs.EOF
        MsgBox rs.Fields.Item("name") & ", " & rs.Fields.Item("age")
        rs.MoveNext
    Loop
    
    GoTo FinSub
ManejoError:
    MsgBox Err.Number & ": " & Err.Description
FinSub:
    Set rs = Nothing
    Set conn = Nothing
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Proveedor** | Microsoft.ACE.OLEDB.12.0 (para archivos .accdb de Access 2007+) |
| **Archivo** | `student.accdb` en la misma carpeta del libro activo |
| **Consulta** | `SELECT * FROM student` - selecciona todos los campos de la tabla |
| **Manejo de errores** | Utiliza `On Error GoTo` para capturar errores de conexión |
| **Limpieza** | Libera los objetos `Recordset` y `Connection` al finalizar |

| Diferencia con Access 2003 | Valor |
|---------------------------|-------|
| **Proveedor** | ACE.OLEDB.12.0 en lugar de Jet.OLEDB.4.0 |
| **Extensión** | .accdb en lugar de .mdb |
| **Formato** | Nuevo formato de base de datos Access 2007+ |

---

*Módulo: 10 - VBA Database | Archivo: 01 - Ketnoi voi access 2007-2010 DB*
