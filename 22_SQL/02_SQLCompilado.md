# SQL Compilado para Excel

Macro que ejecuta consultas SQL sobre el propio libro de Excel usando ADODB, creando una nueva hoja con los resultados.

## Código VBA

```vba
Sub SQL_01(sourceSQL As String)
    Dim orderConn As Object, orderData As Object, orderField As Object
    Set orderConn = CreateObject("ADODB.Connection")
    Set orderData = CreateObject("ADODB.Recordset")
    Dim i As Integer
    
    On Error GoTo close_Connection
    With orderConn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
        .Open
    End With
    
    With orderData
        .activeConnection = orderConn
        .Source = sourceSQL
        .LockType = 1 'adLockReadOnly
        .CursorType = 0 'adForwardOnly
        .Open
    End With
    i = 0
    
    On Error GoTo close_RecordSet
    Worksheets.Add after:=Worksheets(Worksheets.Count)
    For Each orderField In orderData.Fields
        i = i + 1
        Cells(1, i) = orderField.Name
    Next orderField
    Range("A2").CopyFromRecordset orderData
    
    'Renombrar columnas
    'Select [Order Date] as [Date],Item as [test-Item],region as [Departement] from SaleOrders
    'Agregar una columna con valor fijo: Select [Order Date],'SQL Es genial' from SalesOrders
    'Select [Unit Cost] as [Cost],Region as [Departement],'SQL Es genial' as [SQL Jeje] from SalesOrders
    'Se pueden agregar columnas calculadas sobre valores existentes:
    'Select Total,(Total*0.1) as [impuesto 10%] from SalesOrders
    'Select sum(Total) as [Suma Total] from SalesOrders
    'Select Sum (Units) as [Suma Unidades],Sum(Total) as [Suma Total], (Sum (Total)/Sum(Units)) as [Ingreso por Unidad] from SalesOrders
    'Select * from SalesOrders where [Unit Cost] = 1.99
    'Select * from SalesOrders where [Unit Cost] <= 50
    'Select * from SalesOrders where [orderdate] >= #25-06-2016#
    'Select * from SalesOrders where [rep] IN("morgan","parent")
    'Select rep, item,Total as [Total ventas] from SalesOrders order by Total ASC
    'Select rep, item, sum(Total) as [Total ventas] from SalesOrders group by rep,item
    
    'Obtener año del valor en columna
'    Select year(OrderDate) as [Año], month(OrderDate) as [Mes], sum(Total) as [Ventas totales] from SalesOrders group by year(OrderDate),month(OrderDate)
    
close_RecordSet:
    If Err.Number <> 0 Then MsgBox Err.Description
    orderData.Close
    
close_Connection:
    If Err.Number <> 0 Then MsgBox Err.Description
    orderConn.Close
    
    
    Set orderConn = Nothing
    Set orderData = Nothing
End Sub
```

## Descripción

### Función principal

| Elemento | Descripción |
|----------|-------------|
| **Nombre** | `SQL_01` |
| **Parámetro** | `sourceSQL` - Consulta SQL a ejecutar |
| **Resultado** | Crea nueva hoja con los datos de la consulta |

### Configuración de conexión

| Propiedad | Valor | Descripción |
|-----------|-------|-------------|
| Provider | Microsoft.ACE.OLEDB.12.0 | Proveedor OLE DB para Excel |
| HDR=YES | Activo | Primera fila contiene encabezados |
| IMEX=1 | Activo | Modo de importación |

### Ejemplos de consultas SQL

| Tipo | Ejemplo |
|------|---------|
| Renombrar columnas | `SELECT [Fecha] as [Date] FROM [Hoja1$]` |
| Valor fijo | `SELECT [Nombre], 'Activo' as [Estado] FROM [Hoja1$]` |
| Cálculos | `SELECT [Total], [Total]*0.1 as [Impuesto] FROM [Hoja1$]` |
| Agrupación | `SELECT [Rep], SUM([Total]) as [Ventas] FROM [Hoja1$] GROUP BY [Rep]` |
| Filtros | `SELECT * FROM [Hoja1$] WHERE [Costo] <= 50` |
| Fechas | `SELECT * FROM [Hoja1$] WHERE [Fecha] >= #25-06-2016#` |
| IN | `SELECT * FROM [Hoja1$] WHERE [Rep] IN('Ana','Luis')` |
| Orden | `SELECT * FROM [Hoja1$] ORDER BY [Total] ASC` |

### Manejo de errores

- `close_RecordSet`: Cierra el recordset si hay error
- `close_Connection`: Cierra la conexión si hay error
- Muestra mensaje con la descripción del error

---

*Módulo: 16 - VBA SQL | Archivo: 01 - SQL Tong hop.txt*
