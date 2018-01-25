# Objetos Workbook - Referencias y Métodos

Guía de diferentes formas de referenciar y trabajar con objetos Workbook en VBA.

## Código VBA

```vba
Option Explicit

' ============================================================
' DIFERENTES FORMAS DE REFERENCIAR UN WORKBOOK
' ============================================================

' 1. Libro actual (este libro donde está el código)
Sub EjemploLibroActual()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    MsgBox "Libro actual: " & wb.Name
End Sub

' 2. Libro activo (el que está en pantalla)
Sub EjemploLibroActivo()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    MsgBox "Libro activo: " & wb.Name
End Sub

' 3. Abrir un libro específico
Sub EjemploAbrirLibro()
    Dim wb As Workbook
    Set wb = Workbooks.Open("C:\Datos\MiLibro.xlsx", ReadOnly:=True)
    MsgBox "Libro abierto: " & wb.Name
    wb.Close
End Sub

' 4. Referenciar por nombre
Sub EjemploReferenciarNombre()
    Dim wb As Workbook
    Set wb = Workbooks("MiLibro.xlsx")
    MsgBox "Libro encontrado: " & wb.FullName
End Sub

' 5. Referenciar por índice
Sub EjemploReferenciarIndice()
    Dim wb As Workbook
    Set wb = Workbooks(2) ' Segundo libro abierto
    MsgBox "Libro por índice: " & wb.Name
End Sub

' 6. Crear libro nuevo
Sub EjemploCrearLibro()
    Dim wb As Workbook
    Set wb = Workbooks.Add
    MsgBox "Libro nuevo creado: " & wb.Name
End Sub


' ============================================================
' DECLARAR VARIABLES WORKBOOK
' ============================================================
Public Sub EjemploDeclaracion()
    Dim wbEntrada As Workbook
    Dim wbSalida As Workbook
    
    ' Asignar libro activo
    Set wbEntrada = ActiveWorkbook
    
    ' Abrir archivo de salida
    Set wbSalida = Workbooks.Open("C:\Datos\Salida.xlsx")
    
    ' Trabajar con ambos libros...
    
    ' Liberar memoria
    Set wbSalida = Nothing
    Set wbEntrada = Nothing
End Sub


' ============================================================
' GUARDAR WORKBOOK
' ============================================================
Public Sub GuardarLibro()
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    ' Guardar con nombre específico
    wb.SaveAs fileName:=ThisWorkbook.Path & "\NuevoArchivo.xlsx"
    wb.Close
End Sub

Public Sub GuardarLibroActivo()
    ActiveWorkbook.Save
End Sub


' ============================================================
' GUARDAR COPIA DEL WORKBOOK
' ============================================================
Public Sub GuardarCopiaLibro()
    Dim rutaCopia As String
    
    rutaCopia = ThisWorkbook.Path & "\Copia_" & Format(Now, "YYYYMMDD") & "_" & ThisWorkbook.Name
    
    ThisWorkbook.SaveCopyAs rutaCopia
    
    MsgBox "Copia guardada: " & rutaCopia, vbInformation
End Sub


' ============================================================
' EJEMPLO PRÁCTICO: Procesar datos entre libros
' ============================================================
Public Sub ProcesarEntreLibros()
    Dim wbOrigen As Workbook
    Dim wbDestino As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    
    ' Abrir libro origen
    Set wbOrigen = Workbooks.Open("C:\Datos\Origen.xlsx", ReadOnly:=True)
    Set wsOrigen = wbOrigen.Sheets(1)
    
    ' Crear o abrir libro destino
    Set wbDestino = Workbooks.Add
    Set wsDestino = wbDestino.Sheets(1)
    
    ' Copiar datos
    wsOrigen.UsedRange.Copy wsDestino.Range("A1")
    
    ' Guardar y cerrar
    wbDestino.SaveAs ThisWorkbook.Path & "\Resultado.xlsx"
    wbDestino.Close
    wbOrigen.Close SaveChanges:=False
    
    MsgBox "Proceso completado.", vbInformation
End Sub
```

## Resumen de Referencias

| Referencia | Descripción |
|------------|-------------|
| `ThisWorkbook` | Libro que contiene el código VBA |
| `ActiveWorkbook` | Libro actualmente activo/visible |
| `Workbooks("nombre")` | Libro por nombre |
| `Workbooks(indice)` | Libro por posición |
| `Workbooks.Open(ruta)` | Abrir un libro |
| `Workbooks.Add` | Crear libro nuevo |

## Métodos de Workbook

| Método | Descripción |
|--------|-------------|
| `.Save` | Guardar cambios |
| `.SaveAs` | Guardar como nuevo archivo |
| `.SaveCopyAs` | Guardar copia sin cambiar nombre |
| `.Close` | Cerrar libro |
| `.Activate` | Activar el libro |

---

*Módulo: 05_GestionWorkbooks | Archivo: 02_ObjetosWorkbook.md*
