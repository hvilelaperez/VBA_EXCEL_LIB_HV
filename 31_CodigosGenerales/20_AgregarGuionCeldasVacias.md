# Agregar Guion en Celdas Vacías

Marca y colorea celdas que contienen equipos de construcción necesarios.

## Código VBA

```vba
Sub EscribirEquiposNecesarios()
 Dim Rng As Range, sRng As Range, Cls As Range
 Dim MiDireccion As String
 Dim Filas As Long, W As Integer
 
 With Sheets("Diary").[I1]
    Filas = .CurrentRegion.Rows.Count
    Set Rng = .Resize(Filas)
 End With
 Sheets("May Thi Cong").Select
 For Each Cls In Range([B2], [B2].End(xlDown))
    Set sRng = Rng.Find(Cls.Value, , xlFormulas, xlPart)
    If Not sRng Is Nothing Then
        MiDireccion = sRng.Address
        W = W + 1:                  If W = 12 Then W = 0
        Do
            sRng.Interior.ColorIndex = 34 + W
            Sheets("Diary").Cells(sRng.Row, "T").Value = Cls.Offset(, 1).Value
            Set sRng = Rng.FindNext(sRng)
        Loop While Not sRng Is Nothing And sRng.Address <> MiDireccion
    End If
 Next Cls
End Sub

Sub EscribirEquiposNecesarios()
 Dim Rng As Range, sRng As Range, Cls As Range
 Dim MiDireccion As String
 Dim Filas As Long, W As Integer
 
 With Sheets("Diary").[I1]
    Filas = .CurrentRegion.Rows.Count
    Set Rng = .Resize(Filas)
 End With
 Sheets("May Thi Cong").Select
 For Each Cls In Range([B2], [B2].End(xlDown))
    Set sRng = Rng.Find(Cls.Value, , xlFormulas, xlPart)
    If Not sRng Is Nothing Then
        MiDireccion = sRng.Address
        W = W + 1:                  If W = 12 Then W = 0
        Do
            sRng.Interior.ColorIndex = 34 + W
            Sheets("Diary").Cells(sRng.Row, "T").Value = Cls.Offset(, 1).Value
            Set sRng = Rng.FindNext(sRng)
        Loop While Not sRng Is Nothing And sRng.Address <> MiDireccion
    End If
 Next Cls
End Sub


Public Sub HerramientasConstruccion()
Dim Dic As Object, sArr(), dArr(), tArr(), Tmp, Txt As String
Dim I As Long, J As Long, N As Long, R1 As Long, R2 As Long
Set Dic = CreateObject("Scripting.Dictionary")
tArr = Sheets("May thi cong").Range("B2", Sheets("May thi cong").Range("B2").End(xlDown)).Resize(, 2).Value
R2 = UBound(tArr)
With Sheets("DIARY")
    sArr = .Range("I2", .Range("I60000").End(xlUp)).Value
    R1 = UBound(sArr)
    ReDim dArr(1 To R1, 1 To 1)
    For I = 1 To R1
        Dic.RemoveAll
        Tmp = Split(sArr(I, 1), ChrW(10))
        For J = LBound(Tmp) To UBound(Tmp)
            Txt = Tmp(J)
            For N = 1 To R2
                If Txt Like tArr(N, 1) & "*" Then
                    If Not Dic.Exists(tArr(N, 2)) Then
                        Dic.Item(tArr(N, 2)) = ""
                        dArr(I, 1) = dArr(I, 1) & "; " & tArr(N, 2)
                    End If
                End If
            Next N
        Next J
    Next I
    For I = 1 To R1
        If Len(dArr(I, 1)) Then
            dArr(I, 1) = Mid(dArr(I, 1), 3)
        End If
    Next I
    .Range("T2").Resize(R1) = dArr
End With
Set Dic = Nothing
End Sub


Sub BuscarEquiposConstruccion()
Dim sh As Worksheet, Equipos(), sArr(), i As Long, ii As Long, Res()
With Sheets("May Thi Cong")
   Equipos = .Range("A2", .[A65536].End(3)).Resize(, 3).Value
End With
For Each sh In ThisWorkbook.Worksheets
   If Replace(LCase(sh.Name), " ", "") <> "maythicong" Then
      sArr = sh.Range("A2", sh.[A65536].End(3)).Resize(, 9).Value
      ReDim Res(1 To UBound(sArr), 1 To 1)
      For i = 1 To UBound(sArr)
         If sArr(i, 9) <> Empty Then
            For ii = 1 To UBound(Equipos)
               If sArr(i, 9) Like "*" & Equipos(ii, 2) & "*" Then
                  Res(i, 1) = Equipos(ii, 3)
               End If
            Next
         End If
      Next
      sh.[T2].Resize(UBound(Res), UBound(Res, 2)) = Res
   End If
Next
End Sub

Sub Ocultar()
Union([E1], [F1], [K1:L1]).EntireColumn.Hidden = True
End Sub
Sub Mostrar()
Union([E1], [F1], [K1], [L1]).EntireColumn.Hidden = False
End Sub
```

## Descripción

| Sub | Función |
|-----|---------|
| `EscribirEquiposNecesarios` | Busca y colorea equipos de construcción en la hoja Diary |
| `HerramientasConstruccion` | Usa diccionario para identificar herramientas únicas |
| `BuscarEquiposConstruccion` | Busca equipos en todas las hojas del libro |
| `Ocultar` / `Mostrar` | Oculta/muestra columnas específicas |

---

*Módulo: 31_CodigosGenerales | Archivo: 20_AgregarGuionCeldasVacias.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
