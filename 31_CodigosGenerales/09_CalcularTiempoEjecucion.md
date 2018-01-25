# Calcular Tiempo de Ejecución del Código

Macros y funciones para medir el tiempo de ejecución de código VBA utilizando diversas funciones de la API de Windows.

## Código VBA

```vba
Declare Function GetTickCount Lib "kernel32" () As Long
Option Explicit


Sub QUERY_NAME_SHEET()

    Dim sh As Worksheet
    Dim arr As Variant, i As Integer
    Dim T1 As Long, T2 As Long
    T1 = GetTickCount
    ReDim arr(1 To ThisWorkbook.Worksheets.Count)
    i = 1
    For Each sh In Worksheets
        arr(i) = sh.Name
        i = i + 1
    Next sh

    If i = 1 Then Exit Sub
    Sheet2.Range("H2").Resize(UBound(arr), 1).Value = Application.Transpose(arr)
    T2 = GetTickCount
    MsgBox (T2 - T1) / 1000, vbInformation, "Los segundos de ejecución son"
End Sub


'Cuando se necesita medir cambios de tiempo mínimos de hasta 1/1000 de segundo (1 milisegundo)
' se deben usar funciones API. Las funciones de medición de tiempo disponibles son:
' Now, Time, Timer (de VB/VBA)
' GetTickCount
' TimeGetTime
' QueryPerformanceCounter, QueryPerformanceFrequency


Declare Function QueryPerformanceCounter Lib "Kernel32" _
                        (X As Currency) As Boolean
Declare Function QueryPerformanceFrequency Lib "Kernel32" _
                        (X As Currency) As Boolean
Declare Function GetTickCount Lib "Kernel32" () As Long
Declare Function timeGetTime Lib "winmm.dll" () As Long


'Función Timer
Sub Test_Timer()
    Dim Loops&, T1&, T2&
    Debug.Print
    Loops = 0
    T1 = Timer
    Do
        T2 = Timer
        Loops = Loops + 1
    Loop Until T1 <> T2

    Debug.Print "Resolución mínima de Timer: "; _
                (T2 - T1); "segundo(s)"
    Debug.Print "Tomó"; Loops; "iteraciones"
End Sub


'Función GetTickCount
Sub Test_GetTickCount()
    Dim Loops&, T1&, T2&
    Debug.Print
    Loops = 0
    T1 = GetTickCount
    Do
        T2 = GetTickCount
        Loops = Loops + 1
    Loop Until T1 <> T2

    Debug.Print "Resolución mínima de GetTickCount: "; _
                (T2 - T1); "milisegundo(s)"
    Debug.Print "Tomó"; Loops; "iteraciones"
End Sub


'Función timeGetTime
Sub Test_timeGetTime()
    Dim Loops&, T1&, T2&
    Debug.Print
    Loops = 0
    T1 = timeGetTime
    Do
        T2 = timeGetTime
        Loops = Loops + 1
    Loop Until T1 <> T2

    Debug.Print "Resolución mínima de timeGetTime: "; _
                (T2 - T1); "milisegundo(s)"
    Debug.Print "Tomó"; Loops; "iteraciones"
End Sub


'Función QueryPerformanceCounter
Private Sub Test_QueryPerformanceCounter()
    Dim Loops&, T1@, T2@, Freq@, Overhead@, I&

    QueryPerformanceFrequency Freq
    QueryPerformanceCounter T1
    QueryPerformanceCounter T2
    Overhead = T2 - T1
    QueryPerformanceCounter T1

    Do
        QueryPerformanceCounter T2
        Loops = Loops + 1
    Loop Until T1 <> T2

    Debug.Print (T2 - T1 - Overhead) / Freq * 1000; "milisegundos(ms)"
    Debug.Print "Tomó"; Loops; "iteraciones"
End Sub


'Ejemplo rápido con Timer
Dim StartTime As Double

'   Iniciar conteo de tiempo
    StartTime = Timer

'   Tus líneas de código aquí
'
'

' Finalizar tu código
' Calcular y mostrar el tiempo de ejecución

    MsgBox Format(Timer - StartTime, "00.00") & " segundos."


'Otro método con QueryPerformanceCounter
Public Declare Function QueryPerformanceFrequency _
                         Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter _
                         Lib "kernel32.dll" (lpPerformanceCount As Currency) As Long

Sub CalculateTime()
    Dim Ar(1 To 20, 1 To 4) As Currency, WS As Worksheet
    Dim n As Currency, str As Currency, fin As Currency
    Dim y As Currency
    Dim i As Long, j As Long
    Application.ScreenUpdating = False
    For i = 1 To 4
        For j = 1 To 20
            Set WS = ThisWorkbook.Sheets.Add
            WS.Range("A1").Value = 1
            QueryPerformanceFrequency y
            QueryPerformanceCounter str
            Select Case i
                Case 1: Macro1
                Case 2: Macro1_Version2
                Case 3: Macro1_Version3
                Case 4: Macro1_Version4
            End Select
            QueryPerformanceCounter fin
            Application.DisplayAlerts = False
            WS.Delete
            Application.DisplayAlerts = True
            n = (fin - str)
            Ar(j, i) = CCur(Format(n, "##########.############") / y)
        Next j
    Next i
    With Range("A1").Resize(1, 4)
        .Value = Array("Macro1", "Macro2", "Macro3", "Macro4")
        .Font.Bold = True
    End With
    Range("A2").Resize(20, 4).Value = Ar
    With Range("A22").Resize(1, 4)
        .FormulaR1C1 = "=AVERAGE(R2C:R21C)"
        .Offset(1).FormulaR1C1 = "=RANK(R22C,R22C1:R22C4,1)"
        .Resize(2).Font.Bold = True
    End With
    Application.ScreenUpdating = True
End Sub
```

## Descripción

| Función API | Precisión | Descripción |
|-------------|-----------|-------------|
| `Timer` | 1 segundo | Función nativa de VB/VBA, retorna segundos desde medianoche |
| `GetTickCount` | 1 ms | Tiempo desde el inicio del sistema en milisegundos |
| `timeGetTime` | 1 ms | Similar a GetTickCount, de la librería winmm.dll |
| `QueryPerformanceCounter` | < 1 ms | Mayor precisión, usa contadores de alto rendimiento |

| Macro | Descripción |
|-------|-------------|
| `QUERY_NAME_SHEET` | Ejemplo que mide el tiempo para listar nombres de hojas |
| `Test_Timer` | Prueba la resolución mínima de la función Timer |
| `Test_GetTickCount` | Prueba la resolución mínima de GetTickCount |
| `Test_timeGetTime` | Prueba la resolución mínima de timeGetTime |
| `Test_QueryPerformanceCounter` | Prueba la resolución de QueryPerformanceCounter |
| `CalculateTime` | Benchmark que compara el rendimiento de 4 macros diferentes |

---

*Módulo: 31_CodigosGenerales | Archivo: 09_CalcularTiempoEjecucion.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
