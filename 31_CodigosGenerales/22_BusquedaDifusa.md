# Búsqueda Difusa (Fuzzy Search)

Implementa búsqueda aproximada de personas usando distancia de Levenshtein y similitud de cadenas.

## Código VBA

```vba
Option Explicit

Private Type ConsultaPersona
    nombre As String
    apellido As String
    nombreCompleto As String
    nombresRango As Range
    apellidosRango As Range
    nombresCompletosRango As Range
    nombresIntercambiados As Boolean
End Type

Private Type ResultadoPersona
    encontrado As Boolean
    fila As Integer
    metodoBusqueda As String
End Type

Const INTEGER_MIN As Integer = -32768
Private NO_ENCONTRADO As ResultadoPersona

Const UMBRAL_NO_COINCIDENCIA_DIST As Double = 0.5
Const UMBRAL_NO_COINCIDENCIA_RATIO_LEN As Double = 0.25
Const UMBRAL_NO_COINCIDENCIA_UBICACION As Double = 0.25
Const UMBRAL_NO_COINCIDENCIA_SIM_NO_ORDENADA As Double = 0.3

Const ASC_A As Integer = 65
Const NUM_LETRAS As Integer = 26
Const NUM_LETRAS_ESPECIALES As Integer = 27
Const VAL_NO_LETRA As Integer = 26

' Busca una persona por nombre y apellido
' Retorna un arreglo con (id de la persona, nombre completo en BD, método de coincidencia)
' Si no se encuentra, id y nombre serán "N/A"
' Los métodos utilizados son:
'  - buscar nombre completo exacto,
'  - intercambiar nombre y apellido y buscar (a veces se confunden)
'  - filtrar por apellido exacto y buscar similar en nombre (y viceversa)
'  - luego buscar por similitud en nombre completo

Private Function BuscarIDPersona(nombre As String, apellido As String, listaNombres As Range, _
    listaApellidos As Range, listaCompletos As Range, ids As Range)
    
    Dim pr As ResultadoPersona
    Dim pq As ConsultaPersona
    pq.nombre = nombre
    pq.apellido = apellido
    pq.nombreCompleto = nombre + " " + apellido
    Set pq.nombresRango = listaNombres
    Set pq.apellidosRango = listaApellidos
    Set pq.nombresCompletosRango = listaCompletos
    pq.nombresIntercambiados = False
    
    pr = BuscarPersonaInterno(pq)
    
    Dim nombreCom As String
    Dim id As String
    If pr.encontrado Then
        id = ids.Cells(pr.fila, 1).Value
        nombreCom = listaCompletos.Cells(pr.fila, 1).Value
    Else
        id = "N/A"
        nombreCom = "N/A"
    End If
    
    BuscarIDPersona = Array(id, nombreCom, pr.metodoBusqueda
End Function

Private Function BuscarPersonaInterno(pq As ConsultaPersona) As ResultadoPersona
    Dim pr As ResultadoPersona
    
    If Trim(pq.nombreCompleto) = "" Then
        GoTo NoEncontrado
    End If
    
    pr = BuscarExacto(pq)
    If pr.encontrado Then
        pr.metodoBusqueda = "COMPLETO_EXACTO"
        BuscarPersonaInterno = pr
        Exit Function
    End If
    
    pr = BuscarExacto(IntercambiarNombreApellido(pq))
    If pr.encontrado Then
        pr.metodoBusqueda = "COMPLETO_EXACTO_INVERTIDO"
        BuscarPersonaInterno = pr
        Exit Function
    End If
    
    pr = BuscarApellidoExactoNombreSimilar(pq)
    If pr.encontrado Then
        pr.metodoBusqueda = "APELLIDO_EXACTO_NOMBRE_SIMILAR"
        BuscarPersonaInterno = pr
        Exit Function
    End If
    
    pr = BuscarApellidoExactoNombreSimilar(IntercambiarNombreApellido(IntercambiarListasNombreApellido(pq)))
    If pr.encontrado Then
        pr.metodoBusqueda = "NOMBRE_EXACTO_APELLIDO_SIMILAR"
        BuscarPersonaInterno = pr
        Exit Function
    End If
    
    pr = BuscarApellidoExactoNombreSimilar(IntercambiarNombreApellido(pq))
    If pr.encontrado Then
        pr.metodoBusqueda = "APELLIDO_EXACTO_INVERTIDO_NOMBRE_SIMILAR_INVERTIDO"
        BuscarPersonaInterno = pr
        Exit Function
    End If
    
    pr = BuscarApellidoExactoNombreSimilar(IntercambiarListasNombreApellido(pq))
    If pr.encontrado Then
        pr.metodoBusqueda = "NOMBRE_EXACTO_INVERTIDO_APELLIDO_SIMILAR_INVERTIDO"
        BuscarPersonaInterno = pr
        Exit Function
    End If
            
    pr = BuscarSimilarCompleto(pq)
    If pr.encontrado Then
        pr.metodoBusqueda = "COMPLETO_SIMILAR"
        BuscarPersonaInterno = pr
        Exit Function
    End If
    
    pr = BuscarSimilarCompleto(IntercambiarNombreApellido(pq))
    If pr.encontrado Then
        pr.metodoBusqueda = "COMPLETO_SIMILAR_INVERTIDO"
        BuscarPersonaInterno = pr
        Exit Function
    End If
    
NoEncontrado:
    pr.metodoBusqueda = "NO_ENCONTRADO"
    BuscarPersonaInterno = pr
End Function

Private Function BuscarExacto(pq As ConsultaPersona) As ResultadoPersona
    Dim rangoEncontrado As Range
    Dim pr As ResultadoPersona
    Set rangoEncontrado = pq.nombresCompletosRango.Find(pq.nombreCompleto, LookIn:=xlValues, LookAt:=xlWhole)
    pr.encontrado = Not rangoEncontrado Is Nothing
    
    If pr.encontrado Then
        pr.fila = rangoEncontrado.Row
    End If
    
    BuscarExacto = pr
End Function

Private Function BuscarApellidoExactoNombreSimilar(pq As ConsultaPersona) As ResultadoPersona
    Dim pr As ResultadoPersona
    pr.encontrado = False
    
    Dim filasApellido As Collection
    Set filasApellido = BuscarFilasClave(pq.apellido, pq.apellidosRango)
    If filasApellido.Count > 0 Then
        Dim nombresParaApellido As Collection
        Set nombresParaApellido = BuscarValoresFilas(filasApellido, pq.nombresRango)
        
        Dim masSimilar As Variant
        masSimilar = IndiceMasSimilarColeccion(pq.nombre, nombresParaApellido)
        
        If CDbl(masSimilar(1)) / CDbl(Len(pq.nombre)) > UMBRAL_NO_COINCIDENCIA_DIST Then
            pr.encontrado = True
            pr.fila = filasApellido(masSimilar(0))
        End If
    End If
    
    BuscarApellidoExactoNombreSimilar = pr
End Function

Private Function BuscarSimilarCompleto(pq As ConsultaPersona) As ResultadoPersona
    Dim pr As ResultadoPersona
    
    Dim filaYsimilitud As Variant
    filaYsimilitud = CoincidenciaSimilar(pq.nombreCompleto, pq.nombresCompletosRango)
    
    pr.fila = filaYsimilitud(0)
    pr.encontrado = (filaYsimilitud(1) / CDbl(Len(pq.nombreCompleto))) > UMBRAL_NO_COINCIDENCIA_DIST
    
    BuscarSimilarCompleto = pr
End Function

Private Function IntercambiarNombreApellido(pq As ConsultaPersona) As ConsultaPersona
    Dim pqIntercambiado As ConsultaPersona
    pqIntercambiado = pq
    pqIntercambiado.nombre = pq.apellido
    pqIntercambiado.apellido = pq.nombre
    pqIntercambiado.nombreCompleto = pqIntercambiado.nombre + " " + pqIntercambiado.apellido
    pqIntercambiado.nombresIntercambiados = Not pq.nombresIntercambiados
    IntercambiarNombreApellido = pqIntercambiado
End Function

Private Function IntercambiarListasNombreApellido(pq As ConsultaPersona) As ConsultaPersona
    Dim pqListasIntercambiadas As ConsultaPersona
    pqListasIntercambiadas = pq
    Set pqListasIntercambiadas.nombresRango = pq.apellidosRango
    Set pqListasIntercambiadas.apellidosRango = pq.nombresRango
    pqListasIntercambiadas.nombresIntercambiados = Not pq.nombresIntercambiados
    IntercambiarListasNombreApellido = pqListasIntercambiadas
End Function

Function IndiceMasSimilarColeccion(cadena As String, col As Collection)
    Dim mejorSimilitud As Variant
    Dim masSimilarI As Integer
    Dim similitudActual As Integer
    mejorSimilitud = INTEGER_MIN
    masSimilarI = -1
    
    Dim i As Integer
    For i = 1 To col.Count
        similitudActual = SimilitudLevenshtein(cadena, col.Item(i))
        If similitudActual > mejorSimilitud Then
            masSimilarI = i
            mejorSimilitud = similitudActual
        End If
    Next
    
    IndiceMasSimilarColeccion = Array(masSimilarI, mejorSimilitud)
End Function

Private Function BuscarValoresFilas(filas As Collection, rng As Range) As Collection
    Dim i As Integer
    Dim vals As New Collection
    For i = 1 To filas.Count
        vals.Add (rng.Cells(filas(i), 1).Value)
    Next
    Set BuscarValoresFilas = vals
End Function

Private Function BuscarFilasClave(clave As Variant, claves As Range) As Collection
    Dim filas As New Collection
    
    Dim rangoEncontrado As Range
    Set rangoEncontrado = claves.Find(clave, LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim primeraDireccion As Variant
    If Not rangoEncontrado Is Nothing Then
        primeraDireccion = rangoEncontrado.Address
        Do
            filas.Add (rangoEncontrado.Row)
            Set rangoEncontrado = claves.Find(clave, rangoEncontrado, LookIn:=xlValues, LookAt:=xlWhole)
        Loop While Not rangoEncontrado Is Nothing And rangoEncontrado.Address <> primeraDireccion
    End If
    
    Set BuscarFilasClave = filas
End Function

Function CoincidenciaSimilar(busqueda As Variant, lista As Range) As Variant
    Dim maxSim As Integer
    maxSim = 0
    
    Dim i As Integer
    Dim maxSimI As Integer
    maxSimI = -1
    
    Dim ultimaFila As Integer
    ultimaFila = lista.Cells.End(xlDown).Row
    
    Dim primeraFila As Integer
    primeraFila = lista.Cells.End(xlUp).Row
    
    Dim cadenaBusqueda As String
    
    If TypeOf busqueda Is Range Then
        cadenaBusqueda = busqueda.Value
    Else
        cadenaBusqueda = busqueda
    End If
    
    Dim valorEnLista As String
    Dim sim As Integer
    For i = primeraFila To ultimaFila
        valorEnLista = lista.Cells(i, 1).Value
           
        'Para optimizar, verificar longitud, ubicación de espacio y similitud no ordenada
        'antes de ejecutar el algoritmo de distancia de edición
        If Abs(Len(valorEnLista) - Len(cadenaBusqueda)) / Len(cadenaBusqueda) _
              < UMBRAL_NO_COINCIDENCIA_RATIO_LEN Then
            If Abs(InStr(cadenaBusqueda, " ") - InStr(valorEnLista, " ")) / Len(cadenaBusqueda) _
                  < UMBRAL_NO_COINCIDENCIA_UBICACION Then
                If SimilitudNoOrdenada(cadenaBusqueda, valorEnLista) / Len(cadenaBusqueda) _
                      > UMBRAL_NO_COINCIDENCIA_SIM_NO_ORDENADA Then
                    sim = SimilitudLevenshtein(cadenaBusqueda, valorEnLista)
                    If sim > maxSim Then
                        maxSim = sim
                        maxSimI = i
                    End If
                End If
            End If
        End If
    Next
        
    CoincidenciaSimilar = Array(maxSimI, maxSim)
End Function

Private Function ValorLetra(letra As String) As Integer
    Dim ascVal As Integer
    ascVal = Asc(letra)
    
    If ascVal < ASC_A Or ascVal > ASC_A + NUM_LETRAS Then
        ValorLetra = VAL_NO_LETRA
    Else
        ValorLetra = ascVal - ASC_A
    End If
End Function

Private Function SimilitudNoOrdenada(ByVal str1 As String, ByVal str2 As String) As Integer
    Dim i, j, ascVal As Integer
    Dim ascA As Integer
    Dim ascZ As Integer
    Dim conteoDiferenciasLetras(27) As Integer
    Dim diferenciasLetras As Integer
    Dim valorLetra As Integer
    Dim len1 As Integer
    Dim len2 As Integer
    
    len1 = Len(str1)
    len2 = Len(str2)
    str1 = UCase(str1)
    str2 = UCase(str2)
    
    i = 0
    While i < len1
        valorLetra = ValorLetra(CaracterEn(str1, i))
        conteoDiferenciasLetras(valorLetra) = conteoDiferenciasLetras(valorLetra) + 1
        i = i + 1
    Wend
    
    i = 0
    While i < len2
        valorLetra = ValorLetra(CaracterEn(str2, i))
        conteoDiferenciasLetras(valorLetra) = conteoDiferenciasLetras(valorLetra) - 1
        i = i + 1
    Wend
    
    diferenciasLetras = 0
    For i = 0 To UBound(conteoDiferenciasLetras)
        diferenciasLetras = diferenciasLetras + Abs(conteoDiferenciasLetras(i))
    Next
    
    Dim minLen As Integer
    If len1 < len2 Then minLen = len1 Else minLen = len2
        
    SimilitudNoOrdenada = minLen - diferenciasLetras
End Function

'--------------------------------------------------------------------
' Calcula la distancia de edición entre str1 y str2 usando el
' algoritmo de programación dinámica de distancia Levenshtein
' Esto es realmente "similitud de edición" ya que cadenas más similares
' tienen una puntuación mayor.
Private Function SimilitudLevenshtein(str1 As String, str2 As String)
    Dim len1, len2, i, j, puntuacion, simCar, gap1, gap2, valorCoincidencia As Integer
    
    len1 = Len(str1)
    len2 = Len(str2)
    
    Dim puntuacion_gap As Integer
    puntuacion_gap = -1

    Dim D As Variant
    ReDim D(0 To (len1 + 1), 0 To (len2 + 1)) As Integer
    D(0, 0) = 0
 
    For i = 0 To len1
        D(i, 0) = puntuacion_gap * i
    Next
    
    For j = 0 To len2
        D(0, j) = puntuacion_gap * j
    Next
    
    For i = 1 To len1
        For j = 1 To len2
            valorCoincidencia = D(i - 1, j - 1) + SimilitudCaracteres(CaracterEn(str1, i - 1), CaracterEn(str2, j - 1))
            gap2 = D(i, j - 1) + puntuacion_gap
            gap1 = D(i - 1, j) + puntuacion_gap
            D(i, j) = Application.WorksheetFunction.Max(valorCoincidencia, gap2, gap1)
        Next
    Next
    
    i = len1
    j = len2
    puntuacion = 0
    
    Dim alinear As String
    
    While i > 0 And j > 0
        simCar = SimilitudCaracteres(CaracterEn(str1, i - 1), CaracterEn(str2, j - 1))
        If D(i, j) - simCar = D(i - 1, j - 1) Then
            If simCar > 0 Then
                'alinear = "M"
            Else
                'alinear = "C"
            End If
            
            i = i - 1
            j = j - 1
            
            puntuacion = puntuacion + simCar
        ElseIf D(i, j) - puntuacion_gap = D(i, j - 1) Then
            'alinear = "A"
            j = j - 1
        ElseIf D(i, j) - puntuacion_gap = D(i - 1, j) Then
            'alinear = "D"
            i = i - 1
            puntuacion = puntuacion + puntuacion_gap
        Else
            MsgBox "Puntuación inesperada en retroceso"
        End If
    Wend
    
    While j > 0
        j = j - 1
        puntuacion = puntuacion + puntuacion_gap
    Wend
    
    While i > 0
        i = i - 1
        puntuacion = puntuacion + puntuacion_gap
    Wend

    SimilitudLevenshtein = puntuacion
End Function

Function CaracterEn(str, indiceInicioCero)
    CaracterEn = Mid(str, indiceInicioCero + 1, 1)
End Function

Function SimilitudCaracteres(chr1 As Variant, chr2 As Variant)
    If UCase(chr1) = UCase(chr2) Then
        SimilitudCaracteres = 1
    Else
        SimilitudCaracteres = -1
    End If
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `ConsultaPersona` | Tipo de dato para almacenar datos de búsqueda |
| `ResultadoPersona` | Tipo de dato para almacenar resultados |
| `SimilitudLevenshtein` | Calcula distancia de edición entre cadenas |
| `SimilitudNoOrdenada` | Compara letras sin importar el orden |
| `BuscarPersonaInterno` | Orquesta la búsqueda con múltiples métodos |
| UMBRALES | Configuran sensibilidad de la búsqueda |

---

*Módulo: 31_CodigosGenerales | Archivo: 22_BusquedaDifusa.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
