# Detectar Idioma de Excel

Función para determinar el idioma del sistema configurado en Microsoft Excel.

## Código VBA

```vba
Public Function IdiomaSistema()
'Fuente: https://excel-malin.com

On Error GoTo ErrorFuncion

Dim CodigoIdiomaSistema As Single
'Encontrar el código del idioma
CodigoIdiomaSistema = Application.LanguageSettings.LanguageID(msoLanguageIDUI)

'Asociar el código con el idioma (idiomas principales, los demás se reemplazan por inglés)
Select Case CodigoIdiomaSistema
Case 1036, 2060, 11276, 3084, 9228, 12300, 15372, 5132, 13324, 6156, 14348, 58380, 8204, 10252, 4108, 7180: IdiomaSistema = "Francés"
Case 1033, 2057, 3081, 10249, 4105, 9225, 15369, 16393, 14345, 6153, 8201, 17417, 5129, 13321, 18441, 7177, 11273, 12297: IdiomaSistema = "Inglés"
Case 1031, 3079, 5127, 4103, 2055: IdiomaSistema = "Alemán"
Case 2052, 4100, 1028, 3076, 5124: IdiomaSistema = "Chino"
Case 1043, 2067: IdiomaSistema = "Holandés"
Case 1040, 2064: IdiomaSistema = "Italiano"
Case 3082, 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 22538, 2058, 19466, 6154, 15370, 10250, 20490, 21514, 14346, 8202: IdiomaSistema = "Español"
Case 1025, 5121, 15361, 3073, 2049, 11265, 13313, 12289, 4097, 6145, 8193, 16385, 10241, 7169, 14337, 9217: IdiomaSistema = "Árabe"
Case Else: IdiomaSistema = "Otro"
End Select

Exit Function

ErrorFuncion:
IdiomaSistema = ""
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `IdiomaSistema` | Función pública que retorna una cadena con el nombre del idioma |
| `msoLanguageIDUI` | Constante que identifica el idioma de la interfaz de usuario |
| `Application.LanguageSettings.LanguageID` | Obtiene el código numérico del idioma del sistema |

### Idiomas soportados

| Código | Idioma |
|--------|--------|
| 1036, 2060, etc. | Francés |
| 1033, 2057, etc. | Inglés |
| 1031, 3079, etc. | Alemán |
| 2052, 4100, etc. | Chino |
| 1043, 2067 | Holandés |
| 1040, 2064 | Italiano |
| 3082, 1034, etc. | Español |
| 1025, 5121, etc. | Árabe |
| Otros | Otro |

### Manejo de errores

Si ocurre un error al obtener el código de idioma, la función retorna una cadena vacía `""`.

---

*Módulo: 96 - VBA Codes general | Archivo: 26 - Xac dinh vesion ngon ngu su dung trong Excel.txt*
