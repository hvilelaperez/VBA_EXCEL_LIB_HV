# DataObject con Late Binding (MS Forms)

Uso del objeto `DataObject` de Microsoft Forms con late binding para manipular el portapapeles y aplicar expresiones regulares.

## Código VBA

```vba
CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

ejemplo:
With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        rng.Copy
        .getFromClipboard
        Application.CutCopyMode = False
        strCopia = .GetText
        .Clear
    strPatron = "[^0-9A-Za-z\r\n]" 'o = "[^a-zA-Z_0-9]" ' o = "[^0-9A-Za-z\r\n]+" ' porque "[^\w\r\n]+" también coincide con _ y letras Unicode
        With RegEx
            .Pattern = strPatron
            .Global = True
            strCopia = .Replace(strCopia, vbNullString)
        End With
        .SetText strCopia
        .PutInClipBoard
    End With
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `CreateObject("New:{1C3B4210-...}")` | Crea una instancia del objeto DataObject de MS Forms v2.0 con late binding |
| `rng.Copy` | Copia el rango al portapapeles |
| `.getFromClipboard` | Obtiene el contenido del portapapeles al DataObject |
| `.GetText` | Extrae el texto del DataObject como cadena |
| `.Clear` | Limpia el contenido del DataObject |
| `.SetText` | Establece el texto modificado en el DataObject |
| `.PutInClipBoard` | Devuelve el texto al portapapeles |

### Flujo de ejecución

1. Copia un rango al portapapeles
2. Obtiene el contenido del portapapeles al DataObject
3. Extrae el texto como cadena
4. Aplica una expresión regular para eliminar caracteres no deseados
5. Coloca el texto limpio de vuelta en el portapapeles

### Patrones de expresión regular

| Patrón | Descripción |
|--------|-------------|
| `[^0-9A-Za-z\r\n]` | Elimina todo excepto letras, números y saltos de línea |
| `[^a-zA-Z_0-9]` | Elimina todo excepto letras, números y guiones bajos |
| `[^0-9A-Za-z\r\n]+` | Igual que el primero pero con una o más ocurrencias |

### Nota sobre Late Binding

> El GUID `{1C3B4210-F441-11CE-B9EA-00AA006B1A69}` identifica el objeto `Microsoft Forms 2.0 DataObject`. El late binding permite usar el objeto sin agregar una referencia explícita al proyecto.

---

*Módulo: 96 - VBA Codes general | Archivo: 30 - MS form Data object late binding.txt*
