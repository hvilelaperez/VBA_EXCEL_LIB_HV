# Mostrar Hora en Formulario (Splash Screen)

Macro para ocultar automáticamente un formulario de bienvenida (Splash Screen) después de un tiempo determinado.

## Código VBA

```vba
Private Sub UserForm_Initialize()
    Application.OnTime Now + TimeValue("00:00:03"), "DetenerLogo"
End Sub

Sub DetenerLogo()
    LogoSplashScreen.Hide
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `UserForm_Initialize` | Evento que se ejecuta al inicializar el formulario |
| `Application.OnTime` | Programa la ejecución de una macro en un momento específico |
| `Now + TimeValue("00:00:03")` | Fecha y hora actual + 3 segundos |
| `DetenerLogo` | Macro que oculta el Splash Screen |
| `LogoSplashScreen.Hide` | Oculta el formulario de bienvenida |

### Flujo de ejecución

1. Al abrir el formulario, `UserForm_Initialize` se ejecuta automáticamente
2. Se programa la macro `DetenerLogo` para ejecutarse 3 segundos después
3. Después de 3 segundos, `DetenerLogo` oculta el Splash Screen

### Uso común

Este patrón se utiliza para mostrar una pantalla de bienvenida o logotipo durante unos segundos antes de mostrar la interfaz principal de la aplicación.

---

*Módulo: 96 - VBA Codes general | Archivo: 16 - Dat time hine len cho form.txt*
