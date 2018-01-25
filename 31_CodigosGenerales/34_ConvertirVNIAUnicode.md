# Convertir VNI a Unicode

Función que convierte cadenas de texto escritas en codificación VNI a formato Unicode estándar.

## Código VBA

```vba
' Autor: Henry Vilela - hvilelaperez@gmail.com
' Autor original: phantronghiep07 (0915 080 282)
Function GoVni2Uni(CadenaVNI As String) As String
'---------------------------------------------------------------------------------------
' Función: GoVni2Uni
' Convierte cadena escrita en VNI a cadena Vietnamita Unicode
'---------------------------------------------------------------------------------------
    Dim i As Integer
    Dim MaAcii, VNI

    ' Tabla de códigos Unicode correspondientes
    MaAcii = Array(7845, 7847, 7849, 7851, 7853, 226, 225, 224, 7843, 227, 7841, 7855, 7857, 7859, _
                   7861, 7863, 259, 250, 249, 7911, 361, 7909, 7913, 7915, 7917, 7919, 7921, 432, _
                   7871, 7873, 7875, 7877, 7879, 234, 233, 232, 7867, 7869, 7865, 7889, 7891, 7893, _
                   7895, 7897, 244, 243, 242, 7887, 245, 7885, 7899, 7901, 7903, 7905, 7907, 417, _
                   237, 236, 7881, 297, 7883, 253, 7923, 7927, 7929, 7925, 273, 7844, 7846, 7848, _
                   7850, 7852, 194, 193, 192, 7842, 195, 7840, 7854, 7856, 7858, 7860, 7862, 258, _
                   218, 217, 7910, 360, 7908, 7912, 7914, 7916, 7918, 7920, 431, 7870, 7872, 7874, _
                   7876, 7878, 202, 201, 200, 7866, 7868, 7864, 7888, 7890, 7892, 7894, 7896, 212, _
                   211, 210, 7886, 213, 7884, 7898, 7900, 7902, 7904, 7906, 416, 205, 204, 7880, 296, _
                   7882, 221, 7922, 7926, 7928, 7924, 272)

    ' Tabla de códigos VNI equivalentes
    VNI = Array("a61", "a62", "a63", "a64", "a65", "a6", "a1", "a2", "a3", "a4", "a5", "a81", "a82", _
                "a83", "a84", "a85", "a8", "u1", "u2", "u3", "u4", "u5", "u71", "u72", "u73", "u74", _
                "u75", "u7", "e61", "e62", "e63", "e64", "e65", "e6", "e1", "e2", "e3", "e4", "e5", _
                "o61", "o62", "o63", "o64", "o65", "o6", "o1", "o2", "o3", "o4", "o5", "o71", "o72", _
                "o73", "o74", "o75", "o7", "i1", "i2", "i3", "i4", "i5", "y1", "y2", "y3", "y4", "y5", _
                "d9", "A61", "A62", "A63", "A64", "A65", "A6", "A1", "A2", "A3", "A4", "A5", _
                "A81", "A82", "A83", "A84", "A85", "A8", "U1", "U2", "U3", "U4", "U5", "U71", _
                "U72", "U73", "U74", "U75", "U7", "E61", "E62", "E63", "E64", "E65", "E6", "E1", _
                "E2", "E3", "E4", "E5", "O61", "O62", "O63", "O64", "O65", "O6", "O1", "O2", _
                "O3", "O4", "O5", "O71", "O72", "O73", "O74", "O75", "O7", "I1", "I2", "I3", "I4", _
                "I5", "Y1", "Y2", "Y3", "Y4", "Y5", "D9")

    ' Realizar las 134 sustituciones de VNI a Unicode
    GoVni2Uni = CadenaVNI
    For i = 0 To 133
        GoVni2Uni = Replace(GoVni2Uni, VNI(i), ChrW(MaAcii(i)))
    Next i
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Función** | `GoVni2Uni` - Convierte cadena VNI a Unicode |
| **Parámetro** | `CadenaVNI` - Texto en codificación VNI a convertir |
| **Retorno** | String con texto en formato Unicode |
| **Codificación VNI** | Sistema de codificación vietnamita que usa combinaciones de caracteres (ej: `a1` = á) |
| **Proceso** | Realiza 134 reemplazos usando tablas de mapeo VNI → Unicode |
| **Uso** | `=GoVni2Uni(A1)` para convertir celdas con texto VNI a Unicode |

---

*Módulo: 31_CodigosGenerales | Archivo: 45 - VNI to Unicode.txt*
