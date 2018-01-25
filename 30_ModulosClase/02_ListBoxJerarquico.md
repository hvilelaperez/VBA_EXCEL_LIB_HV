# ListBox Jerárquico - Segunda Parte

Continuación de la clase `ajpHList` con las propiedades de configuración, manejo de eventos y ejemplo de uso en un UserForm.

## Código VBA

```vba
Private Sub m_BuildNodeList()
'
' Llenar ListBox con nodos válidos
'
    Dim intNodeIndex As Integer
    Dim intItemIndex As Integer
    Dim strStatus As String
    Dim strBuf As String
    
    m_objListbox.Clear
    For intNodeIndex = 0 To m_intNNodes - 1
        If m_intNodeLevel(intNodeIndex) <= 0 Then
            m_objListbox.AddItem m_MakeNode(intNodeIndex)
            intItemIndex = m_objListbox.ListCount - 1
            m_objListbox.List(intItemIndex, 1) = CStr(intNodeIndex)
        End If
    Next
End Sub

Public Property Get DisplayListCount() As Integer
' número de elementos mostrados actualmente
    DisplayListCount = m_objListbox.ListCount - 1
End Property
Public Property Let NodeClosedText(Text As String)
' usuario establece texto de Nodo Cerrado
    m_strNode_Closed = Text
    
    m_BuildDropNode
End Property
Public Property Get NodeClosedText() As String
    NodeClosedText = m_strNode_Closed
End Property
Public Property Get NodeOpenText() As String
    NodeOpenText = m_strNode_Open
End Property
Public Property Get NodeEndText() As String
    NodeEndText = m_strNode_End
End Property
Public Property Get NodeDropText() As String
    NodeDropText = m_strNode_Drop
End Property
Public Property Let NodeOpenText(Text As String)
' usuario establece texto de Nodo Abierto
    m_strNode_Open = Text
    m_BuildDropNode
End Property
Public Property Let NodeDropText(Text As String)
' usuario establece texto de Nodo de Desplazamiento
    m_strNode_Drop = Text
    
    m_BuildDropNode
End Property
Public Property Let NodeEndText(Text As String)
' usuario establece texto de Nodo Final
    m_strNode_End = Text
    m_BuildDropNode

End Property
Public Property Get NodeCount() As Integer
' conteo de todos los elementos
    NodeCount = m_intNNodes - 2
End Property
Public Sub Refresh()
    m_CreateNodes
    m_BuildNodeList
End Sub

Public Property Set UseListbox(LBox As MSForms.ListBox)
    Set m_objListbox = LBox
End Property
Public Property Get Version() As String
    Version = HLIST_VERSION
End Property

Private Sub Class_Initialize()

    Dim intLen As Integer
    
    m_strNode_Closed = "..+."
    m_strNode_Open = "..-."
    m_strNode_End = "...."
    m_strNode_Drop = "!"

    m_BuildDropNode
    
End Sub

Private Sub Class_Terminate()
    
    Set m_objListbox = Nothing
    
End Sub


Private Sub m_objListbox_Click()

    Dim intIndex As Integer
    Dim intNodeIndex As Integer
    
    intIndex = m_objListbox.ListIndex
    intNodeIndex = CInt(m_objListbox.List(intIndex, 1))
    
    RaiseEvent ClickedNode(intNodeIndex, intIndex)
    
End Sub
Private Sub m_objListbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim intRowIndex As Integer
    Dim intNodeIndex As Integer
    
    intRowIndex = m_objListbox.ListIndex
    intNodeIndex = CInt(m_objListbox.List(intRowIndex, 1))
    
    m_UpdateNodes intNodeIndex, intRowIndex
    RaiseEvent DblClickedNode(intNodeIndex, intRowIndex)
    
End Sub
Private Sub m_objListbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Dim intRowIndex As Integer
    Dim intNodeIndex As Integer
    
    If KeyCode = vbKeyReturn Then
        intRowIndex = m_objListbox.ListIndex
        intNodeIndex = CInt(m_objListbox.List(intRowIndex, 1))
        m_UpdateNodes intNodeIndex, intRowIndex
    End If
    
End Sub

'Para usar la clase anterior. Agregue a un form, en el form agregue
'un ListBox con el nombre ListBox1 y agregue el siguiente código
'al módulo del form:


Private Sub UserForm_Initialize()
    
    ListBox1.Move 0, 0, 200, Me.InsideHeight - 2
    
   
    Set m_objHList = New ajpHList 'Inicializar variable
    Set m_objHList.UseListbox = ListBox1 'ListBox1 es el HList
        
    
    With m_objHList
    
    '
    ' Los niveles comienzan en 0 (cero)
    ' El incremento de nivel debe ser incremental
    '
        .AddNode "Uno", 0
        .AddNode "Dos", 1
        .AddNode "Tres", 1
        .AddNode "Cuatro", 1
        .AddNode "Cinco", 0
        .AddNode "Seis", 1
        .AddNode "Siete", 2
        .AddNode "Ocho", 2
        .AddNode "Nueve", 2
        .AddNode "Diez", 1
        .AddNode "Once", 2
        .AddNode "Doce", 2
        .AddNode "Trece", 2
        .AddNode "Catorce", 2
        .AddNode "Quince", 0
        .AddNode "Dieciséis", 1
        .AddNode "Diecisiete", 2
        .AddNode "Dieciocho", 3
        .AddNode "Diecinueve", 4
        .AddNode "Veinte", 0
        .AddNode "Veintiuno", 1
        .AddNode "Veintidós", 1
        .AddNode "Veintitrés", 2
        .AddNode "Veinticuatro", 3
        .AddNode "Veinticinco", 0
        .Create
    End With
    
End Sub
```

## Descripción

| Propiedad/Método | Descripción |
|---|---|
| `NodeClosedText` | Texto visual para nodo cerrado (por defecto `..+.`) |
| `NodeOpenText` | Texto visual para nodo abierto (por defecto `..-.`) |
| `NodeEndText` | Texto visual para nodo final (por defecto `....`) |
| `NodeDropText` | Carácter de conexión entre niveles (por defecto `!`) |
| `NodeCount` | Número total de nodos en la estructura |
| `DisplayListCount` | Número de elementos visibles actualmente en el ListBox |
| `Refresh` | Reconstruye la estructura visual del árbol |
| `UseListbox` | Propiedad Set para asignar el control ListBox a usar |
| Evento `ClickedNode` | Se dispara al hacer clic en un nodo |
| Evento `DblClickedNode` | Se dispara al hacer doble clic en un nodo |
| `UserForm_Initialize` | Ejemplo de uso: crea 25 nodos jerárquicos en 5 niveles |

**Características principales:**
- Soporta hasta 5 niveles de jerarquía
- Los nodos se expanden/contraen con doble clic o tecla Enter
- Los símbolos `+` y `-` indican estado del nodo
- Los caracteres `!` y espacios crean la conexión visual vertical

---

*Módulo: 95 - VBA Class module | Archivo: 01 - Class Hierarchical Listbox.txt*
