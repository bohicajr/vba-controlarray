VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Binary
Option Explicit

'Requires reference to Microsoft Forms Object Library (MSForms)

Public Property Get Name() As String: End Property
Public Property Get ControlType() As eControlTypes: End Property
Public Property Get Object() As MSForms.Control: End Property
