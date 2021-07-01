VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mBtn As CommandButton
Attribute mBtn.VB_VarHelpID = -1
Private mActionListener As ActionListener

Friend Sub New_(aBtn As CommandButton, Text As String, aActionListener As ActionListener)
    Set mBtn = aBtn: aBtn.Caption = Text: Set mActionListener = aActionListener
End Sub

Private Sub mBtn_Click()
    Dim v
    mActionListener.ActionPerformed v
End Sub

Public Property Get Enabled() As Boolean
    Enabled = mBtn.Enabled
End Property
Public Property Let Enabled(ByVal Value As Boolean)
    mBtn.Enabled = Value
End Property
Public Property Get Text() As String
    Text = mBtn.Caption
End Property
Public Property Let Text(ByVal Value As String)
    mBtn.Caption = Value
End Property

'private static class MyButton extends JButton {
'    public MyButton(String text, ActionListener actionListener) {
'        super(text);
'        addActionListener(actionListener);
'    }
'}
