VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBLambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ActionListener
Private mObj As Object
Private mFncName As String
Private mArgs()

Friend Sub New_(aObj As Object, aFncNam As String) ', ParamArray args())
    Set mObj = aObj: mFncName = aFncNam ': mArgs() = args()
End Sub
Public Sub CallSub()
Attribute CallSub.VB_UserMemId = 0
    CallByName mObj, mFncName, VbMethod ', mArgs
End Sub

Private Sub ActionListener_ActionPerformed(ActionEvent As Variant)
    CallSub
End Sub
