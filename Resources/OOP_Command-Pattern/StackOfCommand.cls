VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StackOfCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Arr() As Command
Private m_Count As Long

Private Sub Class_Initialize()
    ReDim m_Arr(0 To 3)
End Sub

Public Sub Clear()
    Class_Initialize
    m_Count = 0
End Sub
Public Property Get IsEmpty() As Boolean
    IsEmpty = m_Count = 0
End Property

Public Function Peek() As Command
    If m_Count = 0 Then Exit Function
    Set Peek = m_Arr(m_Count - 1)
End Function

Public Function Pop() As Command
    Set Pop = Peek
    m_Count = m_Count - 1
    Set m_Arr(m_Count) = Nothing
End Function

Public Sub Push(aCommand As Command)
    If UBound(m_Arr) < m_Count Then ReDim Preserve m_Arr(0 To m_Count * 2)
    Set m_Arr(m_Count) = aCommand
    m_Count = m_Count + 1
End Sub


