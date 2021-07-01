VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Command"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Execute()
End Sub

Public Sub Undo()
End Sub

Public Sub Redo()
End Sub

Public Property Get IsCollapsible(aCommand As Command) As Boolean
End Property

Public Sub Collapse(aCommand As Command)
End Sub

Public Property Get Name() As String
End Property

'public interface Command {
'
'    void execute();
'
'    void undo();
'
'    void redo();
'
'    boolean isCollapsible(Command command);
'
'    void collapse(Command command);
'
'    String getName();
'
'}
