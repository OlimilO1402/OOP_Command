VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UndoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event PropertyChange(ByVal aPropName As String, ByVal nam1 As String, ByVal nam2 As String)

Private UndoStack As StackOfCommand
Private RedoStack As StackOfCommand

Private Type OldState
    UndoName      As String
    RedoName      As String
    undoAvailable As Boolean
    redoAvailable As Boolean
End Type
Private Function New_OldState() As OldState
    With New_OldState
        .UndoName = Me.UndoName
        .RedoName = Me.RedoName
        .undoAvailable = Me.IsUndoAvailable
        .redoAvailable = Me.IsRedoAvailable
    End With
End Function

Private Sub Class_Initialize()
    Set UndoStack = New StackOfCommand
    Set RedoStack = New StackOfCommand
End Sub

Public Sub Execute(aCommand As Command)
Try: On Error GoTo Catch
    Dim aOldState As OldState: aOldState = New_OldState
    aCommand.Execute
    'Achtung hier:
    
    If Me.IsUndoAvailable Then
        If UndoStack.Peek.IsCollapsible(aCommand) Then
            UndoStack.Peek.Collapse aCommand
        Else
            UndoStack.Push aCommand
        End If
    Else
        UndoStack.Push aCommand
    End If
    RedoStack.Clear
    fireChanges aOldState
    Exit Sub
Catch: '(IllegalStateException e)
       '// report and log
End Sub

Public Sub Undo()
    If Not UndoStack.IsEmpty Then
        Dim aOldState As OldState: aOldState = New_OldState
        Dim aCommand As Command: Set aCommand = UndoStack.Pop
        aCommand.Undo
        RedoStack.Push aCommand
        fireChanges aOldState
        
    End If
End Sub

Public Sub Redo()
    If Not RedoStack.IsEmpty Then
        Dim aOldState As OldState: aOldState = New_OldState
        Dim aCommand As Command: Set aCommand = RedoStack.Pop
        aCommand.Redo
        UndoStack.Push aCommand
        fireChanges aOldState
        
    End If
End Sub

Private Sub fireChanges(aOldState As OldState)
    RaiseEvent PropertyChange("UndoName", aOldState.UndoName, Me.UndoName)
    'RaiseEvent PropertyChange("RedoName", aOldState.RedoName, Me.RedoName)
    'RaiseEvent PropertyChange("CanUndo", aOldState.undoAvailable, Me.IsUndoAvailable)
    'RaiseEvent PropertyChange("CanRedo", aOldState.redoAvailable, Me.IsRedoAvailable)
End Sub

Public Property Get IsUndoAvailable() As Boolean
    IsUndoAvailable = Not UndoStack.IsEmpty
End Property

Public Property Get IsRedoAvailable() As Boolean
    IsRedoAvailable = Not RedoStack.IsEmpty
End Property

Public Property Get UndoName() As String
    If IsUndoAvailable Then
        UndoName = UndoStack.Peek.Name
    End If
End Property

Public Property Get RedoName() As String
    If IsRedoAvailable Then
        RedoName = RedoStack.Peek.Name
    End If
End Property

'public class UndoManager {
'    private PropertyChangeSupport pcs = new PropertyChangeSupport(this);
'
'    public void addPropertyChangeListener(PropertyChangeListener listener) {
'        pcs.addPropertyChangeListener(listener);
'    }
'    public void addPropertyChangeListener(String propertyName, PropertyChangeListener listener) {
'        pcs.addPropertyChangeListener(propertyName, listener);
'    }
'    public void removePropertyChangeListener(PropertyChangeListener listener) {
'        pcs.removePropertyChangeListener(listener);
'    }
'    public void removePropertyChangeListener(String propertyName, PropertyChangeListener listener) {
'        pcs.removePropertyChangeListener(propertyName, listener);
'    }
'
'    private Stack<Command> undoStack = new Stack<Command>();
'    private Stack<Command> redoStack = new Stack<Command>();
'
'    private class OldState {
'        private String undoName = getUndoName();
'        private String redoName = getUndoName();
'        private boolean undoAvailable = isUndoAvailable();
'        private boolean redoAvailable = isRedoAvailable();
'    }
'
'    public void execute(Command command) {
'        try {
'            OldState oldState = new OldState();
'            command.execute();
'            if (isUndoAvailable() && undoStack.peek().isCollapsible(command)) {
'                undoStack.peek().collapse(command);
'            } else {
'                undoStack.push(command);
'            }
'            redoStack.clear();
'            fireChanges(oldState);
'        } catch (IllegalStateException e) {
'            // report and log
'        }
'    }
'
'    public void undo() {
'        if (!undoStack.isEmpty()) {
'            try {
'                OldState oldState = new OldState();
'                Command command = undoStack.pop();
'                command.undo();
'                redoStack.push(command);
'                fireChanges(oldState);
'            } catch (IllegalStateException e) {
'                // report and log
'            }
'        }
'    }
'
'    public void redo() {
'        if (!redoStack.isEmpty()) {
'            try {
'                OldState oldState = new OldState();
'                Command command = redoStack.pop();
'                command.redo();
'                undoStack.push(command);
'                fireChanges(oldState);
'            } catch (IllegalStateException e) {
'                // report and log
'            }
'        }
'    }
'
'    private void fireChanges(OldState oldState) {
'        pcs.firePropertyChange("undoName", oldState.undoName, getUndoName());
'        pcs.firePropertyChange("redoName", oldState.redoName, getRedoName());
'        pcs.firePropertyChange("canUndo", oldState.undoAvailable, isUndoAvailable());
'        pcs.firePropertyChange("canRedo", oldState.redoAvailable, isRedoAvailable());
'    }
'
'    public boolean isUndoAvailable() {
'        return !undoStack.isEmpty();
'    }
'    public boolean isRedoAvailable() {
'        return !redoStack.isEmpty();
'    }
'
'    public String getUndoName() {
'        if (isUndoAvailable()) {
'            return undoStack.peek().getName();
'        }
'        return "";
'    }
'    public String getRedoName() {
'        if (isRedoAvailable()) {
'            return redoStack.peek().getName();
'        }
'        return "";
'    }
'}
