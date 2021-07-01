VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompositeCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Command
Private m_Commands() As Command
Private m_Name       As String

Friend Sub New_(aName As String, ByVal commands)
    m_Name = aName
    Dim u As Long: u = UBound(commands)
    ReDim m_Commands(u)
    Dim i As Long
    For i = 0 To UBound(commands)
        Set m_Commands(i) = commands(i)
    Next
End Sub

Private Property Get Command_IsCollapsible(aCommand As Command) As Boolean
End Property

Private Sub Command_Collapse(aCommand As Command) 'nope
End Sub

Private Property Get Command_Name() As String
    Command_Name = m_Name
End Property

Private Sub Command_Execute()
    Dim aCommand 'As Command
    For Each aCommand In m_Commands
        aCommand.Execute
    Next
End Sub

Private Sub Command_Undo()
    Dim aCommand 'As Command
    For Each aCommand In m_Commands
        aCommand.Undo
    Next
End Sub

Private Sub Command_Redo()
    Command_Execute
End Sub

'public class CompositeCommand implements Command {
'    private Command[] commands;
'    private String name;
'
'    public CompositeCommand(String name, Command... commands) {
'        this.name = name;
'        this.commands = commands;
'    }
'
'    @Override
'    public void execute() {
'        try {
'            // start transaction
'            for (Command command : commands) {
'                command.execute();
'            }
'        } finally {
'            // end transaction
'        }
'    }
'
'    @Override
'    public void undo() {
'        try
'            // start transaction
'            for (int i = commands.length-1; i >= 0; i--) {
'                Command command = commands[i];
'                command.undo();
'            }
'        } finally {
'            // end transaction
'        }
'    }
'
'    @Override
'    public void redo() {
'        execute();
'    }
'
'    @Override
'    public boolean isCollapsible(Command command) {
'        return false;
'    }
'
'    @Override
'    public void collapse(Command command) {
'    }
'
'    @Override
'    public String getName() {
'        return name;
'    }
'}
