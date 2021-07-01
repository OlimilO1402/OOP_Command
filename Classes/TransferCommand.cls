VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransferCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Command
Private m_CompositeCommand As Command 'CompositeCommand
'Private m_From   As Account
'Private m_To     As Account
'Private m_Amount As Double

Friend Sub New_(aFrom As Account, aTo As Account, ByVal aAmount As Double)
    Set m_CompositeCommand = MNew.CompositeCommand("Transfer " & aAmount & " from account " & aFrom.Id & " to account " & aTo.Id, _
        MNew.WithdrawCommand(aFrom, aAmount), MNew.DepositCommand(aTo, aAmount))
End Sub

Private Sub Command_Collapse(aCommand As Command) 'nope
End Sub
Private Property Get Command_IsCollapsible(aCommand As Command) As Boolean 'nope
End Property

Private Sub Command_Execute()
    m_CompositeCommand.Execute
End Sub

Private Property Get Command_Name() As String
    Command_Name = m_CompositeCommand.Name
End Property

Private Sub Command_Undo()
    m_CompositeCommand.Undo
End Sub

Private Sub Command_Redo()
    m_CompositeCommand.Redo
End Sub

'public class TransferCommand implements Command {
'    private Account from;
'    private Account to;
'    private int amount;
'
'    public TransferCommand(Account from, Account to, int amount) {
'        this.from = from;
'        this.to = to;
'        this.amount = amount;
'    }
'
'    @Override
'    public void execute() {
'        try {
'            // start transaction
'            from.withdraw(amount);
'            to.deposit(amount);
'        } finally {
'            // end transaction
'        }
'    }
'
'    @Override
'    public void undo() {
'        try {
'            // start transaction
'            to.withdraw(amount);
'            from.deposit(amount);
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
'    public String getName() {
'        return "Transfer " + amount + " from account " + from.getId() + " to account " + to.getId();
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
'}


'Private m_From   As Account
'Private m_To     As Account
'Private m_Amount As Double
'
'Friend Sub New_(aFrom As Account, aTo As Account, ByVal aAmount As Double)
'    Set m_From = aFrom: Set m_To = aTo: m_Amount = aAmount
'End Sub
'
'Private Sub Command_Collapse(aCommand As Command) 'nope
'End Sub
'Private Property Get Command_IsCollapsible(aCommand As Command) As Boolean 'nope
'End Property
'
'Private Sub Command_Execute()
'    m_From.Withdraw m_Amount
'    m_To.Deposit m_Amount
'End Sub
'
'Private Property Get Command_Name() As String
'    Command_Name = "Transfer " & m_Amount & " from account " & m_From.Id & " to account " & m_To.Id
'End Property
'
'Private Sub Command_Undo()
'    m_To.Withdraw m_Amount
'    m_From.Deposit m_Amount
'End Sub
'
'Private Sub Command_Redo()
'    Command_Execute
'End Sub
'
