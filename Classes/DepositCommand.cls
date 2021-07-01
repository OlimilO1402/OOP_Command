VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DepositCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Command
Private m_Account As Account
Private m_Amount  As Double

Friend Sub New_(aAccount As Account, ByVal aAmount As Double)
    Set m_Account = aAccount
    m_Amount = aAmount
End Sub

Public Property Get Amount() As Double
    Amount = m_Amount
End Property

Public Property Get Account() As Account
    Set Account = m_Account
End Property

Private Sub Command_Collapse(aCommand As Command)
'amount += ((DepositCommand)command).amount;
    m_Amount = m_Amount + CDepositCommand(aCommand).Amount
End Sub

Private Property Get Command_IsCollapsible(aCommand As Command) As Boolean
    'return command.getClass() == DepositCommand.class;
    If TypeName(CDepositCommand(aCommand)) = TypeName(Me) Then
        If CDepositCommand(aCommand).Account Is m_Account Then
            Command_IsCollapsible = True
            'TypeName(CDepositCommand(aCommand)) = TypeName(Me) '"DepositCommand"
        End If
    End If
End Property

Private Property Get Command_Name() As String
    'Command_Name = TypeName(Me) '"DepositCommand"
    Command_Name = "deposit of " & m_Amount & " to " & m_Account.Id
End Property

Private Sub Command_Execute()
    m_Account.Deposit m_Amount
End Sub

Private Sub Command_Undo()
    m_Account.Withdraw m_Amount
End Sub

Private Sub Command_Redo()
    Command_Execute
End Sub

'public class DepositCommand implements Command {
'    private Account account;
'    private int amount;
'
'    public DepositCommand(Account account, int amount) {
'        this.account = account;
'        this.amount = amount;
'    }
'
'    @Override
'    public void execute() {
'        account.deposit(amount);
'    }
'
'    @Override
'    public void undo() {
'        account.withdraw(amount);
'    }
'
'    @Override
'    public void redo() {
'        execute();
'    }
'
'    @Override
'    public String getName() {
'        return "Deposit " + amount + " into account " + account.getId();
'    }
'
'    @Override
'    public boolean isCollapsible(Command command) {
'        return command.getClass() == DepositCommand.class;
'    }
'
'    @Override
'    public void collapse(Command command) {
'        amount += ((DepositCommand)command).amount;
'    }
'}
