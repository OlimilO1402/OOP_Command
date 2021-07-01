VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WithdrawCommand"
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
    m_Amount = m_Amount + CWithdrawCommand(aCommand).Amount
End Sub

Private Property Get Command_IsCollapsible(aCommand As Command) As Boolean
    'return command.getClass() == DepositCommand.class;
    If TypeName(CWithdrawCommand(aCommand)) = TypeName(Me) Then
        If CWithdrawCommand(aCommand).Account Is m_Account Then
            Command_IsCollapsible = True
            'TypeName(CDepositCommand(aCommand)) = TypeName(Me) '"DepositCommand"
        End If
    End If
End Property

Private Property Get Command_Name() As String
    'Command_Name = TypeName(Me) '"WithdrawCommand" 'geht auch mit TypeName(me)
    Command_Name = "withdraw of " & m_Amount & " from " & m_Account.Id
End Property

Private Sub Command_Execute()
    m_Account.Withdraw m_Amount
End Sub

Private Sub Command_Undo()
    m_Account.Deposit m_Amount
End Sub

Private Sub Command_Redo()
    Command_Execute
End Sub

'public class WithdrawCommand implements Command {
'    private Account account;
'    private int amount;
'
'    public WithdrawCommand(Account account, int amount) {
'        this.account = account;
'        this.amount = amount;
'    }
'
'    @Override
'    public void execute() {
'        account.withdraw(amount);
'    }
'
'    @Override
'    public void undo() {
'        account.deposit(amount);
'    }
'
'    @Override
'    public void redo() {
'        execute();
'    }
'
'    @Override
'    public String getName() {
'        return "Withdraw " + amount + " from account " + account.getId();
'    }
'
'    @Override
'    public boolean isCollapsible(Command command) {
'        return command.getClass() == WithdrawCommand.class;
'    }
'
'    @Override
'    public void collapse(Command command) {
'        amount += ((WithdrawCommand)command).amount;
'    }
'}
