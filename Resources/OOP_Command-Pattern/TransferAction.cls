VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransferAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ActionListener
Private m_From   As Account
Private m_To     As Account
Private m_Amount As Double
Private m_UndoManager As UndoManager

Friend Sub New_(aUndoManager As UndoManager, aFrom As Account, aTo As Account, ByVal aAmount As Double)
    Set m_From = aFrom: Set m_To = aTo: m_Amount = aAmount: Set m_UndoManager = aUndoManager
End Sub

Private Sub ActionListener_ActionPerformed(ActionEvent As Variant)
    Dim aCommand As CompositeCommand
    Set aCommand = MNew.CompositeCommand("Transfer " & m_Amount & " from " & m_From.Id & " to " & m_To.Id, _
        MNew.WithdrawCommand(m_From, m_Amount), MNew.DepositCommand(m_To, m_Amount))
    'm_UndoManager.Execute aCommand
    
    m_UndoManager.Execute MNew.TransferCommand(m_From, m_To, m_Amount)
End Sub
'public class TransferAction implements ActionListener {
'    private Account from;
'    private Account to;
'    private int amount;
'    private UndoManager undoManager;
'
'    public TransferAction(UndoManager undoManager, Account from, Account to, int amount) {
'        this.undoManager = undoManager;
'        this.from = from;
'        this.to = to;
'        this.amount = amount;
'    }
'
'    @Override
'    public void actionPerformed(ActionEvent e) {
'        CompositeCommand command = new CompositeCommand("Transfer " + amount + " from " + from.getId() + " to " + to.getId(),
'                new WithdrawCommand(from, amount),
'                new DepositCommand(to, amount)
'                );
'
'        undoManager.execute(command);
'//      undoManager.execute(new TransferCommand(from, to, amount));
'    }
'}
