VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DepositAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ActionListener
Private m_Account As Account
Private m_Amount  As Double
Private m_UndoManager As UndoManager

Friend Sub New_(aUndoManager As UndoManager, aAccount As Account, aAmount As Double)
    Set m_UndoManager = aUndoManager: Set m_Account = aAccount: m_Amount = aAmount
End Sub

Private Sub ActionListener_ActionPerformed(ActionEvent As Variant)
    m_UndoManager.Execute MNew.DepositCommand(m_Account, m_Amount)
End Sub

'public class DepositAction implements ActionListener {
'    private Account account;
'    private int amount;
'    private UndoManager undoManager;
'
'    public DepositAction(UndoManager undoManager, Account account, int amount) {
'        this.undoManager = undoManager;
'        this.account = account;
'        this.amount = amount;
'    }
'
'    @Override
'    public void actionPerformed(ActionEvent e) {
'        undoManager.execute(new DepositCommand(account, amount));
'    }
'}
