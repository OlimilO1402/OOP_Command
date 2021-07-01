VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Id      As Long
Private m_Balance As Double

Public Event PropertyChange(ByVal aPropName As String, ByVal OldBalance As Double, ByVal CurBalance As Double)

Friend Sub New_(ByVal aID As Long, Optional ByVal aBalance As Double)
    m_Id = aID: m_Balance = aBalance
End Sub

Public Property Get Id() As Long
    Id = m_Id
End Property

Public Property Get Balance() As Double
    Balance = m_Balance
End Property

Public Sub Deposit(ByVal Amount As Double)
    Dim OldBalance As Double: OldBalance = m_Balance
    m_Balance = m_Balance + Amount
    RaiseEvent PropertyChange("Balance", OldBalance, m_Balance)
End Sub

Public Sub Withdraw(ByVal Amount As Double)
    'If Amount > m_Balance Then
    '    'throw new IllegalStateException("Insufficient Funds");
    '    'Err.Raise 1, "IllegalStateException", "IllegalStateException Insufficent Funds"
    '    MsgBox "IllegalStateException " & "Insufficent Funds"
    '    Exit Sub
    'End If
    Dim OldBalance As Double: OldBalance = m_Balance
    m_Balance = m_Balance - Amount
    RaiseEvent PropertyChange("Balance", OldBalance, m_Balance)
End Sub

'public class Account {
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
'    private int id;
'    private int balance;
'
'    public Account(int id) {
'        this.id = id;
'    }
'
'    public Account(int id, int balance) {
'        this.id = id;
'        this.balance = balance;
'    }
'
'    public int getId() {
'        return id;
'    }
'
'    public void deposit(int amount) {
'        int oldBalance = balance;
'        balance += amount;
'        pcs.firePropertyChange("balance", oldBalance, balance)
'    }
'    public void withdraw(int amount) {
'        if (amount > balance) {
'            throw new IllegalStateException("Insufficient Funds");
'        }
'        int oldBalance = balance;
'        balance -= amount;
'        pcs.firePropertyChange("balance", oldBalance, balance);
'    }
'
'    public int getBalance() {
'        return balance;
'    }
'}
