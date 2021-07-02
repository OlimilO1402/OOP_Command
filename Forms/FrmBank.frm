VERSION 5.00
Begin VB.Form FrmBank 
   Caption         =   "GoF - Command pattern"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5055
   Icon            =   "FrmBank.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnRedo 
      Caption         =   "Redo"
      Height          =   615
      Left            =   2520
      TabIndex        =   11
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton BtnUndo 
      Caption         =   "Undo"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton BtnTransfer10fromAcc2toAcc1 
      Caption         =   "Transfer 10.- from acc2 to acc1"
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton BtnTransfer10fromAcc1toAcc2 
      Caption         =   "Transfer 10.- from acc1 to acc2"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton BtnWithdraw10Acc2 
      Caption         =   "Withdraw 10.- from account 1"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton BtnWithdraw10Acc1 
      Caption         =   "Withdraw 10.- from account 1"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton BtnDeposit10Acc2 
      Caption         =   "Deposit 10.- to account 2"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton BtnDeposit10Acc1 
      Caption         =   "Deposit 10.- to account 1"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label LblAccount2Balance 
      Caption         =   "- - - - -"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label LblAccount2Id 
      Caption         =   "- - - - -"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label LblAccount1Balance 
      Caption         =   "- - - - -"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label LblAccount1Id 
      Caption         =   "- - - - -"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "FrmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'https://www.youtube.com/watch?v=83dcoRgE0UI
'https://github.com/javadude/patterns.session4/tree/master/Command
'http://www.javadude.com/articles/patterns/index.html
'kleine Vereinfachung für VB
Private WithEvents Account1    As Account
Attribute Account1.VB_VarHelpID = -1
Private WithEvents Account2    As Account
Attribute Account2.VB_VarHelpID = -1
Private WithEvents UndoManager As UndoManager
Attribute UndoManager.VB_VarHelpID = -1

Private Sub Form_Load()
    Set Account1 = MNew.Account(1)
    Set Account2 = MNew.Account(2)
    Set UndoManager = New UndoManager
End Sub

Private Sub BtnDeposit10Acc1_Click()
    UndoManager.Execute MNew.DepositCommand(Account1, 10)
End Sub
Private Sub BtnDeposit10Acc2_Click()
    UndoManager.Execute MNew.DepositCommand(Account2, 10)
End Sub

Private Sub BtnWithdraw10Acc1_Click()
    UndoManager.Execute MNew.WithdrawCommand(Account1, 10)
End Sub
Private Sub BtnWithdraw10Acc2_Click()
    UndoManager.Execute MNew.WithdrawCommand(Account2, 10)
End Sub

Private Sub BtnTransfer10fromAcc1toAcc2_Click()
    UndoManager.Execute MNew.TransferCommand(Account1, Account2, 10)
End Sub
Private Sub BtnTransfer10fromAcc2toAcc1_Click()
    UndoManager.Execute MNew.TransferCommand(Account2, Account1, 10)
End Sub

Private Sub BtnUndo_Click()
    UndoManager.Undo
End Sub
Private Sub BtnRedo_Click()
    UndoManager.Redo
End Sub

Private Sub Account1_PropertyChange()
    LblAccount1Id.Caption = "Account #" & Account1.Id
    LblAccount1Balance.Caption = "Balance: $" & Account1.Balance
End Sub

Private Sub Account2_PropertyChange()
    LblAccount2Id.Caption = "Account #" & Account2.Id
    LblAccount2Balance.Caption = "Balance: $" & Account2.Balance
End Sub

Private Sub UndoManager_PropertyChange()
    BtnUndo.Enabled = UndoManager.IsUndoAvailable
    BtnUndo.Caption = "Undo " & UndoManager.UndoName
    BtnRedo.Enabled = UndoManager.IsRedoAvailable
    BtnRedo.Caption = "Redo " & UndoManager.RedoName
End Sub

'public class BankUI extends JFrame {
'    private Account account1 = new Account(1);
'    private Account account2 = new Account(2);
'    private UndoManager undoManager = new UndoManager();
'    private MyButton undoButton = new MyButton("Undo", e -> {undoManager.undo();});
'    private MyButton redoButton = new MyButton("Redo", e -> {undoManager.redo();});
'
'    public BankUI() {
'        setLayout(new BorderLayout());
'        add(BorderLayout.NORTH, new JPanel(new GridLayout(1, 0, 5, 5))
'        {{
'            add(new AccountUI(account1));
'            add(new AccountUI(account2));
'        }});
'        add(BorderLayout.CENTER, new JButton(new SampleAction()));
'        add(BorderLayout.SOUTH, new JPanel(new GridLayout(0, 1, 5, 5)) {{
'            add(new MyButton("Deposit $10 to account 1", new DepositAction(undoManager, account1, 10)));
'            add(new MyButton("Deposit $10 to account 2", new DepositAction(undoManager, account2, 10)));
'            add(new MyButton("Withdraw $10 from account 1", new WithdrawAction(undoManager, account1, 10)));
'            add(new MyButton("Withdraw $10 from account 2", new WithdrawAction(undoManager, account2, 10)));
'            add(new MyButton("Transfer $10 from account 1 to account 2", new TransferAction(undoManager, account1, account2, 10)));
'            add(new MyButton("Transfer $10 from account 2 to account 1", new TransferAction(undoManager, account2, account1, 10)));
'            add(undoButton);
'            add(redoButton);
'        }});
'        undoManager.addPropertyChangeListener(e -> updateButtons());
'        updateButtons();
'        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
'        pack();
'    }
'
'    private void updateButtons() {
'        undoButton.setVisible(undoManager.isUndoAvailable());
'        undoButton.setText("Undo " + undoManager.getUndoName());
'        redoButton.setVisible(undoManager.isRedoAvailable());
'        redoButton.setText("Redo " + undoManager.getRedoName());
'    }
'
'    private static class MyButton extends JButton {
'        public MyButton(String text, ActionListener actionListener) {
'            super(text);
'            addActionListener(actionListener);
'        }
'    }
'}
