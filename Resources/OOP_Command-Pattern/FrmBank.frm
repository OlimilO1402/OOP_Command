VERSION 5.00
Begin VB.Form FrmBank 
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton CbRedoButton 
      Caption         =   "Redo"
      Height          =   615
      Left            =   2520
      TabIndex        =   11
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton CbUndoButton 
      Caption         =   "Undo"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Transfer 10.- from acc2 to acc1"
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Transfer 10.- from acc1 to acc2"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Withdraw 10.- from account 1"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Withdraw 10.- from account 1"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Deposit 10.- to account 2"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deposit 10.- to account 1"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label LblAccount2Balance 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label LblAccount2Id 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label LblAccount1Balance 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label LblAccount1Id 
      Caption         =   "Label1"
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
'
'https://github.com/javadude/patterns.session4/tree/master/Command
'http://www.javadude.com/articles/patterns/index.html

Private Account1    As Account
Attribute Account1.VB_VarHelpID = -1
Private Account2    As Account
Attribute Account2.VB_VarHelpID = -1
Private WithEvents UndoManager As UndoManager
Attribute UndoManager.VB_VarHelpID = -1
Private UndoButton  As MyButton
Private RedoButton  As MyButton

Private ctrls As Collection

Private Sub Form_Load()
    Set ctrls = New Collection
    Set Account1 = MNew.Account(1)
    Set Account2 = MNew.Account(2)
    Set UndoManager = New UndoManager
    Set UndoButton = MNew.MyButton(CbUndoButton, "Undo", elam(UndoManager, "Undo"))
    Set RedoButton = MNew.MyButton(CbRedoButton, "Redo", elam(UndoManager, "Redo"))
    
    ctrls.Add MNew.AccountUI(LblAccount1Id, LblAccount1Balance, Account1)
    ctrls.Add MNew.AccountUI(LblAccount2Id, LblAccount2Balance, Account2)
    
    ctrls.Add MNew.MyButton(Command1, "Deposit $10 to account 1", MNew.DepositAction(UndoManager, Account1, 10))
    ctrls.Add MNew.MyButton(Command2, "Deposit $10 to account 2", MNew.DepositAction(UndoManager, Account2, 10))
    
    ctrls.Add MNew.MyButton(Command3, "Withdraw $10 from account 1", MNew.WithdrawAction(UndoManager, Account1, 10))
    ctrls.Add MNew.MyButton(Command4, "Withdraw $10 from account 2", MNew.WithdrawAction(UndoManager, Account2, 10))
    
    ctrls.Add MNew.MyButton(Command5, "Transfer $10 from account 1 to account 2", MNew.TransferAction(UndoManager, Account1, Account2, 10))
    ctrls.Add MNew.MyButton(Command6, "Transfer $10 from account 2 to account 1", MNew.TransferAction(UndoManager, Account2, Account1, 10))
    
    UpdateButtons
End Sub

Private Sub UndoManager_PropertyChange(ByVal aPropName As String, ByVal nam1 As String, ByVal nam2 As String)
    UpdateButtons
End Sub

Private Sub UpdateButtons()
    UndoButton.Enabled = UndoManager.IsUndoAvailable
    UndoButton.Text = "Undo " & UndoManager.UndoName
    RedoButton.Enabled = UndoManager.IsRedoAvailable
    RedoButton.Text = "Redo " & UndoManager.RedoName
End Sub

'Private Sub CbRedoButton_Click()
'    UndoManager.Redo
'End Sub
'
'Private Sub CbUndoButton_Click()
'    UndoManager.Undo
'End Sub

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
