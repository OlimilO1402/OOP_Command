VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AccountUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Pattern: Model View Controller
'hier View, die Klasse AccountUI ist der View des Models, das Model ist die Klasse Account
'der View hat einen Verweis auf das Model
'das Model gibt rückmeldung an den View, per Event, ob sich etwas, und was sich geändert hat
Private accountNumField As Label
Private balanceField    As Label
Private WithEvents mModel As Account
Attribute mModel.VB_VarHelpID = -1

Friend Sub New_(aLbl1 As Label, aLbl2 As Label, aModel As Account)
    Set accountNumField = aLbl1: aLbl1.Caption = "-----"
    Set balanceField = aLbl2:    aLbl2.Caption = "-------"
    Set mModel = aModel
End Sub

Public Property Set Model(aModel As Account)
    Set Model = aModel
    UpdateFields
End Property

Private Sub mModel_PropertyChange(ByVal aPropName As String, ByVal OldBalance As Double, ByVal CurBalance As Double)
    'hmm so sollte man es auf keinen Fall machen hier sollte man auch etwas gezielter reagieren
    UpdateFields
End Sub

Private Sub UpdateFields()
    If Not mModel Is Nothing Then
        accountNumField.Caption = "Account #" & mModel.Id
        balanceField.Caption = "Balance $" & mModel.Balance
    End If
End Sub

'public class AccountUI extends JPanel {
'    private JLabel accountNumField = new JLabel("-----");
'    private JLabel balanceField = new JLabel("--------");
'    private Account model;
'
'    public AccountUI(Account model) {
'        this();
'        setModel(model);
'    }
'    public AccountUI() {
'        setLayout(new BorderLayout());
'        add(BorderLayout.NORTH, new JPanel(new BorderLayout()) {{
'            add(BorderLayout.WEST, new JPanel(new GridLayout(0, 1, 5, 5)) {{
'                add(new JLabel("Account #"));
'                add(new JLabel("Balance $"));
'            }});
'            add(BorderLayout.CENTER, new JPanel(new GridLayout(0, 1, 5, 5)) {{
'                add(accountNumField);
'                add(balanceField);
'            }});
'        }});
'    }
'    private PropertyChangeListener listener = e -> updateFields();
'
'    public void setModel(Account model) {
'        if (this.model != null) {
'            this.model.removePropertyChangeListener(listener);
'        }
'        this.model = model;
'        if (this.model != null) {
'            this.model.addPropertyChangeListener(listener);
'        }
'        updateFields();
'    }
'
'    private void updateFields() {
'        accountNumField.setText(this.model != null ? this.model.getId() + "" : "");
'        balanceField.setText(this.model != null ? this.model.getBalance() + "" : "");
'    }
'}
