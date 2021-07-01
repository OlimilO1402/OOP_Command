Attribute VB_Name = "MNew"
Option Explicit

Public Function elam(aObj As Object, aFncNam As String) As VBLambda
    Set elam = New VBLambda: elam.New_ aObj, aFncNam
End Function
Public Function MyButton(aBtn As CommandButton, Text As String, aActionListener As ActionListener) As MyButton
    Set MyButton = New MyButton: MyButton.New_ aBtn, Text, aActionListener
End Function


Public Function Account(ByVal aID As Long, Optional ByVal aBalance As Double) As Account
    Set Account = New Account: Account.New_ aID, aBalance
End Function
'Public Function AccountUI(aFrm As Form) As AccountUI
Public Function AccountUI(aLbl1 As Label, aLbl2 As Label, aModel As Account) As AccountUI
    Set AccountUI = New AccountUI: AccountUI.New_ aLbl1, aLbl2, aModel    'aFrm
End Function


Public Function DepositCommand(aAccount As Account, ByVal aAmount As Double) As DepositCommand
    Set DepositCommand = New DepositCommand: DepositCommand.New_ aAccount, aAmount
End Function
Public Function CDepositCommand(obj As Object) As DepositCommand
    On Error Resume Next
    Set CDepositCommand = obj
End Function
Public Function DepositAction(aUndoManager As UndoManager, aAccount As Account, aAmount As Double) As DepositAction
    Set DepositAction = New DepositAction: DepositAction.New_ aUndoManager, aAccount, aAmount
End Function

Public Function WithdrawCommand(aAccount As Account, ByVal aAmount As Double) As WithdrawCommand
    Set WithdrawCommand = New WithdrawCommand: WithdrawCommand.New_ aAccount, aAmount
End Function
Public Function CWithdrawCommand(obj As Object) As WithdrawCommand
    On Error Resume Next
    Set CWithdrawCommand = obj
End Function
Public Function WithdrawAction(aUndoManager As UndoManager, aAccount As Account, aAmount As Double) As WithdrawAction
    Set WithdrawAction = New WithdrawAction: WithdrawAction.New_ aUndoManager, aAccount, aAmount
End Function

Public Function TransferAction(aUndoManager As UndoManager, aFrom As Account, aTo As Account, ByVal aAmount As Double) As TransferAction
    Set TransferAction = New TransferAction: TransferAction.New_ aUndoManager, aFrom, aTo, aAmount
End Function
Public Function TransferCommand(aFrom As Account, aTo As Account, ByVal aAmount As Double) As TransferCommand
    Set TransferCommand = New TransferCommand: TransferCommand.New_ aFrom, aTo, aAmount
End Function

Public Function CompositeCommand(aName As String, ParamArray commands()) As CompositeCommand
    Set CompositeCommand = New CompositeCommand: CompositeCommand.New_ aName, commands()
End Function

