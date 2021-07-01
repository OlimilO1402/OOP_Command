Attribute VB_Name = "MNew"
Option Explicit

Public Function Account(ByVal aID As Long, Optional ByVal aBalance As Double) As Account
    Set Account = New Account: Account.New_ aID, aBalance
End Function

Public Function DepositCommand(aAccount As Account, ByVal aAmount As Double) As DepositCommand
    Set DepositCommand = New DepositCommand: DepositCommand.New_ aAccount, aAmount
End Function
Public Function CDepositCommand(obj As Object) As DepositCommand
    On Error Resume Next
    Set CDepositCommand = obj
End Function

Public Function WithdrawCommand(aAccount As Account, ByVal aAmount As Double) As WithdrawCommand
    Set WithdrawCommand = New WithdrawCommand: WithdrawCommand.New_ aAccount, aAmount
End Function
Public Function CWithdrawCommand(obj As Object) As WithdrawCommand
    On Error Resume Next
    Set CWithdrawCommand = obj
End Function

Public Function TransferCommand(aFrom As Account, aTo As Account, ByVal aAmount As Double) As TransferCommand
    Set TransferCommand = New TransferCommand: TransferCommand.New_ aFrom, aTo, aAmount
End Function

Public Function CompositeCommand(aName As String, ParamArray commands()) As CompositeCommand
    Set CompositeCommand = New CompositeCommand: CompositeCommand.New_ aName, commands()
End Function

