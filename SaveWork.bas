Attribute VB_Name = "SaveWork"
Option Explicit

Sub SaveWork()

Dim storeValue, Clientname, AccountNum, myValue, payOff, messageM


storeValue = Range("B1").Value
Clientname = Range("B2").Value
AccountNum = Range("B3").Value

'On click of the Save button. Prompt user to process another. 
myValue = "n"
messageM = "Do you want to process another? (y/n)"
myValue = Application.InputBox(messageM)

If myValue = "y" Or myValue = "Y" Or myValue = "Yes" Or myValue = "YES" Or myValue = "yes" Then

'Saves file and opens the other master excel file.
Application.ScreenUpdating = False
Workbooks.Open Filename:=("C:\YourPathToMaster1")
Workbooks("Payoff_Master.xlsm").Activate
ActiveWorkbook.SaveAs Filename:=("C:\YourSaveDesination\" & storeValue & "\" & storeValue & " " & Clientname & " " & AccountNum & " " & Format(Now(), "DD-MMM-YY hh mm AMPM") & ".xlsm")
ActiveWorkbook.Close Filename:=("C:\YourSaveDesination\" & storeValue & "\" & storeValue & " " & Clientname & " " & AccountNum & " " & Format(Now(), "DD-MMM-YY hh mm AMPM") & ".xlsm")
Application.ScreenUpdating = True

'Saves file and closes the workbook.
Else
Application.ScreenUpdating = False
ActiveWorkbook.SaveAs Filename:=("C:\YourSaveDesination\" & storeValue & "\" & storeValue & " " & Clientname & " " & AccountNum & " " & Format(Now(), "DD-MMM-YY hh mm AMPM") & ".xlsm")
ActiveWorkbook.Close
Application.ScreenUpdating = True
End If


End Sub


