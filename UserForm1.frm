VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Final Mark"
   ClientHeight    =   9195.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
'sale
sale1.Value = ""
sale2.Value = ""
sale3.Value = ""
sale4.Value = ""
sale5.Value = ""
sale6.Value = ""
sale7.Value = ""
sale8.Value = ""
sale9.Value = ""

'revenue
rev1.Value = ""
rev2.Value = ""
rev3.Value = ""
rev4.Value = ""
rev5.Value = ""
rev6.Value = ""
rev7.Value = ""
rev8.Value = ""
rev9.Value = ""

'cost
cost1.Value = ""
cost2.Value = ""
cost3.Value = ""
cost4.Value = ""
cost5.Value = ""
cost6.Value = ""
cost7.Value = ""
cost8.Value = ""
cost9.Value = ""

'Item Description
Descript1.Value = ""
Descript2.Value = ""
Descript3.Value = ""
Descript4.Value = ""
Descript5.Value = ""
Descript6.Value = ""
Descript7.Value = ""
Descript8.Value = ""
Descript9.Value = ""

'ticket number
ticketnum.Value = ""

'Account Number
AccountNum.Value = ""

'Agreement Date
AgreeDate.Value = ""

'Same As Cash Days
SacDays.Value = ""

'Client Name
Clientname.Value = ""

'Principal Value
principalC.Value = ""

'Total Value
totalC.Value = ""

'Notes
Notes.Value = ""


Store.AddItem "02"
Store.AddItem "03"
Store.AddItem "04"
Store.AddItem "05"
Store.AddItem "06"
Store.AddItem "07"
Store.AddItem "08"
Store.AddItem "09"
Store.AddItem "10"
Store.AddItem "11"
Store.AddItem "12"
Store.AddItem "13"
Store.AddItem "14"
Store.AddItem "15"
Store.AddItem "16"
Store.AddItem "17"
Store.AddItem "18"
Store.AddItem "19"
Store.AddItem "20"
Store.AddItem "21"
Store.AddItem "22"
Store.AddItem "23"
Store.AddItem "24"

'Set Focus
Store.SetFocus



End Sub
'This will put the data on the screen. 
Private Sub CommandButton1_Click()
    
    Sheet1.Activate
    
    Range("B1").Value = Store.Value
    Range("B2").Value = Clientname.Value
    Range("B3").Value = AccountNum.Value
    Range("B4").Value = AgreeDate.Value
    Range("B5").Value = SacDays.Value
    Range("B7").Value = ticketnum.Value
    Range("B9").Value = Descript1.Value
    Range("C9").Value = sale1.Value
    Range("D9").Value = rev1.Value
    Range("E9").Value = cost1.Value
    Range("B10").Value = Descript2.Value
    Range("C10").Value = sale2.Value
    Range("D10").Value = rev2.Value
    Range("E10").Value = cost2.Value
    Range("B11").Value = Descript3.Value
    Range("C11").Value = sale3.Value
    Range("D11").Value = rev3.Value
    Range("E11").Value = cost3.Value
    Range("B12").Value = Descript4.Value
    Range("C12").Value = sale4.Value
    Range("D12").Value = rev4.Value
    Range("E12").Value = cost4.Value
    Range("B13").Value = Descript5.Value
    Range("C13").Value = sale5.Value
    Range("D13").Value = rev5.Value
    Range("E13").Value = cost5.Value
    Range("B14").Value = Descript6.Value
    Range("C14").Value = sale6.Value
    Range("D14").Value = rev6.Value
    Range("E14").Value = cost6.Value
    Range("B15").Value = Descript7.Value
    Range("C15").Value = sale7.Value
    Range("D15").Value = rev7.Value
    Range("E15").Value = cost7.Value
    Range("B16").Value = Descript8.Value
    Range("C16").Value = sale8.Value
    Range("D16").Value = rev8.Value
    Range("E16").Value = cost8.Value
    Range("B17").Value = Descript9.Value
    Range("C17").Value = sale9.Value
    Range("D17").Value = rev9.Value
    Range("E17").Value = cost9.Value
    Range("C22").Value = principalC.Value
    Range("D22").Value = totalC.Value
    Range("B24").Value = Notes.Value
    
    
    

End Sub
