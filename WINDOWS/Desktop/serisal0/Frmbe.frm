VERSION 5.00
Begin VB.Form frmbe 
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmbe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MAKE SURE YOU HAVE THE COMMAND1'S VISIBLE PROPERTY SET TO FALSE
'YOU CAN CHANGE THE TRIAL LENTH IF YOU WANT TO, ALSO ANY FEEDBACK WOULD BE VERY
'MUCH APPRECIATED, I KNOW THAT THIS IS VERY SIMPLE BUT MY MOTTO IS
'IF IT WORKS THEN HELL....JUST LEAVE IT ALONE!
'BUT IF ANYONE ELSE CAN THINK OF A WAY TO IMPROOVE THIS THEN PLEASE FEEL
'FREE TO EMAIL ME AT
'raif.jackson@eltek-semi.com
'THANKS
Option Explicit

Public Sub Command1_Click()
Dim str1 As String
Dim str2 As String
Dim dlen
Dim regi
Dim pass As String

Close
'OPEN THE FILE AND GET BOTH STRINGS OF INFO
Open App.Path & "\Reg.ini" For Input As #1
Do Until EOF(1)
Line Input #1, str1
Line Input #1, str2
Loop
'CHECK TO SEE IF IT HAS BEEN RIGISTERED
'IF IT HAS THERE IS NO NEED TO GO ON JUST EXIT SUB
If str1 = "Registered" And str2 = " " Then
Exit Sub
Else
End If
'SO IT HASNT BEEN REGISTERED CHECK TO SEE IF IT HAS RAN
'ITS TRIAL PERIOD
dlen = Mid(str1, 1, 2)
'DLEN IS THE NUMBER OF USES
regi = Mid(str2, 1, 12)
'REGI IS IF IT HAS BEEN REGISTERED OR NOT
Close #1

'IF THE NUMBER OF USES EXCEEDS NINE ASK FOR THE REGISTRATION CODE
If dlen > 9 Then
    pass = InputBox("Please enter Registration Code", "Please Register")
'CHECK TO SEE IF THE CODE IS CORRECT
        If pass = "XmTrC1598" Then
'IT IS? OH GOOD THEN OPEN THE REG.INI FILE AND MARK IT SO
            Open App.Path & "\Reg.ini" For Output As #1
            Print #1, "Registered" & vbCrLf & " "
            Close #1
        Else
'IT ISN'T? OH YOU ARE NAUGHTY WELL HAVE TO STOP THE PROGRAM

        MsgBox "Invalid Registration Code", vbCritical, "WRONG"
        End
        
        End If
Else
'SO IT HASN'T RAN ITS TRAIL PEROID? WELL HE HAD BETTER
'TELL THE REG.INI FILE THAT THE TRIAL PERIOD IS GETTING
'CLOSER TO EXPIRING

    Open App.Path & "\Reg.ini" For Output As #1
    Print #1, dlen + 1 & vbCrLf & "Unregistered"
    Close #1
End If
MsgBox App.Path
'AND THAT IS IT! I'M SURE IT COULD BE MADE SIMPLER BUT I DONT SEE HOW!
End Sub


Private Sub Form_Load()
'CHECK TO SEE IF REG.INI EXISTS IF NOT
'IF NOT THEN CREATE IT WITH 1 TRY USED
'IF IT IS THERE CHECK TO SEE IF
'A) IT NEEDS REGISTERING
'B) IT HAS BEEN REGITERED
Dim i As Long
'CHECK THAT THIS IS THE FIRST TIME THIS FORM HAS BEEN OPENED
'SO IF I IS GREATER THAN ONE IT HAS BEEN LOADED B4 SO IGNORE
'ALL THE CHECKING
i = i + 1
If i > 1 Then
MsgBox "Allready loaded once wont add another try"
Exit Sub
End If

If Dir$(App.Path & "\Reg.ini") <> "" Then
GoTo bext
Else
Open App.Path & "\Reg.ini" For Output As #1
Print #1, "1  Try Gone" & vbCrLf & "Unregistered"
Close #1
End If
bext:
Call Command1_Click
End Sub
