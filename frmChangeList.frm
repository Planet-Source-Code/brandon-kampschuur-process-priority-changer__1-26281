VERSION 5.00
Begin VB.Form frmChangeList 
   Caption         =   "Add Task"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox combxTo 
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Text            =   "Combo2"
      Top             =   360
      Width           =   855
   End
   Begin VB.ComboBox combxFrom 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Process 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "To:"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "From:"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Process Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmChangeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Dim listx As ListItem

ChangeList(frmPriority.lstvwChangeList.ListItems.Count + 1).Process = Trim(Process.Text)
Set listx = frmPriority.lstvwChangeList.ListItems.Add(frmPriority.lstvwChangeList.ListItems.Count + 1, , ChangeList(frmPriority.lstvwChangeList.ListItems.Count + 1).Process)

Select Case combxFrom.ListIndex
    Case 0:
        listx.SubItems(1) = "Idle"
    Case 1:
        listx.SubItems(1) = "Normal"
    Case 2:
        listx.SubItems(1) = "High"
    Case 3:
        listx.SubItems(1) = "Highest"
    Case Else
        MsgBox ("Invalid data selected. Try again.")
        combxFrom.SetFocus
        Exit Sub
End Select

Select Case combxTo.ListIndex
    Case 0:
        listx.SubItems(2) = "Idle"
    Case 1:
        listx.SubItems(2) = "Normal"
    Case 2:
        listx.SubItems(2) = "High"
    Case 3:
        listx.SubItems(2) = "Highest"
    Case Else
        MsgBox ("Invalid data selected. Try again.")
        combxTo.SetFocus
        Exit Sub
End Select

frmPriority.ReDefineChangeList
frmPriority.SaveChangeList
frmPriority.lstbxSystemDialog.AddItem Time & " : Added processes " & ChangeList(frmPriority.lstvwChangeList.ListItems.Count).Process & " to Change List..."
Unload frmChangeList
End Sub

Private Sub cmdCancel_Click()
Unload frmChangeList                                                            'Exit without doing anything
End Sub

Private Sub Form_Load()
combxFrom.Clear                                                                 'Populate Combo Boxes
combxFrom.AddItem "Idle"
combxFrom.AddItem "Normal"
combxFrom.AddItem "High"
combxFrom.AddItem "Highest"
combxTo.Clear
combxTo.AddItem "Idle"
combxTo.AddItem "Normal"
combxTo.AddItem "High"
combxTo.AddItem "Highest"

End Sub
