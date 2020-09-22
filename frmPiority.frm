VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPriority 
   Caption         =   "Priority Changer (v1.2) rev. 8-16-2001"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstvwProcesses 
      Height          =   3975
      Left            =   4560
      TabIndex        =   8
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   7011
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame fmeOptions 
      Caption         =   "Options"
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4335
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   495
         Left            =   1320
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin MSComctlLib.ListView lstvwChangeList 
         Height          =   1215
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2143
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame fmeUpdateFrequency 
         Caption         =   "Update Frequency"
         Height          =   1575
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   1815
         Begin VB.OptionButton opt5Seconds 
            Caption         =   "Every 5 Seconds"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1575
         End
         Begin VB.OptionButton opt2Seconds 
            Caption         =   "Every 2 Seconds"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton opt1Second 
            Caption         =   "Every Second"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optCustom 
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label lblCustom 
            Caption         =   "[ custom ]"
            Height          =   375
            Left            =   360
            TabIndex        =   16
            Top             =   1080
            Width           =   1335
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Change Process Settings:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Timer Timer1 
      Left            =   6240
      Top             =   0
   End
   Begin VB.ListBox lstbxSystemDialog 
      Height          =   1230
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Kittrich Corporation 2001"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Programmed by: Brandon Kampschuur"
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Current Processes:"
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "System Dialog:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmPriority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*Priority: Process Priority Changer program for win98, winNT kernels           *
'*    The program operates on a few main principles. The handles and processes  *
'*    arrays declared at the beginning are used by the program to store that    *
'*    information whenever it updates.  The core of the program are the subs    *
'*    that change priority and what is found in the timer1.timer routine. Most  *
'*    of the rest of the code is just standard Windows event handling.          *
'********************************************************************************

Dim Handles(1 To 200) As Long                                                   'stores the current process names
Dim Processes(1 To 200) As String                                               'stores the corresponding process handle

Private Sub cmdAdd_Click()
frmChangeList.Show 1                                                            'shows the other form for task addition
End Sub

Private Sub cmdRemove_Click()
Dim Itmx As ListItem
Dim reply As Long

Set Itmx = lstvwChangeList.SelectedItem                                         'Get selected List Item
reply = MsgBox("Remove item: " & Itmx.Text, vbYesNo)                            'Verify remove
If reply = vbYes Then
    lstvwChangeList.ListItems.Remove (Itmx.Index)                               'remove Item
Else
    Exit Sub                                                                    'user cancelled
End If

ReDefineChangeList                                                              'update ChangeList
SaveChangeList                                                                  'save ChangeList

lstbxSystemDialog.AddItem Time & " : Removed process " & Itmx.Text & " from Change List..."

End Sub

Private Sub Form_Load()

lstvwProcesses.ListItems.Clear                                                  'setup List Controls
lstvwProcesses.ColumnHeaders.Clear
lstvwChangeList.ColumnHeaders.Add 1, "Name", "Name", 1440
lstvwChangeList.ColumnHeaders.Add 2, "From", "From", 700
lstvwChangeList.ColumnHeaders.Add 3, "To", "To", 700
lstvwProcesses.ColumnHeaders.Add 1, "Name", "Name", 1440
lstvwProcesses.ColumnHeaders.Add 2, "Priority", "Priority", 700
lstvwProcesses.SortKey = 1
lstbxSystemDialog.AddItem Time & " : Starting program..."

Dim Itmx As ListItem
Dim j As Long
Dim cb As Long
Dim reply, reply2 As Long
Dim temp, temp2, temp3 As String
Dim cbNeeded As Long
Dim NumElements As Long
Dim ProcessIDs() As Long
Dim cbNeeded2 As Long
Dim NumElements2 As Long
Dim Modules(1 To 200) As Long
Dim lRet As Long
Dim ModuleName As String
Dim nSize As Long
Dim hProcess As Long
Dim i As Long
Dim Priority As Long
Dim fNum As Integer

On Error GoTo Errorhandler
fNum = FreeFile
Open (App.Path & "\setup.txt") For Input As fNum
  
Do Until EOF(fNum)
     Line Input #fNum, temp
     reply = InStr(1, temp, Chr(9))
     ChangeList(frmPriority.lstvwChangeList.ListItems.Count + 1).Process = Trim(Left(temp, reply - 1))
     Set Itmx = lstvwChangeList.ListItems.Add(frmPriority.lstvwChangeList.ListItems.Count + 1, , ChangeList(frmPriority.lstvwChangeList.ListItems.Count + 1).Process)
     reply2 = InStr(reply + 1, temp, Chr(9))
     temp3 = Trim(Right(temp, Len(temp) - reply2))
     temp2 = Trim(Mid(temp, reply + 1, (Len(temp) - (Len(ChangeList(frmPriority.lstvwChangeList.ListItems.Count).Process) + Len(temp3)) - 2)))

     Select Case temp2
     Case "High":
         Itmx.SubItems(1) = "High"
     Case "Idle":
         Itmx.SubItems(1) = "Idle"
     Case "Normal":
         Itmx.SubItems(1) = "Normal"
     Case "Highest":
         Itmx.SubItems(1) = "Highest"
     End Select
     
     Select Case temp3
     Case "High":
         Itmx.SubItems(2) = "High"
     Case "Idle":
         Itmx.SubItems(2) = "Idle"
     Case "Normal":
         Itmx.SubItems(2) = "Normal"
     Case "Highest":
         Itmx.SubItems(2) = "Highest"
     End Select
Loop

ReDefineChangeList
Close fNum

Rest:
                                                                                'Get the array containing the process id's for each process object
cb = 8
cbNeeded = 96
Do While cb <= cbNeeded
    cb = cb * 2
    ReDim ProcessIDs(cb / 4) As Long
    lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
Loop
         
NumElements = cbNeeded / 4
j = 1                                                                           'j keeps track of index for lstvwProcesses
For i = 1 To NumElements
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, ProcessIDs(i))                'Get a handle to the Process
    If hProcess <> 0 Then                                                       'Got a Process handle
        lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)         'Get an array of the module handles for the specified process
        If lRet <> 0 Then                                                       'If the Module Array is retrieved, Get the ModuleFileName
            ModuleName = Space(MAX_PATH)                                        'Prepare variables...
            nSize = 500
            lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize) 'Get process name
            Handles(i) = hProcess                                               'Assign handle to array
            Processes(i) = Trim(Right(ModuleName, _
                (Len(ModuleName) - InStrRev(ModuleName, "\"))))                 'Assign name to array
            Set Itmx = lstvwProcesses.ListItems.Add(j, , Processes(i))          'Add process to Process List
            Priority = GetPriorityClass(Handles(i))                             'Retrieve process priority
            AddProcessListSubItems Priority, Itmx                               'Add sub item info to Process List
            ChangePriority i, Priority                                          'Call Change Priority Sub Routine
            j = j + 1
        End If
        lRet = CloseHandle(hProcess)                                            'Close the handle to the process
    End If
Next
lstbxSystemDialog.ListIndex = lstbxSystemDialog.ListCount - 1                   'select the most recent entry
Timer1.Interval = 1000
Exit Sub

Errorhandler:                                                                   'Deals with no setup.txt
    MsgBox ("There is no setup file. Please setup the Change List.")
    GoTo Rest
End Sub

Private Sub opt1Second_Click()
If opt1Second.Value = True Then Timer1.Interval = 1000
lstbxSystemDialog.AddItem Time & " : Changed Update Frequency to 1 Second."
End Sub

Private Sub opt2Seconds_Click()
If opt2Seconds.Value = True Then Timer1.Interval = 2000
lstbxSystemDialog.AddItem Time & " : Changed Update Frequency to 2 Seconds."
End Sub

Private Sub opt5Seconds_Click()
If opt5Seconds.Value = True Then Timer1.Interval = 5000
lstbxSystemDialog.AddItem Time & " : Changed Update Frequency to 5 Seconds."
End Sub

Private Sub optCustom_Click()
Dim reply As String
Dim temp As Double
Comehere:
If optCustom.Value = True Then
reply = InputBox("Enter desired update frequency in seconds.")
If IsNumeric(reply) Then
    If CLng(reply) > 65 Then
        MsgBox "You have entered a value larger than 65 seconds. This is not supported. Using 65 seconds as max."
        reply = CStr(65)
    End If
    temp = CLng(reply) * 1000
    Timer1.Interval = temp
    lstbxSystemDialog.AddItem Time & " : Changed update frequency to every " & reply & " seconds."
    lblCustom.Caption = "Every " & reply & " Seconds"
Else
    MsgBox "You did not enter a number. Try again"
    GoTo Comehere
End If
End If
End Sub

Private Sub Timer1_Timer()

lstvwProcesses.ListItems.Clear

Dim Itmx As ListItem
Dim j, k As Long
Dim cb As Long
Dim reply As Long
Dim temp As String * 200
Dim cbNeeded As Long
Dim NumElements As Long
Dim ProcessIDs() As Long
Dim cbNeeded2 As Long
Dim NumElements2 As Long
Dim Modules(1 To 200) As Long
Dim lRet As Long
Dim ModuleName As String
Dim nSize As Long
Dim hProcess As Long
Dim i As Long
Dim Priority As Long

                                                                               'Get the array containing the process id's for each process object
cb = 8
cbNeeded = 96
Do While cb <= cbNeeded
    cb = cb * 2
    ReDim ProcessIDs(cb / 4) As Long
    lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
Loop
         
NumElements = cbNeeded / 4
j = 1                                                                           'j keeps track of index for lstvwProcesses
For i = 1 To NumElements
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, ProcessIDs(i))                'Get a handle to the Process
    If hProcess <> 0 Then                                                       'Got a Process handle
        lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)         'Get an array of the module handles for the specified process
        If lRet <> 0 Then                                                       'If the Module Array is retrieved, Get the ModuleFileName
            ModuleName = Space(MAX_PATH)                                        'Prepare variables...
            nSize = 500
            lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize) 'Get process name
            Handles(i) = hProcess                                               'Assign handle to array
            Processes(i) = Trim(Right(ModuleName, _
                (Len(ModuleName) - InStrRev(ModuleName, "\"))))                 'Assign name to array
            Set Itmx = lstvwProcesses.ListItems.Add(j, , Processes(i))          'Add process to Process List
            Priority = GetPriorityClass(Handles(i))                             'Retrieve process priority
            AddProcessListSubItems Priority, Itmx                               'Add sub item info to Process List
            ChangePriority i, Priority                                          'Call Change Priority Sub Routine
            j = j + 1
        End If
        lRet = CloseHandle(hProcess)                                            'Close the handle to the process
    End If
Next
lstbxSystemDialog.ListIndex = lstbxSystemDialog.ListCount - 1                   'select the most recent entry
End Sub

Sub ReDefineChangeList()                                                        'This sub updates the Changelist type
Dim Itmx As ListItem                                                            'so that it corresponds to what is seen
Dim temp As String                                                              'in the lstvwChangeList

For i = 1 To lstvwChangeList.ListItems.Count
    Set Itmx = lstvwChangeList.ListItems.Item(i)
    ChangeList(i).Process = Itmx.Text
    
    temp = Itmx.SubItems(1)
    Select Case temp
        Case "Idle":
            ChangeList(i).From = IDLE_PRIORITY_CLASS
        Case "Normal":
            ChangeList(i).From = NORMAL_PRIORITY_CLASS
        Case "High":
            ChangeList(i).From = HIGH_PRIORITY_CLASS
        Case "Highest":
            ChangeList(i).From = REALTIME_PRIORITY_CLASS
    End Select
    
    temp = Itmx.SubItems(2)
    Select Case temp
        Case "Idle":
            ChangeList(i).To = IDLE_PRIORITY_CLASS
        Case "Normal":
            ChangeList(i).To = NORMAL_PRIORITY_CLASS
        Case "High":
            ChangeList(i).To = HIGH_PRIORITY_CLASS
        Case "Highest":
            ChangeList(i).To = REALTIME_PRIORITY_CLASS
    End Select
Next i

End Sub

Sub AddProcessListSubItems(Priority As Long, Itmx As ListItem)                      'This routine adds to the lstvwProcesses priority header
Select Case Priority                                                                'Add Item case
    Case NORMAL_PRIORITY_CLASS:
         Itmx.SubItems(1) = "Normal"
    Case IDLE_PRIORITY_CLASS:
         Itmx.SubItems(1) = "Idle"
    Case HIGH_PRIORITY_CLASS:
         Itmx.SubItems(1) = "High"
    Case REALTIME_PRIORITY_CLASS:
         Itmx.SubItems(1) = "Highest"
End Select
End Sub

Sub ChangePriority(i As Long, Priority As Long)                                     'This routine checks the priority with the Changelist
Dim k As Integer                                                                    'and changes the priority if necessary

For k = 1 To lstvwChangeList.ListItems.Count                                        'Begin loop that cycles through Change List
    If InStr(1, UCase(Processes(i)), UCase(ChangeList(k).Process)) Then             'Is Process in Change List?
        If Priority = ChangeList(k).From Then                                       'Does Priority need to be changed?
            reply = SetPriorityClass(Handles(i), ChangeList(k).To)                  'change priority
            If reply = 0 Then                                                       'If error, then explain..
                reply = GetLastError
                If reply = 5 Then
                    lstbxSystemDialog.AddItem Time & " : Error #: " & reply & " occured for process: " & Processes(i)
                End If
            Else
                lstbxSystemDialog.AddItem Time & " : Changed priority of process: " & _
                    Processes(i)                                                    'Successfully changed
            End If
        End If
    End If
Next k
End Sub

Sub SaveChangeList()                                                                'This routine saves the Change List based on
Dim fNum As Integer                                                                 'the information in the lstvwChangeList treeview
Dim Itmx As ListItem
Dim temp As String
Dim i As Integer

fNum = FreeFile                                                                     'Get available File number

Open (App.Path & "\setup.txt") For Output As fNum                                   'Open File

For i = 1 To lstvwChangeList.ListItems.Count                                        'Cylce through list items and save
    Set Itmx = lstvwChangeList.ListItems.Item(i)
    temp = Itmx.Text & Chr(9) & Itmx.SubItems(1) & Chr(9) & Itmx.SubItems(2)
    Print #fNum, temp                                                               'Print to file
Next i

lstbxSystemDialog.AddItem Time & " : Saved new Change List..."                      'Update console
Close fNum                                                                          'Close file

End Sub

