VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Autostart Selector"
   ClientHeight    =   3120
   ClientLeft      =   2805
   ClientTop       =   4275
   ClientWidth     =   10155
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Add"
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Info"
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Update"
      Height          =   300
      Left            =   3000
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear list 2"
      Height          =   300
      Left            =   5520
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   300
      Left            =   8160
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save to registry"
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "=>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.ListBox AppList2 
      Height          =   2205
      Left            =   5520
      TabIndex        =   3
      Top             =   240
      Width           =   4335
   End
   Begin VB.ListBox AppList 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Programs that starts with Windows"
      Height          =   2655
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Programs that you DONT want to start with Windows"
      Height          =   2655
      Left            =   5400
      TabIndex        =   8
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub setup()
Dim temp, temp2 As Variant
Dim i, j, x, y As Integer
Dim Program As String
AppList.Clear
AppList2.Clear

temp = GetAllValues(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\")
temp2 = GetAllValues(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices\")


For x = 0 To UBound(temp)
 Program = temp(x, 0)
 If Mid(GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\", Program), 1, 4) = "rem " Then
  AppList2.AddItem temp(x, 0)
  temp(x, 0) = "%" & temp(x, 0)
 End If
Next

For y = 0 To UBound(temp2)
 Program = temp2(y, 0)
 If Mid(GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices\", Program), 1, 4) = "rem " Then
  AppList2.AddItem temp2(y, 0)
  temp2(y, 0) = "%" & temp2(y, 0)
 End If
Next


For i = 0 To UBound(temp)
 If Not (Mid(temp(i, 0), 1, 1)) = "%" Then AppList.AddItem temp(i, 0)
Next

For j = 0 To UBound(temp2)
 If Not (Mid(temp2(j, 0), 1, 1)) = "%" Then AppList.AddItem temp2(j, 0)
Next




End Sub


Private Sub Command1_Click()
If AppList2.ListIndex > -1 Then
 AppList.AddItem AppList2.List(AppList2.ListIndex)
 AppList2.RemoveItem (AppList2.ListIndex)
End If
End Sub

Private Sub Command2_Click()
If AppList.ListIndex > -1 Then
 AppList2.AddItem AppList.List(AppList.ListIndex)
 AppList.RemoveItem (AppList.ListIndex)
End If




End Sub

Private Sub Command3_Click()
Dim i, j As Integer
Dim temp As String
Dim Service As Boolean

If AppList2.ListCount = 0 Then GoTo 1:


For i = 0 To AppList2.ListCount - 1



If AppList2.ListCount > 0 Then
 Service = False
 temp = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\", AppList2.List(i))
 If temp = "" Then
  temp = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices\", AppList2.List(i))
  Service = True
 End If
End If

If Not (Mid(temp, 1, 4)) = "rem " Then temp = "rem " & temp

If Service = False Then
 SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\", AppList2.List(i), temp
Else
 SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices\", AppList2.List(i), temp
End If





Next
1:

For j = 0 To AppList.ListCount - 1
 Service = False
 temp = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\", AppList.List(j))
 If temp = "" Then
  temp = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices\", AppList.List(j))
  Service = True
 End If

 If Mid(temp, 1, 3) = "rem" Then
  temp = Mid(temp, 5, Len(temp))
   If Service = False Then
   SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\", AppList.List(j), temp
  Else
   SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices\", AppList.List(j), temp
  End If
 End If

Next






End Sub

Private Sub Command4_Click()
End
End Sub


Private Sub Command5_Click()
AppList2.Clear
End Sub

Private Sub Command6_Click()
setup
End Sub

Private Sub Command7_Click()
MsgBox "This program was made by Martin Dahl, you can contact me on ICQ:1380740", vbInformation, "About"
End Sub

Private Sub Command8_Click()
Dim Name As String
Dim Path As String

Name = InputBox("Enter the name of the program that you want to autostart with windows", "Name input")
Path = InputBox("Enter the program path", "Path input")

If Path = "" Or Name = "" Then
 MsgBox "Please enter both path and name", vbCritical, "error"
 Exit Sub
End If

SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\", Name, Path


End Sub

Private Sub Form_Load()
Call setup



End Sub


