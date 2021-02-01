VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Most Recently Used Items v1.0"
   ClientHeight    =   2685
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   4620
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   405
      Left            =   1680
      TabIndex        =   1
      Top             =   2205
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   405
      Left            =   3045
      TabIndex        =   3
      Top             =   2205
      Width           =   1275
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   405
      Left            =   315
      TabIndex        =   2
      Top             =   2205
      Width           =   1275
   End
   Begin VB.Frame fraMain 
      Caption         =   "Select or Input Item"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   105
      TabIndex        =   4
      Top             =   105
      Width           =   4425
      Begin VB.ComboBox cboItems 
         Height          =   1545
         Left            =   105
         Style           =   1  'Simple Combo
         TabIndex        =   0
         Text            =   "cboItems"
         Top             =   315
         Width           =   4215
      End
   End
   Begin VB.Menu mnuItems 
      Caption         =   "&Items"
      Begin VB.Menu mnuItemsSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuItemsClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuItemsClearAll 
         Caption         =   "Clear &All"
      End
      Begin VB.Menu mnuItemsLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuItemsLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemsRecent 
         Caption         =   "Recent"
         Index           =   1
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContent 
         Caption         =   "&Content"
      End
      Begin VB.Menu mnuHelpLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mruItems As New classMRU

Private Sub Form_Load()
Dim i%
    mruItems.MaxNum = 5
    mruItems.Parent = "Company"
    mruItems.Child = "Application"
    mruItems.Prefix = "Item"
    '
    For i = 2 To mruItems.MaxNum
            Load mnuItemsRecent(i)
    Next i
    '
    FillCombo
    FillMenus
End Sub


Private Sub FillCombo()
Dim Items() As String
Dim i%, num%
    cboItems.Clear
    mruItems.GetArray Items()
    num = mruItems.num
    If num > 0 Then
        For i = 1 To num
            cboItems.AddItem Items(i)
        Next i
    End If
End Sub

Private Sub FillMenus()
Dim Items() As String
Dim i%, num%
On Error Resume Next
    mruItems.GetArray Items()
    num = mruItems.num
    If num > 0 Then
        For i = 1 To num
            mnuItemsRecent(i).Caption = Items(i)
            mnuItemsRecent(i).Visible = True
        Next i
    End If
On Error GoTo 0
End Sub

Private Sub SaveItem(item)
    mruItems.Save item
    FillCombo
    FillMenus
End Sub

Private Sub cboItems_Click()
    SaveItem (cboItems.Text)
End Sub

Private Sub cmdClear_Click()
Dim i%
    If MsgBox("Clear list of most recently used items?", vbYesNo + vbDefaultButton1 + vbExclamation, "Clear...") = vbYes Then
        mruItems.Clear
        cboItems.Clear
        For i = 1 To mruItems.MaxNum
            mnuItemsRecent(i).Visible = False
        Next i
    End If
End Sub

Private Sub cmdSave_Click()
    If cboItems.Text <> "" Then
        SaveItem (cboItems.Text)
    End If
    cboItems.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub mnuItemsSave_Click()
    cmdSave_Click
End Sub
Private Sub mnuItemsClear_Click()
    cmdClear_Click
End Sub
Private Sub mnuItemsClearAll_Click()
    cmdClear_Click
    mruItems.ClearParent
End Sub
Private Sub mnuItemExit_Click()
    Unload Me
End Sub

Private Sub mnuItemsRecent_Click(Index As Integer)
    SaveItem (mnuItemsRecent(Index).Caption)
End Sub

Private Sub mnuHelpContent_Click()
    MsgBox "Help Content", vbOKOnly + vbInformation, "Content..."
End Sub
Private Sub mnuHelpAbout_Click()
    MsgBox "Most Recently Used Items v1.0", vbOKOnly + vbInformation, "About..."
End Sub

