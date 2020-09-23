VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SuperGT"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3525
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Game Setup"
      Height          =   2040
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   3390
      Begin VB.CheckBox chkJoystick 
         Caption         =   "Use Joystick"
         Height          =   195
         Left            =   1755
         TabIndex        =   11
         Top             =   1185
         Width           =   1380
      End
      Begin VB.Frame Frame3 
         Caption         =   "Race"
         Height          =   705
         Left            =   1740
         TabIndex        =   8
         Top             =   255
         Width           =   1530
         Begin VB.TextBox txtLaps 
            Height          =   285
            Left            =   885
            TabIndex        =   9
            Text            =   "0"
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label2 
            Caption         =   "Laps:"
            Height          =   225
            Left            =   135
            TabIndex        =   10
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Graphics"
         Height          =   1680
         Left            =   135
         TabIndex        =   3
         Top             =   255
         Width           =   1530
         Begin VB.CheckBox chkFlat 
            Caption         =   "Render Flat"
            Height          =   210
            Left            =   105
            TabIndex        =   13
            Top             =   465
            Width           =   1350
         End
         Begin VB.CheckBox chkFiltering 
            Caption         =   "Bilinear Filtering"
            Height          =   225
            Left            =   105
            TabIndex        =   12
            Top             =   240
            Width           =   1395
         End
         Begin VB.CheckBox chkFS 
            Caption         =   "Use Fullscreen"
            Height          =   330
            Left            =   105
            TabIndex        =   7
            Top             =   630
            Width           =   1380
         End
         Begin VB.OptionButton optScreenRes 
            Caption         =   "320x240"
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   6
            Top             =   930
            Width           =   1005
         End
         Begin VB.OptionButton optScreenRes 
            Caption         =   "512x384"
            Height          =   225
            Index           =   1
            Left            =   105
            TabIndex        =   5
            Top             =   1155
            Width           =   1080
         End
         Begin VB.OptionButton optScreenRes 
            Caption         =   "640x480"
            Height          =   210
            Index           =   2
            Left            =   105
            TabIndex        =   4
            Top             =   1365
            Width           =   1050
         End
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   360
         Left            =   2535
         TabIndex        =   2
         Top             =   1545
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   360
         Left            =   1725
         TabIndex        =   1
         Top             =   1545
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ScreenX As String
Dim ScreenY As String

Dim UseFS As String
Dim UseJoystick As String
Dim UseFiltering As String
Dim UseFlat As String
Dim TotalLaps As String


Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    ChDir App.Path
    
    If chkFS.Value = 1 Then UseFS = "TRUE" Else UseFS = "FALSE"
    TotalLaps = txtLaps.Text
    If optScreenRes(0).Value = True Then
        ScreenX = 320
        ScreenY = 240
    ElseIf optScreenRes(1).Value = True Then
        ScreenX = 512
        ScreenY = 384
    ElseIf optScreenRes(2).Value = True Then
        ScreenX = 640
        ScreenY = 480
    End If
    If chkJoystick.Value = 1 Then UseJoystick = "TRUE" Else UseJoystick = "FALSE"
    If chkFiltering.Value = 1 Then UseFiltering = "TRUE" Else UseFiltering = "FALSE"
    If chkFlat.Value = 1 Then UseFlat = "TRUE" Else UseFlat = "FALSE"
    On Error GoTo Oops
    
    Open "supergt.cfg" For Output As #1
    Print #1, UseFS
    Print #1, TotalLaps
    Print #1, ScreenX
    Print #1, ScreenY
    Print #1, UseJoystick
    Print #1, UseFiltering
    Print #1, UseFlat
    
    Close #1
    
    GoTo Success
    
Oops:
    MsgBox "Sorry, config file not found", , "Error"
    End
Success:

End Sub

Private Sub Form_Load()
    ChDir App.Path
    
    On Error GoTo NoSuchFile
    Open "supergt.cfg" For Input As #1
    
    Input #1, UseFS
    Input #1, TotalLaps
    Input #1, ScreenX
    Input #1, ScreenY
    Input #1, UseJoystick
    Input #1, UseFiltering
    Input #1, UseFlat
    
    If UseFS = "TRUE" Then chkFS.Value = 1 Else chkFS.Value = 0
    If UseJoystick = "TRUE" Then chkJoystick.Value = 1
    If UseFiltering = "TRUE" Then chkFiltering.Value = 1 Else chkFiltering.Value = 0
    If UseFlat = "TRUE" Then chkFlat.Value = 1
    txtLaps.Text = TotalLaps
    
    If ScreenX = "320" Then optScreenRes(0).Value = True
    If ScreenX = "512" Then optScreenRes(1).Value = True
    If ScreenX = "640" Then optScreenRes(2).Value = True
    
    Close #1
    
    GoTo Success
    
NoSuchFile:
    MsgBox "Sorry, config file not found", , "Error"
    End
Success:

End Sub

Private Sub optScreenRes_Click(Index As Integer)
    If optScreenRes(Index).Value = True Then Exit Sub
    If Index = 0 Then
        ScreenX = 320
        ScreenY = 240
    ElseIf Index = 1 Then
        ScreenX = 512
        ScreenY = 384
    ElseIf Index = 2 Then
        ScreenX = 640
        ScreenY = 480
    End If
        
End Sub
