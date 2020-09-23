VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form TabStrips1 
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   4920
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Cmd3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4815
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8493
      TabWidthStyle   =   2
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "One"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Two"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "There"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "TabStrips1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TabStrip1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    If TabStrip1.SelectedItem = TabStrip1.Tabs.Item(1) Then
        Cmd1.Visible = True
        Cmd2.Visible = False
        Cmd3.Visible = False
    ElseIf TabStrip1.SelectedItem = TabStrip1.Tabs.Item(2) Then
        Cmd1.Visible = False
        Cmd2.Visible = True
        Cmd3.Visible = False
    ElseIf TabStrip1.SelectedItem = TabStrip1.Tabs.Item(3) Then
        Cmd1.Visible = False
        Cmd2.Visible = False
        Cmd3.Visible = True
    End If
    Timer1.Enabled = False
End Sub
