VERSION 5.00
Begin VB.Form FrmComPort 
   Caption         =   "设置电视串口"
   ClientHeight    =   1935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3735
   Icon            =   "FrmComPort.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "ComSet"
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton cmdSet 
         Caption         =   "设置"
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "取消"
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "TV"
         ForeColor       =   &H000000C0&
         Height          =   1455
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2175
         Begin VB.ComboBox cmbComBaud 
            Height          =   300
            Left            =   960
            TabIndex        =   3
            Text            =   "9600"
            Top             =   840
            Width           =   975
         End
         Begin VB.ComboBox cmbComID 
            Height          =   300
            ItemData        =   "FrmComPort.frx":1DF72
            Left            =   960
            List            =   "FrmComPort.frx":1DF74
            TabIndex        =   2
            Text            =   "COM1"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "波特率:"
            Height          =   200
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   900
            Width           =   700
         End
         Begin VB.Label Label1 
            Caption         =   "串口:"
            Height          =   200
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   450
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "FrmComPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error GoTo ErrExit
    Dim i As Integer

    cmbComID.Text = "COM" & CStr(gintTVComID)
    cmbComBaud.Text = CStr(glngTVComBaud)

    For i = 1 To 20
        cmbComID.AddItem "COM" & i
    Next i

    cmbComBaud.AddItem "9600"
    cmbComBaud.AddItem "19200"
    cmbComBaud.AddItem "38400"
    cmbComBaud.AddItem "57600"
    cmbComBaud.AddItem "115200"

    Exit Sub
ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Form1.ZOrder (0)
End Sub

Private Sub cmdSet_Click()
    Dim clsSaveConfigData As ProjectConfig
    
    Set clsSaveConfigData = New ProjectConfig

    clsSaveConfigData.ComBaud = cmbComBaud.Text
    clsSaveConfigData.ComID = Val(Replace(cmbComID.Text, "COM", ""))
    
    clsSaveConfigData.SaveConfigData
    
    Set clsSaveConfigData = Nothing

    Unload Me
    
    Form1.SubInit
    Form1.Show
End Sub

