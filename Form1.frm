VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Letv Max65 Burning Mode Tool"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnMode03 
      Caption         =   "Mode 03"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6050
      TabIndex        =   4
      Top             =   1080
      Width           =   2500
   End
   Begin VB.CommandButton BtnMode02 
      Caption         =   "Mode 02"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Width           =   2500
   End
   Begin VB.CommandButton BtnMode01 
      Caption         =   "Mode 01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2500
   End
   Begin VB.PictureBox PictureLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   758
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   735
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   120
      Width           =   2528
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3960
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label LbToolName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Max65 Burning Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   5900
   End
   Begin VB.Menu MenuItemSetting 
      Caption         =   "Setting"
      Begin VB.Menu MenuItemComSetting 
         Caption         =   "COM Setting"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SubInit()
    Dim clsConfigData As ProjectConfig

    Set clsConfigData = New ProjectConfig
    clsConfigData.LoadConfigData
    
    glngTVComBaud = clsConfigData.ComBaud
    gintTVComID = clsConfigData.ComID
    SubInitComPort

    Set clsConfigData = Nothing
End Sub

Private Sub SubInitComPort()
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

    MSComm1.CommPort = gintTVComID
    MSComm1.Settings = glngTVComBaud & ",N,8,1"
    MSComm1.InputLen = 0

    MSComm1.InBufferCount = 0
    MSComm1.OutBufferCount = 0
    MSComm1.InputMode = comInputModeBinary

    MSComm1.NullDiscard = False
    MSComm1.DTREnable = False
    MSComm1.EOFEnable = False
    MSComm1.RTSEnable = False
    MSComm1.SThreshold = 1
    MSComm1.RThreshold = 1
    MSComm1.InBufferSize = 1024
    MSComm1.OutBufferSize = 512
End Sub

Private Sub BtnMode01_Click()
On Error GoTo ErrExit
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    BurningMode 0
    Exit Sub
ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub BtnMode02_Click()
On Error GoTo ErrExit
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    BurningMode 1
    Exit Sub
ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub BtnMode03_Click()
On Error GoTo ErrExit
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    BurningMode 2
    Exit Sub
ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Form_Load()
    SubInit
End Sub

Private Sub MenuItemComSetting_Click()
    FrmComPort.Show
End Sub
