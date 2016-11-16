VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProjectConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************
' Class module for handling config.xml of the
' application.
'**********************************************

Option Explicit

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Type udtConfigData
    strComBaud As String
    intComID As Integer
End Type

Private mConfigData As udtConfigData
Private mstrConfigFilePath As String

Private Sub Class_Initialize()
    mstrConfigFilePath = App.Path & "\config.xml"
    mConfigData.strComBaud = "115200"
    mConfigData.intComID = 1
End Sub

Public Sub LoadConfigData()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean

    If Not CBool(PathFileExists(mstrConfigFilePath)) Then
        MsgBox "Cannot open " & mstrConfigFilePath & " file."
        End
    End If
    
    success = xmlDoc.Load(mstrConfigFilePath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        mConfigData.strComBaud = xmlDoc.selectSingleNode("/config/common").selectSingleNode("@baud").Text
        mConfigData.intComID = Val(xmlDoc.selectSingleNode("/config/common").selectSingleNode("@id").Text)
    End If
End Sub

Public Sub SaveConfigData()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(mstrConfigFilePath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        xmlDoc.selectSingleNode("/config/common").selectSingleNode("@baud").Text = mConfigData.strComBaud
        xmlDoc.selectSingleNode("/config/common").selectSingleNode("@id").Text = CStr(mConfigData.intComID)
        
        xmlDoc.Save mstrConfigFilePath
    End If
End Sub

Public Property Get ComBaud() As String
    ComBaud = mConfigData.strComBaud
End Property

Public Property Let ComBaud(strComBaud As String)
    mConfigData.strComBaud = strComBaud
End Property

Public Property Get ComID() As Integer
    ComID = mConfigData.intComID
End Property

Public Property Let ComID(intComID As Integer)
    mConfigData.intComID = intComID
End Property
