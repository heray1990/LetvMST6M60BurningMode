Attribute VB_Name = "ModuleUartCmd"
'**********************************************
' Module for burning mode of Letv MST6M60.
'**********************************************
Option Explicit

Private mSendDataBuf(0 To 10) As Byte
Private mDDCDataWithoutChksum(0 To 5) As Byte
Private i As Integer

Private Sub SendCmd()
    Form1.MSComm1.Output = mSendDataBuf
    DelayMS 500
End Sub

Private Function CalDDCChkSum(ByRef data() As Byte) As Byte
    Dim tmp As Integer

    tmp = 0
    CalDDCChkSum = &H0

    For i = 0 To 5
        tmp = tmp + data(i)
    Next i
    
    CalDDCChkSum = tmp And &HF
End Function

Private Function CalChkSum(ByRef data() As Byte) As Byte
    Dim tmp As Integer

    tmp = 0
    CalChkSum = &H0

    For i = 0 To 9
        tmp = tmp + data(i)
    Next i
    
    CalChkSum = &HFF - tmp And &HFF
End Function

Private Sub DataToDDC()
    For i = 0 To 5
        mDDCDataWithoutChksum(i) = mSendDataBuf(i + 4)
    Next i
End Sub

Public Sub PanelPattern(intBurningMode As Integer)
    'E0 0B 40 X7 0F 01 XX 00 00 00 CHK
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &HF
    mSendDataBuf(5) = &H1
    mSendDataBuf(6) = intBurningMode
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0

    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &H7
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)

    SendCmd
End Sub

Public Sub BurningMode(intOnOff As Integer)
    'E0 0B 40 XD 10 XX 00 00 00 00 CHK
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &H10
    mSendDataBuf(5) = intOnOff
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0

    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &HD
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)

    SendCmd
End Sub

Public Sub RebootMonitor()
    'E0 0B 40 2D 12 00 00 00 00 00 95
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(3) = &H2D
    mSendDataBuf(4) = &H12
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = &H95

    SendCmd
End Sub
