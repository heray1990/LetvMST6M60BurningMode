Attribute VB_Name = "ModuleUartCmd"
'**********************************************
' Module for burning mode of Letv MST6M60.
'**********************************************
Option Explicit

Private mSendDataBuf(0 To 10) As Byte

Private Sub SendCmd()
    Form1.MSComm1.Output = mSendDataBuf
    DelayMS 500
End Sub

Public Sub BurningMode01()
    'E0 0B 40 17 0F 01 01 00 00 00 AC
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(3) = &H17
    mSendDataBuf(4) = &HF
    mSendDataBuf(5) = &H1
    mSendDataBuf(6) = &H1
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = &HAC

    SendCmd
End Sub

Public Sub BurningMode02()
    'E0 0B 40 27 0F 01 02 00 00 00 9B
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(3) = &H27
    mSendDataBuf(4) = &HF
    mSendDataBuf(5) = &H1
    mSendDataBuf(6) = &H2
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = &H9B

    SendCmd
End Sub
