VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private DX As New DirectX7
Private DI As DirectInput
Private diDev As DirectInputDevice
Private diState As DIKEYBOARDSTATE

Public Sub Init(hWnd As Long)

Set DI = DX.DirectInputCreate()

Set diDev = DI.CreateDevice("GUID_SysKeyboard")

diDev.SetCommonDataFormat DIFORMAT_KEYBOARD

diDev.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE

diDev.Acquire

End Sub

Public Function State() As DIKEYBOARDSTATE
    diDev.GetDeviceStateKeyboard State
End Function
