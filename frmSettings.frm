VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Device Settings"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   315
      Left            =   2340
      TabIndex        =   3
      Top             =   780
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Top             =   780
      Width           =   1095
   End
   Begin VB.ComboBox cmbVideo 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   4515
   End
   Begin VB.Label lblVideo 
      AutoSize        =   -1  'True
      Caption         =   "Video device:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   975
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sGUID() As String
Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdStart_Click()
    Call Start(sGUID(cmbVideo.ListIndex))
End Sub

Private Sub Form_Load()
    Dim d3dEnum As Direct3DEnumDevices
    Dim lK As Long
    
    '// Enumerate Devices
    Set mDDraw = mDX.DirectDrawCreate("")
    Set mD3D = mDDraw.GetDirect3D()
    Set d3dEnum = mD3D.GetDevicesEnum()
    
    cmbVideo.Clear
    ReDim sGUID(d3dEnum.GetCount() - 1)
    For lK = 1 To d3dEnum.GetCount()
        cmbVideo.AddItem d3dEnum.GetName(lK)
        sGUID(lK - 1) = d3dEnum.GetGuid(lK)
    Next lK
    
    If cmbVideo.ListCount > 0 Then cmbVideo.ListIndex = cmbVideo.ListCount - 1
    
    Set d3dEnum = Nothing
    Set mD3D = Nothing
    Set mDDraw = Nothing
End Sub


