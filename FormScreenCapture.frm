VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FormScreenCapture 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Screen Capture Demo "
   ClientHeight    =   8505
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleMode       =   0  'User
   ScaleWidth      =   20804.05
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   14160
      Top             =   7680
   End
   Begin VB.PictureBox Picture2 
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   4515
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Portcheck"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   2
      Top             =   6120
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   1
      Top             =   6120
      Width           =   615
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   14760
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   7320
      ScaleHeight     =   3075
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   1320
      Width           =   4815
   End
End
Attribute VB_Name = "FormScreenCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal opCode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Dim hwdc As Long
Dim startcap As Boolean
Dim buffer As Variant
Dim aa As Integer
Dim str As String
 Dim temp() As Byte
 Dim val As Integer
 Dim go, n
Private Sub Command3_Click()
 If (MSComm1.PortOpen = True) Then
        MsgBox "Com1 Port is open"
    Else
        MsgBox "Port is not open"
 End If
End Sub
Private Sub Form_Load()

MSComm1.RThreshold = 1
MSComm1.InputLen = 1
MSComm1.Settings = "9600,N,8,1"
MSComm1.DTREnable = True
MSComm1.CommPort = 1
MSComm1.PortOpen = True

Dim temp1 As Long
    hwdc = capCreateCaptureWindow("Dixanta Vision System", ws_child Or ws_visible, 0, 0, 320, 240, Picture1.hWnd, 0)
  If (hwdc <> 0) Then
    temp1 = SendMessage(hwdc, wm_cap_driver_connect, 0, 0)
    temp1 = SendMessage(hwdc, wm_cap_set_preview, 1, 0)
    temp1 = SendMessage(hwdc, WM_CAP_SET_PREVIEWRATE, 30, 0)
    startcap = True
  Else
    MsgBox ("No Webcam found")
  End If
End Sub
Private Sub Timer1_Timer()
    
    If (MSComm1.CommEvent = comEvReceive) Then
            buffer = MSComm1.Input
            temp = buffer
            val = temp(0)
    
        If (val = 1) Then
                Text1.Text = "1"
             Dim scrHwnd As Long
                scrHwnd = GetDesktopWindow
             Dim shDC As Long
                    shDC = GetDC(scrHwnd)
                    CreateCompatibleBitmap shDC, 500, 500
                    BitBlt FormScreenCapture.hDC, 0, 0, 300, 300, shDC, 500, 200, vbSrcCopy
                
                    FormScreenCapture.Picture = FormScreenCapture.Image
                    Picture2.Picture = FormScreenCapture.Picture
                    aa = aa + abc()
                    str = App.Path & "\Picture\" & "Snapshot" & aa & ".jpg"
                    SavePicture Picture2.Picture, str
     
        Else
               Text1.Text = "0"
        End If
  
     End If

End Sub
