VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":0000
   MousePointer    =   1  'Arrow
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Blackness 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2880
      Top             =   1080
   End
   Begin VB.Timer BaskInGlory 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   1680
      Top             =   1080
   End
   Begin VB.Timer ShutDown 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   1080
   End
   Begin VB.Timer PhysicalMemoryDump 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   1080
   End
   Begin VB.Label Message 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BaskInGlory_Timer()

    ShutDown.Enabled = True
    BaskInGlory.Enabled = False

End Sub

Private Sub Blackness_Timer()

    BlockInput False
    Blackness.Enabled = False
    End
    
End Sub

Private Sub Form_Load()
    PhysicalMemoryDump.Enabled = True

    Me.BackColor = RGB(0, 0, 255)
    Me.Message.Width = Screen.Width
    Me.Message.Height = Screen.Height
    Me.WindowState = 2
    Me.MousePointer = vbCustom
    BlockInput True
    Call SetError(vbNullString)

End Sub

Public Sub SetError(ByVal X As String)
Dim Text As String

    Text = Text + "*** Stop: 0x0000001E <0xF24A447A, 0x00000001, 0x0000000>" & vbCrLf & "KMDE_EXEPTION_NOT_HANDLED" & vbCrLf & vbCrLf & vbCrLf
    
    Text = Text + "A problem has been detected and windows has been shut down to prevent damage to your computer." & vbCrLf & vbCrLf
    Text = Text + "A process or thread crucial to system operation has unexpectedly exited or been terminated." & vbCrLf
    Text = Text + "If this is the first time you've seen this Stop error screen," & vbCrLf
    Text = Text + "restart your computer. If this screen appears again, follow" & vbCrLf
    Text = Text + "these steps:" & vbCrLf & vbCrLf
    Text = Text + "Check to make sure any new hardware or software is properly installed." & vbCrLf
    Text = Text + "If this is a new installation, ask your hardware or software manufacturer" & vbCrLf
    Text = Text + "for any windows updates you might need." & vbCrLf & vbCrLf
    
    Text = Text + "If problems continue, disable or remove any newly installed hardware" & vbCrLf
    Text = Text + "or software. Disable BIOS memory options such as caching or shadowing." & vbCrLf
    Text = Text + "If you need to use Safe Mode to remove or disable components, restart" & vbCrLf
    Text = Text + "your computer. Press F8 to select Advanced Startup Options, and then" & vbCrLf
    Text = Text + "select Safe Mode." & vbCrLf & vbCrLf & vbCrLf
    
    Text = Text + "Technical information:" & vbCrLf & vbCrLf
    Text = Text + "*** STOP: 0x00000C#4  (0x00000003, 0x81C99020, 0x81C99194, 0x80SFAI9A)" & vbCrLf
    Text = Text + "*** ati3diag.dll - Address ED80AC55 base at ED88f000, Date Stamp 3dcb24d0" & vbCrLf & vbCrLf
    Text = Text + "Kernel Debugger Using: COM2 <Port 0x2f8, Baud Rate 19200>" & vbCrLf & vbCrLf & vbCrLf
    Text = Text + X
    Me.Message.Caption = Text
End Sub

Private Sub Form_Unload(Cancel As Integer)

    BlockInput False

End Sub

Private Sub PhysicalMemoryDump_Timer()
Dim Text As String

    Tick = Tick + RAND(1, 6)

    If Tick >= 100 Then
        Tick = 100
        PhysicalMemoryDump.Enabled = False
        Call SetError("Collecting data for crash dump ..." & vbCrLf & "Initializing disk for crash dump ..." & vbCrLf & "Beginning dump of physical memory..." & vbCrLf & "Dumping physical memory to disk: " & Tick & vbCrLf & "Physical memory dump complete." & vbCrLf & "Contact your system administrator, your technical support group, or Bill Gates" & vbCrLf & "for further assistance." & vbCrLf)
        RGBValue = 255
        BaskInGlory.Enabled = True
    Else
        PhysicalMemoryDump.Interval = RAND(1000, 2250)
        Call SetError("Collecting data for crash dump ..." & vbCrLf & "Initializing disk for crash dump ..." & vbCrLf & "Beginning dump of physical memory..." & vbCrLf & "Dumping physical memory to disk: " & Tick)
    End If

End Sub

Private Sub ShutDown_Timer()
RGBValue = RGBValue - 3
If RGBValue <= 0 Then
    ShutDown.Enabled = False
    Blackness.Enabled = True
    Exit Sub
Else
    Me.BackColor = RGB(0, 0, RGBValue)
    Me.Message.ForeColor = RGB(RGBValue, RGBValue, RGBValue)
End If
End Sub
