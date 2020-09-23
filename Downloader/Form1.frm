VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Downloader"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSave 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Text            =   "C:\Temp\Test.exe"
      Top             =   1140
      Width           =   6195
   End
   Begin MSComctlLib.ProgressBar pB 
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin Downloader.BajaArchivos j 
      Height          =   720
      Left            =   5460
      TabIndex        =   2
      Top             =   180
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download"
      Height          =   435
      Left            =   6180
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtURL 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Text            =   "http://download.microsoft.com/download/speechSDK/SDK/5.1/WXP/EN-US/speechsdk51msm.exe"
      Top             =   0
      Width           =   7395
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Save As..."
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   1260
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   2760
      TabIndex        =   5
      Top             =   420
      Width           =   480
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   420
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
If Command1.Caption = "Cancel" Then
    j.CancelDownload
    Command1.Caption = "Download"
Else
    j.Download txtURL.Text, txtSave.Text
    Command1.Caption = "Cancel"
End If
End Sub

Private Sub j_Completed(Bytes As Long, sId As String)
 Me.Caption = "DOWNLOAD COMPLETED" & Bytes & sId
End Sub

Private Sub j_Progress(DownLoadedBytes As Long, TotalBytes As Long, sId As String)
If DownLoadedBytes > 0 Then
    pB.Value = (DownLoadedBytes * 100) / TotalBytes
    lbl1.Caption = DownLoadedBytes & "/" & TotalBytes
    Label1.Caption = pB.Value & "%"
End If
End Sub
