VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmASCII 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PIC2ASCII - Arya Software Group"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11520
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "XRtf Method"
      Height          =   345
      Left            =   5235
      TabIndex        =   8
      Top             =   8025
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Second Method"
      Height          =   345
      Left            =   3855
      TabIndex        =   7
      Top             =   8025
      Width           =   1320
   End
   Begin VB.TextBox TxtTab 
      Height          =   330
      Left            =   7350
      TabIndex        =   5
      Text            =   "0"
      Top             =   8025
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Rtf"
      Height          =   345
      Left            =   1305
      TabIndex        =   4
      Top             =   8025
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Graphics"
      Height          =   345
      Left            =   45
      TabIndex        =   3
      Top             =   8025
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   10005
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "First Method"
      Height          =   345
      Left            =   2580
      TabIndex        =   2
      Top             =   8025
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox Rtf 
      Height          =   7860
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   13864
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   2.25
         Charset         =   2
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   -2160
      Picture         =   "Form1.frx":0390
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   0
      Top             =   45
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Offset : "
      Height          =   195
      Left            =   6690
      TabIndex        =   6
      Top             =   8085
      Width           =   555
   End
End
Attribute VB_Name = "frmASCII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long

Private Sub cmdOpen_Click()
    On Error GoTo Ha
    
    CD.CancelError = True
    CD.Filter = "Image Files|*.bmp;*.jpg;*.gif;*.ico;*.jpeg"
    CD.ShowOpen
    Pic.Picture = LoadPicture(CD.FileName)
Ha:
   Exit Sub
End Sub

Private Sub Command1_Click()
   Dim I As Integer, J As Integer, k
   '/////////////// Initial Value ///////////////
   k = Timer
   Rtf.Font.Size = 2
   Rtf.Font.Bold = False
   Rtf.Font.Underline = False
   Rtf.Font.Strikethrough = False
   Rtf.SelText = String(Val(TxtTab.Text), vbTab)
   Rtf.Visible = False
   '/////////////////////////////////////////////
   For I = 0 To Pic.ScaleHeight - 1
       Rtf.SelUnderline = True
       Me.Caption = "PIC2ASCII - " & (I / (Pic.ScaleHeight - 1)) * 100 & "%"
       For J = 0 To Pic.ScaleWidth - 1
           Rtf.SelStart = Len(Rtf.Text)
           Rtf.SelLength = 1
           Rtf.SelColor = GetPixel(Pic.hdc, J, I)
           If J < Pic.ScaleWidth - 1 Then
              Rtf.SelText = "g"
           Else
              Rtf.SelText = "g" & vbCrLf
              Rtf.SelUnderline = False
              Rtf.SelText = Rtf.SelText & String(Val(TxtTab.Text), vbTab)
              Rtf.SelUnderline = True
           End If
       Next
   Next
   
   '/////////////////////////////////////////////
   Rtf.SelStart = Len(Rtf.Text)
   Rtf.SelLength = 1
   Rtf.SelFontSize = 24
   Rtf.SelText = " "
   Rtf.Visible = True
   '/////////////////////////////////////////////
   MsgBox Timer - k
End Sub

Private Sub Command2_Click()
   On Error GoTo Ha
   CD.CancelError = True
   CD.Filter = "Rich Text Document|*.Rtf"
   CD.FileName = ""
   CD.ShowSave
   Rtf.SaveFile CD.FileName
Ha:
   Exit Sub
End Sub

Private Sub Command3_Click()
    Dim I As Long, S As String, t As Double
    Dim SW As Long, SH As Long, J As Long
    Dim Index As Long
    t = Timer
    SW = Pic.ScaleWidth
    SH = Pic.ScaleHeight
    S = String$(SH * (SW + 2), "g")
    For I = 1 To SH
        Mid$(S, (I * (SW + 2)) - 1) = vbCrLf
    Next
    Rtf.Text = S
    Rtf.Visible = False
    Index = 0
    For I = 0 To SH - 1
        Me.Caption = "PIC2ASCII - " & (I / (Pic.ScaleHeight - 1)) * 100 & "%"
        For J = 0 To SW - 1
            Index = Index + 1
            Rtf.SelStart = Index
            Rtf.SelLength = 1
            Rtf.SelColor = GetPixel(Pic.hdc, J, I)
        Next
        Index = Index + 2
    Next
    Rtf.Visible = True
    MsgBox Timer - t
End Sub

Private Sub Command4_Click()
    Dim I As Long, J As Long
    Dim SW As Long, SH As Long
    Dim CT As String, CHT As String, S As String
    Dim Index As Long, t As Double
    
    SW = Pic.ScaleWidth
    SH = Pic.ScaleHeight
    
    Dim CL As Long, R As Long, G As Long, B As Long
    
    t = Timer
    CT = "{\rtf1\fbidis\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\froman\fprq2\fcharset2 Webdings;}}" + vbCrLf + "{\colortbl ;"

    For I = 0 To SH - 1
        Me.Caption = "PIC2ASCII - " & (I / (SH - 1)) * 100 & "%"
        S = ""
        For J = 0 To SW - 1
            CL = GetPixel(Pic.hdc, J, I)
            
            R = CL Mod 256
            CL = CL \ 256
            G = CL Mod 256
            CL = CL \ 256
            B = CL Mod 256
            
            S = S & "\red" & R & "\green" & G & "\blue" & B & ";"
        Next
        CT = CT & S & vbCrLf
    Next
    
    CHT = "}" & vbCrLf & "\viewkind4\uc1\pard\ul\f0\fs4" & vbCrLf
    For I = 0 To SH - 1
        S = ""
        For J = 0 To SW - 1
            Index = Index + 1
            S = S & "\cf" & Index & " g"
        Next
        CHT = CHT & S & "\par" & vbCrLf
    Next
    CHT = CHT & "}" & vbCrLf
    
    Rtf.Visible = False
    Rtf.TextRTF = CT & CHT
    Rtf.Visible = True
    
    MsgBox Timer - t
End Sub

