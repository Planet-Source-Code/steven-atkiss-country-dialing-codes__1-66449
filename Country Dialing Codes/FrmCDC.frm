VERSION 5.00
Begin VB.Form FrmCDC 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PCS: Country Dialing Codes"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCDC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   540
      Width           =   855
   End
   Begin VB.ComboBox CmbCDC 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   540
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Enter code or first few letters of  the country."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4680
   End
End
Attribute VB_Name = "FrmCDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmbCDC_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        CmdFind_Click
    End If
    
End Sub

Private Sub CmdFind_Click()
    
    Dim Country As String
    Dim Code As Integer
    Dim FCountry As String
    
    CmbCDC.Text = Trim(CmbCDC.Text)
    
    If Left$(CmbCDC.Text, 1) = "+" Then CmbCDC.Text = Right$(CmbCDC.Text, Len(CmbCDC.Text) - 1)
    
    
    Select Case UCase(CmbCDC.Text)
        Case "UK", "GB", "BRITAIN", "ENGLAND", "SCOTLAND", "WALES"
            FCountry = "United Kingdom"
        Case "AMERICA", "USA", "US"
            FCountry = "United States"
        Case Else
            FCountry = CmbCDC.Text
    End Select
    
    
    
    
    
    If IsNumeric(CmbCDC.Text) = True Then
        Code = CmbCDC.Text
        
        CmbCDC.Clear
        CmbCDC.Text = "Please Wait..."
        
        Country = GetREGSZVal("SOFTWARE\Microsoft\Windows\CurrentVersion\Telephony\Country List\" & Code & "\", "Name")
        Country = StripNonChar(Country)
        
        If Country = "" Then Country = "Unknown Code"
        
        CmbCDC.AddItem "(" & Format(Code, "000") & ") " & Country
        CmbCDC.ListIndex = 0
        
    Else
    
        CmbCDC.Clear
        CmbCDC.Text = "Please Wait..."
        
        For Code = 1 To 1000
            Country = GetREGSZVal("SOFTWARE\Microsoft\Windows\CurrentVersion\Telephony\Country List\" & Code & "\", "Name")
            Country = StripNonChar(Country)
            
            If UCase(Left$(Country, Len(FCountry))) = UCase(FCountry) Then
                CmbCDC.AddItem "(" & Format(Code, "000") & ") " & Country
            End If
            
            DoEvents
            
        Next Code
        
        If CmbCDC.ListCount = 0 Then
            CmbCDC.Text = "No Entries Found!"
            CmbCDC.SelStart = 0
            CmbCDC.SelLength = Len(CmbCDC.Text)
        Else
            CmbCDC.ListIndex = 0
        End If
        
    End If
    
    
    
End Sub

Private Sub Form_Load()
    
    CmbCDC.Text = ""
    
End Sub
