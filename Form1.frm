VERSION 5.00
Object = "{26585398-BFBA-11D3-AA4D-000000000000}#3.0#0"; "quickQuoter2.0"
Begin VB.Form Form1 
   Caption         =   "Stock Design"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&View Graph"
      Height          =   310
      Left            =   6360
      TabIndex        =   39
      Top             =   0
      Width           =   1455
   End
   Begin Project1.QuickQuoter QuickQuoter1 
      Height          =   480
      Left            =   7800
      TabIndex        =   35
      Top             =   5760
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Get Quote"
      Height          =   310
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "Stock Symbol"
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label21 
      Caption         =   "Graph Type:"
      Height          =   255
      Left            =   3360
      TabIndex        =   38
      Top             =   30
      Width           =   975
   End
   Begin VB.Label Label20 
      Caption         =   "Symbol:"
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   25
      Width           =   975
   End
   Begin VB.Label Label19 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5280
      TabIndex        =   36
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label18 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   34
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   33
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   32
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   31
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   30
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   29
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   28
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   27
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Stock Exchange:"
      Height          =   255
      Index           =   15
      Left            =   3960
      TabIndex        =   26
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Market Capitilization:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   14
      Left            =   3960
      TabIndex        =   25
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Ask:"
      Height          =   255
      Index           =   13
      Left            =   3960
      TabIndex        =   24
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Bid:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   12
      Left            =   3960
      TabIndex        =   23
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "52 Week Low:"
      Height          =   255
      Index           =   11
      Left            =   3960
      TabIndex        =   22
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "52 Week High:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   10
      Left            =   3960
      TabIndex        =   21
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Change(%):"
      Height          =   255
      Index           =   9
      Left            =   3960
      TabIndex        =   20
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Change:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   8
      Left            =   3960
      TabIndex        =   19
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "P/E Ratio:"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   10
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Per Share Outstanding:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Per Share Earning:"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Volume:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Low:"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "High:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Open:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Last:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim symbols As String


Private Sub Command1_Click()
symbols = Text1.Text
QuickQuoter1.GetQuote symbols
End Sub

Private Sub Command2_Click()
'MsgBox Len(Combo1.ListIndex)
QuickQuoter1.GetGraph Combo1.ListIndex + 1, Text1.Text  'for 'graphType' enter 1 for 1 year big graph, 2 for 2 year big graph, 3 for 3 months big and 4 for 6 months small
End Sub
Private Sub Form_Load()
Combo1.AddItem "1 Year Big"
Combo1.AddItem "2 Year Big"
Combo1.AddItem "3 Months Big"
Combo1.AddItem "6 Months Small"
End Sub
Private Sub QuickQuoter1_GraphDownloadCompleted(filename As String)
Form2.Picture1.Picture = LoadPicture(filename)
Form2.Show
End Sub
Private Sub QuickQuoter1_QuoteDownloadFinish()
Label1.Caption = QuickQuoter1.CompanyName & "(" & symbols & ")"
Label19.Caption = "As of " & QuickQuoter1.DateTime
Label3.Caption = QuickQuoter1.LastPrice
Label4.Caption = QuickQuoter1.OpenPrice
Label5.Caption = QuickQuoter1.High
Label6.Caption = QuickQuoter1.Low
Label7.Caption = QuickQuoter1.Volume
Label8.Caption = QuickQuoter1.PerShareProfit
Label9.Caption = QuickQuoter1.ShareOutstanding
Label10.Caption = QuickQuoter1.PERatio
Label11.Caption = QuickQuoter1.Change
Label12.Caption = QuickQuoter1.PercentChange
Label13.Caption = QuickQuoter1.FiftyTwoWeekHigh
Label14.Caption = QuickQuoter1.FiftyTwoWeeksLow
Label15.Caption = QuickQuoter1.Bid
Label16.Caption = QuickQuoter1.Ask
Label17.Caption = QuickQuoter1.MarketCapitilization
Label18.Caption = QuickQuoter1.StockExchange
End Sub
