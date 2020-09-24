VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.UserControl QuickQuoter 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   DrawStyle       =   3  'Dash-Dot
   ScaleHeight     =   495
   ScaleWidth      =   510
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   360
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      Picture         =   "UserControl1.ctx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   0
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox stock 
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"UserControl1.ctx":030A
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2160
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "QuickQuoter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Dim compname, datetime1, last1, open1, high1, low1, high52, changeper1, volume1, exchange, change1, marketcap, ask1, bid1, low52, peratio1, pershareprofit1, shareoutstanding1 As String

Public Event GraphDownloadCompleted(filename As String)

Public Event QuoteDownloadStart()

Public Event QuoteDownloadFinish()

Public Event GraphDownloadStart()

Public Event OnError(errordescription As String)
Public Function GetQuote(symbol As String)
RaiseEvent QuoteDownloadStart
stock.Text = Inet1.OpenURL("http://www.stockpoint.com/quote.asp?Exchange=US&Symbol=" & symbol & "&Company=&x=0&y=0")
If InStr(1, stock.Text, "Symbol not found in database.") <> 0 Then
RaiseEvent OnError("Stock symbol not found.")
GoTo endIT
End If
'*************COMPANY NAME************
find1 = InStr(1, stock.Text, "<FONT COLOR=WHITE><B>") + 21
find2 = InStr(1, stock.Text, "&nbsp;") - 1
compname = Mid(stock.Text, find1, Len(stock.Text) - find2)
find1 = InStr(1, compname, "&nbsp;")
compname = Left(compname, find1 - 1)
'****************STOCK DATE AND TIME***
find1 = InStr(1, stock.Text, "As of ") + 6
find2 = InStr(1, stock.Text, "(E.T.)") + 6
datetime1 = Mid(stock.Text, find1, find2 - find1)
temp = datetime1
find1 = InStr(1, datetime1, "&nbsp;&nbsp;") - 1
datetime1 = Left(datetime1, find1)
find1 = find1 + 12
datetime1 = datetime1 & " " & Right(temp, find1)
temp = Right(datetime1, 11)
datetime1 = Left(datetime1, Len(datetime1) - 11)
datetime1 = datetime1 & " " & Right(temp, 6)
'************LAST*********************
find1 = InStr(1, stock.Text, "Last") + 4
find2 = InStr(find1, stock.Text, "<B>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</B>") - 1
last1 = Left(stock.Text, find1)
'**********OPEN**********************
find1 = InStr(find1, stock.Text, "Open") + 4
find2 = InStr(find1, stock.Text, "-1>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</FONT>") - 1
open1 = Left(stock.Text, find1)
Text1.Text = open1
open1 = Right(Text1.Text, Len(Text1.Text) - 2)
'*********high1***********************
find1 = InStr(find1, stock.Text, "high1") + 4
find2 = InStr(find1, stock.Text, "-1>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</FONT>") - 1
high1 = Left(stock.Text, find1)
Text1.Text = high1
high1 = Right(Text1.Text, Len(Text1.Text) - 2)
'************low1*******************
find1 = InStr(find1, stock.Text, "low1") + 6
find2 = InStr(find1, stock.Text, "-1>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</FONT>") - 1
low1 = Left(stock.Text, find1)
Text1.Text = low1
low1 = Right(Text1.Text, Len(Text1.Text) - 2)
'************low1*******************
find1 = InStr(find1, stock.Text, "volume") + 6
find2 = InStr(find1, stock.Text, "-1>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</FONT>") - 1
volume1 = Left(stock.Text, find1)
Text1.Text = volume1
volume1 = Right(Text1.Text, Len(Text1.Text) - 2)
'************EARNING PER SHARE*******************
find1 = InStr(find1, stock.Text, "P/Share") + 7
find2 = InStr(find1, stock.Text, "-1>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</FONT>") - 1
pershareprofit1 = Left(stock.Text, find1)
Text1.Text = pershareprofit1
pershareprofit1 = Right(Text1.Text, Len(Text1.Text) - 2)
'************SHARE OUTSTANDING*******************
find1 = InStr(find1, stock.Text, "Outstanding") + 11
find2 = InStr(find1, stock.Text, "-1>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</FONT>") - 1
shareoutstanding1 = Left(stock.Text, find1)
Text1.Text = shareoutstanding1
shareoutstanding1 = Right(Text1.Text, Len(Text1.Text) - 2)
'************P/E RATIO*******************
find1 = InStr(find1, stock.Text, "Ratio") + 5
find2 = InStr(find1, stock.Text, "-1>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</FONT>") - 1
peratio1 = Left(stock.Text, find1)
Text1.Text = peratio1
peratio1 = Right(Text1.Text, Len(Text1.Text) - 2)
'***************change1*******************
find1 = 0
find1 = InStr(1, stock.Text, "#339933")
If find1 <> 0 Then
find1 = find1 + 9
ElseIf InStr(1, stock.Text, "RED") <> 0 Then
find1 = InStr(1, stock.Text, "RED") + 6
ans = 1
Else
find1 = InStr(1, stock.Text, "SIZE=-1>") + 8
ans = 2
End If
stock.Text = Mid(stock.Text, find1, Len(stock.Text) - find1)
find2 = InStr(1, stock.Text, "</FONT>")
change1 = Left(stock.Text, find2 - 1)
If ans = 2 Then change1 = Mid(change1, 3, Len(change1) - 3)
'change1 = Right(Text1.Text, Len(Text1.Text) - 2)
If ans = 1 Then change1 = "-" & change1

'*************PERCENT CHANGE******************
find1 = 0
find1 = InStr(1, stock.Text, "#339933")
If find1 <> 0 Then
find1 = find1 + 9
ElseIf InStr(1, stock.Text, "RED") <> 0 Then
find1 = InStr(1, stock.Text, "RED") + 6
ans = 1
Else
find1 = InStr(1, stock.Text, "SIZE=-1>") + 8
ans = 2
End If
stock.Text = Mid(stock.Text, find1, Len(stock.Text) - find1)
find2 = InStr(1, stock.Text, "</FONT>")
changeper1 = Left(stock.Text, find2 - 1)
Text1.Text = changeper1
'changeper1 = Right(Text1.Text, Len(Text1.Text) - 2)
If ans = 2 Then changeper1 = Mid(changeper1, 3, Len(changeper1) - 3)
If ans = 1 Then
changeper1 = "-" & changeper1
ans = 0
End If
GoAhead:
'**********52 WEEK HIGH**********************
find1 = InStr(find1, stock.Text, "High") + 4
find2 = InStr(find1, stock.Text, "-1>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</FONT>") - 1
high52 = Left(stock.Text, find1)
Text1.Text = high52
high52 = Right(Text1.Text, Len(Text1.Text) - 2)
'**********52 WEEK LOW**********************
find1 = InStr(find1, stock.Text, "Low") + 3
find2 = InStr(find1, stock.Text, "-1>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</FONT>") - 1
low52 = Left(stock.Text, find1)
Text1.Text = low52
low52 = Right(Text1.Text, Len(Text1.Text) - 2)
'**********bid1**********************
find1 = InStr(find1, stock.Text, "bid1") + 3
find2 = InStr(find1, stock.Text, "-1>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</FONT>") - 1
bid1 = Left(stock.Text, find1)
Text1.Text = bid1
bid1 = Right(Text1.Text, Len(Text1.Text) - 2)
'**********ask1**********************
find1 = InStr(find1, stock.Text, "ask1") + 3
find2 = InStr(find1, stock.Text, "-1>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</FONT>") - 1
ask1 = Left(stock.Text, find1)
Text1.Text = ask1
ask1 = Right(Text1.Text, Len(Text1.Text) - 2)
'**********MARKET CAPITILIZATION**********************
find1 = InStr(find1, stock.Text, "Capitalization") + 14
find2 = InStr(find1, stock.Text, "-1>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</FONT>") - 1
marketcap = Left(stock.Text, find1)
Text1.Text = marketcap
marketcap = Right(Text1.Text, Len(Text1.Text) - 2)
'**********Stock Exchange **********************
find1 = InStr(find1, stock.Text, "Exchange") + 8
find2 = InStr(find1, stock.Text, "-1>") + 3
stock.Text = Mid(stock.Text, find2, Len(stock.Text) - find2)
find1 = InStr(1, stock.Text, "</FONT>") - 1
exchange = Left(stock.Text, find1)
Text1.Text = exchange
exchange = Right(Text1.Text, Len(Text1.Text) - 2)
endIT:
RaiseEvent QuoteDownloadFinish
End Function
   

Public Property Get CompanyName() As String
CompanyName = compname
End Property

Public Property Let CompanyName(ByVal vNewValue As String)
compname = vNewValue
PropertyChanged CompanyName
End Property

Public Property Get DateTime() As String
DateTime = datetime1
End Property

Public Property Let DateTime(ByVal vNewValue As String)
datetime1 = vNewValue
PropertyChanged DateTime
End Property

Public Property Get LastPrice() As String
LastPrice = last1
End Property

Public Property Let LastPrice(ByVal vNewValue As String)
last1 = vNewValue
PropertyChanged
End Property

Public Property Get OpenPrice() As Variant
OpenPrice = open1
End Property

Public Property Let OpenPrice(ByVal vNewValue As Variant)
open1 = vNewValue
PropertyChanged OpenPrice
End Property

Public Property Get High() As Variant
High = high1
End Property

Public Property Let High(ByVal vNewValue As Variant)
high1 = vNewValue
PropertyChanged High
End Property

Public Property Get Low() As Variant
Low = low1
End Property

Public Property Let Low(ByVal vNewValue As Variant)
low1 = vNewValue
End Property

Public Property Get PerShareProfit() As Variant
PerShareProfit = pershareprofit1
End Property

Public Property Let PerShareProfit(ByVal vNewValue As Variant)
pershareprofit1 = vNewValue
PropertyChanged PerShareProfit
End Property

Public Property Get ShareOutstanding() As Variant
ShareOutstanding = shareoutstanding1
End Property

Public Property Let ShareOutstanding(ByVal vNewValue As Variant)
shareoutstanding1 = vNewValue
PropertyChanged ShareOutstanding
End Property


Public Property Get PERatio() As Variant
PERatio = peratio1
End Property

Public Property Let PERatio(ByVal vNewValue As Variant)
peratio1 = vNewValue
PropertyChanged PERatio
End Property

Public Property Get Change() As Variant
Change = change1
End Property

Public Property Let Change(ByVal vNewValue As Variant)
change1 = vNewValue
PropertyChanged Change
End Property

Public Property Get PercentChange() As Variant
PercentChange = changeper1
End Property

Public Property Let PercentChange(ByVal vNewValue As Variant)
changeper1 = vNewValue
PropertyChanged PercentChange
End Property

Public Property Get FiftyTwoWeekHigh() As Variant
FiftyTwoWeekHigh = high52
End Property

Public Property Let FiftyTwoWeekHigh(ByVal vNewValue As Variant)
high52 = vNewValue
PropertyChanged FiftyTwoWeekHigh
End Property

Public Property Get FiftyTwoWeeksLow() As Variant
FiftyTwoWeeksLow = low52
End Property

Public Property Let FiftyTwoWeeksLow(ByVal vNewValue As Variant)
low52 = vNewValue
PropertyChanged FiftyTwoWeeksLow
End Property

Public Property Get Bid() As Variant
Bid = bid1
End Property

Public Property Let Bid(ByVal vNewValue As Variant)
bid1 = vNewValue
PropertyChanged Bid
End Property

Public Property Get Ask() As Variant
Ask = ask1
End Property

Public Property Let Ask(ByVal vNewValue As Variant)
ask1 = vNewValue
PropertyChanged vNewValue
End Property

Public Property Get MarketCapitilization() As Variant
MarketCapitilization = marketcap
End Property

Public Property Let MarketCapitilization(ByVal vNewValue As Variant)
marketcap = vNewValue
PropertyChanged MarketCapitilization
End Property

Public Property Get StockExchange() As Variant
StockExchange = exchange
End Property

Public Property Let StockExchange(ByVal vNewValue As Variant)
exchange = vNewValue
PropertyChanged StockExchange
End Property

Private Sub UserControl_Resize()
UserControl.Width = Picture1.Width
UserControl.Height = Picture1.Height
End Sub

Public Property Get Volume() As Variant
Volume = volume1
End Property

Public Property Let Volume(ByVal vNewValue As Variant)
volume1 = vNewValue
PropertyChanged Volume
End Property

Public Function GetGraph(GraphType As Integer, symbol As String)
RaiseEvent GraphDownloadStart
If GraphType > 4 Then
RaiseEvent OnError("Graph type should be an integer less than or equal to 4")
GoTo exitFast
End If
Dim url As String
Dim Bilden() As Byte
Select Case GraphType
Case 1:
url = "http://chart.yahoo.com/c/1y/" & Left(symbol, 1) & "/" & symbol & ".gif" 'big
Case 2:
url = "http://chart.yahoo.com/c/2y/" & Left(symbol, 1) & "/" & symbol & ".gif" 'big
Case 3:
url = "http://chart.yahoo.com/c/3m/" & Left(symbol, 1) & "/" & symbol & ".gif" 'big
Case 4:
url = "http://chart.yahoo.com/c/0b/" & Left(symbol, 1) & "/" & symbol & ".gif" 'small
End Select
Bilden() = Inet2.OpenURL(url, icByteArray) ' Download picture.s = Bilden()
s = Bilden()
If Len(s) <> 75 Then
Open "C:\stock.gif" For Binary Access Write As #1 ' Save the file.
Put #1, , Bilden()
Close #1
Else
RaiseEvent OnError("Symbol no found.")
End If
'picturebox1.Picture = LoadPicture("c:/stock.gif")
RaiseEvent GraphDownloadCompleted("c:/stock.gif")
exitFast:
End Function


