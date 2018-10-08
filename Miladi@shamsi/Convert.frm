VERSION 5.00
Begin VB.Form Taqvim 
   Caption         =   "Form1"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10740
   Icon            =   "Convert.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label WEEk 
      AutoSize        =   -1  'True
      Caption         =   "WEE"
      Height          =   195
      Left            =   9600
      TabIndex        =   3
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label KKK 
      Height          =   855
      Left            =   4920
      TabIndex        =   2
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label Tarikh 
      AutoSize        =   -1  'True
      Caption         =   "1390/25/520"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   7560
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "—÷« „ÕÌÿÌ 09369154258"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   72
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "Taqvim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function TarikhShamsi(Optional date1 As String, Optional SmallDate1 As Boolean) As String

      '====================================================
      Dim d, p, w, mon, MM, Ym, u, v, rp, X, I, Ys, Ms, Dm, P1, D1, Ds, DateShamsi
      d = Array(20, 19, 20, 20, 21, 21, 22, 22, 22, 22, 21, 21)
      p = Array(11, 12, 10, 12, 11, 11, 10, 10, 10, 9, 10, 10)
      w = Array("Ìò‘‰»Â", "œÊ‘‰»Â", "”Â ‘‰»Â", "çÂ«—‘‰»Â", "Å‰Ã‘‰»Â", "Ã„⁄Â", "‘‰»Â")
      
      If SmallDate1 = True Then
            mon = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
      Else
            mon = Array("›—Ê—œÌ‰", "«—œÌ»Â‘ ", "Œ—œ«œ", " Ì—", "„—œ«œ", "‘Â—ÌÊ—", "„Â—", "¬»«‰", "¬–—", "œÌ", "»Â„‰", "«”›‰œ")
      End If
      
      If date1 = "" Then date1 = Date
      
      Dm = Day(date1) '»œ”  ¬Ê—œ‰ —Ê“
      MM = Month(date1) '»œ”  ¬Ê—œ‰ „«Â
      Ym = Year(date1) '»œ”  ¬Ê—œ‰ ”«·
      u = 0
      rp = 0
      If (Ym Mod 4) = 0 Then u = 1 ' ‘ŒÌ’ ò»Ì”Â »Êœ‰
      If ((Ym Mod 100) = 0 And (Ym Mod 400) <> 0) Then u = 0 ' ‘ŒÌ’ ò»Ì”Â ‰»Êœ‰
      Ys = Ym - 622 ' »œÌ· ”«· „Ì·«œÌ »Â ‘„”Ì
      X = Ys - 22
      X = X Mod 33
      If ((X Mod 4) = 0 And X <> 32) Then rp = 1
      I = Not (rp - 2) + Not (u - 2) * 2
      X = 0
      If (I = 0 And MM = 3) Then X = 1
      If I = 0 Then I = 3
      Ms = (9 + MM) Mod 13
      If Ms < 10 Then Ms = Ms + 1
      D1 = d(MM - 1)
      If (I = 1 And MM > 2) Then D1 = D1 - 1
      If (I = 2 And MM < 3) Then D1 = D1 - 1
      P1 = p(MM - 1)
      If (I = 1 And MM > 2) Then P1 = P1 + 1
      If (I = 2 And MM < 4) Then P1 = P1 + 1
      If (Dm > 0 And Dm <= D1) Then
             Ds = P1 + Dm + X - 1
          X = 1
      Else
          Ds = Dm - D1
          Ms = Ms + 1
          If Ms = 13 Then Ms = 1
          X = 2
      End If
      If ((MM = 3 And X = 2) Or MM > 3) Then Ys = Ys + 1
      If SmallDate1 = True Then
'     ??? ??? ?? ???? ???? ???????? ???????? ?? ??? ?? ?? ???? ????? ?? ?????
'            TarikhShamsi = Trim(Str(Ys)) + "/" + Trim(mon(Ms - 1)) + "/" + Trim(Str(Ds))
            TarikhShamsi = Mid(Trim(str(Ys)), 3, 2) + "/" + Trim(mon(Ms - 1)) + "/" + Trim(str(Ds))
            Tarikh.Caption = str(Ys) & "/" & (Ms) & "/" & str(Ds)
      Else
            TarikhShamsi = w(Weekday(Date) - 1) + " " + str(Ds) + " " + mon(Ms - 1) + " " + str(Ys)
            Tarikh.Caption = (Ys) & "/" & (Ms) & "/" & (Val(Ds))
            If Val(Ms) < 10 Then Ms = "0" & Ms
            
            If Val(Ds) < 10 Then Ds = "0" & Ds
             KKK.Caption = Ys & Ms & Ds
      End If

End Function

Private Sub Form_Load()
'MsgBox TarikhShamsi(Date)
'Form1.Caption = TarikhShamsi(Date)
Label1.Caption = TarikhShamsi(Date)
WEEk.Caption = Weekday(Date)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
