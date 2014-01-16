VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTLetra 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Palabra"
      Height          =   255
      Left            =   1320
      TabIndex        =   22
      Top             =   120
      Width           =   2175
   End
   Begin VB.Shape a62 
      Height          =   135
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape a61 
      Height          =   135
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line a6 
      Visible         =   0   'False
      X1              =   5640
      X2              =   5880
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line a5 
      Visible         =   0   'False
      X1              =   5760
      X2              =   6120
      Y1              =   2160
      Y2              =   2520
   End
   Begin VB.Line a4 
      Visible         =   0   'False
      X1              =   5760
      X2              =   5400
      Y1              =   2160
      Y2              =   2520
   End
   Begin VB.Line a3 
      Visible         =   0   'False
      X1              =   5280
      X2              =   6240
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line a2 
      Visible         =   0   'False
      X1              =   5760
      X2              =   5760
      Y1              =   1440
      Y2              =   2160
   End
   Begin VB.Shape a1 
      Height          =   855
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   11
      Left            =   2640
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   12
      Left            =   2880
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   13
      Left            =   3120
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   14
      Left            =   3360
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   15
      Left            =   3600
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   16
      Left            =   3840
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   17
      Left            =   4080
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   18
      Left            =   4320
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   19
      Left            =   4560
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   20
      Left            =   4800
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   10
      Left            =   2400
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   9
      Left            =   2160
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   8
      Left            =   1920
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   7
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RS As New Recordset
Public Conexion As New Connection
Public cuanto As Integer
Dim i As Integer
Public Acierto As Boolean
Public IntCorrecta As Integer, IntInCorrecta As Integer
Private Sub Form_Load()
Randomize
Conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ahorcado.mdb;Persist Security Info=False"
RS.CursorLocation = adUseClient
RS.CursorType = adOpenDynamic
RS.LockType = adLockOptimistic
Call Juego
End Sub
Private Sub Juego()
Dim Azar As Integer
a1.Visible = False
a6.Visible = False
a61.Visible = False
a62.Visible = False
a2.Visible = False
a3.Visible = False
a4.Visible = False
a5.Visible = False
For i = 1 To Len(Me.Tag)
    LblPalabra(i).Visible = False
Next
If RS.State = 1 Then RS.Close
RS.Open "select count(palabra) as cuanto from tahorcado", Conexion, adOpenDynamic
cuanto = RS!cuanto
Azar = 0
Do While Azar < 1 Or Azar > cuanto
    Azar = Int(Rnd() * cuanto) + 1
    If RS.State = 1 Then RS.Close
    RS.Open "select * from tahorcado where id=" & Azar, Conexion, adOpenDynamic
    If RS!ya = True Then
    Call Juego
    Exit Sub
    End If
    DoEvents
    If RS.State = 1 Then RS.Close
    RS.Open "select * from tahorcado where id=" & Azar, Conexion, adOpenDynamic
    Me.Tag = RS!palabra
    If RS.State = 1 Then RS.Close
    RS.Open "update tahorcado set ya=true where id=" & Azar, Conexion, adOpenDynamic
Loop
For i = 1 To Len(Me.Tag)
    LblPalabra(i).Visible = True
Next
TXTLetra = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
RS.Open "update tahorcado set ya=false", Conexion, adOpenDynamic
End Sub

Private Sub TXTLetra_Change()
If TXTLetra = "" Then Exit Sub
Acierto = False
If IsNumeric(TXTLetra) = True Then
    MsgBox "Solo puedes escribir letras", vbQuestion
    SendKeys "{BS}"
End If
For i = 1 To Len(Me.Tag)
If UCase(TXTLetra) = UCase(Mid(Me.Tag, i, 1)) Then
    LblPalabra(i).Caption = TXTLetra
    Acierto = True
End If
Next
If Acierto = False Then
    IntInCorrecta = IntInCorrecta + 1
    Select Case IntInCorrecta
        Case 1
            a1.Visible = True
        Case 2
            a6.Visible = True
            a61.Visible = True
            a62.Visible = True
        Case 3
            a2.Visible = True
        Case 4
            a3.Visible = True
        Case 5
            a4.Visible = True
        Case 6
            a5.Visible = True
    End Select
End If
For i = 1 To Len(Me.Tag)
    If LblPalabra(i).Caption = "_" Then IntCorrecta = Len(Me.Tag)
Next
    
If IntCorrecta <> Len(Me.Tag) Then
    MsgBox "Ganaste", vbInformation
    IntCorrecta = 0
    IntInCorrecta = 0
    For i = 1 To 20
        LblPalabra(i).Caption = "_"
    Next
    Label1.Caption = ""
    Call Juego
    Exit Sub
End If
If IntInCorrecta = 6 Then
    MsgBox "Perdiste" & vbCrLf & "La respuesta es: " & Me.Tag, vbCritical, "Error"
    IntInCorrecta = 0
    IntCorrecta = 0
    For i = 1 To 20
        LblPalabra(i).Caption = "_"
    Next
    Label1.Caption = ""
    Call Juego
    Exit Sub
End If
Label1.Caption = Label1.Caption & TXTLetra.Text
TXTLetra.SelStart = 0
TXTLetra.SelLength = 1
IntCorrecta = 0
End Sub

