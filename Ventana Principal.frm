VERSION 5.00
Begin VB.Form FrmPrincipal 
   Caption         =   "Ventana Principal"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdJugar 
      Caption         =   "&Jugar de Nuevo"
      Height          =   315
      Left            =   2640
      TabIndex        =   16
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton CmdLetra 
      Caption         =   "&Letra para usar"
      Height          =   315
      Left            =   840
      TabIndex        =   15
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label LblLetra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1560
      TabIndex        =   14
      Top             =   2520
      Width           =   195
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   14
      Left            =   5040
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   13
      Left            =   4680
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   12
      Left            =   4320
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   11
      Left            =   3960
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   10
      Left            =   3600
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   9
      Left            =   3240
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   8
      Left            =   2880
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   7
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   6
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   5
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   4
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LblPalabra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IntRegistros As Integer, Cantidad As Integer, Cuanto As Integer, IntPerdida As Integer
Dim i As Integer
Private Sub CmdJugar_Click()
For i = 1 To 14
    LblPalabra(i).Tag = ""
    LblPalabra(i).Caption = ""
    LblPalabra(i).Visible = False
Next
Cantidad = 0
Cuanto = 0
IntPerdida = 0
Call Generar_Palabra
End Sub
Private Sub CmdLetra_Click()
LblLetra.Caption = UCase(InputBox("Escribe la letra que desees", "El ahorcado"))
End Sub
Private Sub Form_Load()
Randomize
Call AbrirBD
Call Generar_Palabra
End Sub
Private Sub Generar_Palabra()
Dim IntAzar As Integer
IntAzar = 0
Sql = "select count(palabra) as Registros from t1"
Set RS = DB.OpenRecordset(Sql)
IntRegistros = RS!registros

Do While IntAzar > IntRegistros Or IntAzar = 0
    IntAzar = Int(Rnd() * IntRegistros) + 1
    If IntAzar <= IntRegistros And IntAzar >= 1 Then
        Sql = "Select * from t1 where codigo= " & IntAzar
        Set RS = DB.OpenRecordset(Sql)
        If RS!hecho = True Then
            Call Generar_Palabra
            Exit Sub
        Else
            RS.Edit
            RS!hecho = True
            RS.Update
        End If
            Cantidad = Len(RS!palabra)
            For i = 1 To Len(RS!palabra)
                LblPalabra(i).Visible = True
                LblPalabra(i).Caption = "_"
                LblPalabra(i).Tag = UCase(Mid(RS!palabra, i, 1))
            Next
    End If
Loop
End Sub
Private Sub Form_Unload(Cancel As Integer)
Sql = "select * from t1"
Set RS = DB.OpenRecordset(Sql)
With RS
    .MoveFirst
    While Not RS.EOF
        .Edit
        !hecho = False
        .Update
        .MoveNext
    Wend
End With
End Sub
Private Sub LblPalabra_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If Source.Caption = LblPalabra(Index).Tag Then
    LblPalabra(Index).Caption = LblPalabra(Index).Tag
    Cuanto = Cuanto + 1
Else
    IntPerdida = IntPerdida + 1
End If
If Cuanto = Cantidad Then
    MsgBox "Ganaste", vbQuestion, "Eres el papa de los helados"
    CmdJugar_Click
ElseIf Cuanto <> Cantidad And IntPerdida = 6 Then
    MsgBox "El juego Termino", vbQuestion, "Perdiste"
    Unload Me
End If
End Sub
