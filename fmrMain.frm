VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6210
   ClientLeft      =   150
   ClientTop       =   195
   ClientWidth     =   5535
   Icon            =   "fmrMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "fmrMain.frx":169B2
   ScaleHeight     =   6210
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBackSpace 
      Caption         =   "Back Space"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   15
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   5400
   End
   Begin VB.CommandButton Command9 
      Caption         =   "V=I*R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdResultado 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   5295
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton Command11 
      Caption         =   "P=V*R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   11
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "P=(Ve2)/R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   10
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
      Caption         =   "P=(Ie2)*R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "V=r(P*R)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "V=P/I"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "R=P/(Ie2)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "R=V/I"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "R=(Ve2)/P"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "I=r(P/R)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "I=P/V"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I=V/R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Text            =   "0.000"
      Top             =   240
      Width           =   5295
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5760
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   5640
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Image Image 
      Height          =   5220
      Left            =   5760
      Picture         =   "fmrMain.frx":17A7C
      ToolTipText     =   "Click na fórmula para entrada de dados"
      Top             =   240
      Width           =   5595
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flagErro As Boolean
Dim strCaption As String
Dim virgula As Boolean
Dim step As Integer
Dim formula As Integer
Dim value1, value2, result As Single

Private Sub Form_Load()
   strCaption = App.Title & " - " & "v" & App.Major & "." & App.Minor & "." & App.Revision
   Me.Caption = strCaption
   
   step = 0
   Call clearValores
   cmdResultado.Enabled = False
   cmdBackSpace.Enabled = False
   txtDisplay.Enabled = False
   
End Sub

Private Sub Command1_Click()
   If step = 0 Then
      step = 1
      Call startFormula(1)
   ElseIf step = 1 Then
      step = 2
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "1"
      value1 = txtDisplay.Text
      Command12.Enabled = True
   ElseIf step = 2 Then
      txtDisplay.Text = txtDisplay.Text + "1"
      value1 = txtDisplay.Text
   ElseIf step = 4 Then
      step = 5
      Call startFormula(1)
   ElseIf step = 5 Then
      step = 6
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "1"
      value2 = txtDisplay.Text
      cmdResultado.Enabled = True
   ElseIf step = 6 Then
      txtDisplay.Text = txtDisplay.Text + "1"
      value2 = txtDisplay.Text
   End If
   
End Sub

Private Sub Command2_Click()
   If step = 0 Then
      step = 1
      Call startFormula(2)
   ElseIf step = 1 Then
      step = 2
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "2"
      value1 = txtDisplay.Text
      Command12.Enabled = True
   ElseIf step = 2 Then
      txtDisplay.Text = txtDisplay.Text + "2"
      value1 = txtDisplay.Text
   ElseIf step = 4 Then
      step = 5
      Call startFormula(2)
   ElseIf step = 5 Then
      step = 6
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "2"
      value2 = txtDisplay.Text
      cmdResultado.Enabled = True
   ElseIf step = 6 Then
      txtDisplay.Text = txtDisplay.Text + "2"
      value2 = txtDisplay.Text
   End If
   
End Sub

Private Sub Command3_Click()
   If step = 0 Then
      step = 1
      Call startFormula(3)
   ElseIf step = 1 Then
      step = 2
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "3"
      value1 = txtDisplay.Text
      Command12.Enabled = True
   ElseIf step = 2 Then
      txtDisplay.Text = txtDisplay.Text + "3"
      value1 = txtDisplay.Text
   ElseIf step = 4 Then
      step = 5
      Call startFormula(3)
   ElseIf step = 5 Then
      step = 6
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "3"
      value2 = txtDisplay.Text
      cmdResultado.Enabled = True
   ElseIf step = 6 Then
      txtDisplay.Text = txtDisplay.Text + "3"
      value2 = txtDisplay.Text
   End If
   
End Sub

Private Sub Command4_Click()
   If step = 0 Then
      step = 1
      Call startFormula(4)
   ElseIf step = 1 Then
      step = 2
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "4"
      value1 = txtDisplay.Text
      Command12.Enabled = True
   ElseIf step = 2 Then
      txtDisplay.Text = txtDisplay.Text + "4"
      value1 = txtDisplay.Text
   ElseIf step = 4 Then
      step = 5
      Call startFormula(4)
   ElseIf step = 5 Then
      step = 6
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "4"
      value2 = txtDisplay.Text
      cmdResultado.Enabled = True
   ElseIf step = 6 Then
      txtDisplay.Text = txtDisplay.Text + "4"
      value2 = txtDisplay.Text
   End If
   
End Sub

Private Sub Command5_Click()
   If step = 0 Then
      step = 1
      Call startFormula(5)
   ElseIf step = 1 Then
      step = 2
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "5"
      value1 = txtDisplay.Text
      Command12.Enabled = True
   ElseIf step = 2 Then
      txtDisplay.Text = txtDisplay.Text + "5"
      value1 = txtDisplay.Text
   ElseIf step = 4 Then
      step = 5
      Call startFormula(5)
   ElseIf step = 5 Then
      step = 6
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "5"
      value2 = txtDisplay.Text
      cmdResultado.Enabled = True
   ElseIf step = 6 Then
      txtDisplay.Text = txtDisplay.Text + "5"
      value2 = txtDisplay.Text
   End If
   
End Sub

Private Sub Command6_Click()
   If step = 0 Then
      step = 1
      Call startFormula(6)
   ElseIf step = 1 Then
      step = 2
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "6"
      value1 = txtDisplay.Text
      Command12.Enabled = True
   ElseIf step = 2 Then
      txtDisplay.Text = txtDisplay.Text + "6"
      value1 = txtDisplay.Text
   ElseIf step = 4 Then
      step = 5
      Call startFormula(6)
   ElseIf step = 5 Then
      step = 6
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "6"
      value2 = txtDisplay.Text
      cmdResultado.Enabled = True
   ElseIf step = 6 Then
      txtDisplay.Text = txtDisplay.Text + "6"
      value2 = txtDisplay.Text
   End If
   
End Sub

Private Sub Command7_Click()
   If step = 0 Then
      step = 1
      Call startFormula(7)
   ElseIf step = 1 Then
      step = 2
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "7"
      value1 = txtDisplay.Text
      Command12.Enabled = True
   ElseIf step = 2 Then
      txtDisplay.Text = txtDisplay.Text + "7"
      value1 = txtDisplay.Text
   ElseIf step = 4 Then
      step = 5
      Call startFormula(7)
   ElseIf step = 5 Then
      step = 6
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "7"
      value2 = txtDisplay.Text
      cmdResultado.Enabled = True
   ElseIf step = 6 Then
      txtDisplay.Text = txtDisplay.Text + "1"
      value2 = txtDisplay.Text
   End If
   
End Sub

Private Sub Command8_Click()
   If step = 0 Then
      step = 1
      Call startFormula(8)
   ElseIf step = 1 Then
      step = 2
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "8"
      value1 = txtDisplay.Text
      Command12.Enabled = True
   ElseIf step = 2 Then
      txtDisplay.Text = txtDisplay.Text + "8"
      value1 = txtDisplay.Text
   ElseIf step = 4 Then
      step = 5
      Call startFormula(8)
   ElseIf step = 5 Then
      step = 6
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "8"
      value2 = txtDisplay.Text
      cmdResultado.Enabled = True
   ElseIf step = 6 Then
      txtDisplay.Text = txtDisplay.Text + "8"
      value2 = txtDisplay.Text
   End If
   
End Sub

Private Sub Command9_Click()
   If step = 0 Then
      step = 1
      formula = 9
      Call startFormula(9)
   ElseIf step = 1 Then
      step = 2
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "9"
      value1 = txtDisplay.Text
      Command12.Enabled = True
   ElseIf step = 2 Then
      txtDisplay.Text = txtDisplay.Text + "9"
      value1 = txtDisplay.Text
   ElseIf step = 4 Then
      step = 5
      Call startFormula(9)
   ElseIf step = 5 Then
      step = 6
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "9"
      value2 = txtDisplay.Text
      cmdResultado.Enabled = True
   ElseIf step = 6 Then
      txtDisplay.Text = txtDisplay.Text + "9"
      value2 = txtDisplay.Text
   End If
   
End Sub

Private Sub Command10_Click()
   If step = 0 Then
      step = 1
      formula = 10
      Call startFormula(10)
   ElseIf step = 1 Then
      step = 2
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "0"
      value1 = txtDisplay.Text
      Command12.Enabled = True
   ElseIf step = 2 Then
      txtDisplay.Text = txtDisplay.Text + "0"
      value1 = txtDisplay.Text
   ElseIf step = 4 Then
      step = 5
      Call startFormula(10)
   ElseIf step = 5 Then
      step = 6
      txtDisplay.Font.Size = 24
      txtDisplay.Text = "0"
      value2 = txtDisplay.Text
      cmdResultado.Enabled = True
   ElseIf step = 6 Then
      txtDisplay.Text = txtDisplay.Text + "0"
      value2 = txtDisplay.Text
   End If
   
End Sub

Private Sub Command11_Click()
   
   If step = 0 Then
      step = 1
      Call startFormula(11)
   ElseIf step = 1 Then
      step = 2
      txtDisplay.Font.Size = 24
      txtDisplay.Text = ","
      value1 = txtDisplay.Text
      virgula = True
      Command12.Enabled = True
   ElseIf step = 2 Then
      If virgula = True Then Exit Sub
      txtDisplay.Text = txtDisplay.Text + ","
      value1 = txtDisplay.Text
      virgula = True
   ElseIf step = 4 Then
      step = 5
      Call startFormula(11)
   ElseIf step = 5 Then
      step = 6
      txtDisplay.Font.Size = 24
      txtDisplay.Text = ","
      value2 = txtDisplay.Text
      virgula = True
      cmdResultado.Enabled = True
   ElseIf step = 6 Then
      If virgula = True Then Exit Sub
      txtDisplay.Text = txtDisplay.Text + ","
      value2 = txtDisplay.Text
      virgula = True
   End If
   
End Sub

Private Sub Command12_Click()
   If step = 0 Then
      step = 1
      Call startFormula(12)
   Else
      If txtDisplay.Text = Empty Then
         Beep
         Exit Sub
      End If
      
      step = 4
      If formula = 1 Then Call Command1_Click
      If formula = 2 Then Call Command2_Click
      If formula = 3 Then Call Command3_Click
      If formula = 4 Then Call Command4_Click
      If formula = 5 Then Call Command5_Click
      If formula = 6 Then Call Command6_Click
      If formula = 7 Then Call Command7_Click
      If formula = 8 Then Call Command8_Click
      If formula = 9 Then Call Command9_Click
      If formula = 10 Then Call Command10_Click
      If formula = 11 Then Call Command11_Click
      
      If formula = 12 Then
         step = 5
         Call startFormula(12)
      End If
      
      virgula = False
   End If
   
End Sub

Private Sub cmdClear_Click()
   Call clearValores
   updateCommand (False)
   Me.Caption = strCaption
   
   If txtDisplay.Text = "Error" Then
      If flagErro = False Then
         flagErro = True
      Else
         flagErro = False
         txtDisplay.Text = "0.000"
      End If
   Else
      txtDisplay.Text = "0.000"
   End If
   
End Sub

Private Sub cmdBackSpace_Click()
Dim ultimoCaracter As String

    If Len(txtDisplay.Text) > 0 Then
        ultimoCaracter = Right(txtDisplay.Text, 1)
        txtDisplay.Text = Left(txtDisplay.Text, Len(txtDisplay.Text) - 1)
    End If
    
    If ultimoCaracter = "," Then virgula = False
    
    If step = 1 Or step = 2 Then
      If txtDisplay.Text = Empty Then
         value1 = 0
      Else
         value1 = txtDisplay.Text
      End If
    End If
    
    If step = 5 Or step = 6 Then
      If txtDisplay.Text = Empty Then
         value2 = 0
      Else
         value2 = txtDisplay.Text
      End If
    End If
    
    
End Sub

Private Sub cmdResultado_Click()
On Error GoTo Erro
   
   If txtDisplay.Text = Empty Then
      Beep
      Exit Sub
   End If
   
   'CORRENTE
   If formula = 1 Then
      result = CDbl(value1 / value2)
      txtDisplay.Text = Format(result, "0.000") & "A"
   ElseIf formula = 2 Then
      result = value1 / value2
      txtDisplay.Text = Format(result, "0.000") & "A"
   ElseIf formula = 3 Then
      result = Math.Sqr(value1 / value2)
      txtDisplay.Text = Format(result, "0.000") & "A"
      
   'RESISTÊNCIA
   ElseIf formula = 4 Then
      result = (value1 * value1) / value2
      txtDisplay.Text = Format(result, "0.000") & "R"
   ElseIf formula = 5 Then
      result = value1 / value2
      txtDisplay.Text = Format(result, "0.000") & "R"
   ElseIf formula = 6 Then
      result = value1 / (value2 * value2)
      txtDisplay.Text = Format(result, "0.000") & "R"
      
   'VOLTAGEM
   ElseIf formula = 7 Then
      result = value1 / value2
      txtDisplay.Text = Format(result, "0.000") & "V"
   ElseIf formula = 8 Then
      result = Math.Sqr(value1 * value2)
      txtDisplay.Text = Format(result, "0.000") & "V"
   ElseIf formula = 9 Then
      result = value1 * value2
      txtDisplay.Text = Format(result, "0.000") & "V"
   
   'POTÊNCIA
   ElseIf formula = 10 Then
      result = (value1 * value1) / value2
      txtDisplay.Text = Format(result, "0.000") & "W"
   ElseIf formula = 11 Then
      result = value1 * value2
      txtDisplay.Text = Format(result, "0.000") & "W"
   ElseIf formula = 12 Then
      result = (value1 * value1) * value2
      txtDisplay.Text = Format(result, "0.000") & "W"
   End If
   
   Call Timer1_Timer
   step = 0
   updateCommand (False)
   
Exit Sub

Erro:
   txtDisplay.Text = "Error"
   Call cmdClear_Click
   Beep
   
End Sub

Private Sub startFormula(varFormula As Integer)
      formula = varFormula
      txtDisplay.Font.Size = 12
      Command12.Enabled = False
      updateCommand (True)
      
   If step = 1 Then
      'CORRENTE
      If formula = 1 Then txtDisplay.Text = "Digite o valor de TENSÃO(V)"
      If formula = 2 Then txtDisplay.Text = "Digite o valor de POTÊNCIA(W)"
      If formula = 3 Then txtDisplay.Text = "Digite o valor de POTÊNCIA(W)"
      'RESISTÊNCIA
      If formula = 4 Then txtDisplay.Text = "Digite o valor de TENSÃO(V)"
      If formula = 5 Then txtDisplay.Text = "Digite o valor de TENSÃO(V)"
      If formula = 6 Then txtDisplay.Text = "Digite o valor de POTÊNCIA(W)"
      'VOLTAGEM
      If formula = 7 Then txtDisplay.Text = "Digite o valor de POTÊNCIA(W)"
      If formula = 8 Then txtDisplay.Text = "Digite o valor de POTÊNCIA(W)"
      If formula = 9 Then txtDisplay.Text = "Digite o valor de CORRENTE(A)"
      'POTÊNCIA
      If formula = 10 Then txtDisplay.Text = "Digite o valor de TENSÃO(V)"
      If formula = 11 Then txtDisplay.Text = "Digite o valor de TENSÃO(V)"
      If formula = 12 Then txtDisplay.Text = "Digite o valor de CORRENTE(A)"
   End If
   
   If step = 5 Then
      'CORRENTE
      If formula = 1 Then txtDisplay.Text = "Digite o valor de RESISTÊNCIA(R)"
      If formula = 2 Then txtDisplay.Text = "Digite o valor de TENSÃO(V)"
      If formula = 3 Then txtDisplay.Text = "Digite o valor de RESISTÊNCIA(R)"
      'RESISTÊNCIA
      If formula = 4 Then txtDisplay.Text = "Digite o valor de POTÊNCIA(W)"
      If formula = 5 Then txtDisplay.Text = "Digite o valor de CORRENTE(A)"
      If formula = 6 Then txtDisplay.Text = "Digite o valor de CORRENTE(A)"
      'VOLTAGEM
      If formula = 7 Then txtDisplay.Text = "Digite o valor de CORRENTE(A)"
      If formula = 8 Then txtDisplay.Text = "Digite o valor de RESISTÊNCIA(R)"
      If formula = 9 Then txtDisplay.Text = "Digite o valor de RESISTÊNCIA(R)"
      'POTÊNCIA
      If formula = 10 Then txtDisplay.Text = "Digite o valor de RESISTÊNCIA(R)"
      If formula = 11 Then txtDisplay.Text = "Digite o valor de RESISTÊNCIA(R)"
      If formula = 12 Then txtDisplay.Text = "Digite o valor de RESISTÊNCIA(R)"
   End If
   
End Sub

Private Sub updateCommand(flag As Boolean)
   If flag = True Then
      Command1.Caption = "1"
      Command2.Caption = "2"
      Command3.Caption = "3"
      Command4.Caption = "4"
      Command5.Caption = "5"
      Command6.Caption = "6"
      Command7.Caption = "7"
      Command8.Caption = "8"
      Command9.Caption = "9"
      Command10.Caption = "0"
      Command11.Caption = ","
      Command12.Caption = "Next"
      cmdResultado.Enabled = False
   Else
      Command1.Caption = "I=V/R"
      Command2.Caption = "I=P/V"
      Command3.Caption = "I=r(P/R)"
      Command4.Caption = "R=(Ve2)/P"
      Command5.Caption = "R=V/I"
      Command6.Caption = "R=P/(Ie2)"
      Command7.Caption = "V=P/I"
      Command8.Caption = "V=r(P*R)"
      Command9.Caption = "V=I*R"
      Command10.Caption = "P=(Ve2)/R"
      Command11.Caption = "P=V*R"
      Command12.Caption = "P=(Ie2)*R"
      
      If step > 0 Then
         step = 0
         Call clearValores
         virgula = False
         cmdResultado.Enabled = False
         Command12.Enabled = True
         txtDisplay.Font.Size = 24
         If txtDisplay.Text = "Error" Then
            Me.Caption = strCaption
         Else
            txtDisplay.Text = "0.000"
         End If
      Else
         Call clearValores
         virgula = False
         cmdResultado.Enabled = False
         Command12.Enabled = True
         txtDisplay.Font.Size = 24
      End If
      
   End If
   
End Sub

Private Sub clearValores()
   value1 = 0
   value2 = 0
   result = 0
         
End Sub

Private Sub Timer1_Timer()

   If step > 0 Then
      If formula = 1 Then Me.Caption = "V:" & value1 & "   " & "R:" & value2 & "   " & "A:" & result
      If formula = 2 Then Me.Caption = "P:" & value1 & "   " & "V:" & value2 & "   " & "A:" & result
      If formula = 3 Then Me.Caption = "P:" & value1 & "   " & "R:" & value2 & "   " & "A:" & result
      If formula = 4 Then Me.Caption = "V:" & value1 & "   " & "P:" & value2 & "   " & "R:" & result
      If formula = 5 Then Me.Caption = "V:" & value1 & "   " & "I:" & value2 & "   " & "R:" & result
      If formula = 6 Then Me.Caption = "P:" & value1 & "   " & "I:" & value2 & "   " & "R:" & result
      If formula = 7 Then Me.Caption = "P:" & value1 & "   " & "I:" & value2 & "   " & "V:" & result
      If formula = 8 Then Me.Caption = "P:" & value1 & "   " & "R:" & value2 & "   " & "V:" & result
      If formula = 9 Then Me.Caption = "I:" & value1 & "   " & "R:" & value2 & "   " & "V:" & result
      If formula = 10 Then Me.Caption = "V:" & value1 & "   " & "R:" & value2 & "   " & "W:" & result
      If formula = 11 Then Me.Caption = "V:" & value1 & "   " & "R:" & value2 & "   " & "W:" & result
      If formula = 12 Then Me.Caption = "I:" & value1 & "   " & "R:" & value2 & "   " & "W:" & result
   End If
   
   'Trata comando backspace
   If txtDisplay.Text = Empty Then
      cmdBackSpace.Enabled = False
   ElseIf Command12.Enabled = True And Command12.Caption = "Next" Or cmdResultado.Enabled = True Then
      cmdBackSpace.Enabled = True
   Else
      cmdBackSpace.Enabled = False
   End If
   
End Sub
