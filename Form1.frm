VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   4410
   ClientTop       =   2670
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   6525
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   16
      Left            =   3840
      TabIndex        =   35
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   5520
      TabIndex        =   34
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5520
      TabIndex        =   33
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   15
      Left            =   3840
      TabIndex        =   16
      Text            =   "0"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   14
      Left            =   3840
      TabIndex        =   15
      Text            =   "0"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   13
      Left            =   3840
      TabIndex        =   14
      Text            =   "0"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   12
      Left            =   3840
      TabIndex        =   13
      Text            =   "0"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   11
      Left            =   2280
      TabIndex        =   12
      Text            =   "0"
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   10
      Left            =   2280
      TabIndex        =   11
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   9
      Left            =   2280
      TabIndex        =   10
      Text            =   "0"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   8
      Left            =   2280
      TabIndex        =   9
      Text            =   "0"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   7
      Left            =   2280
      TabIndex        =   8
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   6
      Left            =   2280
      TabIndex        =   7
      Text            =   "0"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   5
      Left            =   840
      TabIndex        =   6
      Text            =   "0"
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   4
      Left            =   840
      TabIndex        =   5
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   3
      Left            =   840
      TabIndex        =   4
      Text            =   "0"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   2
      Left            =   840
      TabIndex        =   3
      Text            =   "0"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "dTm"
      Height          =   255
      Left            =   3120
      TabIndex        =   36
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "method"
      Height          =   375
      Index           =   15
      Left            =   3120
      TabIndex        =   32
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "K"
      Height          =   375
      Index           =   14
      Left            =   3120
      TabIndex        =   31
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "phai"
      Height          =   375
      Index           =   13
      Left            =   3120
      TabIndex        =   30
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "A"
      Height          =   375
      Index           =   12
      Left            =   3120
      TabIndex        =   29
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "thickpipe"
      Height          =   375
      Index           =   11
      Left            =   1680
      TabIndex        =   28
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "apipe"
      Height          =   375
      Index           =   10
      Left            =   1680
      TabIndex        =   27
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "ahot"
      Height          =   375
      Index           =   9
      Left            =   1680
      TabIndex        =   26
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "acool"
      Height          =   375
      Index           =   8
      Left            =   1680
      TabIndex        =   25
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Cph"
      Height          =   375
      Index           =   7
      Left            =   1680
      TabIndex        =   24
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Cpc"
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   23
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Th2"
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   22
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Th1"
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   21
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Tc2"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Tc1"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "qmlh"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "qvlc"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call initial_Form1textV0
End Sub

Private Sub Command2_Click()
    Call Variable_Change
    Call Formula
    Call Variable_OUT
   
End Sub

Private Sub Command3_Click()
    Call Variable_INPUT
End Sub

Private Sub Form_Load()
  
  Call initial_HeatV0
  Call initial_Form1textV0
End Sub

Private Sub Text_Change(Index As Integer)
  
    Call Variable_Change
End Sub

    Public Function Formula()
        Dim CircleNumber As Integer
        Dim FstCVarNum As Integer
        Dim LstCVarNum As Integer
        CircleNumber = 0
        FstCVarNum = 0
        LstCVarNum = 0
        CircleNumber = 0
        
        Do While CircleNumber <= 10
            FstCVarNum = NumVariable()
            If Not (K * A * dTm) Then
            Call Formula_PhaiK      'Phai = KAdTm
            End If
            If Not (dTm * dT2 * dT1) Then
            Call Formula_dTm        'dTm = (dT2 - dT1) / (ln(dT1 / dT2))
            End If
            If Not (Phai * qmLh * Cph * Th1 * Th2) Then
            Call Formula_Phaih      'Phai = qmLh * Cph * (Th1 - Th2)
            End If
            If Not (Phai * qmLc * Cpc * Tc2 * Tc1) Then
            Call Formula_Phaic      'Phai = qmLc * Cpc * (Tc2 - Tc1)
            End If
            If Not (dT2 * Th1 * Tc2) Then
            Call Formula_dT2        'dT2 = Th1 - Tc2
            End If
            If Not (dT1 * Th2 * Tc1) Then
            Call Formula_dT1        'dT1 = Th2 - Tc1
            End If
            If Not (K * aCool * aHot * ThickPipe * aPipe) Then
            Call Formula_K          'K = 1 / ((1 / aCool) + (1 / aHot) + (ThickPipe / aPipe))
            End If
            LstCVarNum = NumVariable()
            If FstCVarNum = LstCVarNum Then
                CircleNumber = CircleNumber + 1
            Else: CircleNumber = 0
            End If
           Call Variable_Change
        Loop
        
       Call Variable_OUT
       
   End Function

    Public Function NumVariable() As Integer
         
        Dim i As Integer
        NumVariable = 0
        For i = 1 To 14
            If IsNumeric(Text(i)) And (Text(i)) Then
            '�����������
            NumVariable = NumVariable + 1
            End If
        Next
    End Function
    
    
    
    
    
    
    
