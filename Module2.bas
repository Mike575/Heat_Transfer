Attribute VB_Name = "Module2"
    Public dbLow As Double
    Public dbHigh As Double
    Public dbMiddle As Double
    Public dbLow_fvalue As Double
    Public dbHigh_fvalue As Double
    Public dbMiddle_fvalue As Double
    Public str_Formula As String
    Public dbVariable As Double
    

    
    Public Function Formula_PhaiK()        'Phai = KAdTm
        
         If (Phai = 0) And (K <> 0) And (A <> 0) And (dTm <> 0) Then
             Phai = K * A * dTm
         ElseIf (Phai <> 0) And (K = 0) And (A <> 0) And (dTm <> 0) Then
             K = Phai / (A * dTm)
         ElseIf (Phai <> 0) And (K <> 0) And (A = 0) And (dTm <> 0) Then
             A = Phai / (K * dTm)
         ElseIf (Phai <> 0) And (K <> 0) And (A <> 0) And (dTm = 0) Then
             dTm = Phai / (K * A)

         End If
    End Function
    
    Public Function Formula_dTm()      'dTm = (dT2 - dT1) / (ln(dT1 / dT2))
        
        If (dTm = 0) And (dT2 <> 0) And (dT1 <> 0) Then
            dTm = (dT2 - dT1) / (Log(dT1 / dT2))
        End If
    End Function

   Public Function Formula_Phaih()        '0 = qmLh * Cph * (Th1 - Th2)-Phai
     
        If (Phai <> 0) And (qmLh = 0) And (Cph <> 0) And (Th1 <> 0) And (Th2 <> 0) Then
            qmLh = Phai / (Cph * (Th1 - Th2))
        ElseIf (Phai <> 0) And (qmLh <> 0) And (Cph = 0) And (Th1 <> 0) And (Th2 <> 0) Then
            Cph = Phai / (qmLh * (Th1 - Th2))
        ElseIf (Phai <> 0) And (qmLh <> 0) And (Cph <> 0) And (Th1 = 0) And (Th2 <> 0) Then
            Th1 = Phai / (qmLh * Cph) + Th2
        ElseIf (Phai <> 0) And (qmLh <> 0) And (Cph <> 0) And (Th1 <> 0) And (Th2 = 0) Then
            Th2 = Th1 - Phai / (qmLh * Cph)
        ElseIf (Phai = 0) And (qmLh <> 0) And (Cph <> 0) And (Th1 <> 0) And (Th2 <> 0) Then
            Phai = qmLh * Cph * (Th1 - Th2)
        End If
    End Function

   Public Function Formula_Phaic()        '0 = qmLc * Cpc * (Tc2 - Tc1)-Phai
        
        If (Phai <> 0) And (qmLc = 0) And (Cpc <> 0) And (Tc1 <> 0) And (Tc2 <> 0) Then
            qmLc = Phai / (Cpc * (Tc2 - Tc1))
        ElseIf (Phai <> 0) And (qmLc <> 0) And (Cpc = 0) And (Tc1 <> 0) And (Tc2 <> 0) Then
            Cpc = Phai / (qmLc * (Tc2 - Tc1))
        ElseIf (Phai <> 0) And (qmLc <> 0) And (Cpc <> 0) And (Tc1 = 0) And (Tc2 <> 0) Then
            Tc1 = Tc2 - Phai / (qmLc * Cpc)
        ElseIf (Phai <> 0) And (qmLc <> 0) And (Cpc <> 0) And (Tc1 <> 0) And (Tc2 = 0) Then
            Tc2 = Tc1 + Phai / (qmLc * Cpc)
        ElseIf (Phai = 0) And (qmLc <> 0) And (Cpc <> 0) And (Tc1 <> 0) And (Tc2 <> 0) Then
            Phai = qmLc * Cpc * (Tc2 - Tc1)
        End If
    End Function

    Public Function Formula_dT2()      'dT2 = Th1 - Tc2
        
        If (dT2 = 0) And (Th1 <> 0) And (Tc2 <> 0) Then
            dT2 = Th1 - Tc2
        ElseIf (dT2 <> 0) And (Th1 = 0) And (Tc2 <> 0) Then
            Th1 = dT2 + Tc2
        ElseIf (dT2 <> 0) And (Th1 <> 0) And (Tc2 = 0) Then
            Tc2 = Th1 - dT2
        End If
    End Function
    
    Public Function Formula_dT1()      'dT1 = Th2 - Tc1
        
        If (dT1 = 0) And (Th2 <> 0) And (Tc1 <> 0) Then
            dT1 = Th2 - Tc1
        ElseIf (dT1 <> 0) And (Th2 = 0) And (Tc1 <> 0) Then
            Th2 = dT1 + Tc1
        ElseIf (dT1 <> 0) And (Th2 <> 0) And (Tc1 = 0) Then
            Tc1 = Th2 - dT1
        End If
    End Function
    
    Public Function Formula_K()      '1 / ((1 / aCool) + (1 / aHot) + (ThickPipe / aPipe)) - K = 0
        
        If (K = 0) And (aCool <> 0) And (aHot <> 0) And (ThickPipe <> 0) And (aPipe <> 0) Then
            K = 1 / ((1 / aCool) + (1 / aHot) + (ThickPipe / aPipe))
        ElseIf (K <> 0) And (aCool = 0) And (aHot <> 0) And (ThickPipe <> 0) And (aPipe <> 0) Then
            aCool = K / (1 - K * (1 / aHot) - K * (ThickPipe / aPipe))
        ElseIf (K <> 0) And (aCool <> 0) And (aHot = 0) And (ThickPipe <> 0) And (aPipe <> 0) Then
            aHot = K / (1 - K * (1 / aCool) - K * (ThickPipe / aPipe))
        ElseIf (K <> 0) And (aCool <> 0) And (aHot <> 0) And (ThickPipe = 0) And (aPipe <> 0) Then
            ThickPipe = (aPipe / K) * (1 - K * (1 / aCool) - K * (1 / aHot))
        ElseIf (K <> 0) And (aCool <> 0) And (aHot = 0) And (ThickPipe <> 0) And (aPipe = 0) Then
            aPipe = (1 - K * (1 / aCool) - K * (1 / aHot)) / (K * ThickPipe)
            
        End If
    End Function

    Public Function Formula()
        Dim CircleNumber As Integer
        Dim FstCVarNum As Integer
        Dim LstCVarNum As Integer
        CircleNumber = 0
        FstCVarNum = 0
        LstCVarNum = 0
        CircleNumber = 0
        Form2.Show
        
        Do While (CircleNumber <= 4)
            FstCVarNum = NumVariable()
            Call Formula_PhaiK      'Phai = KAdTm
            Call Formula_dTm        'dTm = (dT2 - dT1) / (ln(dT1 / dT2))
            Call Formula_Phaih      'Phai = qmLh * Cph * (Th1 - Th2)
            Call Formula_Phaic      'Phai = qmLc * Cpc * (Tc2 - Tc1)
            Call Formula_dT2        'dT2 = Th1 - Tc2
            Call Formula_dT1        'dT1 = Th2 - Tc1
            Call Formula_K          'K = 1 / ((1 / aCool) + (1 / aHot) + (ThickPipe / aPipe))
            Form2.Print "K:"; K
            Form2.Print "Phai:"; Phai
            Form2.Print "dT1:"; dT1
            Form2.Print "dT2:"; dT2
            Form2.Print "Tc2:"; Tc2
            Form2.Print "dTm:"; dTm
            Form2.Print "A:"; A
            
            LstCVarNum = NumVariable()
            If FstCVarNum = LstCVarNum Then
                CircleNumber = CircleNumber + 1
            Else: CircleNumber = 0
            End If
        Loop
        
       Call Variable_OUT
       
   End Function
   
    Public Function NumVariable() As Integer
         
        Dim i As Integer
        NumVariable = 0
        For i = 1 To 14
            If IsNumeric(Form1.Text(i)) And (Form1.Text(i)) Then
            '输入的是数字
            NumVariable = NumVariable + 1
            End If
        Next
    End Function
    
