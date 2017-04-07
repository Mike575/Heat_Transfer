Attribute VB_Name = "Module2"
    Public dbLow As Double
    Public dbHigh As Double
    Public dbMiddle As Double
    Public dbLow_fvalue As Double
    Public dbHigh_fvalue As Double
    Public dbMiddle_fvalue As Double
    Public str_Formula As String
    Public dbVariable As Double
    
    Declare Function EbExecuteLine Lib "VBA6.dll" (ByVal pStringToExec As Long, ByVal Unknownn1 As Long, ByVal Unknownn2 As Long, ByVal fCheckOnly As Long) As Long

    Public Function ELine(sCode As String, Optional fCheckOnly As Boolean) As Boolean
        ELine = EbExecuteLine(StrPtr(sCode), 0&, 0&, Abs(fCheckOnly)) = 0
    End Function

    
    Public Function Formula_PhaiK()        'Phai = KAdTm
        
         If (Phai = 0) And (K <> 0) And (A <> 0) And (dTm <> 0) Then
             Phai = K * A * dTm
         ElseIf (Phai <> 0) And (K = 0) And (A <> 0) And (dTm <> 0) Then
             K = Phai / (A * dTm)
         ElseIf (Phai <> 0) And (K <> 0) And (A = 0) And (dTm <> 0) Then
             A = Phai / (K * dTm)
         ElseIf (Phai <> 0) And (K <> 0) And (A <> 0) And (dTm = 0) Then
             dTm = Phai / (K * A)
 '        Else
  '           If (K * A * dTm - Phai) > 0.2 Or (K * A * dTm - Phai) < -0.2 Then
    '         MsgBox "K,A,dTm,Phai collide!", vbCritical, "heat_exchange"
   '          End If
         End If
    End Function
    
    Public Function Formula_dTm()      'dTm = (dT2 - dT1) / (ln(dT1 / dT2))
        
        If (dTm = 0) And (dT2 <> 0) And (dT1 <> 0) Then
            dTm = (dT2 - dT1) / (Log(dT1 / dT2))
        End If
    End Function

   Public Function Formula_Phaih()        '0 = qmLh * Cph * (Th1 - Th2)-Phai
     
        If (Phai <> 0) And (qmLh = 0) And (Cph <> 0) And (Th1 <> 0) And (Th2 <> 0) Then
'            str_Formula = "dbvariable * Cph * (Th1 - Th2)-Phai"
'            Call BiSection
'            qmLh = dbVariable
            qmLh = Phai / (Cph * (Th1 - Th2))
        ElseIf (Phai <> 0) And (qmLh <> 0) And (Cph = 0) And (Th1 <> 0) And (Th2 <> 0) Then
'            str_Formula = "qmLh * dbvariable * (Th1 - Th2)-Phai"
'            Call BiSection
'            Cph = dbVariable
            Cph = Phai / (qmLh * (Th1 - Th2))
        ElseIf (Phai <> 0) And (qmLh <> 0) And (Cph <> 0) And (Th1 = 0) And (Th2 <> 0) Then
'            str_Formula = "qmLh * Cph * (dbvariable - Th2)-Phai"
'            Call BiSection
'            Th1 = dbVariable
            Th1 = Phai / (qmLh * Cph) + Th2
        ElseIf (Phai <> 0) And (qmLh <> 0) And (Cph <> 0) And (Th1 <> 0) And (Th2 = 0) Then
'            str_Formula = "qmLh * Cph * (Th1 - dbvariable)-Phai"
'            Call BiSection
'            Th2 = dbVariable
            Th2 = Th1 - Phai / (qmLh * Cph)
        ElseIf (Phai = 0) And (qmLh <> 0) And (Cph <> 0) And (Th1 <> 0) And (Th2 <> 0) Then
            Phai = qmLh * Cph * (Th1 - Th2)
'        Else
 '           If (qmLh * Cph * (Th1 - Th2) - Phai) > 0.2 Or (qmLh * Cph * (Th1 - Th2) - Phai) < -0.2 Then
  '          MsgBox "qmLh,Cph,Th1,Th2,Phai collide!", vbCritical, "heat_exchange"
   '         End If
        End If
    End Function

   Public Function Formula_Phaic()        '0 = qmLc * Cpc * (Tc2 - Tc1)-Phai
        
        If (Phai <> 0) And (qmLc = 0) And (Cpc <> 0) And (Tc1 <> 0) And (Tc2 <> 0) Then
 '           str_Formula = "dbvariable * Cpc * (Tc1 - Tc2)-Phai"        'This formular is wrong. Excuteline is also not effective
 '           Call BiSection
 '           qmLc = dbVariable
            qmLc = Phai / (Cpc * (Tc2 - Tc1))
        ElseIf (Phai <> 0) And (qmLc <> 0) And (Cpc = 0) And (Tc1 <> 0) And (Tc2 <> 0) Then
 '           str_Formula = "qmLc * dbvariable * (Tc1 - Tc2)-Phai"
 '           Call BiSection
 '           Cpc = dbVariable
            Cpc = Phai / (qmLc * (Tc2 - Tc1))
        ElseIf (Phai <> 0) And (qmLc <> 0) And (Cpc <> 0) And (Tc1 = 0) And (Tc2 <> 0) Then
 '           str_Formula = "qmLc * Cpc * (dbvariable - Tc2)-Phai"
 '           Call BiSection
 '           Tc1 = dbVariable
            Tc1 = Tc2 - Phai / (qmLc * Cpc)
        ElseIf (Phai <> 0) And (qmLc <> 0) And (Cpc <> 0) And (Tc1 <> 0) And (Tc2 = 0) Then
 '           str_Formula = "qmLc * Cpc * (Tc1 - dbvariable)-Phai"
 '           Call BiSection
 '           Tc2 = dbVariable
            Tc2 = Tc1 + Phai / (qmLc * Cpc)
        ElseIf (Phai = 0) And (qmLc <> 0) And (Cpc <> 0) And (Tc1 <> 0) And (Tc2 <> 0) Then
            Phai = qmLc * Cpc * (Tc2 - Tc1)
  '      Else
   '         If (qmLc * Cpc * (Tc1 - Tc2) - Phai) > 0.2 Or (qmLc * Cpc * (Tc1 - Tc2) - Phai) < -0.2 Then
    '        MsgBox "qmLC,Cpc,Tc1,Tc2,Phai collide!", vbCritical, "heat_exchange"
     '       End If
        End If
    End Function

    Public Function Formula_dT2()      'dT2 = Th1 - Tc2
        
        If (dT2 = 0) And (Th1 <> 0) And (Tc2 <> 0) Then
            dT2 = Th1 - Tc2
        ElseIf (dT2 <> 0) And (Th1 = 0) And (Tc2 <> 0) Then
            Th1 = dT2 + Tc2
        ElseIf (dT2 <> 0) And (Th1 <> 0) And (Tc2 = 0) Then
            Tc2 = Th1 - dT2
 '       Else
  '          If (Th1 - Tc2 - dT2) > 0.2 Or (Th1 - Tc2 - dT2) < -0.2 Then
   '         MsgBox "Th1,Tc2,dT2 collide!", vbCritical, "heat_exchange"
    '        End If
        End If
    End Function
    
    Public Function Formula_dT1()      'dT1 = Th2 - Tc1
        
        If (dT1 = 0) And (Th2 <> 0) And (Tc1 <> 0) Then
            dT1 = Th2 - Tc1
        ElseIf (dT1 <> 0) And (Th2 = 0) And (Tc1 <> 0) Then
            Th2 = dT1 + Tc1
        ElseIf (dT1 <> 0) And (Th2 <> 0) And (Tc1 = 0) Then
            Tc1 = Th2 - dT1
'        Else
 '           If (Th2 - Tc1 - dT1) > 0.2 Or (Th2 - Tc1 - dT1) < -0.2 Then
  '          MsgBox "Th2,Tc1,dT1 collide!", vbCritical, "heat_exchange"
   '         End If
        End If
    End Function
    
    Public Function Formula_K()      '1 / ((1 / aCool) + (1 / aHot) + (ThickPipe / aPipe)) - K = 0
        
        If (K = 0) And (aCool <> 0) And (aHot <> 0) And (ThickPipe <> 0) And (aPipe <> 0) Then
            K = 1 / ((1 / aCool) + (1 / aHot) + (ThickPipe / aPipe))
        ElseIf (K <> 0) And (aCool = 0) And (aHot <> 0) And (ThickPipe <> 0) And (aPipe <> 0) Then
 '           str_Formula = "1 / ((1 / dbvariable) + (1 / aHot) + (ThickPipe / aPipe))-K"
 '           Call BiSection
 '           aCool = dbVariable
            aCool = K / (1 - K * (1 / aHot) - K * (ThickPipe / aPipe))
        ElseIf (K <> 0) And (aCool <> 0) And (aHot = 0) And (ThickPipe <> 0) And (aPipe <> 0) Then
 '           str_Formula = "1 / ((1 / aCool) + (1 / dbVariable) + (ThickPipe / aPipe))-K"
 '           Call BiSection
 '           aHot = dbVariable
            aHot = K / (1 - K * (1 / aCool) - K * (ThickPipe / aPipe))
        ElseIf (K <> 0) And (aCool <> 0) And (aHot <> 0) And (ThickPipe = 0) And (aPipe <> 0) Then
 '           str_Formula = "1 / ((1 / aCool) + (1 / aHot) + (dbVariable / aPipe))-K"
 '           Call BiSection
 '           ThickPipe = dbVariable
            ThickPipe = (aPipe / K) * (1 - K * (1 / aCool) - K * (1 / aHot))
        ElseIf (K <> 0) And (aCool <> 0) And (aHot = 0) And (ThickPipe <> 0) And (aPipe = 0) Then
 '           str_Formula = "1 / ((1 / aCool) + (1 / aHot) + (ThickPipe / aPipe))-K"
 '           Call BiSection
 '           aPipe = dbVariable
            aPipe = (1 - K * (1 / aCool) - K * (1 / aHot)) / (K * ThickPipe)
'        Else
 '           If (1 / ((1 / aCool) + (1 / aHot) + (ThickPipe / aPipe)) - K) > 0.2 Or (1 / ((1 / aCool) + (1 / aHot) + (ThickPipe / aPipe)) - K) < -0.2 Then
  '          MsgBox "K,aCool,aHot,ThickPipe,aPipe collide!", vbCritical, "heat_exchange"
   '         End If
        End If
    End Function

