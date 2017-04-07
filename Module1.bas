Attribute VB_Name = "Module1"
Public Method As Integer
Public qvLc As Double      '体积流量
Public qvLh As Double
Public qvVc As Double
Public qvVh As Double
Public qmLc As Double       '质量流量
Public qmLh As Double
Public qmVc As Double
Public qmVh As Double
Public Tc1 As Double
Public Tc2 As Double
Public dT1 As Double
Public dT2 As Double
Public Th1 As Double
Public Th2 As Double
Public dTm As Double        '对数平均温度差
Public Cpc As Double
Public Cph As Double
Public aCool As Double      '有效传热膜系数
Public aHot As Double
Public aPipe As Double      '实际管子导热系数用入表示
Public dPipe As Double       '管子的直径
Public ThickPipe As Double   '管子的厚度
Public K As Double          '总传热系数
Public A As Double          '传热面积
Public q As Double          '对流传热的面积热流量
Public Phai As Double       '热流量，单位W
Public NumVariable As Integer

Public Function initial_HeatV0()
        
        Method = 0
        NumVariable = 0
        qvLc = 0       '体积流量Text(i)
        qvLh = 0
        qvVc = 0
        qvVh = 0
        qmLc = 0        '质量流量
        qmLh = 0
        qmVc = 0
        qmVh = 0
        Tc1 = 0         '温度
        Tc2 = 0
        dT1 = 0
        dT2 = 0
        Th1 = 0
        Th2 = 0
        dTm = 0         '对数平均温度差
        Cpc = 0         '定压热容
        Cph = 0
        aCool = 0       '有效传热膜系数
        aHot = 0
        aPipe = 0       '实际管子导热系数用入表示
        dPipe = 0        '管子的直径
        ThickPipe = 0    '管子的厚度
        K = 0           '总传热系数
        A = 0           '传热面积
        q = 0           '对流传热的面积热流量
        Phai = 0        '热流量，单位W
End Function

Public Function initial_Form1textV0()
        
        Form1.Text(0) = 0     '体积流量
        Form1.Text(1) = 0
        Form1.Text(2) = 0
        Form1.Text(3) = 0
        Form1.Text(4) = 0
        Form1.Text(5) = 0
        Form1.Text(6) = 0
        Form1.Text(7) = 0
        Form1.Text(8) = 0     '有效传热膜系数
        Form1.Text(9) = 0
        Form1.Text(10) = 0     '实际管子导热系数用入表示
        Form1.Text(11) = 0  '管子的厚度
        Form1.Text(12) = 0         '传热面积
        Form1.Text(13) = 0      '热流量，单位W
        Form1.Text(14) = 0         '总传热系数
        Form1.Text(15) = 0
End Function

Public Function Variable_Change()
  
  Dim i As Integer
  
  For i = 0 To 15
   
    If Not IsNumeric(Form1.Text(i)) Then
        Form1.Text(i) = 0
    End If
  Next i
  
  qvLc = Form1.Text(0)       '体积流量
  qmLh = Form1.Text(1)
  Tc1 = Form1.Text(2)
  Tc2 = Form1.Text(3)
  Th1 = Form1.Text(4)
  Th2 = Form1.Text(5)
  Cpc = Form1.Text(6)
  Cph = Form1.Text(7)
  aCool = Form1.Text(8)       '有效传热膜系数
  aHot = Form1.Text(9)
  aPipe = Form1.Text(10)       '实际管子导热系数用入表示
  ThickPipe = Form1.Text(11)    '管子的厚度
  A = Form1.Text(12)           '传热面积
  Phai = Form1.Text(13)        '热流量，单位W
  K = Form1.Text(14)           '总传热系数
  Method = Form1.Text(15)
  dTm = Form1.Text(16)

End Function

Public Function Variable_OUT()
  
  Form1.Text(0) = qvLc     '体积流量
  Form1.Text(1) = qmLh
  Form1.Text(2) = Tc1
  Form1.Text(3) = Tc2
  Form1.Text(4) = Th1
  Form1.Text(5) = Th2
  Form1.Text(6) = Cpc
  Form1.Text(7) = Cph
  Form1.Text(8) = aCool      '有效传热膜系数
  Form1.Text(9) = aHot
  Form1.Text(10) = aPipe     '实际管子导热系数用入表示
  Form1.Text(11) = ThickPipe  '管子的厚度
  Form1.Text(12) = A         '传热面积
  Form1.Text(13) = Phai      '热流量，单位W
  Form1.Text(14) = K         '总传热系数
  Form1.Text(15) = Method
  Form1.Text(16) = dTm

End Function

Public Function Variable_INPUT()
  
  Form1.Text(0) = 30000     'qvLc体积流量
  Form1.Text(1) = 9075      'qmLh
  Form1.Text(2) = 20        'Tc1
  Form1.Text(3) = 0         'Tc2
  Form1.Text(4) = 90        'Th1
  Form1.Text(5) = 40        'Th2
  Form1.Text(6) = 4.18      'Cpc
  Form1.Text(7) = 3.35      'Cph
  Form1.Text(8) = 1000      'aCool      '有效传热膜系数
  Form1.Text(9) = 300       'aHot
  Form1.Text(10) = 49       'aPipe     '实际管子导热系数用入表示
  Form1.Text(11) = 0.0025   'ThickPipe  '管子的厚度
  Form1.Text(12) = 0        'A         '传热面积
  Form1.Text(13) = 0        'Phai      '热流量，单位W
  Form1.Text(14) = 0        'K         '总传热系数
  Form1.Text(15) = 0        'Method
  Form1.Text(16) = 0        'dTm


End Function




