Attribute VB_Name = "Module1"
Public Method As Integer
Public qvLc As Double      '�������
Public qvLh As Double
Public qvVc As Double
Public qvVh As Double
Public qmLc As Double       '��������
Public qmLh As Double
Public qmVc As Double
Public qmVh As Double
Public Tc1 As Double
Public Tc2 As Double
Public dT1 As Double
Public dT2 As Double
Public Th1 As Double
Public Th2 As Double
Public dTm As Double        '����ƽ���¶Ȳ�
Public Cpc As Double
Public Cph As Double
Public aCool As Double      '��Ч����Ĥϵ��
Public aHot As Double
Public aPipe As Double      'ʵ�ʹ��ӵ���ϵ�������ʾ
Public dPipe As Double       '���ӵ�ֱ��
Public ThickPipe As Double   '���ӵĺ��
Public K As Double          '�ܴ���ϵ��
Public A As Double          '�������
Public q As Double          '�������ȵ����������
Public Phai As Double       '����������λW
Public NumVariable As Integer

Public Function initial_HeatV0()
        
        Method = 0
        NumVariable = 0
        qvLc = 0       '�������Text(i)
        qvLh = 0
        qvVc = 0
        qvVh = 0
        qmLc = 0        '��������
        qmLh = 0
        qmVc = 0
        qmVh = 0
        Tc1 = 0         '�¶�
        Tc2 = 0
        dT1 = 0
        dT2 = 0
        Th1 = 0
        Th2 = 0
        dTm = 0         '����ƽ���¶Ȳ�
        Cpc = 0         '��ѹ����
        Cph = 0
        aCool = 0       '��Ч����Ĥϵ��
        aHot = 0
        aPipe = 0       'ʵ�ʹ��ӵ���ϵ�������ʾ
        dPipe = 0        '���ӵ�ֱ��
        ThickPipe = 0    '���ӵĺ��
        K = 0           '�ܴ���ϵ��
        A = 0           '�������
        q = 0           '�������ȵ����������
        Phai = 0        '����������λW
End Function

Public Function initial_Form1textV0()
        
        Form1.Text(0) = 0     '�������
        Form1.Text(1) = 0
        Form1.Text(2) = 0
        Form1.Text(3) = 0
        Form1.Text(4) = 0
        Form1.Text(5) = 0
        Form1.Text(6) = 0
        Form1.Text(7) = 0
        Form1.Text(8) = 0     '��Ч����Ĥϵ��
        Form1.Text(9) = 0
        Form1.Text(10) = 0     'ʵ�ʹ��ӵ���ϵ�������ʾ
        Form1.Text(11) = 0  '���ӵĺ��
        Form1.Text(12) = 0         '�������
        Form1.Text(13) = 0      '����������λW
        Form1.Text(14) = 0         '�ܴ���ϵ��
        Form1.Text(15) = 0
End Function

Public Function Variable_Change()
  
  Dim i As Integer
  
  For i = 0 To 15
   
    If Not IsNumeric(Form1.Text(i)) Then
        Form1.Text(i) = 0
    End If
  Next i
  
  qvLc = Form1.Text(0)       '�������
  qmLh = Form1.Text(1)
  Tc1 = Form1.Text(2)
  Tc2 = Form1.Text(3)
  Th1 = Form1.Text(4)
  Th2 = Form1.Text(5)
  Cpc = Form1.Text(6)
  Cph = Form1.Text(7)
  aCool = Form1.Text(8)       '��Ч����Ĥϵ��
  aHot = Form1.Text(9)
  aPipe = Form1.Text(10)       'ʵ�ʹ��ӵ���ϵ�������ʾ
  ThickPipe = Form1.Text(11)    '���ӵĺ��
  A = Form1.Text(12)           '�������
  Phai = Form1.Text(13)        '����������λW
  K = Form1.Text(14)           '�ܴ���ϵ��
  Method = Form1.Text(15)
  dTm = Form1.Text(16)

End Function

Public Function Variable_OUT()
  
  Form1.Text(0) = qvLc     '�������
  Form1.Text(1) = qmLh
  Form1.Text(2) = Tc1
  Form1.Text(3) = Tc2
  Form1.Text(4) = Th1
  Form1.Text(5) = Th2
  Form1.Text(6) = Cpc
  Form1.Text(7) = Cph
  Form1.Text(8) = aCool      '��Ч����Ĥϵ��
  Form1.Text(9) = aHot
  Form1.Text(10) = aPipe     'ʵ�ʹ��ӵ���ϵ�������ʾ
  Form1.Text(11) = ThickPipe  '���ӵĺ��
  Form1.Text(12) = A         '�������
  Form1.Text(13) = Phai      '����������λW
  Form1.Text(14) = K         '�ܴ���ϵ��
  Form1.Text(15) = Method
  Form1.Text(16) = dTm

End Function

Public Function Variable_INPUT()
  
  Form1.Text(0) = 30000     'qvLc�������
  Form1.Text(1) = 9075      'qmLh
  Form1.Text(2) = 20        'Tc1
  Form1.Text(3) = 0         'Tc2
  Form1.Text(4) = 90        'Th1
  Form1.Text(5) = 40        'Th2
  Form1.Text(6) = 4.18      'Cpc
  Form1.Text(7) = 3.35      'Cph
  Form1.Text(8) = 1000      'aCool      '��Ч����Ĥϵ��
  Form1.Text(9) = 300       'aHot
  Form1.Text(10) = 49       'aPipe     'ʵ�ʹ��ӵ���ϵ�������ʾ
  Form1.Text(11) = 0.0025   'ThickPipe  '���ӵĺ��
  Form1.Text(12) = 0        'A         '�������
  Form1.Text(13) = 0        'Phai      '����������λW
  Form1.Text(14) = 0        'K         '�ܴ���ϵ��
  Form1.Text(15) = 0        'Method
  Form1.Text(16) = 0        'dTm


End Function




