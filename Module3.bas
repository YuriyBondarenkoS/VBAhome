Attribute VB_Name = "Module3"
'������������ �������������� �������
Public Function TGInterp(MasX As Variant, MasY As Variant, x As Double, Flag As String) As Double
    For i = 1 To 71
        If CDbl(MasY(i)) <> 0 Then
            If x = CDbl(MasX(i)) Then
                TGInterp = CDbl(MasY(i))
                Exit Function
            End If
            If x < CDbl(MasX(i)) And x > CDbl(MasX(i + 1)) Then
                TGInterp = CDbl(MasY(i)) + (CDbl(MasY(i + 1)) - CDbl(MasY(i))) * ((x - CDbl(MasX(i))) / (CDbl(MasX(i + 1)) - CDbl(MasX(i))))
                Exit Function
            End If
            If x > CDbl(MasX(i)) And x > CDbl(MasX(i + 1)) Then
                If Flag = "int" Then
                    TGInterp = CDbl(MasY(i))
                    Exit Function
                End If
                If Flag = "ext" Then
                    TGInterp = CDbl(MasY(i)) + (CDbl(MasY(i + 1)) - CDbl(MasY(i))) * ((x - CDbl(MasX(i))) / (CDbl(MasX(i + 1)) - CDbl(MasX(i))))
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

'������� ������������ �������� �������� ������
Function interpolationNorms(x, y, TR)
    If x >= TR.Cells(2, 1) And y >= TR.Cells(1, 2) Then
        For i = 1 To TR.Rows.Count - 1 Step 1
            xc = TR.Cells(i, 1)
            If x >= xc Then
                imin = i
                imax = i + 1
                x1 = TR.Cells(imin, 1)
                x2 = TR.Cells(imax, 1)
            End If
        Next i
        For j = 1 To TR.Columns.Count - 1 Step 1
            yc = TR.Cells(1, j)
            If y >= yc Then
                jmin = j
                jmax = j + 1
                y1 = TR.Cells(1, jmin)
                y2 = TR.Cells(1, jmax)
            End If
        Next j
        f11 = TR.Cells(imin, jmin)
        f12 = TR.Cells(imin, jmax)
        f21 = TR.Cells(imax, jmin)
        f22 = TR.Cells(imax, jmax)
        interpolationNorms = (f11 * (x2 - x) * (y2 - y) + f12 * (x2 - x) * (y - y1) + f21 * (x - x1) * (y2 - y) + f22 * (x - x1) * (y - y1)) / (x2 - x1) / (y2 - y1)
    End If
End Function

'�������� ��� ����������� �������� ���� �������� ������ ��� ��������� ��������� � ������������ ��������� ��������� �� ��������� � ��������� ������������
Function rangeNorms(Year, Direction, Typ, Period) 'As Object
    If Direction = "�������" And Year <= 1989 And Typ = "��������� ���������" Then Set rangeNorms = Sheets("1989N").Range("H107:L134")
    If Direction = "�������" And Year <= 1989 And Typ = "��������� ������������" Then Set rangeNorms = Sheets("1989N").Range("H107:L134")
    If Direction = "�������" And Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And Period <= 5000 Then Set rangeNorms = Sheets("1997N").Range("M116:P144")
    If Direction = "�������" And Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And Period > 5000 Then Set rangeNorms = Sheets("1997N").Range("R116:U144")
    If Direction = "�������" And Year > 1989 And Year <= 1997 And Typ = "��������� ���������" And Period <= 5000 Then Set rangeNorms = Sheets("1997N").Range("P151:T179")
    If Direction = "�������" And Year > 1989 And Year <= 1997 And Typ = "��������� ���������" And Period > 5000 Then Set rangeNorms = Sheets("1997N").Range("V151:Z179")
    If Direction = "�������" And Year > 1997 And Year <= 2003 And Typ = "��������� ������������" And Period <= 5000 Then Set rangeNorms = Sheets("2003N").Range("P117:T145")
    If Direction = "�������" And Year > 1997 And Year <= 2003 And Typ = "��������� ������������" And Period > 5000 Then Set rangeNorms = Sheets("2003N").Range("V117:Z145")
    If Direction = "�������" And Year > 1997 And Year <= 2003 And Typ = "��������� ���������" And Period <= 5000 Then Set rangeNorms = Sheets("2003N").Range("P117:T145")
    If Direction = "�������" And Year > 1997 And Year <= 2003 And Typ = "��������� ���������" And Period > 5000 Then Set rangeNorms = Sheets("2003N").Range("V117:Z145")
End Function

'��������� ��������� � ������������ ����� ��� ������
Function rangeNorms1(Year, Direction, Typ, Period)

    If Direction = "�������" And Year <= 1989 And Typ = "��������� ���������" Then Set rangeNorms1 = Sheets("1989N").Range("N107:R134")
    If Direction = "�������" And Year <= 1989 And Typ = "��������� ������������" Then Set rangeNorms1 = Sheets("1989N").Range("N107:R134")
    If Direction = "�������" And Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And Period <= 5000 Then Set rangeNorms1 = Sheets("1997N").Range("W116:Z144")
    If Direction = "�������" And Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And Period > 5000 Then Set rangeNorms1 = Sheets("1997N").Range("AB116:AE144")
    If Direction = "�������" And Year > 1989 And Year <= 1997 And Typ = "��������� ���������" And Period <= 5000 Then Set rangeNorms1 = Sheets("1997N").Range("AB151:AF179")
    If Direction = "�������" And Year > 1989 And Year <= 1997 And Typ = "��������� ���������" And Period > 5000 Then Set rangeNorms1 = Sheets("1997N").Range("AH151:AL179")
    If Direction = "�������" And Year > 1997 And Year <= 2003 And Typ = "��������� ������������" And Period <= 5000 Then Set rangeNorms1 = Sheets("2003N").Range("AB117:AF145")
    If Direction = "�������" And Year > 1997 And Year <= 2003 And Typ = "��������� ������������" And Period > 5000 Then Set rangeNorms1 = Sheets("2003N").Range("AH117:AL145")
    If Direction = "�������" And Year > 1997 And Year <= 2003 And Typ = "��������� ���������" And Period <= 5000 Then Set rangeNorms1 = Sheets("2003N").Range("AB117:AF145")
    If Direction = "�������" And Year > 1997 And Year <= 2003 And Typ = "��������� ���������" And Period > 5000 Then Set rangeNorms1 = Sheets("2003N").Range("AH117:AL145")
End Function

'���������, ���������, �������
Function rangeNorms3(Year, Direction, Typ, Period)
    If Year <= 1989 And Typ = "���������" Then Set rangeNorms3 = Sheets("1989N").Range("B5:N33")
    If Year <= 1989 And Typ = "�������" Then Set rangeNorms3 = Sheets("1989N").Range("B38:N66")
    If Year <= 1989 And Typ = "���������" Then Set rangeNorms3 = Sheets("1989N").Range("B71:N99")
    If Direction = "������" And Year <= 1989 And Typ = "��������� ���������" Then Set rangeNorms3 = Sheets("1989N").Range("N107:R135")
    If Direction = "������" And Year <= 1989 And Typ = "��������� ������������" Then Set rangeNorms3 = Sheets("1989N").Range("N107:R135")
    If Year > 1989 And Year <= 1997 And Typ = "���������" And Period <= 5000 Then Set rangeNorms3 = Sheets("1997N").Range("B6:M34")
    If Year > 1989 And Year <= 1997 And Typ = "���������" And Period > 5000 Then Set rangeNorms3 = Sheets("1997N").Range("N6:Y34")
    If Year > 1989 And Year <= 1997 And Typ = "���������" And Period <= 5000 Then Set rangeNorms3 = Sheets("1997N").Range("B42:L69")
    If Year > 1989 And Year <= 1997 And Typ = "���������" And Period > 5000 Then Set rangeNorms3 = Sheets("1997N").Range("M42:W69")
    If Year > 1989 And Year <= 1997 And Typ = "�������" And Period <= 5000 Then Set rangeNorms3 = Sheets("1997N").Range("B78:L105")
    If Year > 1989 And Year <= 1997 And Typ = "�������" And Period > 5000 Then Set rangeNorms3 = Sheets("1997N").Range("M78:W105")
    If Direction = "������" And Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And Period <= 5000 Then Set rangeNorms3 = Sheets("1997N").Range("W116:Z144")
    If Direction = "������" And Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And Period > 5000 Then Set rangeNorms3 = Sheets("1997N").Range("AB116:AE144")
    If Direction = "������" And Year > 1989 And Year <= 1997 And Typ = "��������� ���������" And Period <= 5000 Then Set rangeNorms3 = Sheets("1997N").Range("AB151:AF179")
    If Direction = "������" And Year > 1989 And Year <= 1997 And Typ = "��������� ���������" And Period > 5000 Then Set rangeNorms3 = Sheets("1997N").Range("AH151:AL179")
    If Year > 1997 And Year <= 2003 And Typ = "���������" And Period <= 5000 Then Set rangeNorms3 = Sheets("2003N").Range("B6:M34")
    If Year > 1997 And Year <= 2003 And Typ = "���������" And Period > 5000 Then Set rangeNorms3 = Sheets("2003N").Range("N6:Y34")
    If Year > 1997 And Year <= 2003 And Typ = "���������" And Period <= 5000 Then Set rangeNorms3 = Sheets("2003N").Range("B42:L70")
    If Year > 1997 And Year <= 2003 And Typ = "���������" And Period > 5000 Then Set rangeNorms3 = Sheets("2003N").Range("M42:W70")
    If Year > 1997 And Year <= 2003 And Typ = "�������" And Period <= 5000 Then Set rangeNorms3 = Sheets("2003N").Range("B79:L107")
    If Year > 1997 And Year <= 2003 And Typ = "�������" And Period > 5000 Then Set rangeNorms3 = Sheets("2003N").Range("M79:W107")
    If Direction = "������" And Year > 1997 And Year <= 2003 And Typ = "��������� ������������" And Period <= 5000 Then Set rangeNorms3 = Sheets("2003N").Range("AB117:AF145")
    If Direction = "������" And Year > 1997 And Year <= 2003 And Typ = "��������� ������������" And Period > 5000 Then Set rangeNorms3 = Sheets("2003N").Range("AH117:AL145")
    If Direction = "������" And Year > 1997 And Year <= 2003 And Typ = "��������� ���������" And Period <= 5000 Then Set rangeNorms3 = Sheets("2003N").Range("AB117:AF145")
    If Direction = "������" And Year > 1997 And Year <= 2003 And Typ = "��������� ���������" And Period > 5000 Then Set rangeNorms3 = Sheets("2003N").Range("AH117:AL145")
    If Year >= 2004 And Typ = "���������" And Period <= 5000 Then Set rangeNorms3 = Sheets("2004N").Range("B6:M34")
    If Year >= 2004 And Typ = "���������" And Period > 5000 Then Set rangeNorms3 = Sheets("2004N").Range("N6:Y34")
    If Year >= 2004 And Typ = "���������" And Period <= 5000 Then Set rangeNorms3 = Sheets("2004N").Range("B42:L70")
    If Year >= 2004 And Typ = "���������" And Period > 5000 Then Set rangeNorms3 = Sheets("2004N").Range("M42:W70")
    If Year >= 2004 And Typ = "�������" And Period <= 5000 Then Set rangeNorms3 = Sheets("2004N").Range("B78:L106")
    If Year >= 2004 And Typ = "�������" And Period > 5000 Then Set rangeNorms3 = Sheets("2004N").Range("M78:W106")
End Function

'��� ������. ��������� ��������� � ������������ ����� ��� ������ � ������� ��������� (2004)
Function rangeNorms4(Year, Direction, Typ, Period)
    If Year > 2003 And Typ = "��������� ���������" Then Set rangeNorms4 = Sheets("1989N").Range("H107:L134")
    If Year > 2003 And Typ = "��������� ������������" Then Set rangeNorms4 = Sheets("1989N").Range("H107:L134")
End Function

'��������� ��������� � ������������ ������� ����� ��� ������ (2004)
Function rangeNorms5(Year, Direction, Typ, Period)
    If Year > 2003 And Typ = "��������� ���������" Then Set rangeNorms5 = Sheets("1989N").Range("N107:R134")
    If Year > 2003 And Typ = "��������� ������������" Then Set rangeNorms5 = Sheets("1989N").Range("N107:R134")
End Function

Function rangeNorms6(Year, Direction, Typ, Period)
    If Year > 2003 And Typ = "��������� ���������" And Period <= 5000 Then Set rangeNorms6 = Sheets("2004N").Range("J117:N145")
    If Year > 2003 And Typ = "��������� ���������" And Period > 5000 Then Set rangeNorms6 = Sheets("2004N").Range("P117:T145")
    If Year > 2003 And Typ = "��������� ������������" And Period <= 5000 Then Set rangeNorms6 = Sheets("2004N").Range("J154:N182")
    If Year > 2003 And Typ = "��������� ������������" And Period > 5000 Then Set rangeNorms6 = Sheets("2004N").Range("P154:T182")
End Function

Function rangeNorms7(Year, Direction, Typ, Period)
    If Year > 2003 And Typ = "��������� ���������" And Period <= 5000 Then Set rangeNorms7 = Sheets("2004N").Range("J117:N145")
    If Year > 2003 And Typ = "��������� ���������" And Period > 5000 Then Set rangeNorms7 = Sheets("2004N").Range("P117:T145")
    If Year > 2003 And Typ = "��������� ������������" And Period <= 5000 Then Set rangeNorms7 = Sheets("2004N").Range("J154:N182")
    If Year > 2003 And Typ = "��������� ������������" And Period > 5000 Then Set rangeNorms7 = Sheets("2004N").Range("P154:T182")
End Function

'������ ������ �������� ����
Function periodWork(Chart, Regime)
    If Chart = "������ 1" And Regime = "���" Then j = 22                'Period = timework(Chart, Regime)
    If Chart = "������ 2" And Regime = "���" Then j = 23
    If Chart = "������ 3" And Regime = "���" Then j = 24
    If Chart = "������ 4" And Regime = "���" Then j = 25
    If Chart = "������ 5" And Regime = "���" Then j = 26
    If Chart = "������ 6" And Regime = "���" Then j = 27
    If Chart = "������ 7" And Regime = "���" Then j = 28
    If Chart = "������ 8" And Regime = "���" Then j = 29
    If Chart = "������ 9" And Regime = "���" Then j = 30
    If Chart = "������ 10" And Regime = "���" Then j = 31
        If Chart = "������ 1" And Regime = "��" Then j = 2
        If Chart = "������ 2" And Regime = "��" Then j = 4
        If Chart = "������ 3" And Regime = "��" Then j = 6
        If Chart = "������ 4" And Regime = "��" Then j = 8
        If Chart = "������ 5" And Regime = "��" Then j = 10
        If Chart = "������ 6" And Regime = "��" Then j = 12
        If Chart = "������ 7" And Regime = "��" Then j = 14
        If Chart = "������ 8" And Regime = "��" Then j = 16
        If Chart = "������ 9" And Regime = "��" Then j = 18
        If Chart = "������ 10" And Regime = "��" Then j = 20
            If Chart = "������ 1" And Regime = "����" Then j = 3
            If Chart = "������ 2" And Regime = "����" Then j = 5
            If Chart = "������ 3" And Regime = "����" Then j = 7
            If Chart = "������ 4" And Regime = "����" Then j = 9
            If Chart = "������ 5" And Regime = "����" Then j = 11
            If Chart = "������ 6" And Regime = "����" Then j = 13
            If Chart = "������ 7" And Regime = "����" Then j = 15
            If Chart = "������ 8" And Regime = "����" Then j = 17
            If Chart = "������ 9" And Regime = "����" Then j = 19
            If Chart = "������ 10" And Regime = "����" Then j = 21
                periodWork = Sheets("temperature").Cells(19, j)
End Function

'��������� ����������� ��������� ������������
Function flowTemperature(Typ, Chart, Regime, Direction)
    If Typ = "���������" And Regime = "��" Then i = 26                     'y = Tpodacha(Typ, Chart, Regime, Direction)
    If Typ = "���������" And Regime = "���" Then i = 25
    If Typ = "���������" And Regime = "����" Then i = 27
    If Typ = "��������� ���������" And Regime = "��" Then i = 29
    If Typ = "��������� ���������" And Regime = "���" Then i = 28
    If Typ = "��������� ���������" And Regime = "����" Then i = 30
    If Typ = "��������� ������������" And Regime = "��" Then i = 29
    If Typ = "��������� ������������" And Regime = "���" Then i = 28
    If Typ = "��������� ������������" And Regime = "����" Then i = 30
    If Typ = "���������" And Regime = "��" Then i = 32
    If Typ = "���������" And Regime = "���" Then i = 31
    If Typ = "���������" And Regime = "����" Then i = 33
    If Typ = "�������" And Regime = "��" Then i = 35
    If Typ = "�������" And Regime = "���" Then i = 34
    If Typ = "�������" And Regime = "����" Then i = 36
        If Direction = "������" And Chart = "������ 1" Then j = 34
        If Direction = "������" And Chart = "������ 2" Then j = 36
        If Direction = "������" And Chart = "������ 3" Then j = 38
        If Direction = "������" And Chart = "������ 4" Then j = 40
        If Direction = "������" And Chart = "������ 5" Then j = 42
        If Direction = "������" And Chart = "������ 6" Then j = 44
        If Direction = "������" And Chart = "������ 7" Then j = 46
        If Direction = "������" And Chart = "������ 8" Then j = 48
        If Direction = "������" And Chart = "������ 9" Then j = 50
        If Direction = "������" And Chart = "������ 10" Then j = 52
            If Direction = "�������" And Chart = "������ 1" Then j = 35
            If Direction = "�������" And Chart = "������ 2" Then j = 37
            If Direction = "�������" And Chart = "������ 3" Then j = 39
            If Direction = "�������" And Chart = "������ 4" Then j = 41
            If Direction = "�������" And Chart = "������ 5" Then j = 43
            If Direction = "�������" And Chart = "������ 6" Then j = 45
            If Direction = "�������" And Chart = "������ 7" Then j = 47
            If Direction = "�������" And Chart = "������ 8" Then j = 49
            If Direction = "�������" And Chart = "������ 9" Then j = 51
            If Direction = "�������" And Chart = "������ 10" Then j = 53
                flowTemperature = Sheets("temperature").Cells(i, j)
End Function

'��������� ����������� ��������� ������������
Function returnTemperature(Typ, Chart, Regime, Direction)
    If Typ = "���������" And Regime = "��" Then i = 26                     'y = Tpodacha(Typ, Chart, Regime, Direction)
    If Typ = "���������" And Regime = "���" Then i = 25
    If Typ = "���������" And Regime = "����" Then i = 27
    If Typ = "��������� ���������" And Regime = "��" Then i = 29
    If Typ = "��������� ���������" And Regime = "���" Then i = 28
    If Typ = "��������� ���������" And Regime = "����" Then i = 30
    If Typ = "��������� ������������" And Regime = "��" Then i = 29
    If Typ = "��������� ������������" And Regime = "���" Then i = 28
    If Typ = "��������� ������������" And Regime = "����" Then i = 30
    If Typ = "���������" And Regime = "��" Then i = 32
    If Typ = "���������" And Regime = "���" Then i = 31
    If Typ = "���������" And Regime = "����" Then i = 33
    If Typ = "�������" And Regime = "��" Then i = 35
    If Typ = "�������" And Regime = "���" Then i = 34
    If Typ = "�������" And Regime = "����" Then i = 36
        If Direction = "������" And Chart = "������ 1" Then j = 34
        If Direction = "������" And Chart = "������ 2" Then j = 36
        If Direction = "������" And Chart = "������ 3" Then j = 38
        If Direction = "������" And Chart = "������ 4" Then j = 40
        If Direction = "������" And Chart = "������ 5" Then j = 42
        If Direction = "������" And Chart = "������ 6" Then j = 44
        If Direction = "������" And Chart = "������ 7" Then j = 46
        If Direction = "������" And Chart = "������ 8" Then j = 48
        If Direction = "������" And Chart = "������ 9" Then j = 50
        If Direction = "������" And Chart = "������ 10" Then j = 52
            If Direction = "�������" And Chart = "������ 1" Then j = 35
            If Direction = "�������" And Chart = "������ 2" Then j = 37
            If Direction = "�������" And Chart = "������ 3" Then j = 39
            If Direction = "�������" And Chart = "������ 4" Then j = 41
            If Direction = "�������" And Chart = "������ 5" Then j = 43
            If Direction = "�������" And Chart = "������ 6" Then j = 45
            If Direction = "�������" And Chart = "������ 7" Then j = 47
            If Direction = "�������" And Chart = "������ 8" Then j = 49
            If Direction = "�������" And Chart = "������ 9" Then j = 51
            If Direction = "�������" And Chart = "������ 10" Then j = 53
                If Direction = "������" Then j1 = 1
                If Direction = "�������" Then j1 = -1
                    k = j + j1
                    returnTemperature = Sheets("temperature").Cells(i, k)
End Function

'������ ���� �������� ������ ��� ������� calculation. �������� ����� �������� ������
Function qn(Year As Integer, Direction As String, Typ As String, x As Double, Chart As String, Regime As String)
Dim TR As Object
Dim TR1 As Object
Dim TR3 As Object
Dim qn1 As Double
Dim qn2 As Double
Dim qn3 As Double
Dim Period As Integer
'������
    If Typ <> "��������� ���������" And Typ <> "��������� ������������" And Typ <> "���������" And Typ <> "���������" And Typ <> "�������" Then qn = "typ"
    If Chart <> "������ 1" And Chart <> "������ 2" And Chart <> "������ 3" And Chart <> "������ 4" And Chart <> "������ 5" And Chart <> "������ 6" And Chart <> "������ 7" _
    And Chart <> "������ 8" And Chart <> "������ 9" And Chart <> "������ 10" Then qn = "chart"
    If Regime <> "���" And Regime <> "��" And Regime <> "����" Then qn = "regime"
    If Direction <> "������" And Direction <> "�������" Then qn = "direction"
    If Year < 1900 Or Year > 2020 Then qn = "year"
    '�����������.Range("BB7:BB18").Interior.ColorIndex = 35
    'If x = 0 Then qn = "diameter"
    'If qn = "typ" Or qn = "chart" Or qn = "regime" Or qn = "direction" Or qn = "year" Or qn = "diameter" Then MsgBox ("Wrong " + qn), vbInformation, "Data entered incorrectly"
'        If qn = "typ" Or qn = "chart" Or qn = "regime" Or qn = "direction" Or qn = "year" Or qn = "diameter" Then Finderer01
            If qn = "typ" Or qn = "chart" Or qn = "regime" Or qn = "direction" Or qn = "year" Or qn = "diameter" Then Exit Function
'������ ������
    Period = periodWork(Chart, Regime) 'Sheets("temperature").Cells(19, j)
'��������� ��������� � ������������ ����� ��� ������ � ������� ���������
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Direction = "�������" And Year > 2003 And Typ = "��������� ���������" _
Or Direction = "�������" And Year > 2003 And Typ = "��������� ������������" Then
    qn1 = 0
    Else: qn1 = interpolationNorms(x, flowTemperature(Typ, Chart, Regime, Direction), rangeNorms(Year, Direction, Typ, Period))
End If
'��������� ��������� � ������������ ����� ��� ������
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Direction = "�������" And Year > 2003 And Typ = "��������� ���������" _
Or Direction = "�������" And Year > 2003 And Typ = "��������� ������������" Then
    qn2 = 0
    Else: qn2 = interpolationNorms(x, returnTemperature(Typ, Chart, Regime, Direction), rangeNorms1(Year, Direction, Typ, Period))
End If
'���������, ���������, �������
If Direction = "�������" And Typ = "��������� ���������" Or Direction = "�������" And Typ = "��������� ������������" Or Typ = "��������� ���������" _
And Direction = "������" And Year >= 2004 Or Typ = "��������� ������������" And Direction = "������" And Year >= 2004 Then
    qn3 = 0
    Else: qn3 = interpolationNorms(x, flowTemperature(Typ, Chart, Regime, Direction), rangeNorms3(Year, Direction, Typ, Period))
End If
'��� ������
'��������� ��������� � ������������ ����� ��� ������ � ������� ��������� (2004)
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn4 = 0
    Else: qn4 = interpolationNorms(x, returnTemperature(Typ, Chart, Regime, Direction), rangeNorms4(Year, Direction, Typ, Period))
End If
'��������� ��������� � ������������ ������� ����� ��� ������ (2004)
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn5 = 0
    Else: qn5 = interpolationNorms(x, flowTemperature(Typ, Chart, Regime, Direction), rangeNorms5(Year, Direction, Typ, Period))
End If
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    K2004p = 0
    Else: K2004p = qn5 / qn4
End If
'��� �������
'��������� ��������� � ������������ ����� ��� ������ � ������� ��������� (2004)
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn8 = 0
    Else: qn8 = interpolationNorms(x, flowTemperature(Typ, Chart, Regime, Direction), rangeNorms4(Year, Direction, Typ, Period))
End If
'��������� ��������� � ������������ ������� ����� ��� ������ (2004)
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn9 = 0
    Else: qn9 = interpolationNorms(x, returnTemperature(Typ, Chart, Regime, Direction), rangeNorms5(Year, Direction, Typ, Period))
End If
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    K2004o = 0
    Else: K2004o = 1 - qn9 / qn8
End If
'��� ������
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    qn6 = 0
    Else: qn6 = interpolationNorms(x, returnTemperature(Typ, Chart, Regime, Direction), rangeNorms6(Year, Direction, Typ, Period))
End If
'��� �������
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    qn7 = 0
    Else: qn7 = interpolationNorms(x, flowTemperature(Typ, Chart, Regime, Direction), rangeNorms7(Year, Direction, Typ, Period))
End If
    qn = qn1 - qn2 + qn3 + K2004p * qn6 + K2004o * qn7
End Function

'����������� ���
Function kiz(Year As Integer, Typ As String, x As Double, insulation As String) As Double
    If Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And x >= 25 And x <= 70 And insulation = "��������������" Then
    kiz = 0.5
    ElseIf Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And x >= 80 And x <= 175 And insulation = "��������������" Then
    kiz = 0.6
    ElseIf Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And x >= 200 And x <= 300 And insulation = "��������������" Then
    kiz = 0.7
    ElseIf Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And x >= 350 And x <= 500 And insulation = "��������������" Then
    kiz = 0.8
    ElseIf Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And x >= 25 And x <= 70 And insulation = "��������� ��������� ��" Then
    kiz = 0.5
    ElseIf Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And x >= 80 And x <= 175 And insulation = "��������� ��������� ��" Then
    kiz = 0.6
    ElseIf Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And x >= 200 And x <= 300 And insulation = "��������� ��������� ��" Then
    kiz = 0.7
    ElseIf Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And x >= 350 And x <= 500 And insulation = "��������� ��������� ��" Then
    kiz = 0.8
    ElseIf Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And x >= 25 And x <= 70 And insulation = "������������" Then
    kiz = 0.7
    ElseIf Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And x >= 80 And x <= 175 And insulation = "������������" Then
    kiz = 0.8
    ElseIf Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And x >= 200 And x <= 300 And insulation = "������������" Then
    kiz = 0.9
    ElseIf Year > 1989 And Year <= 1997 And Typ = "��������� ������������" And x >= 350 And x <= 500 And insulation = "������������" Then
    kiz = 1
    Else: kiz = 1
End If
End Function

'����������� ������� �������� ������
Function betta(Typ As String, x As Double) As Double
    If x < 150 And Typ <> "��������� ������������" Then
        betta = 1.2
        ElseIf Typ = "��������� ������������" Then
        betta = 1.15
        Else: betta = 1.15
    End If
End Function

'�������� �������
Function condition_diameter(Dy As Double) As Double
If Dy <= 10 Then
condition_diameter = 15
ElseIf Dy > 10 And Dy <= 18 Then
condition_diameter = 15
  ElseIf Dy > 18 And Dy <= 25 Then
  condition_diameter = 25
ElseIf Dy > 25 And Dy <= 28 Then
condition_diameter = 25
  ElseIf Dy > 28 And Dy <= 32 Then
  condition_diameter = 32
ElseIf Dy > 32 And Dy <= 38 Then
condition_diameter = 32
  ElseIf Dy > 38 And Dy <= 40 Then
  condition_diameter = 40
ElseIf Dy > 40 And Dy <= 45 Then
condition_diameter = 40
  ElseIf Dy > 45 And Dy <= 50 Then
  condition_diameter = 50
ElseIf Dy > 50 And Dy <= 57 Then
condition_diameter = 50
  ElseIf Dy > 57 And Dy <= 65 Then
  condition_diameter = 65
ElseIf Dy > 65 And Dy <= 72 Then
condition_diameter = 65
  ElseIf Dy > 72 And Dy <= 80 Then
  condition_diameter = 70
ElseIf Dy > 80 And Dy <= 90 Then
condition_diameter = 80
  ElseIf Dy > 90 And Dy <= 100 Then
  condition_diameter = 100
ElseIf Dy > 100 And Dy <= 113 Then
condition_diameter = 100
  ElseIf Dy > 113 And Dy <= 125 Then
  condition_diameter = 125
ElseIf Dy > 125 And Dy <= 137 Then
condition_diameter = 125
  ElseIf Dy > 137 And Dy <= 150 Then
  condition_diameter = 150
ElseIf Dy > 150 And Dy <= 167 Then
condition_diameter = 150
  ElseIf Dy > 167 And Dy <= 175 Then
  condition_diameter = 175
ElseIf Dy > 175 And Dy <= 194 Then
condition_diameter = 175
  ElseIf Dy > 194 And Dy <= 200 Then
  condition_diameter = 200
ElseIf Dy > 200 And Dy <= 219 Then
condition_diameter = 200
  ElseIf Dy > 219 And Dy <= 225 Then
  condition_diameter = 225
ElseIf Dy > 225 And Dy <= 237 Then
condition_diameter = 225
  ElseIf Dy > 237 And Dy <= 250 Then
  condition_diameter = 250
ElseIf Dy > 250 And Dy <= 275 Then
condition_diameter = 250
  ElseIf Dy > 275 And Dy <= 300 Then
  condition_diameter = 300
ElseIf Dy > 300 And Dy <= 325 Then
condition_diameter = 300
  ElseIf Dy > 325 And Dy <= 350 Then
  condition_diameter = 350
ElseIf Dy > 350 And Dy <= 380 Then
condition_diameter = 350
  ElseIf Dy > 380 And Dy <= 400 Then
  condition_diameter = 400
ElseIf Dy > 400 And Dy <= 425 Then
condition_diameter = 400
  ElseIf Dy > 425 And Dy <= 450 Then
  condition_diameter = 400
ElseIf Dy > 450 And Dy <= 480 Then
condition_diameter = 450
  ElseIf Dy > 480 And Dy <= 500 Then
  condition_diameter = 500
ElseIf Dy > 500 And Dy <= 550 Then
condition_diameter = 500
ElseIf Dy > 550 And Dy <= 600 Then
condition_diameter = 600
ElseIf Dy > 600 And Dy <= 650 Then
condition_diameter = 600
ElseIf Dy > 650 And Dy <= 700 Then
condition_diameter = 700
ElseIf Dy > 700 And Dy <= 750 Then
condition_diameter = 700
ElseIf Dy > 750 And Dy <= 800 Then
condition_diameter = 800
ElseIf Dy > 800 And Dy <= 850 Then
condition_diameter = 800
ElseIf Dy > 850 And Dy <= 900 Then
condition_diameter = 900
ElseIf Dy > 900 And Dy <= 950 Then
condition_diameter = 900
ElseIf Dy > 950 And Dy <= 1000 Then
condition_diameter = 1000
ElseIf Dy > 1000 And Dy <= 1050 Then
condition_diameter = 1000
ElseIf Dy > 1050 And Dy <= 1150 Then
condition_diameter = 1100
ElseIf Dy > 1150 And Dy <= 1200 Then
condition_diameter = 1200
ElseIf Dy > 1200 And Dy <= 1300 Then
condition_diameter = 1200
ElseIf Dy > 1300 And Dy <= 1450 Then
condition_diameter = 1400
End If
End Function

'Sub gjgh()
'F = Qyn(667.309, "������ 1", "���")
'End Sub
'������ ����� � ��������
Function Qyn(gyn As Double, Chart As String, Regime As String) As Double
Dim t1 As Double
Dim t2 As Double
Dim tx As Double
Dim b As Double
    b = Sheets("temperature").Cells(37, 34)
        If Chart = "������ 1" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 34)
            t2 = Sheets("temperature").Cells(19, 35)
            tx = Sheets("temperature").Cells(19, 54)
            ElseIf Chart = "������ 2" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 36)
            t2 = Sheets("temperature").Cells(19, 37)
            tx = Sheets("temperature").Cells(19, 55)
            ElseIf Chart = "������ 3" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 38)
            t2 = Sheets("temperature").Cells(19, 39)
            tx = Sheets("temperature").Cells(19, 56)
            ElseIf Chart = "������ 4" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 40)
            t2 = Sheets("temperature").Cells(19, 41)
            tx = Sheets("temperature").Cells(19, 57)
            ElseIf Chart = "������ 5" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 42)
            t2 = Sheets("temperature").Cells(19, 43)
            tx = Sheets("temperature").Cells(19, 58)
            ElseIf Chart = "������ 6" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 44)
            t2 = Sheets("temperature").Cells(19, 45)
            tx = Sheets("temperature").Cells(19, 59)
            ElseIf Chart = "������ 7" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 46)
            t2 = Sheets("temperature").Cells(19, 47)
            tx = Sheets("temperature").Cells(19, 60)
            ElseIf Chart = "������ 8" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 48)
            t2 = Sheets("temperature").Cells(19, 49)
            tx = Sheets("temperature").Cells(19, 61)
            ElseIf Chart = "������ 9" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 50)
            t2 = Sheets("temperature").Cells(19, 51)
            tx = Sheets("temperature").Cells(19, 62)
            ElseIf Chart = "������ 10" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 52)
            t2 = Sheets("temperature").Cells(19, 53)
            tx = Sheets("temperature").Cells(19, 63)
            ElseIf Chart = "������ 1" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 34)
            t2 = Sheets("temperature").Cells(20, 35)
            tx = Sheets("temperature").Cells(20, 54)
            ElseIf Chart = "������ 2" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 36)
            t2 = Sheets("temperature").Cells(20, 37)
            tx = Sheets("temperature").Cells(20, 55)
            ElseIf Chart = "������ 3" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 38)
            t2 = Sheets("temperature").Cells(20, 39)
            tx = Sheets("temperature").Cells(20, 56)
            ElseIf Chart = "������ 4" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 40)
            t2 = Sheets("temperature").Cells(20, 41)
            tx = Sheets("temperature").Cells(20, 57)
            ElseIf Chart = "������ 5" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 42)
            t2 = Sheets("temperature").Cells(20, 43)
            tx = Sheets("temperature").Cells(20, 58)
            ElseIf Chart = "������ 6" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 44)
            t2 = Sheets("temperature").Cells(20, 45)
            tx = Sheets("temperature").Cells(20, 59)
            ElseIf Chart = "������ 7" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 46)
            t2 = Sheets("temperature").Cells(20, 47)
            tx = Sheets("temperature").Cells(20, 60)
            ElseIf Chart = "������ 8" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 48)
            t2 = Sheets("temperature").Cells(20, 49)
            tx = Sheets("temperature").Cells(20, 61)
            ElseIf Chart = "������ 9" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 50)
            t2 = Sheets("temperature").Cells(20, 51)
            tx = Sheets("temperature").Cells(20, 62)
            ElseIf Chart = "������ 10" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 52)
            t2 = Sheets("temperature").Cells(20, 53)
            tx = Sheets("temperature").Cells(20, 63)
            ElseIf Chart = "������ 1" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 34)
            t2 = Sheets("temperature").Cells(21, 35)
            tx = Sheets("temperature").Cells(21, 54)
            ElseIf Chart = "������ 2" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 36)
            t2 = Sheets("temperature").Cells(21, 37)
            tx = Sheets("temperature").Cells(21, 55)
            ElseIf Chart = "������ 3" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 38)
            t2 = Sheets("temperature").Cells(21, 39)
            tx = Sheets("temperature").Cells(21, 56)
            ElseIf Chart = "������ 4" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 40)
            t2 = Sheets("temperature").Cells(21, 41)
            tx = Sheets("temperature").Cells(21, 57)
            ElseIf Chart = "������ 5" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 42)
            t2 = Sheets("temperature").Cells(21, 43)
            tx = Sheets("temperature").Cells(21, 58)
            ElseIf Chart = "������ 6" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 44)
            t2 = Sheets("temperature").Cells(21, 45)
            tx = Sheets("temperature").Cells(21, 59)
            ElseIf Chart = "������ 7" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 46)
            t2 = Sheets("temperature").Cells(21, 47)
            tx = Sheets("temperature").Cells(21, 60)
            ElseIf Chart = "������ 8" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 48)
            t2 = Sheets("temperature").Cells(21, 49)
            tx = Sheets("temperature").Cells(21, 61)
            ElseIf Chart = "������ 9" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 50)
            t2 = Sheets("temperature").Cells(21, 51)
            tx = Sheets("temperature").Cells(21, 62)
            ElseIf Chart = "������ 10" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 52)
            t2 = Sheets("temperature").Cells(21, 53)
            tx = Sheets("temperature").Cells(21, 63)
        End If
            Qyn = (gyn * (b * t1 + (1 - b) * t2 - tx) * wskDSWT((t1 * b + t2 * (1 - b) - tx))) / 1000000
End Function

'������ ����� �� ������������ ���������
Function Qisp(gisp As Double, Chart As String, Regime As String) As Double
Dim tisp1 As Double
Dim tx As Double
    tisp1 = Sheets("temperature").Cells(38, 34)
        If Chart = "������ 1" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 54)
        If Chart = "������ 2" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 55)
        If Chart = "������ 3" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 56)
        If Chart = "������ 4" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 57)
        If Chart = "������ 5" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 58)
        If Chart = "������ 6" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 59)
        If Chart = "������ 7" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 60)
        If Chart = "������ 8" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 61)
        If Chart = "������ 9" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 62)
        If Chart = "������ 10" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 63)
            If Chart = "������ 1" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 54)
            If Chart = "������ 2" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 55)
            If Chart = "������ 3" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 56)
            If Chart = "������ 4" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 57)
            If Chart = "������ 5" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 58)
            If Chart = "������ 6" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 59)
            If Chart = "������ 7" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 60)
            If Chart = "������ 8" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 61)
            If Chart = "������ 9" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 62)
            If Chart = "������ 10" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 63)
                If Chart = "������ 1" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 54)
                If Chart = "������ 2" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 55)
                If Chart = "������ 3" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 56)
                If Chart = "������ 4" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 57)
                If Chart = "������ 5" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 58)
                If Chart = "������ 6" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 59)
                If Chart = "������ 7" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 60)
                If Chart = "������ 8" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 61)
                If Chart = "������ 9" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 62)
                If Chart = "������ 10" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 63)
                    Qisp = gisp * (tisp1 - tx) * wskDSWT(tisp1 - tx) / 1000000
End Function

'������ ����� ��� ����������
Function Qzp(gzp, Chart, Regime)
Dim tzp1 As Double
Dim tx As Double
    tzp1 = Sheets("temperature").Cells(39, 34)
        If Chart = "������ 1" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 54)
        If Chart = "������ 2" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 55)
        If Chart = "������ 3" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 56)
        If Chart = "������ 4" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 57)
        If Chart = "������ 5" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 58)
        If Chart = "������ 6" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 59)
        If Chart = "������ 7" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 60)
        If Chart = "������ 8" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 61)
        If Chart = "������ 9" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 62)
        If Chart = "������ 10" And Regime = "���" Then tx = Sheets("temperature").Cells(19, 63)
            If Chart = "������ 1" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 54)
            If Chart = "������ 2" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 55)
            If Chart = "������ 3" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 56)
            If Chart = "������ 4" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 57)
            If Chart = "������ 5" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 58)
            If Chart = "������ 6" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 59)
            If Chart = "������ 7" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 60)
            If Chart = "������ 8" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 61)
            If Chart = "������ 9" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 62)
            If Chart = "������ 10" And Regime = "��" Then tx = Sheets("temperature").Cells(20, 63)
                If Chart = "������ 1" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 54)
                If Chart = "������ 2" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 55)
                If Chart = "������ 3" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 56)
                If Chart = "������ 4" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 57)
                If Chart = "������ 5" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 58)
                If Chart = "������ 6" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 59)
                If Chart = "������ 7" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 60)
                If Chart = "������ 8" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 61)
                If Chart = "������ 9" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 62)
                If Chart = "������ 10" And Regime = "����" Then tx = Sheets("temperature").Cells(21, 63)
                    Qzp = gzp * (tzp1 - tx) * wskDSWT(tzp1 - tx) / 1000000
End Function

'����������� � �������� ������������
Function Tpodacha(Typ As String, Chart As String, Regime As String, Direction As String) As Double
Dim i As Byte
Dim j As Byte
Dim j1 As Byte
    If Typ = "���������" And Regime = "��" Then i = 26
    If Typ = "���������" And Regime = "���" Then i = 25
    If Typ = "���������" And Regime = "����" Then i = 27
    If Typ = "��������� ���������" And Regime = "��" Then i = 29
    If Typ = "��������� ���������" And Regime = "���" Then i = 28
    If Typ = "��������� ���������" And Regime = "����" Then i = 30
    If Typ = "��������� ������������" And Regime = "��" Then i = 29
    If Typ = "��������� ������������" And Regime = "���" Then i = 28
    If Typ = "��������� ������������" And Regime = "����" Then i = 30
    If Typ = "���������" And Regime = "��" Then i = 32
    If Typ = "���������" And Regime = "���" Then i = 31
    If Typ = "���������" And Regime = "����" Then i = 33
    If Typ = "�������" And Regime = "��" Then i = 35
    If Typ = "�������" And Regime = "���" Then i = 34
    If Typ = "�������" And Regime = "����" Then i = 36
        If Direction = "������" And Chart = "������ 1" Then j = 34
        If Direction = "������" And Chart = "������ 2" Then j = 36
        If Direction = "������" And Chart = "������ 3" Then j = 38
        If Direction = "������" And Chart = "������ 4" Then j = 40
        If Direction = "������" And Chart = "������ 5" Then j = 42
        If Direction = "������" And Chart = "������ 6" Then j = 44
        If Direction = "������" And Chart = "������ 7" Then j = 46
        If Direction = "������" And Chart = "������ 8" Then j = 48
        If Direction = "������" And Chart = "������ 9" Then j = 50
        If Direction = "������" And Chart = "������ 10" Then j = 52
            If Direction = "�������" And Chart = "������ 1" Then j = 35
            If Direction = "�������" And Chart = "������ 2" Then j = 37
            If Direction = "�������" And Chart = "������ 3" Then j = 39
            If Direction = "�������" And Chart = "������ 4" Then j = 41
            If Direction = "�������" And Chart = "������ 5" Then j = 43
            If Direction = "�������" And Chart = "������ 6" Then j = 45
            If Direction = "�������" And Chart = "������ 7" Then j = 47
            If Direction = "�������" And Chart = "������ 8" Then j = 49
            If Direction = "�������" And Chart = "������ 9" Then j = 51
            If Direction = "�������" And Chart = "������ 10" Then j = 53
                Tpodacha = Sheets("temperature").Cells(i, j)
End Function

'����������� � �������� ������������
Function Tobratka(Typ As String, Chart As String, Regime As String, Direction As String) As Double
Dim i As Byte
Dim j As Byte
Dim j1 As Byte
Dim k As Byte
    If Typ = "���������" And Regime = "��" Then i = 26
    If Typ = "���������" And Regime = "���" Then i = 25
    If Typ = "���������" And Regime = "����" Then i = 27
    If Typ = "��������� ���������" And Regime = "��" Then i = 29
    If Typ = "��������� ���������" And Regime = "���" Then i = 28
    If Typ = "��������� ���������" And Regime = "����" Then i = 30
    If Typ = "��������� ������������" And Regime = "��" Then i = 29
    If Typ = "��������� ������������" And Regime = "���" Then i = 28
    If Typ = "��������� ������������" And Regime = "����" Then i = 30
    If Typ = "���������" And Regime = "��" Then i = 32
    If Typ = "���������" And Regime = "���" Then i = 31
    If Typ = "���������" And Regime = "����" Then i = 33
    If Typ = "�������" And Regime = "��" Then i = 35
    If Typ = "�������" And Regime = "���" Then i = 34
    If Typ = "�������" And Regime = "����" Then i = 36
        If Direction = "������" And Chart = "������ 1" Then j = 34
        If Direction = "������" And Chart = "������ 2" Then j = 36
        If Direction = "������" And Chart = "������ 3" Then j = 38
        If Direction = "������" And Chart = "������ 4" Then j = 40
        If Direction = "������" And Chart = "������ 5" Then j = 42
        If Direction = "������" And Chart = "������ 6" Then j = 44
        If Direction = "������" And Chart = "������ 7" Then j = 46
        If Direction = "������" And Chart = "������ 8" Then j = 48
        If Direction = "������" And Chart = "������ 9" Then j = 50
        If Direction = "������" And Chart = "������ 10" Then j = 52
            If Direction = "�������" And Chart = "������ 1" Then j = 35
            If Direction = "�������" And Chart = "������ 2" Then j = 37
            If Direction = "�������" And Chart = "������ 3" Then j = 39
            If Direction = "�������" And Chart = "������ 4" Then j = 41
            If Direction = "�������" And Chart = "������ 5" Then j = 43
            If Direction = "�������" And Chart = "������ 6" Then j = 45
            If Direction = "�������" And Chart = "������ 7" Then j = 47
            If Direction = "�������" And Chart = "������ 8" Then j = 49
            If Direction = "�������" And Chart = "������ 9" Then j = 51
            If Direction = "�������" And Chart = "������ 10" Then j = 53
                If Direction = "������" Then j1 = 1
                If Direction = "�������" Then j1 = -1
                   k = j + j1
                   Tobratka = Sheets("temperature").Cells(i, k)
End Function

'����� ������
Function timework(Chart As String, Regime As String) As Integer
Dim j As Byte
    If Chart = "������ 1" And Regime = "���" Then j = 22
    If Chart = "������ 2" And Regime = "���" Then j = 23
    If Chart = "������ 3" And Regime = "���" Then j = 24
    If Chart = "������ 4" And Regime = "���" Then j = 25
    If Chart = "������ 5" And Regime = "���" Then j = 26
    If Chart = "������ 6" And Regime = "���" Then j = 27
    If Chart = "������ 7" And Regime = "���" Then j = 28
    If Chart = "������ 8" And Regime = "���" Then j = 29
    If Chart = "������ 9" And Regime = "���" Then j = 30
    If Chart = "������ 10" And Regime = "���" Then j = 31
        If Chart = "������ 1" And Regime = "��" Then j = 2
        If Chart = "������ 2" And Regime = "��" Then j = 4
        If Chart = "������ 3" And Regime = "��" Then j = 6
        If Chart = "������ 4" And Regime = "��" Then j = 8
        If Chart = "������ 5" And Regime = "��" Then j = 10
        If Chart = "������ 6" And Regime = "��" Then j = 12
        If Chart = "������ 7" And Regime = "��" Then j = 14
        If Chart = "������ 8" And Regime = "��" Then j = 16
        If Chart = "������ 9" And Regime = "��" Then j = 18
        If Chart = "������ 10" And Regime = "��" Then j = 20
            If Chart = "������ 1" And Regime = "����" Then j = 3
            If Chart = "������ 2" And Regime = "����" Then j = 5
            If Chart = "������ 3" And Regime = "����" Then j = 7
            If Chart = "������ 4" And Regime = "����" Then j = 9
            If Chart = "������ 5" And Regime = "����" Then j = 11
            If Chart = "������ 6" And Regime = "����" Then j = 13
            If Chart = "������ 7" And Regime = "����" Then j = 15
            If Chart = "������ 8" And Regime = "����" Then j = 17
            If Chart = "������ 9" And Regime = "����" Then j = 19
            If Chart = "������ 10" And Regime = "����" Then j = 21
                timework = Sheets("temperature").Cells(19, j)
End Function

'���������� �������
Function Dv(Dnar As Double) As Double
Dim bt As Double
    If Dnar = Sheets("chart").Cells(6, 25) Then bt = Sheets("chart").Cells(6, 29)
    If Dnar = Sheets("chart").Cells(7, 25) Then bt = Sheets("chart").Cells(7, 29)
    If Dnar = Sheets("chart").Cells(8, 25) Then bt = Sheets("chart").Cells(8, 29)
    If Dnar = Sheets("chart").Cells(9, 25) Then bt = Sheets("chart").Cells(9, 29)
    If Dnar = Sheets("chart").Cells(10, 25) Then bt = Sheets("chart").Cells(10, 29)
    If Dnar = Sheets("chart").Cells(11, 25) Then bt = Sheets("chart").Cells(11, 29)
    If Dnar = Sheets("chart").Cells(12, 25) Then bt = Sheets("chart").Cells(12, 29)
    If Dnar = Sheets("chart").Cells(13, 25) Then bt = Sheets("chart").Cells(13, 29)
    If Dnar = Sheets("chart").Cells(15, 25) Then bt = Sheets("chart").Cells(15, 29)
    If Dnar = Sheets("chart").Cells(17, 25) Then bt = Sheets("chart").Cells(17, 29)
    If Dnar = Sheets("chart").Cells(19, 25) Then bt = Sheets("chart").Cells(19, 29)
    If Dnar = Sheets("chart").Cells(20, 25) Then bt = Sheets("chart").Cells(20, 29)
    If Dnar = Sheets("chart").Cells(22, 25) Then bt = Sheets("chart").Cells(22, 29)
    If Dnar = Sheets("chart").Cells(23, 25) Then bt = Sheets("chart").Cells(23, 29)
    If Dnar = Sheets("chart").Cells(25, 25) Then bt = Sheets("chart").Cells(25, 29)
    If Dnar = Sheets("chart").Cells(26, 25) Then bt = Sheets("chart").Cells(26, 29)
    If Dnar = Sheets("chart").Cells(28, 25) Then bt = Sheets("chart").Cells(28, 29)
    If Dnar = Sheets("chart").Cells(29, 25) Then bt = Sheets("chart").Cells(29, 29)
    If Dnar = Sheets("chart").Cells(33, 25) Then bt = Sheets("chart").Cells(33, 29)
    If Dnar = Sheets("chart").Cells(34, 25) Then bt = Sheets("chart").Cells(34, 29)
    If Dnar = Sheets("chart").Cells(38, 25) Then bt = Sheets("chart").Cells(38, 29)
    If Dnar = Sheets("chart").Cells(44, 25) Then bt = Sheets("chart").Cells(44, 29)
    If Dnar = Sheets("chart").Cells(50, 25) Then bt = Sheets("chart").Cells(50, 29)
    If Dnar = Sheets("chart").Cells(56, 25) Then bt = Sheets("chart").Cells(56, 29)
    If Dnar = Sheets("chart").Cells(62, 25) Then bt = Sheets("chart").Cells(62, 29)
    If Dnar = Sheets("chart").Cells(68, 25) Then bt = Sheets("chart").Cells(68, 29)
    If Dnar = Sheets("chart").Cells(69, 25) Then bt = Sheets("chart").Cells(69, 29)
    If Dnar = Sheets("chart").Cells(71, 25) Then bt = Sheets("chart").Cells(71, 29)
        Dv = Dnar - bt * 2
End Function

'��������� ����������� �� ���� �������� ������ ����������� ���������� ����� �� ����� ���������
Function tokr(Month As String, xgr As Double, xvoz As Double) As Double
Dim i As Byte
Dim j As Byte
Dim tgr As Double
Dim tvoz As Double
    If Month = "������" Then i = 7
    If Month = "�������" Then i = 8
    If Month = "����" Then i = 9
    If Month = "������" Then i = 10
    If Month = "���" Then i = 11
    If Month = "����" Then i = 12
    If Month = "����" Then i = 13
    If Month = "������" Then i = 14
    If Month = "��������" Then i = 15
    If Month = "�������" Then i = 16
    If Month = "������" Then i = 17
    If Month = "�������" Then i = 18
        If Month = "������" Then j = 7
        If Month = "�������" Then j = 8
        If Month = "����" Then j = 9
        If Month = "������" Then j = 10
        If Month = "���" Then j = 11
        If Month = "����" Then j = 12
        If Month = "����" Then j = 13
        If Month = "������" Then j = 14
        If Month = "��������" Then j = 15
        If Month = "�������" Then j = 16
        If Month = "������" Then j = 17
        If Month = "�������" Then j = 18
            tgr = Sheets("temperature").Cells(i, 32)
            tvoz = Sheets("temperature").Cells(j, 33)
            tokr = (tgr * xgr + tvoz * xvoz) / (xgr + xvoz)
End Function

'������������� ����������� ���������� �����
Function tsrokr(Chart As String, Regime As String, xgr As Double, xvoz As Double) As Double
Dim i As Byte
Dim j As Byte
Dim tgr As Double
Dim tvoz As Double
    If Chart = "������ 1" Then j = 34
    If Chart = "������ 2" Then j = 36
    If Chart = "������ 3" Then j = 38
    If Chart = "������ 4" Then j = 40
    If Chart = "������ 5" Then j = 42
    If Chart = "������ 6" Then j = 44
    If Chart = "������ 7" Then j = 46
    If Chart = "������ 8" Then j = 48
    If Chart = "������ 9" Then j = 50
    If Chart = "������ 10" Then j = 52
    If Regime = "���" Then i = 22
    If Regime = "��" Then i = 23
    If Regime = "����" Then i = 24
        If Chart = "������ 1" Then l = 35
        If Chart = "������ 2" Then l = 37
        If Chart = "������ 3" Then l = 39
        If Chart = "������ 4" Then l = 41
        If Chart = "������ 5" Then l = 43
        If Chart = "������ 6" Then l = 45
        If Chart = "������ 7" Then l = 47
        If Chart = "������ 8" Then l = 49
        If Chart = "������ 9" Then l = 51
        If Chart = "������ 10" Then l = 53
        If Regime = "���" Then k = 22
        If Regime = "��" Then k = 23
        If Regime = "����" Then k = 24
            tgr = Sheets("temperature").Cells(i, j)
            tvoz = Sheets("temperature").Cells(k, l)
            tsrokr = (tgr * xgr + tvoz * xvoz) / (xgr + xvoz)
End Function

'����������� ���� � ������ ������� ��� ���������
Function tpisp(Chart As String, Regime As String, tisp As Double, tokr As Double, tsrokr As Double) As Double
Dim i As Byte
Dim j As Byte
Dim l As Byte
Dim t1 As Double
Dim t2 As Double
    If Chart = "������ 1" Then j = 34
    If Chart = "������ 2" Then j = 36
    If Chart = "������ 3" Then j = 38
    If Chart = "������ 4" Then j = 40
    If Chart = "������ 5" Then j = 42
    If Chart = "������ 6" Then j = 44
    If Chart = "������ 7" Then j = 46
    If Chart = "������ 8" Then j = 48
    If Chart = "������ 9" Then j = 50
    If Chart = "������ 10" Then j = 52
    If Regime = "���" Then i = 19
    If Regime = "��" Then i = 20
    If Regime = "����" Then i = 21
        If Chart = "������ 1" Then l = 35
        If Chart = "������ 2" Then l = 37
        If Chart = "������ 3" Then l = 39
        If Chart = "������ 4" Then l = 41
        If Chart = "������ 5" Then l = 43
        If Chart = "������ 6" Then l = 45
        If Chart = "������ 7" Then l = 47
        If Chart = "������ 8" Then l = 49
        If Chart = "������ 9" Then l = 51
        If Chart = "������ 10" Then l = 53
        If Regime = "���" Then k = 19
        If Regime = "��" Then k = 20
        If Regime = "����" Then k = 21
            t1 = Sheets("temperature").Cells(i, j)
            t2 = Sheets("temperature").Cells(k, l)
            tpisp = (t1 + t2) / 2 + tisp / 2 + tokr - tsrokr
End Function

'����������������� ������� ������
Function tprob(t1 As Double, t2 As Double, V As Double, G As Double) As Double
    tprob = V * wskDSWT((t1 + t2) / 2) / 1000 / G
End Function
'    Sub fghdg()
'    f = qni(2014, "������", "���������", 1000, 1)
'    End Sub
'������ qn ��� ������� ���������
Function qni(Year As Integer, Direction As String, Typ As String, x As Double, ring As Integer) As Double
Dim i As Byte
Dim j As Byte
Dim l As Byte
Dim TR As Object
Dim TR1 As Object
Dim TR3 As Object
Dim TR4 As Object
Dim TR5 As Object
Dim TR6 As Object
Dim TR7 As Object
Dim TR8 As Object
Dim TR9 As Object
Dim qn1 As Double
Dim qn2 As Double
Dim qn3 As Double
Dim qn4 As Double
Dim qn5 As Double
Dim qn6 As Double
Dim qn7 As Double
Dim qn8 As Double
Dim qn9 As Double
Dim Period As Integer
Dim Regime As String
Dim Chart As String
    Regime = Sheets("isptemp").Cells(ring + 3, 11)
    Chart = Sheets("isptemp").Cells(ring + 3, 10)
    Period = periodWork(Chart, Regime)
'����������� ��������� ����������� ��� ������� ���������
        tpod = Sheets("isptemp").Cells(ring + 3, 3)
        tobr = Sheets("isptemp").Cells(ring + 3, 4)
            If Sheets("isptemp").Cells(ring + 3, 12) = "������" Then i = 7
            If Sheets("isptemp").Cells(ring + 3, 12) = "�������" Then i = 8
            If Sheets("isptemp").Cells(ring + 3, 12) = "����" Then i = 9
            If Sheets("isptemp").Cells(ring + 3, 12) = "������" Then i = 10
            If Sheets("isptemp").Cells(ring + 3, 12) = "���" Then i = 11
            If Sheets("isptemp").Cells(ring + 3, 12) = "����" Then i = 12
            If Sheets("isptemp").Cells(ring + 3, 12) = "����" Then i = 13
            If Sheets("isptemp").Cells(ring + 3, 12) = "������" Then i = 14
            If Sheets("isptemp").Cells(ring + 3, 12) = "��������" Then i = 15
            If Sheets("isptemp").Cells(ring + 3, 12) = "�������" Then i = 16
            If Sheets("isptemp").Cells(ring + 3, 12) = "������" Then i = 17
            If Sheets("isptemp").Cells(ring + 3, 12) = "�������" Then i = 18
                tgr = Sheets("temperature").Cells(i, 32) '����������� ������
                tvoz = Sheets("temperature").Cells(i, 33) '����������� �������
                    If Typ = "��������� ���������" And Direction = "������" Then n = tpod - tgr
                    If Typ = "��������� ���������" And Direction = "�������" Then n = (tpod + tobr) / 2 - tgr
                    If Typ = "��������� ������������" And Direction = "������" Then n = tpod - tgr
                    If Typ = "��������� ������������" And Direction = "�������" Then n = (tpod + tobr) / 2 - tgr
                    If Typ = "���������" And Direction = "������" Then n = tpod - tvoz
                    If Typ = "���������" And Direction = "�������" Then n = tobr - tvoz
                    If Typ = "���������" And Direction = "������" Then n = tpod - Sheets("temperature").Cells(22, 32)
                    If Typ = "���������" And Direction = "�������" Then n = tobr - Sheets("temperature").Cells(22, 32)
                    If Typ = "�������" And Direction = "������" Then n = tpod - Sheets("temperature").Cells(23, 32)
                    If Typ = "�������" And Direction = "�������" Then n = tobr - Sheets("temperature").Cells(23, 32)
                        If Typ = "��������� ���������" And Direction = "������" Then nk = (tpod + tobr) / 2 - tgr
                        If Typ = "��������� ���������" And Direction = "�������" Then nk = tpod - tgr
                        If Typ = "��������� ������������" And Direction = "������" Then nk = (tpod + tobr) / 2 - tgr
                        If Typ = "��������� ������������" And Direction = "�������" Then nk = tpod - tgr
'��������� ��������� � ������������ ������� ����� ��� ������ � ������� ���������
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Direction = "�������" And Year > 2003 And Typ = "��������� ���������" _
Or Direction = "�������" And Year > 2003 And Typ = "��������� ������������" Then
    qn1 = 0
    Else: qn1 = interpolationNorms(x, n, rangeNorms(Year, Direction, Typ, Period))
End If
'��������� ��������� � ������������ ������� ����� ��� ������
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Direction = "�������" And Year > 2003 And Typ = "��������� ���������" _
Or Direction = "�������" And Year > 2003 And Typ = "��������� ������������" Then
    qn2 = 0
    Else: qn2 = interpolationNorms(x, nk, rangeNorms1(Year, Direction, Typ, Period))
End If
'���������, ���������, �������
If Direction = "�������" And Typ = "��������� ���������" Or Direction = "�������" And Typ = "��������� ������������" Or Typ = "��������� ���������" _
And Direction = "������" And Year >= 2004 Or Typ = "��������� ������������" And Direction = "������" And Year >= 2004 Then
    qn3 = 0
    Else: qn3 = interpolationNorms(x, n, rangeNorms3(Year, Direction, Typ, Period))
End If
'��� ������
'��������� ��������� � ������������ ����� ��� ������ � ������� ��������� (2004)
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn4 = 0
    Else: qn4 = interpolationNorms(x, nk, rangeNorms4(Year, Direction, Typ, Period))
End If
'��������� ��������� � ������������ ������� ����� ��� ������ (2004)
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn5 = 0
    Else: qn5 = interpolationNorms(x, n, rangeNorms5(Year, Direction, Typ, Period))
End If
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    K2004p = 0
    Else: K2004p = qn5 / qn4
End If
'��� �������
'��������� ��������� � ������������ ����� ��� ������ � ������� ��������� (2004)
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn8 = 0
    Else: qn8 = interpolationNorms(x, n, rangeNorms4(Year, Direction, Typ, Period))
End If
'��������� ��������� � ������������ ������� ����� ��� ������ (2004)
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn9 = 0
    Else: qn9 = interpolationNorms(x, nk, rangeNorms5(Year, Direction, Typ, Period))
End If
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    K2004o = 0
    Else: K2004o = 1 - qn9 / qn8
End If
'��� ������ 2004
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    qn6 = 0
    Else: qn6 = interpolationNorms(x, nk, rangeNorms6(Year, Direction, Typ, Period))
End If
'��� ������� 2004
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    qn7 = 0
    Else: qn7 = interpolationNorms(x, n, rangeNorms7(Year, Direction, Typ, Period))
End If
    qni = qn1 - qn2 + qn3 + K2004p * qn6 + K2004o * qn7
End Function
'    Sub dfgndghmn()
'    f = qnn(2013, "������", "��������� ������������", 1, 70, 54.07, 52.27)
'    End Sub
'������ qn ��� ������������� ��������
Function qnn(Year As Integer, Direction As String, Typ As String, ring As Integer, x, tpnx As Double, tonx As Double) 'y As Double, yk As Double Period As Integer, tonx As Double
Dim TR As Object
Dim TR1 As Object
Dim TR3 As Object
Dim qn11 As Double
Dim qn12 As Double
Dim qn13 As Double
Dim Period As Integer
Dim Regime As String
Dim Chart As String
    If Typ = "��������� ���������" Or Typ = "��������� ������������" Then Z = Sheets("isptemp").Cells(ring + 3, 14) 'Z - ����������� ������
    If Typ = "���������" Or Typ = "���������" Or Typ = "�������" Then Z = Sheets("isptemp").Cells(ring + 3, 13) 'Z - ����������� ���������� �����
        If Typ = "��������� ���������" And Direction = "������" Then m = tpnx - Z 'Nexar_t(ring, characterSection, rangeConditions, Direction)
        If Typ = "��������� ���������" And Direction = "������" Then mk = (tpnx + tonx) / 2 - Z
        If Typ = "��������� ������������" And Direction = "������" Then m = tpnx - Z 'Nexar_t(ring, characterSection, rangeConditions, Direction)
        If Typ = "��������� ������������" And Direction = "������" Then mk = (tpnx + tonx) / 2 - Z
            If Typ = "��������� ���������" And Direction = "�������" Then m = (tpnx + tonx) / 2 - Z
            If Typ = "��������� ���������" And Direction = "�������" Then mk = tpnx - Z
            If Typ = "��������� ������������" And Direction = "�������" Then m = (tpnx + tonx) / 2 - Z
            If Typ = "��������� ������������" And Direction = "�������" Then mk = tpnx - Z
                If Typ = "���������" And Direction = "������" Then m = tpnx - Z
                If Typ = "���������" And Direction = "������" Then mk = tonx - Z
                If Typ = "���������" And Direction = "������" Then m = tpnx - Z
                If Typ = "���������" And Direction = "������" Then mk = tonx - Z
                If Typ = "�������" And Direction = "������" Then m = tpnx - Z
                If Typ = "�������" And Direction = "������" Then mk = tonx - Z
                    If Typ = "���������" And Direction = "�������" Then m = tonx - Z
                    If Typ = "���������" And Direction = "�������" Then mk = tpnx - Z
                    If Typ = "���������" And Direction = "�������" Then m = tonx - Z
                    If Typ = "���������" And Direction = "�������" Then mk = tpnx - Z
                    If Typ = "�������" And Direction = "�������" Then m = tonx - Z
                    If Typ = "�������" And Direction = "�������" Then mk = tpnx - Z
                        Regime = Sheets("isptemp").Cells(ring + 3, 11)
                        Chart = Sheets("isptemp").Cells(ring + 3, 10)
                        Period = periodWork(Chart, Regime)
'��������� ��������� � ������������ ������� ����� ��� ������ � ������� ���������
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Direction = "�������" And Year > 2003 And Typ = "��������� ���������" _
Or Direction = "�������" And Year > 2003 And Typ = "��������� ������������" Then
    qn1 = 0
    Else: qn1 = interpolationNorms(x, m, rangeNorms(Year, Direction, Typ, Period))
End If
'��������� ��������� � ������������ ������� ����� ��� ������
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Direction = "�������" And Year > 2003 And Typ = "��������� ���������" _
Or Direction = "�������" And Year > 2003 And Typ = "��������� ������������" Then
    qn2 = 0
    Else: qn2 = interpolationNorms(x, mk, rangeNorms1(Year, Direction, Typ, Period))
End If
'���������, ���������, �������
If Direction = "�������" And Typ = "��������� ���������" Or Direction = "�������" And Typ = "��������� ������������" Or Typ = "��������� ���������" _
And Direction = "������" And Year >= 2004 Or Typ = "��������� ������������" And Direction = "������" And Year >= 2004 Then
    qn3 = 0
    Else: qn3 = interpolationNorms(x, m, rangeNorms3(Year, Direction, Typ, Period))
End If
'��� ������
'��������� ��������� � ������������ ����� ��� ������ � ������� ��������� (2004)
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn4 = 0
    Else: qn4 = interpolationNorms(x, mk, rangeNorms4(Year, Direction, Typ, Period))
End If
'��������� ��������� � ������������ ������� ����� ��� ������ (2004)
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn5 = 0
    Else: qn5 = interpolationNorms(x, m, rangeNorms5(Year, Direction, Typ, Period))
End If
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    K2004p = 0
    Else: K2004p = qn5 / qn4
End If
'��� �������
'��������� ��������� � ������������ ����� ��� ������ � ������� ��������� (2004)
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn8 = 0
    Else: qn8 = interpolationNorms(x, m, rangeNorms4(Year, Direction, Typ, Period))
End If
'��������� ��������� � ������������ ������� ����� ��� ������ (2004)
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn9 = 0
    Else: qn9 = interpolationNorms(x, mk, rangeNorms5(Year, Direction, Typ, Period))
End If
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    K2004o = 0
    Else: K2004o = 1 - qn9 / qn8
End If
'��� ������
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    qn6 = 0
    Else: qn6 = interpolationNorms(x, mk, rangeNorms6(Year, Direction, Typ, Period))
End If
'��� �������
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    qn7 = 0
    Else: qn7 = interpolationNorms(x, m, rangeNorms7(Year, Direction, Typ, Period))
End If
    qnn = qn1 - qn2 + qn3 + K2004p * qn6 + K2004o * qn7
End Function

'������ qn ��� ������������� ��������
Function qnsr(Year As Integer, Direction As String, Typ, x, ring) 'y As Double, yk As Double Period As Integer
Dim TR As Object
Dim TR1 As Object
Dim TR3 As Object
Dim qn1 As Double
Dim qn2 As Double
Dim qn3 As Double
Dim Period As Integer
Dim Regime As String
Dim Chart As String
    Regime = Sheets("isptemp").Cells(ring + 3, 11)
    Chart = Sheets("isptemp").Cells(ring + 3, 10)
    Period = periodWork(Chart, Regime)
'��������� ��������� � ������������ ������� ����� ��� ������ � ������� ���������
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Direction = "�������" And Year > 2003 And Typ = "��������� ���������" _
Or Direction = "�������" And Year > 2003 And Typ = "��������� ������������" Then
    qn1 = 0
    Else: qn1 = interpolationNorms(x, flowTemperature(Typ, Chart, Regime, Direction), rangeNorms(Year, Direction, Typ, Period))
End If
'��������� ��������� � ������������ ������� ����� ��� ������
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Direction = "�������" And Year > 2003 And Typ = "��������� ���������" _
Or Direction = "�������" And Year > 2003 And Typ = "��������� ������������" Then
    qn2 = 0
    Else: qn2 = interpolationNorms(x, returnTemperature(Typ, Chart, Regime, Direction), rangeNorms1(Year, Direction, Typ, Period))
End If
'���������, ���������, �������
If Direction = "�������" And Typ = "��������� ���������" Or Direction = "�������" And Typ = "��������� ������������" Or Typ = "��������� ���������" _
And Direction = "������" And Year >= 2004 Or Typ = "��������� ������������" And Direction = "������" And Year >= 2004 Then
    qn3 = 0
    Else: qn3 = interpolationNorms(x, flowTemperature(Typ, Chart, Regime, Direction), rangeNorms3(Year, Direction, Typ, Period))
End If
'��� ������
'��������� ��������� � ������������ ����� ��� ������ � ������� ��������� (2004)
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn4 = 0
    Else: qn4 = interpolationNorms(x, returnTemperature(Typ, Chart, Regime, Direction), rangeNorms4(Year, Direction, Typ, Period))
End If
'��������� ��������� � ������������ ������� ����� ��� ������ (2004)
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn5 = 0
    Else: qn5 = interpolationNorms(x, flowTemperature(Typ, Chart, Regime, Direction), rangeNorms5(Year, Direction, Typ, Period))
End If
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    K2004p = 0
    Else: K2004p = qn5 / qn4
End If
'��� �������
'��������� ��������� � ������������ ����� ��� ������ � ������� ��������� (2004)
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn8 = 0
    Else: qn8 = interpolationNorms(x, flowTemperature(Typ, Chart, Regime, Direction), rangeNorms4(Year, Direction, Typ, Period))
End If
'��������� ��������� � ������������ ������� ����� ��� ������ (2004)
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 Then
    qn9 = 0
    Else: qn9 = interpolationNorms(x, returnTemperature(Typ, Chart, Regime, Direction), rangeNorms5(Year, Direction, Typ, Period))
End If
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    K2004o = 0
    Else: K2004o = 1 - qn9 / qn8
End If
'��� ������
If Direction = "�������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    qn6 = 0
    Else: qn6 = interpolationNorms(x, returnTemperature(Typ, Chart, Regime, Direction), rangeNorms6(Year, Direction, Typ, Period))
End If
'��� �������
If Direction = "������" Or Typ = "���������" Or Typ = "���������" Or Typ = "�������" Or Year <= 2003 And Typ = "��������� ���������" Or Year <= 2003 And Typ = "��������� ������������" Then
    qn7 = 0
    Else: qn7 = interpolationNorms(x, flowTemperature(Typ, Chart, Regime, Direction), rangeNorms7(Year, Direction, Typ, Period))
End If
    qnsr = qn1 - qn2 + qn3 + K2004p * qn6 + K2004o * qn7
End Function
'    Sub hui()
'    a = Qpodz(25995, 25985, 76.59, 74.59, 74.59, 72.59, 2, 2, "��������� ���������")
'    End Sub
'����� ����������� �������� ������, ����������� � ������������� �������� ��� ���������� ������������
'(�� ������� ������ �� ������������� ��������)
Function Qpodz(Qp1 As Double, Qo1 As Double, t1 As Double, t2 As Double, t3 As Double, t4 As Double, ring As Integer, site As Integer, Typ As String) As Double
Dim list0 As Object
Dim list1 As Object
Dim list2 As Object
Dim list3 As Object
Dim Regime As String
Dim Chart As String
    Regime = Sheets("isptemp").Cells(ring + 3, 11)
    Chart = Sheets("isptemp").Cells(ring + 3, 10)
        If Regime = "" Then
            MsgBox "�� ������ ����� �� ������� isptemp ��� ��������������� ������ � " & ring
            Exit Function
        ElseIf Chart = "" Then
            MsgBox "�� ������ ������������� ������ �� ������� isptemp ��� ��������������� ������ � " & ring
            Exit Function
        End If
            Set list0 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 21), Sheets("table_2").Cells(15000, 21))
            Set list1 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 1), Sheets("table_2").Cells(15000, 1))
            Set list2 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 2), Sheets("table_2").Cells(15000, 2))
            Set list3 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 4), Sheets("table_2").Cells(15000, 4))
'����������� ��������� ����������� ��� ������� ���������
If Chart = "������ 1" Then j = 34
If Chart = "������ 2" Then j = 36
If Chart = "������ 3" Then j = 38
If Chart = "������ 4" Then j = 40
If Chart = "������ 5" Then j = 42
If Chart = "������ 6" Then j = 44
If Chart = "������ 7" Then j = 46
If Chart = "������ 8" Then j = 48
If Chart = "������ 9" Then j = 50
If Chart = "������ 10" Then j = 52
    If Chart = "������ 1" Then k = 35
    If Chart = "������ 2" Then k = 37
    If Chart = "������ 3" Then k = 39
    If Chart = "������ 4" Then k = 41
    If Chart = "������ 5" Then k = 43
    If Chart = "������ 6" Then k = 45
    If Chart = "������ 7" Then k = 47
    If Chart = "������ 8" Then k = 49
    If Chart = "������ 9" Then k = 51
    If Chart = "������ 10" Then k = 53
        If Regime = "���" Then i = 19
        If Regime = "��" Then i = 20
        If Regime = "����" Then i = 21
        If Regime = "���" Then l = 22
        If Regime = "��" Then l = 23
        If Regime = "����" Then l = 24
            tpod = Sheets("temperature").Cells(i, j)
            tobr = Sheets("temperature").Cells(i, k)
            tgr = Sheets("temperature").Cells(l, j) '����������� ������
            tvoz = Sheets("temperature").Cells(l, k) '����������� �������
            Qp2 = Application.SumIfs(list0, list1, ring, list2, site, list3, "������")
            Qp = Qp1 - Qp2
            Qo2 = Application.SumIfs(list0, list1, ring, list2, site, list3, "�������")
            Qo = Qo1 - Qo2
            Qpodz = (Qp * (tpod - tgr) + Qo * (tobr - tgr)) / (0.25 * (t1 + t2 + t3 + t4) - tokrsre(ring, Typ))
End Function

'����� ����������� �������� ������, ����������� � ������������� �������� ��� ���������� ��������� ������������
'(�� ������� ������ �� ������������� ��������)
Function Qnadpod(Qp1, t1, t2, ring, site, Typ)
Dim list0 As Object
Dim list1 As Object
Dim list2 As Object
Dim list3 As Object
Dim Regime As String
Dim Chart As String
    Set list0 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 21), Sheets("table_2").Cells(25000, 21))
    Set list1 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 1), Sheets("table_2").Cells(25000, 1))
    Set list2 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 2), Sheets("table_2").Cells(25000, 2))
    Set list3 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 4), Sheets("table_2").Cells(25000, 4))
        Regime = Sheets("isptemp").Cells(ring + 3, 11)
        Chart = Sheets("isptemp").Cells(ring + 3, 10)
            If Regime = "" Then
                MsgBox "�� ������ ����� �� ������� isptemp ��� ��������������� ������ � " & ring
                Exit Function
            ElseIf Chart = "" Then
                MsgBox "�� ������ ������������� ������ �� ������� isptemp ��� ��������������� ������ � " & ring
                Exit Function
            End If
'����������� ��������� ����������� ��� ������� ���������
If Chart = "������ 1" Then j = 34
If Chart = "������ 2" Then j = 36
If Chart = "������ 3" Then j = 38
If Chart = "������ 4" Then j = 40
If Chart = "������ 5" Then j = 42
If Chart = "������ 6" Then j = 44
If Chart = "������ 7" Then j = 46
If Chart = "������ 8" Then j = 48
If Chart = "������ 9" Then j = 50
If Chart = "������ 10" Then j = 52
    If Chart = "������ 1" Then k = 35
    If Chart = "������ 2" Then k = 37
    If Chart = "������ 3" Then k = 39
    If Chart = "������ 4" Then k = 41
    If Chart = "������ 5" Then k = 43
    If Chart = "������ 6" Then k = 45
    If Chart = "������ 7" Then k = 47
    If Chart = "������ 8" Then k = 49
    If Chart = "������ 9" Then k = 51
    If Chart = "������ 10" Then k = 53
        If Regime = "���" Then i = 19
        If Regime = "��" Then i = 20
        If Regime = "����" Then i = 21
        If Regime = "���" Then l = 22
        If Regime = "��" Then l = 23
        If Regime = "����" Then l = 24
            tpod = Sheets("temperature").Cells(i, j)
            tobr = Sheets("temperature").Cells(i, k)
            tgr = Sheets("temperature").Cells(l, j) '����������� ������
            tvoz = Sheets("temperature").Cells(l, k) '����������� �������
            Qp2 = Application.SumIfs(list0, list1, ring, list2, site, list3, "������")
            Qp = Qp1 - Qp2
            Qnadpod = (Qp * (tpod - tvoz)) / (0.5 * (t1 + t2) - tokrsre(ring, Typ))
End Function

'����� ����������� �������� ������, ����������� � ������������� �������� ��� ���������� ��������� ������������
'(�� ������� ������ �� ������������� ��������)
Function Qnadobr(Qo1 As Double, t1 As Double, t2 As Double, ring As Integer, site As Integer, Typ As String) As Double
Dim list0 As Object
Dim list1 As Object
Dim list2 As Object
Dim list3 As Object
Dim Regime As String
Dim Chart As String
    Set list0 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 21), Sheets("table_2").Cells(25000, 21))
    Set list1 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 1), Sheets("table_2").Cells(25000, 1))
    Set list2 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 2), Sheets("table_2").Cells(25000, 2))
    Set list3 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 4), Sheets("table_2").Cells(25000, 4))
        Regime = Sheets("isptemp").Cells(ring + 3, 11)
        Chart = Sheets("isptemp").Cells(ring + 3, 10)
'����������� ��������� ����������� ��� ������� ���������
If Chart = "������ 1" Then j = 34
If Chart = "������ 2" Then j = 36
If Chart = "������ 3" Then j = 38
If Chart = "������ 4" Then j = 40
If Chart = "������ 5" Then j = 42
If Chart = "������ 6" Then j = 44
If Chart = "������ 7" Then j = 46
If Chart = "������ 8" Then j = 48
If Chart = "������ 9" Then j = 50
If Chart = "������ 10" Then j = 52
    If Chart = "������ 1" Then k = 35
    If Chart = "������ 2" Then k = 37
    If Chart = "������ 3" Then k = 39
    If Chart = "������ 4" Then k = 41
    If Chart = "������ 5" Then k = 43
    If Chart = "������ 6" Then k = 45
    If Chart = "������ 7" Then k = 47
    If Chart = "������ 8" Then k = 49
    If Chart = "������ 9" Then k = 51
    If Chart = "������ 10" Then k = 53
        If Regime = "���" Then i = 19
        If Regime = "��" Then i = 20
        If Regime = "����" Then i = 21
        If Regime = "���" Then l = 22
        If Regime = "��" Then l = 23
        If Regime = "����" Then l = 24
            tpod = Sheets("temperature").Cells(i, j)
            tobr = Sheets("temperature").Cells(i, k)
            tgr = Sheets("temperature").Cells(l, j) '����������� ������
            tvoz = Sheets("temperature").Cells(l, k) '����������� �������
            Qo2 = Application.SumIfs(list0, list1, ring, list2, site, list3, "�������")
            Qo = Qo1 - Qo2
            Qnadobr = (Qo * (tobr - tvoz)) / (0.5 * (t1 + t2) - tokrsre(ring, Typ))
End Function

'����������� ���������� ����� ��� ���������
Function tokrsre(ring, Typ)
    If Typ = "��������� ���������" Then tokrsre = Sheets("isptemp").Cells(ring + 3, 14)
    If Typ = "��������� ������������" Then tokrsre = Sheets("isptemp").Cells(ring + 3, 14)
    If Typ = "���������" Then tokrsre = Sheets("isptemp").Cells(ring + 3, 13)
End Function

'��������� ����������� ��� qnn
Function ty(Typ As String, tpnx As Double, tonx As Double, Direction As String, ring As Integer) As Double
    If Typ = "��������� ���������" And Direction = "������" Then ty = tpnx - tokrsre(ring, Typ)
    If Typ = "��������� ���������" And Direction = "�������" Then ty = (tpnx + tonx) / 2 - tokrsre(ring, Typ)
    If Typ = "���������" And Direction = "������" Then ty = tpnx - tokrsre(ring, Typ)
    If Typ = "���������" And Direction = "�������" Then ty = tonx - tokrsre(ring, Typ)
End Function

'��������� ����������� ��� qnn
Function tyk(Typ As String, tpnx As Double, tonx As Double, Direction As String, ring As Integer) As Double
    If Typ = "��������� ���������" And Direction = "������" Then tyk = (tpnx + tonx) / 2 - tokrsre(ring, Typ)
    If Typ = "��������� ���������" And Direction = "�������" Then tyk = tpnx - tokrsre(ring, Typ)
    If Typ = "���������" And Direction = "������" Then tyk = tonx - tokrsre(ring, Typ)
    If Typ = "���������" And Direction = "�������" Then tyk = tpnx - tokrsre(ring, Typ)
End Function
'    Sub huit()
'    'Dim b As Range
'    a = Nexar_t(1, 1, "�����.", "������")
'    End Sub
'����������� ������������� �� ������������� ������� �� ��� ���������
Function Nexar_t(ring As Integer, site, characterSectionStraight, characterSectionReturn, Direction As String)  'characterSectionrangeConditions As Range,
Dim cellNumber As Long
Dim list0 As Object
Dim list1 As Object
Dim list2 As Object
Dim list3 As Object
Dim list10 As Object
Dim list11 As Object
Dim list12 As Object
Dim list13 As Object
Dim list20 As Object
Dim list21 As Object
Dim list22 As Object
Dim list23 As Object
Dim list24 As Object
    cellNumber = characterSectionStraight.Row 'lineSequenceNumber = cellNumber.Row
Set list20 = Sheets("table_4").Range(Sheets("table_4").Cells(5, 10), Sheets("table_4").Cells(25000, 10))
Set list21 = Sheets("table_4").Range(Sheets("table_4").Cells(5, 1), Sheets("table_4").Cells(25000, 1))
Set list22 = Sheets("table_4").Range(Sheets("table_4").Cells(5, 2), Sheets("table_4").Cells(25000, 2))
Set list23 = Sheets("table_4").Range(Sheets("table_4").Cells(5, 4), Sheets("table_4").Cells(25000, 4))
Set list24 = Sheets("table_4").Range(Sheets("table_4").Cells(5, 11), Sheets("table_4").Cells(25000, 11))
    Set list0 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 11), Sheets("table_2").Cells(25000, 11))
    Set list1 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 1), Sheets("table_2").Cells(25000, 1))
    Set list2 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 2), Sheets("table_2").Cells(25000, 2))
    Set list3 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 4), Sheets("table_2").Cells(25000, 4))
        Set list10 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 11), Sheets("table_2").Cells(cellNumber - 1, 11))
        Set list11 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 1), Sheets("table_2").Cells(cellNumber - 1, 1))
        Set list12 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 2), Sheets("table_2").Cells(cellNumber - 1, 2))
        Set list13 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 4), Sheets("table_2").Cells(cellNumber - 1, 4))
            If characterSectionStraight = "�����." Or characterSectionReturn = "�����." Then 'Or characterSectionReturn = "�����."
                Mych = Application.SumIfs(list0, list1, ring, list2, site, list3, Direction)
                    If cellNumber = 5 Then
                        Mnexych = 0
                        ElseIf cellNumber = 6 Then
                            Mnexych = 0
                            Else: Mnexych = Application.SumIfs(list10, list11, ring, list12, site, list13, Direction)
                    End If
                tn1 = Application.SumIfs(list20, list21, ring, list22, site, list23, Direction)
                tk1 = Application.SumIfs(list24, list21, ring, list22, site, list23, Direction)
            End If
            If Direction = "������" Then
                If cellNumber = 5 Then
                    s = Sheets("table_2").Cells(5, 11)
                    Else: s = Sheets("table_2").Cells(cellNumber, 11)
                End If
                Nexar_t = tn1 - ((tn1 - tk1) * ((Mnexych + 0.5 * s) / Mych))
            ElseIf Direction = "�������" Then
                If cellNumber = 6 Then
                    s = Sheets("table_2").Cells(6, 11)
                    Else: s = Sheets("table_2").Cells(cellNumber, 11)
                End If
                Nexar_t = tk1 + ((tn1 - tk1) * ((Mnexych + 0.5 * s) / Mych))
            End If
End Function

'����� ����������� �������� ������ ���������� ������������ ��� ������������� ��������
Function SummaQnadzem(ring As Integer, site As Integer, Direction As String) As Double
Dim list0 As Object
Dim list1 As Object
Dim list2 As Object
Dim list3 As Object
Dim list4 As Object
    Set list0 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 17), Sheets("table_2").Cells(25000, 17))
    Set list1 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 1), Sheets("table_2").Cells(25000, 1))
    Set list2 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 2), Sheets("table_2").Cells(25000, 2))
    Set list3 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 4), Sheets("table_2").Cells(25000, 4))
        SummaQnadzem = Application.SumIfs(list0, list1, ring, list2, site, list3, Direction)
End Function
'    Sub hui()
'    a = SummaQpodzem(1, 2)
'    End Sub

'����� ����������� �������� ������ ���������� ������������ ��� ������������� ��������
Function SummaQpodzem(ring As Integer, site As Integer) As Double
Dim list0 As Object
Dim list1 As Object
Dim list2 As Object
Dim list3 As Object
Dim list4 As Object
    Set list0 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 17), Sheets("table_2").Cells(25000, 17))
    Set list1 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 1), Sheets("table_2").Cells(25000, 1))
    Set list2 = Sheets("table_2").Range(Sheets("table_2").Cells(5, 2), Sheets("table_2").Cells(25000, 2))
        SummaQpodzem = Application.SumIfs(list0, list1, ring, list2, site)
End Function

' ����������� ��������� ��������
Function aq(G)
If G = 0 Then
aq = 0
Exit Function
End If

y1 = 1.2
y2 = 1.3
y3 = 1.35
y4 = 1.4
y5 = 1.4
y6 = 1.4

z1 = 90
z2 = 80
z3 = 60
z4 = 40
z5 = 30
z6 = 20

Z = G
    If Z > z1 Then
        ������� = y1
        End If
        If Z = z1 Then
        ������� = y1
        End If
    If (z2 < Z) And (Z < z1) Then
        ������� = y1 - ((y1 - y2) * (z1 - Z)) / (z1 - z2)
        End If
        If Z = z2 Then
        ������� = y2
        End If
    If (z3 < Z) And (Z < z2) Then
        ������� = y2 - ((y2 - y3) * (z2 - Z)) / (z2 - z3)
        End If
        If Z = z3 Then
        ������� = y3
        End If
    If (z4 < Z) And (Z < z3) Then
        ������� = y3 - ((y3 - y4) * (z3 - Z)) / (z3 - z4)
        End If
        If Z = z4 Then
        ������� = y4
        End If
   If (z5 < Z) And (Z < z4) Then
        ������� = y4 - ((y4 - y5) * (z4 - Z)) / (z4 - z5)
        End If
        If Z = z5 Then
        ������� = y5
        End If
   If (z6 < Z) And (Z < z5) Then
        ������� = y5 - ((y5 - y6) * (z5 - Z)) / (z5 - z6)
        End If
        If Z = z6 Then
        ������� = y6
        End If
    If Z < z6 Then
        ������� = y6
        End If
    aq = �������
End Function

' ����������� ��������� ��������
Function aq_1(G)
If G = 0 Then
aq_1 = 0
Exit Function
End If

y1 = 1.4
y2 = 1.4
y3 = 1.5
y4 = 1.6
y5 = 1.7
y6 = 1.7

z1 = 80
z2 = 70
z3 = 60
z4 = 40
z5 = 20
z6 = 10

Z = G
    If Z > z1 Then
        ������� = y1
        End If
        If Z = z1 Then
        ������� = y1
        End If
    If (z2 < Z) And (Z < z1) Then
        ������� = y1 - ((y1 - y2) * (z1 - Z)) / (z1 - z2)
        End If
        If Z = z2 Then
        ������� = y2
        End If
    If (z3 < Z) And (Z < z2) Then
        ������� = y2 - ((y2 - y3) * (z2 - Z)) / (z2 - z3)
        End If
        If Z = z3 Then
        ������� = y3
        End If
    If (z4 < Z) And (Z < z3) Then
        ������� = y3 - ((y3 - y4) * (z3 - Z)) / (z3 - z4)
        End If
        If Z = z4 Then
        ������� = y4
        End If
   If (z5 < Z) And (Z < z4) Then
        ������� = y4 - ((y4 - y5) * (z4 - Z)) / (z4 - z5)
        End If
        If Z = z5 Then
        ������� = y5
        End If
   If (z6 < Z) And (Z < z5) Then
        ������� = y5 - ((y5 - y6) * (z5 - Z)) / (z5 - z6)
        End If
        If Z = z6 Then
        ������� = y6
        End If
    If Z < z6 Then
        ������� = y6
        End If
    aq_1 = �������
End Function

'��� ��������� �������������
Function length(source As String, Chart As String, Regime As Integer) As Double
Dim list0 As Object
Dim list1 As Object
Dim list2 As Object
Dim list3 As Object
    Set list0 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 7), Sheets("calculation").Cells(25000, 7))
    Set list1 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 12), Sheets("calculation").Cells(25000, 12))
    Set list2 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 11), Sheets("calculation").Cells(25000, 11))
    Set list3 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 15), Sheets("calculation").Cells(25000, 15))
        length = Application.SumIfs(list0, list1, source, list2, Chart, list3, Regime)
End Function

'��� ������������ ��������������
Function MX(source As String, Chart As String, Regime As Integer) As Double
Dim list0 As Object
Dim list1 As Object
Dim list2 As Object
Dim list3 As Object
    Set list0 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 10), Sheets("calculation").Cells(25000, 10))
    Set list1 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 12), Sheets("calculation").Cells(25000, 12))
    Set list2 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 11), Sheets("calculation").Cells(25000, 11))
    Set list3 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 15), Sheets("calculation").Cells(25000, 15))
        MX = Application.SumIfs(list0, list1, source, list2, Chart, list3, Regime)
End Function

'��� �����
Function volume(source, Chart, Regime)
Dim list0 As Object
Dim list1 As Object
Dim list2 As Object
Dim list3 As Object
    Set list0 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 21), Sheets("calculation").Cells(25000, 21))
    Set list1 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 12), Sheets("calculation").Cells(25000, 12))
    Set list2 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 11), Sheets("calculation").Cells(25000, 11))
    Set list3 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 14), Sheets("calculation").Cells(25000, 14))
        volume = Application.SumIfs(list0, list1, source, list2, Chart, list3, Regime)
End Function

'����� ������������� ������� ������ �������� ������� ����� ��������
Function Qizol(source, Chart, Regime, Typ)
Dim list0 As Object
Dim list1 As Object
Dim list2 As Object
Dim list3 As Object
Dim list4 As Object 'Direction As String
    Set list0 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 17), Sheets("calculation").Cells(25000, 17))
    Set list1 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 12), Sheets("calculation").Cells(25000, 12))
    Set list2 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 11), Sheets("calculation").Cells(25000, 11))
    Set list3 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 14), Sheets("calculation").Cells(25000, 14))
    Set list4 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 8), Sheets("calculation").Cells(25000, 8))
        Qizol = Application.SumIfs(list0, list1, source, list2, Chart, list3, Regime, list4, Typ) * 1000000
End Function

'����� ������������� ������� ������ �������� ������� ����� �������� ��� ������ ��������
Function Qizolnadz(source, Chart, Regime, Typ, Direction)
Dim list0 As Object
Dim list1 As Object
Dim list2 As Object
Dim list3 As Object
Dim list4 As Object
Dim list5 As Object
    Set list0 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 17), Sheets("calculation").Cells(25000, 17))
    Set list1 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 12), Sheets("calculation").Cells(25000, 12))
    Set list2 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 11), Sheets("calculation").Cells(25000, 11))
    Set list3 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 14), Sheets("calculation").Cells(25000, 14))
    Set list4 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 8), Sheets("calculation").Cells(25000, 8))
    Set list5 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 3), Sheets("calculation").Cells(25000, 3))
        Qizolnadz = Application.SumIfs(list0, list1, source, list2, Chart, list3, Regime, list4, Typ, list5, Direction) * 1000000
End Function

'����������� ������
Function tyn(Chart As String, Regime As String) As Double
Dim t1 As Double
Dim t2 As Double
Dim tx As Double
Dim b As Double
    b = Sheets("temperature").Cells(37, 34)
        If Chart = "������ 1" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 34)
            t2 = Sheets("temperature").Cells(19, 35)
            tx = Sheets("temperature").Cells(19, 54)
            ElseIf Chart = "������ 2" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 36)
            t2 = Sheets("temperature").Cells(19, 37)
            tx = Sheets("temperature").Cells(19, 55)
            ElseIf Chart = "������ 3" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 38)
            t2 = Sheets("temperature").Cells(19, 39)
            tx = Sheets("temperature").Cells(19, 56)
            ElseIf Chart = "������ 4" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 40)
            t2 = Sheets("temperature").Cells(19, 41)
            tx = Sheets("temperature").Cells(19, 57)
            ElseIf Chart = "������ 5" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 42)
            t2 = Sheets("temperature").Cells(19, 43)
            tx = Sheets("temperature").Cells(19, 58)
            ElseIf Chart = "������ 6" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 44)
            t2 = Sheets("temperature").Cells(19, 45)
            tx = Sheets("temperature").Cells(19, 59)
            ElseIf Chart = "������ 7" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 46)
            t2 = Sheets("temperature").Cells(19, 47)
            tx = Sheets("temperature").Cells(19, 60)
            ElseIf Chart = "������ 8" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 48)
            t2 = Sheets("temperature").Cells(19, 49)
            tx = Sheets("temperature").Cells(19, 61)
            ElseIf Chart = "������ 9" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 50)
            t2 = Sheets("temperature").Cells(19, 51)
            tx = Sheets("temperature").Cells(19, 62)
            ElseIf Chart = "������ 10" And Regime = "���" Then
            t1 = Sheets("temperature").Cells(19, 52)
            t2 = Sheets("temperature").Cells(19, 53)
            tx = Sheets("temperature").Cells(19, 63)
            ElseIf Chart = "������ 1" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 34)
            t2 = Sheets("temperature").Cells(20, 35)
            tx = Sheets("temperature").Cells(20, 54)
            ElseIf Chart = "������ 2" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 36)
            t2 = Sheets("temperature").Cells(20, 37)
            tx = Sheets("temperature").Cells(20, 55)
            ElseIf Chart = "������ 3" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 38)
            t2 = Sheets("temperature").Cells(20, 39)
            tx = Sheets("temperature").Cells(20, 56)
            ElseIf Chart = "������ 4" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 40)
            t2 = Sheets("temperature").Cells(20, 41)
            tx = Sheets("temperature").Cells(20, 57)
            ElseIf Chart = "������ 5" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 42)
            t2 = Sheets("temperature").Cells(20, 43)
            tx = Sheets("temperature").Cells(20, 58)
            ElseIf Chart = "������ 6" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 44)
            t2 = Sheets("temperature").Cells(20, 45)
            tx = Sheets("temperature").Cells(20, 59)
            ElseIf Chart = "������ 7" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 46)
            t2 = Sheets("temperature").Cells(20, 47)
            tx = Sheets("temperature").Cells(20, 60)
            ElseIf Chart = "������ 8" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 48)
            t2 = Sheets("temperature").Cells(20, 49)
            tx = Sheets("temperature").Cells(20, 61)
            ElseIf Chart = "������ 9" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 50)
            t2 = Sheets("temperature").Cells(20, 51)
            tx = Sheets("temperature").Cells(20, 62)
            ElseIf Chart = "������ 10" And Regime = "��" Then
            t1 = Sheets("temperature").Cells(20, 52)
            t2 = Sheets("temperature").Cells(20, 53)
            tx = Sheets("temperature").Cells(20, 63)
            ElseIf Chart = "������ 1" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 34)
            t2 = Sheets("temperature").Cells(21, 35)
            tx = Sheets("temperature").Cells(21, 54)
            ElseIf Chart = "������ 2" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 36)
            t2 = Sheets("temperature").Cells(21, 37)
            tx = Sheets("temperature").Cells(21, 55)
            ElseIf Chart = "������ 3" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 38)
            t2 = Sheets("temperature").Cells(21, 39)
            tx = Sheets("temperature").Cells(21, 56)
            ElseIf Chart = "������ 4" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 40)
            t2 = Sheets("temperature").Cells(21, 41)
            tx = Sheets("temperature").Cells(21, 57)
            ElseIf Chart = "������ 5" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 42)
            t2 = Sheets("temperature").Cells(21, 43)
            tx = Sheets("temperature").Cells(21, 58)
            ElseIf Chart = "������ 6" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 44)
            t2 = Sheets("temperature").Cells(21, 45)
            tx = Sheets("temperature").Cells(21, 59)
            ElseIf Chart = "������ 7" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 46)
            t2 = Sheets("temperature").Cells(21, 47)
            tx = Sheets("temperature").Cells(21, 60)
            ElseIf Chart = "������ 8" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 48)
            t2 = Sheets("temperature").Cells(21, 49)
            tx = Sheets("temperature").Cells(21, 61)
            ElseIf Chart = "������ 9" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 50)
            t2 = Sheets("temperature").Cells(21, 51)
            tx = Sheets("temperature").Cells(21, 62)
            ElseIf Chart = "������ 10" And Regime = "����" Then
            t1 = Sheets("temperature").Cells(21, 52)
            t2 = Sheets("temperature").Cells(21, 53)
            tx = Sheets("temperature").Cells(21, 63)
        End If
    tyn = b * t1 + (1 - b) * t2 - tx
End Function

'����������� ����������
Function tzp(Chart As String, Regime As String) As Double
    If Chart = "������ 1" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 54)
        ElseIf Chart = "������ 2" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 55)
        ElseIf Chart = "������ 3" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 56)
        ElseIf Chart = "������ 4" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 57)
        ElseIf Chart = "������ 5" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 58)
        ElseIf Chart = "������ 6" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 59)
        ElseIf Chart = "������ 7" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 60)
        ElseIf Chart = "������ 8" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 61)
        ElseIf Chart = "������ 9" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 62)
        ElseIf Chart = "������ 10" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 63)
        ElseIf Chart = "������ 1" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 54)
        ElseIf Chart = "������ 2" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 55)
        ElseIf Chart = "������ 3" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 56)
        ElseIf Chart = "������ 4" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 57)
        ElseIf Chart = "������ 5" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 58)
        ElseIf Chart = "������ 6" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 59)
        ElseIf Chart = "������ 7" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 60)
        ElseIf Chart = "������ 8" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 61)
        ElseIf Chart = "������ 9" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 62)
        ElseIf Chart = "������ 10" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 63)
        ElseIf Chart = "������ 1" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 54)
        ElseIf Chart = "������ 2" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 55)
        ElseIf Chart = "������ 3" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 56)
        ElseIf Chart = "������ 4" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 57)
        ElseIf Chart = "������ 5" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 58)
        ElseIf Chart = "������ 6" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 59)
        ElseIf Chart = "������ 7" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 60)
        ElseIf Chart = "������ 8" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 61)
        ElseIf Chart = "������ 9" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 62)
        ElseIf Chart = "������ 10" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 63)
    End If
tzp = Sheets("temperature").Cells(39, 34) - tx
End Function

'����������� ������������� ��� ���������
Function tisp(Chart As String, Regime As String) As Double
If Chart = "������ 1" And Regime = "���" Then
    tx = Sheets("temperature").Cells(19, 54)
    ElseIf Chart = "������ 2" And Regime = "���" Then
    tx = Sheets("temperature").Cells(19, 55)
    ElseIf Chart = "������ 3" And Regime = "���" Then
    tx = Sheets("temperature").Cells(19, 56)
    ElseIf Chart = "������ 4" And Regime = "���" Then
    tx = Sheets("temperature").Cells(19, 57)
    ElseIf Chart = "������ 5" And Regime = "���" Then
    tx = Sheets("temperature").Cells(19, 58)
    ElseIf Chart = "������ 6" And Regime = "���" Then
    tx = Sheets("temperature").Cells(19, 59)
    ElseIf Chart = "������ 7" And Regime = "���" Then
    tx = Sheets("temperature").Cells(19, 60)
    ElseIf Chart = "������ 8" And Regime = "���" Then
    tx = Sheets("temperature").Cells(19, 61)
    ElseIf Chart = "������ 9" And Regime = "���" Then
    tx = Sheets("temperature").Cells(19, 62)
    ElseIf Chart = "������ 10" And Regime = "���" Then
    tx = Sheets("temperature").Cells(19, 63)
    ElseIf Chart = "������ 1" And Regime = "��" Then
    tx = Sheets("temperature").Cells(20, 54)
    ElseIf Chart = "������ 2" And Regime = "��" Then
    tx = Sheets("temperature").Cells(20, 55)
    ElseIf Chart = "������ 3" And Regime = "��" Then
    tx = Sheets("temperature").Cells(20, 56)
    ElseIf Chart = "������ 4" And Regime = "��" Then
    tx = Sheets("temperature").Cells(20, 57)
    ElseIf Chart = "������ 5" And Regime = "��" Then
    tx = Sheets("temperature").Cells(20, 58)
    ElseIf Chart = "������ 6" And Regime = "��" Then
    tx = Sheets("temperature").Cells(20, 59)
    ElseIf Chart = "������ 7" And Regime = "��" Then
    tx = Sheets("temperature").Cells(20, 60)
    ElseIf Chart = "������ 8" And Regime = "��" Then
    tx = Sheets("temperature").Cells(20, 61)
    ElseIf Chart = "������ 9" And Regime = "��" Then
    tx = Sheets("temperature").Cells(20, 62)
    ElseIf Chart = "������ 10" And Regime = "��" Then
    tx = Sheets("temperature").Cells(20, 63)
    ElseIf Chart = "������ 1" And Regime = "����" Then
    tx = Sheets("temperature").Cells(21, 54)
    ElseIf Chart = "������ 2" And Regime = "����" Then
    tx = Sheets("temperature").Cells(21, 55)
    ElseIf Chart = "������ 3" And Regime = "����" Then
    tx = Sheets("temperature").Cells(21, 56)
    ElseIf Chart = "������ 4" And Regime = "����" Then
    tx = Sheets("temperature").Cells(21, 57)
    ElseIf Chart = "������ 5" And Regime = "����" Then
    tx = Sheets("temperature").Cells(21, 58)
    ElseIf Chart = "������ 6" And Regime = "����" Then
    tx = Sheets("temperature").Cells(21, 59)
    ElseIf Chart = "������ 7" And Regime = "����" Then
    tx = Sheets("temperature").Cells(21, 60)
    ElseIf Chart = "������ 8" And Regime = "����" Then
    tx = Sheets("temperature").Cells(21, 61)
    ElseIf Chart = "������ 9" And Regime = "����" Then
    tx = Sheets("temperature").Cells(21, 62)
    ElseIf Chart = "������ 10" And Regime = "����" Then
    tx = Sheets("temperature").Cells(21, 63)
End If
    tisp = Sheets("temperature").Cells(38, 34) - tx
End Function
'Sub dgdfg()
'df = Qynm(45121.3, "������ 1", "��", "������")
'End Sub
'������ ����� � �������� ������� ���� ��� �������
Function Qynm(gyn, Chart, Regime, Month)
'Dim b As Double
    If Month = "������" Then i = 7
    If Month = "�������" Then i = 8
    If Month = "����" Then i = 9
    If Month = "������" Then i = 10
    If Month = "���" Then i = 11
    If Month = "����" Then i = 12
    If Month = "����" Then i = 13
    If Month = "������" Then i = 14
    If Month = "��������" Then i = 15
    If Month = "�������" Then i = 16
    If Month = "������" Then i = 17
    If Month = "�������" Then i = 18
        If Chart = "������ 1" Then k = 34
        If Chart = "������ 2" Then k = 36
        If Chart = "������ 3" Then k = 38
        If Chart = "������ 4" Then k = 40
        If Chart = "������ 5" Then k = 42
        If Chart = "������ 6" Then k = 44
        If Chart = "������ 7" Then k = 46
        If Chart = "������ 8" Then k = 48
        If Chart = "������ 9" Then k = 50
        If Chart = "������ 10" Then k = 52
            If Chart = "������ 1" Then l = 35
            If Chart = "������ 2" Then l = 37
            If Chart = "������ 3" Then l = 39
            If Chart = "������ 4" Then l = 41
            If Chart = "������ 5" Then l = 43
            If Chart = "������ 6" Then l = 45
            If Chart = "������ 7" Then l = 47
            If Chart = "������ 8" Then l = 49
            If Chart = "������ 9" Then l = 51
            If Chart = "������ 10" Then l = 53
                If Chart = "������ 1" Then e = 54
                If Chart = "������ 2" Then e = 55
                If Chart = "������ 3" Then e = 56
                If Chart = "������ 4" Then e = 57
                If Chart = "������ 5" Then e = 58
                If Chart = "������ 6" Then e = 59
                If Chart = "������ 7" Then e = 60
                If Chart = "������ 8" Then e = 61
                If Chart = "������ 9" Then e = 62
                If Chart = "������ 10" Then e = 63
                    If Regime = "���" Then j = 22
                    If Regime = "��" Then j = 23
                    If Regime = "����" Then j = 24
                    If Regime = "���" Then u = 19
                    If Regime = "��" Then u = 20
                    If Regime = "����" Then u = 21
                        tgrmo = Sheets("temperature").Cells(i, 32) '����������� ������ ��������������
                        tvozmo = Sheets("temperature").Cells(i, 33) '����������� ������� ��������������
                        tgrsr = Sheets("temperature").Cells(j, k) '����������� ������ �������������
                        tvozsr = Sheets("temperature").Cells(j, l) '����������� ������� �������������
                        t1mo = Sheets("temperature").Cells(i, k) '����������� � ������ ��������������
                        t2mo = Sheets("temperature").Cells(i, l) '����������� � ������� ��������������
                        t1sr = Sheets("temperature").Cells(u, k) '����������� � ������ �������������
                        t2sr = Sheets("temperature").Cells(u, l) '����������� � ������� �������������
                        tpom = Sheets("temperature").Cells(22, 32) '����������� � ���������
                        tton = Sheets("temperature").Cells(23, 32) '���������� � �������
                        tx = Sheets("temperature").Cells(u, e) '����������� �������� ����
                        b = Sheets("temperature").Cells(37, 34) '���� ��������� ������� �������������
                        Qynm = (gyn * (b * t1mo + (1 - b) * t2mo - tx) * wskDSWT((t1mo * b + t2mo * (1 - b) - tx))) / 1000000
End Function
'Sub dfgdgfhjg()
'gbj = Qzpm(36388.14, "������ 1", "��", "����")
'End Sub
'������ ����� � �������� ������� ���� ��� ���������� �������������
Function Qzpm(gzp, Chart, Regime, Month)
Dim tzp As Double
    If Month = "������" Then i = 7
    If Month = "�������" Then i = 8
    If Month = "����" Then i = 9
    If Month = "������" Then i = 10
    If Month = "���" Then i = 11
    If Month = "����" Then i = 12
    If Month = "����" Then i = 13
    If Month = "������" Then i = 14
    If Month = "��������" Then i = 15
    If Month = "�������" Then i = 16
    If Month = "������" Then i = 17
    If Month = "�������" Then i = 18
        If Chart = "������ 1" Then e = 54
        If Chart = "������ 2" Then e = 55
        If Chart = "������ 3" Then e = 56
        If Chart = "������ 4" Then e = 57
        If Chart = "������ 5" Then e = 58
        If Chart = "������ 6" Then e = 59
        If Chart = "������ 7" Then e = 60
        If Chart = "������ 8" Then e = 61
        If Chart = "������ 9" Then e = 62
        If Chart = "������ 10" Then e = 63
            tzp = Sheets("temperature").Cells(39, 34) '����������� ������������� ��� ����������
            tx = Sheets("temperature").Cells(i, e) '����������� �������� ����
            Qzpm = gzp * (tzp - tx) * wskDSWT(tzp - tx) / 1000000
End Function

'������ ����� � �������� ������� ���� ��� ���������
Function Qispm(gisp, Chart, Regime, Month) As Double
Dim tisp As Double
    If Month = "������" Then i = 7
    If Month = "�������" Then i = 8
    If Month = "����" Then i = 9
    If Month = "������" Then i = 10
    If Month = "���" Then i = 11
    If Month = "����" Then i = 12
    If Month = "����" Then i = 13
    If Month = "������" Then i = 14
    If Month = "��������" Then i = 15
    If Month = "�������" Then i = 16
    If Month = "������" Then i = 17
    If Month = "�������" Then i = 18
        If Chart = "������ 1" Then e = 54
        If Chart = "������ 2" Then e = 55
        If Chart = "������ 3" Then e = 56
        If Chart = "������ 4" Then e = 57
        If Chart = "������ 5" Then e = 58
        If Chart = "������ 6" Then e = 59
        If Chart = "������ 7" Then e = 60
        If Chart = "������ 8" Then e = 61
        If Chart = "������ 9" Then e = 62
        If Chart = "������ 10" Then e = 63
            tisp = Sheets("temperature").Cells(38, 34) '����������� ������������� ��� ������������ ����������
            tx = Sheets("temperature").Cells(i, e) '����������� �������� ����
                Qispm = gisp * (tisp - tx) * wskDSWT(tisp - tx) / 1000000
End Function

'�������������� � ������������� ������� �� ����� ��������, ����/�
Function Qizolsr(source, Chart, Regime, Typ, Month, Direction) As Double
If Month = "������" Then i = 7
If Month = "�������" Then i = 8
If Month = "����" Then i = 9
If Month = "������" Then i = 10
If Month = "���" Then i = 11
If Month = "����" Then i = 12
If Month = "����" Then i = 13
If Month = "������" Then i = 14
If Month = "��������" Then i = 15
If Month = "�������" Then i = 16
If Month = "������" Then i = 17
If Month = "�������" Then i = 18
    If Chart = "������ 1" Then k = 34
    If Chart = "������ 2" Then k = 36
    If Chart = "������ 3" Then k = 38
    If Chart = "������ 4" Then k = 40
    If Chart = "������ 5" Then k = 42
    If Chart = "������ 6" Then k = 44
    If Chart = "������ 7" Then k = 46
    If Chart = "������ 8" Then k = 48
    If Chart = "������ 9" Then k = 50
    If Chart = "������ 10" Then k = 52
        If Chart = "������ 1" Then l = 35
        If Chart = "������ 2" Then l = 37
        If Chart = "������ 3" Then l = 39
        If Chart = "������ 4" Then l = 41
        If Chart = "������ 5" Then l = 43
        If Chart = "������ 6" Then l = 45
        If Chart = "������ 7" Then l = 47
        If Chart = "������ 8" Then l = 49
        If Chart = "������ 9" Then l = 51
        If Chart = "������ 10" Then l = 53
            If Regime = "���" Then j = 22
            If Regime = "��" Then j = 23
            If Regime = "����" Then j = 24
            If Regime = "���" Then u = 19
            If Regime = "��" Then u = 20
            If Regime = "����" Then u = 21
                tgrmo = Sheets("temperature").Cells(i, 32) '����������� ������ ��������������
                tvozmo = Sheets("temperature").Cells(i, 33) '����������� ������� ��������������
                tgrsr = Sheets("temperature").Cells(j, k) '����������� ������ �������������
                tvozsr = Sheets("temperature").Cells(j, l) '����������� ������� �������������
                t1mo = Sheets("temperature").Cells(i, k) '����������� � ������ ��������������
                t2mo = Sheets("temperature").Cells(i, l) '����������� � ������� ��������������
                t1sr = Sheets("temperature").Cells(u, k) '����������� � ������ �������������
                t2sr = Sheets("temperature").Cells(u, l) '����������� � ������� �������������
                tpom = Sheets("temperature").Cells(22, 32) '����������� � ���������
                tton = Sheets("temperature").Cells(23, 32) '���������� � �������
                    If Typ = "��������� ���������" And t1mo <> "" And t2mo <> "" Then
                        Qizolsr = Qizol(source, Chart, Regime, "��������� ���������") * (t1mo + t2mo - 2 * tgrmo) / (t1sr + t2sr - 2 * tgrsr) + _
                        Qizol(source, Chart, Regime, "��������� ������������") * (t1mo + t2mo - 2 * tgrmo) / (t1sr + t2sr - 2 * tgrsr) + _
                        Qizol(source, Chart, Regime, "���������") * (t1mo + t2mo - 2 * tpom) / (t1sr + t2sr - 2 * tpom) + _
                        Qizol(source, Chart, Regime, "�������") * (t1mo + t2mo - 2 * tton) / (t1sr + t2sr - 2 * tton)
                    ElseIf Typ = "���������" And Direction = "������" And t1mo <> "" Then
                        Qizolsr = Qizolnadz(source, Chart, Regime, "���������", "������") * (t1mo - tvozmo) / (t1sr - tvozsr)
                    ElseIf Typ = "���������" And Direction = "�������" And t2mo <> "" Then
                        Qizolsr = Qizolnadz(source, Chart, Regime, "���������", "�������") * (t2mo - tvozmo) / (t2sr - tvozsr)
                    End If
End Function

'������ ����� � �������� ������� ���� � ����
Function Qsarz(n1, n2, k1, k2, gyn1, gyn2, Chart, Regime, Month)
Dim b As Double
    If Month = "������" Then i = 7
    If Month = "�������" Then i = 8
    If Month = "����" Then i = 9
    If Month = "������" Then i = 10
    If Month = "���" Then i = 11
    If Month = "����" Then i = 12
    If Month = "����" Then i = 13
    If Month = "������" Then i = 14
    If Month = "��������" Then i = 15
    If Month = "�������" Then i = 16
    If Month = "������" Then i = 17
    If Month = "�������" Then i = 18
        If Chart = "������ 1" Then k = 34
        If Chart = "������ 2" Then k = 36
        If Chart = "������ 3" Then k = 38
        If Chart = "������ 4" Then k = 40
        If Chart = "������ 5" Then k = 42
        If Chart = "������ 6" Then k = 44
        If Chart = "������ 7" Then k = 46
        If Chart = "������ 8" Then k = 48
        If Chart = "������ 9" Then k = 50
        If Chart = "������ 10" Then k = 52
            If Chart = "������ 1" Then l = 35
            If Chart = "������ 2" Then l = 37
            If Chart = "������ 3" Then l = 39
            If Chart = "������ 4" Then l = 41
            If Chart = "������ 5" Then l = 43
            If Chart = "������ 6" Then l = 45
            If Chart = "������ 7" Then l = 47
            If Chart = "������ 8" Then l = 49
            If Chart = "������ 9" Then l = 51
            If Chart = "������ 10" Then l = 53
                t1mo = Sheets("temperature").Cells(i, k) '����������� � ������ ��������������
                t2mo = Sheets("temperature").Cells(i, l) '����������� � ������� ��������������
                Qsarz = (n1 + n2) * ((gyn1 * k1 * (t1mo - tx) * wskDSWT(t1mo - tx)) + (gyn2 * k2 * (t2mo - tx) * wskDSWT(t2mo - tx))) / 1000000
End Function

'���� ������������ ��������������
Function MXratio(source As String, Chart As String, Regime As String, Typ As String, Year As Integer) As Double
Dim list0 As Object
Dim list1 As Object
Dim list2 As Object
Dim list3 As Object
Dim list4 As Object
Dim list5 As Object
    Set list0 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 10), Sheets("calculation").Cells(25000, 10)) 'Range(Cells(7, 10), Cells(15000, 10))
    Set list1 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 12), Sheets("calculation").Cells(25000, 12))
    Set list2 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 11), Sheets("calculation").Cells(25000, 11))
    Set list3 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 14), Sheets("calculation").Cells(25000, 14))
    Set list4 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 8), Sheets("calculation").Cells(25000, 8))
    Set list5 = Sheets("calculation").Range(Sheets("calculation").Cells(7, 32), Sheets("calculation").Cells(25000, 32))
        If Typ = "��������� ���������" And Year = 1 Then
            MXratio = Application.SumIfs(list0, list1, source, list2, Chart, list3, Regime, list4, Typ, list5, Year) + Application.SumIfs(list0, list1, source, list2, Chart, list3, Regime, list4, "��������� ������������", list5, Year)
        ElseIf Typ = "��������� ������������" And Year = 3 Then
            MXratio = Application.SumIfs(list0, list1, source, list2, Chart, list3, Regime, list4, Typ, list5, Year) + Application.SumIfs(list0, list1, source, list2, Chart, list3, Regime, list4, "��������� ���������", list5, Year)
        Else: MXratio = Application.SumIfs(list0, list1, source, list2, Chart, list3, Regime, list4, Typ, list5, Year)
        End If
End Function
'Sub dfgbdf()
'    E = coefficientLosses(2, "������ 2", "���", "�������", 2004, "�������")
'End Sub
'����������� ������
Function coefficientLosses(source As String, Chart As String, Regime As String, Typ As String, Year As Integer, Direction As String)
If source = 1 And Year <= 1989 And Typ = "���������" And Direction = "������" Then Line = 4 '+ (source * 30)
If source = 1 And Year <= 1989 And Typ = "���������" And Direction = "�������" Then Line = 5
If source = 1 And Year <= 1989 And Typ = "�������" And Direction = "������" Then Line = 6 '+ (source * 30)
If source = 1 And Year <= 1989 And Typ = "�������" And Direction = "�������" Then Line = 7
If source = 1 And Year <= 1989 And Typ = "���������" And Direction = "������" Then Line = 8
If source = 1 And Year <= 1989 And Typ = "���������" And Direction = "�������" Then Line = 9
If source = 1 And Year <= 1989 And Typ = "��������� ���������" Then Line = 10
If source = 1 And Year <= 1989 And Typ = "��������� ������������" Then Line = 10
    If source = 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "������" Then Line = 11
    If source = 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "�������" Then Line = 12
    If source = 1 And Year > 1989 And Year < 1998 And Typ = "�������" And Direction = "������" Then Line = 13
    If source = 1 And Year > 1989 And Year < 1998 And Typ = "�������" And Direction = "�������" Then Line = 14
    If source = 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "������" Then Line = 15
    If source = 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "�������" Then Line = 16
    If source = 1 And Year > 1989 And Year < 1998 And Typ = "��������� ������������" Then Line = 17
    If source = 1 And Year > 1989 And Year < 1998 And Typ = "��������� ���������" Then Line = 18
        If source = 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "������" Then Line = 19
        If source = 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "�������" Then Line = 20
        If source = 1 And Year > 1997 And Year < 2004 And Typ = "�������" And Direction = "������" Then Line = 21
        If source = 1 And Year > 1997 And Year < 2004 And Typ = "�������" And Direction = "�������" Then Line = 22
        If source = 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "������" Then Line = 23
        If source = 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "�������" Then Line = 24
        If source = 1 And Year > 1997 And Year < 2004 And Typ = "��������� ���������" Then Line = 25
        If source = 1 And Year > 1997 And Year < 2004 And Typ = "��������� ������������" Then Line = 25
            If source = 1 And Year > 2003 And Typ = "���������" And Direction = "������" Then Line = 26
            If source = 1 And Year > 2003 And Typ = "���������" And Direction = "�������" Then Line = 27
            If source = 1 And Year > 2003 And Typ = "�������" And Direction = "������" Then Line = 28
            If source = 1 And Year > 2003 And Typ = "�������" And Direction = "�������" Then Line = 29
            If source = 1 And Year > 2003 And Typ = "���������" And Direction = "������" Then Line = 30
            If source = 1 And Year > 2003 And Typ = "���������" And Direction = "�������" Then Line = 31
            If source = 1 And Year > 2003 And Typ = "��������� ������������" Then Line = 32
            If source = 1 And Year > 2003 And Typ = "��������� ���������" Then Line = 33
                If source <> 1 And Year <= 1989 And Typ = "���������" And Direction = "������" Then Line = 4 + ((source - 1) * 30)
                If source <> 1 And Year <= 1989 And Typ = "���������" And Direction = "�������" Then Line = 5 + ((source - 1) * 30)
                If source <> 1 And Year <= 1989 And Typ = "�������" And Direction = "������" Then Line = 6 + ((source - 1) * 30)
                If source <> 1 And Year <= 1989 And Typ = "�������" And Direction = "�������" Then Line = 7 + ((source - 1) * 30)
                If source <> 1 And Year <= 1989 And Typ = "���������" And Direction = "������" Then Line = 8 + ((source - 1) * 30)
                If source <> 1 And Year <= 1989 And Typ = "���������" And Direction = "�������" Then Line = 9 + ((source - 1) * 30)
                If source <> 1 And Year <= 1989 And Typ = "��������� ���������" Then Line = 10 + ((source - 1) * 30)
                If source <> 1 And Year <= 1989 And Typ = "��������� ������������" Then Line = 10 + ((source - 1) * 30)
                    If source <> 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "������" Then Line = 11 + ((source - 1) * 30)
                    If source <> 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "�������" Then Line = 12 + ((source - 1) * 30)
                    If source <> 1 And Year > 1989 And Year < 1998 And Typ = "�������" And Direction = "������" Then Line = 13 + ((source - 1) * 30)
                    If source <> 1 And Year > 1989 And Year < 1998 And Typ = "�������" And Direction = "�������" Then Line = 14 + ((source - 1) * 30)
                    If source <> 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "������" Then Line = 15 + ((source - 1) * 30)
                    If source <> 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "�������" Then Line = 16 + ((source - 1) * 30)
                    If source <> 1 And Year > 1989 And Year < 1998 And Typ = "��������� ������������" Then Line = 17 + ((source - 1) * 30)
                    If source <> 1 And Year > 1989 And Year < 1998 And Typ = "��������� ���������" Then Line = 18 + ((source - 1) * 30)
                        If source <> 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "������" Then Line = 19 + ((source - 1) * 30)
                        If source <> 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "�������" Then Line = 20 + ((source - 1) * 30)
                        If source <> 1 And Year > 1997 And Year < 2004 And Typ = "�������" And Direction = "������" Then Line = 21 + ((source - 1) * 30)
                        If source <> 1 And Year > 1997 And Year < 2004 And Typ = "�������" And Direction = "�������" Then Line = 22 + ((source - 1) * 30)
                        If source <> 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "������" Then Line = 23 + ((source - 1) * 30)
                        If source <> 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "�������" Then Line = 24 + ((source - 1) * 30)
                        If source <> 1 And Year > 1997 And Year < 2004 And Typ = "��������� ���������" Then Line = 25 + ((source - 1) * 30)
                        If source <> 1 And Year > 1997 And Year < 2004 And Typ = "��������� ������������" Then Line = 25 + ((source - 1) * 30)
                            If source <> 1 And Year > 2003 And Typ = "���������" And Direction = "������" Then Line = 26 + ((source - 1) * 30)
                            If source <> 1 And Year > 2003 And Typ = "���������" And Direction = "�������" Then Line = 27 + ((source - 1) * 30)
                            If source <> 1 And Year > 2003 And Typ = "�������" And Direction = "������" Then Line = 28 + ((source - 1) * 30)
                            If source <> 1 And Year > 2003 And Typ = "�������" And Direction = "�������" Then Line = 29 + ((source - 1) * 30)
                            If source <> 1 And Year > 2003 And Typ = "���������" And Direction = "������" Then Line = 30 + ((source - 1) * 30)
                            If source <> 1 And Year > 2003 And Typ = "���������" And Direction = "�������" Then Line = 31 + ((source - 1) * 30)
                            If source <> 1 And Year > 2003 And Typ = "��������� ������������" Then Line = 32 + ((source - 1) * 30)
                            If source <> 1 And Year > 2003 And Typ = "��������� ���������" Then Line = 33 + ((source - 1) * 30)
                                If Chart = "������ 1" And Regime = "���" Then Column = 6
                                If Chart = "������ 2" And Regime = "���" Then Column = 9
                                If Chart = "������ 3" And Regime = "���" Then Column = 12
                                If Chart = "������ 4" And Regime = "���" Then Column = 15
                                If Chart = "������ 5" And Regime = "���" Then Column = 18
                                If Chart = "������ 6" And Regime = "���" Then Column = 21
                                If Chart = "������ 7" And Regime = "���" Then Column = 24
                                If Chart = "������ 8" And Regime = "���" Then Column = 27
                                If Chart = "������ 9" And Regime = "���" Then Column = 30
                                If Chart = "������ 10" And Regime = "���" Then Column = 33
                                    If Chart = "������ 1" And Regime = "��" Then Column = 7
                                    If Chart = "������ 2" And Regime = "��" Then Column = 10
                                    If Chart = "������ 3" And Regime = "��" Then Column = 13
                                    If Chart = "������ 4" And Regime = "��" Then Column = 16
                                    If Chart = "������ 5" And Regime = "��" Then Column = 19
                                    If Chart = "������ 6" And Regime = "��" Then Column = 22
                                    If Chart = "������ 7" And Regime = "��" Then Column = 25
                                    If Chart = "������ 8" And Regime = "��" Then Column = 28
                                    If Chart = "������ 9" And Regime = "��" Then Column = 31
                                    If Chart = "������ 10" And Regime = "��" Then Column = 34
                                        If Chart = "������ 1" And Regime = "����" Then Column = 8
                                        If Chart = "������ 2" And Regime = "����" Then Column = 11
                                        If Chart = "������ 3" And Regime = "����" Then Column = 14
                                        If Chart = "������ 4" And Regime = "����" Then Column = 17
                                        If Chart = "������ 5" And Regime = "����" Then Column = 20
                                        If Chart = "������ 6" And Regime = "����" Then Column = 23
                                        If Chart = "������ 7" And Regime = "����" Then Column = 26
                                        If Chart = "������ 8" And Regime = "����" Then Column = 29
                                        If Chart = "������ 9" And Regime = "����" Then Column = 32
                                        If Chart = "������ 10" And Regime = "����" Then Column = 35
            If Line = Empty Then
                coefficientLosses = 1
            Else: coefficientLosses = Sheets("coefficient").Cells(Line, Column)
            End If
End Function

'������ �������� ��������� ���������� �������� ���� ��� ���������� ����������� ������
Sub coldWater()
'������ 1
If Sheets("temperature").Cells(22, 54) = "���" Then
    �����������.Range("BB20") = 5#
    �����������.Range("BB21") = 15#
    �����������.Range("BB20:BB21").Interior.ColorIndex = 35
    �����������.Range("BB7:BB18").Interior.ColorIndex = 0
    For i = 7 To 18
        �����������.Cells(i, 54).FormulaLocal = "=����(V" + CStr(i) + "="""";"""";($BB$20*B" + CStr(i) + "+$BB$21*C" + CStr(i) + ")/V" + CStr(i) + ")"
    Next i
End If
If Sheets("temperature").Cells(22, 54) = "����" Then
    �����������.Range("BB19:BB21").Interior.ColorIndex = 0
    �����������.Range("BB7:BB18").ClearContents
    �����������.Range("BB7:BB18").Interior.ColorIndex = 35
    �����������.Range("BB20").FormulaLocal = "=����($B$19="""";"""";����������(BB7:BB18;$B$7:$B$18)/$B$19)"
    �����������.Range("BB21").FormulaLocal = "=����($C$19="""";"""";����������(BB7:BB18;$C$7:$C$18)/$C$19)"
End If
'������ 2
If Sheets("temperature").Cells(22, 55) = "���" Then
    �����������.Range("BC20") = 5#
    �����������.Range("BC21") = 15#
    �����������.Range("BC20:BC21").Interior.ColorIndex = 35
    �����������.Range("BC7:BC18").Interior.ColorIndex = 0
    For i = 7 To 18
        �����������.Cells(i, 55).FormulaLocal = "=����(W" + CStr(i) + "="""";"""";($BC$20*D" + CStr(i) + "+$BC$21*E" + CStr(i) + ")/W" + CStr(i) + ")"
    Next i
End If
If Sheets("temperature").Cells(22, 55) = "����" Then
    �����������.Range("BC19:BC21").Interior.ColorIndex = 0
    �����������.Range("BC7:BC18").ClearContents
    �����������.Range("BC7:BC18").Interior.ColorIndex = 35
    �����������.Range("BC20").FormulaLocal = "=����($D$19="""";"""";����������(BC7:BC18;$D$7:$D$18)/$D$19)"
    �����������.Range("BC21").FormulaLocal = "=����($E$19="""";"""";����������(BC7:BC18;$E$7:$E$18)/$E$19)"
End If
'������ 3
If Sheets("temperature").Cells(22, 56) = "���" Then
    �����������.Range("BD20") = 5#
    �����������.Range("BD21") = 15#
    �����������.Range("BD20:BD21").Interior.ColorIndex = 35
    �����������.Range("BD7:BD18").Interior.ColorIndex = 0
    For i = 7 To 18
        �����������.Cells(i, 56).FormulaLocal = "=����(X" + CStr(i) + "="""";"""";($BD$20*F" + CStr(i) + "+$BD$21*G" + CStr(i) + ")/X" + CStr(i) + ")"
    Next i
End If
If Sheets("temperature").Cells(22, 56) = "����" Then
    �����������.Range("BD19:BD21").Interior.ColorIndex = 0
    �����������.Range("BD7:BD18").ClearContents
    �����������.Range("BD7:BD18").Interior.ColorIndex = 35
    �����������.Range("BD20").FormulaLocal = "=����($F$19="""";"""";����������(BD7:BD18;$F$7:$F$18)/$F$19)"
    �����������.Range("BD21").FormulaLocal = "=����($G$19="""";"""";����������(BD7:BD18;$G$7:$G$18)/$G$19)"
End If
'������ 4
If Sheets("temperature").Cells(22, 57) = "���" Then
    �����������.Range("BE20") = 5#
    �����������.Range("BE21") = 15#
    �����������.Range("BE20:BE21").Interior.ColorIndex = 35
    �����������.Range("BE7:BE18").Interior.ColorIndex = 0
    For i = 7 To 18
        �����������.Cells(i, 57).FormulaLocal = "=����(Y" + CStr(i) + "="""";"""";($BE$20*H" + CStr(i) + "+$BE$21*I" + CStr(i) + ")/Y" + CStr(i) + ")"
    Next i
End If
If Sheets("temperature").Cells(22, 57) = "����" Then
    �����������.Range("BE19:BE21").Interior.ColorIndex = 0
    �����������.Range("BE7:BE18").ClearContents
    �����������.Range("BE7:BE18").Interior.ColorIndex = 35
    �����������.Range("BE20").FormulaLocal = "=����($H$19="""";"""";����������(BE7:BE18;$H$7:$H$18)/$H$19)"
    �����������.Range("BE21").FormulaLocal = "=����($I$19="""";"""";����������(BE7:BE18;$I$7:$I$18)/$I$19)"
End If
'������ 5
If Sheets("temperature").Cells(22, 58) = "���" Then
    �����������.Range("BF20") = 5#
    �����������.Range("BF21") = 15#
    �����������.Range("BF20:BF21").Interior.ColorIndex = 35
    �����������.Range("BF7:BF18").Interior.ColorIndex = 0
    For i = 7 To 18
        �����������.Cells(i, 58).FormulaLocal = "=����(Z" + CStr(i) + "="""";"""";($BF$20*J" + CStr(i) + "+$BF$21*K" + CStr(i) + ")/Z" + CStr(i) + ")"
    Next i
End If
If Sheets("temperature").Cells(22, 58) = "����" Then
    �����������.Range("BF19:BF21").Interior.ColorIndex = 0
    �����������.Range("BF7:BF18").ClearContents
    �����������.Range("BF7:BF18").Interior.ColorIndex = 35
    �����������.Range("BF20").FormulaLocal = "=����($J$19="""";"""";����������(BF7:BF18;$J$7:$J$18)/$J$19)"
    �����������.Range("BF21").FormulaLocal = "=����($K$19="""";"""";����������(BF7:BF18;$K$7:$K$18)/$K$19)"
End If
'������ 6
If Sheets("temperature").Cells(22, 59) = "���" Then
    �����������.Range("BG20") = 5#
    �����������.Range("BG21") = 15#
    �����������.Range("BG20:BG21").Interior.ColorIndex = 35
    �����������.Range("BG7:BG18").Interior.ColorIndex = 0
    For i = 7 To 18
        �����������.Cells(i, 59).FormulaLocal = "=����(AA" + CStr(i) + "="""";"""";($BG$20*L" + CStr(i) + "+$BG$21*M" + CStr(i) + ")/AA" + CStr(i) + ")"
    Next i
End If
If Sheets("temperature").Cells(22, 59) = "����" Then
    �����������.Range("BG19:BG21").Interior.ColorIndex = 0
    �����������.Range("BG7:BG18").ClearContents
    �����������.Range("BG7:BG18").Interior.ColorIndex = 35
    �����������.Range("BG20").FormulaLocal = "=����($L$19="""";"""";����������(BG7:BG18;$L$7:$L$18)/$L$19)"
    �����������.Range("BG21").FormulaLocal = "=����($M$19="""";"""";����������(BG7:BG18;$M$7:$M$18)/$M$19)"
End If
'������ 7
If Sheets("temperature").Cells(22, 60) = "���" Then
    �����������.Range("BH20") = 5#
    �����������.Range("BH21") = 15#
    �����������.Range("BH20:BH21").Interior.ColorIndex = 35
    �����������.Range("BH7:BH18").Interior.ColorIndex = 0
    For i = 7 To 18
        �����������.Cells(i, 60).FormulaLocal = "=����(AB" + CStr(i) + "="""";"""";($BH$20*N" + CStr(i) + "+$BH$21*O" + CStr(i) + ")/AB" + CStr(i) + ")"
    Next i
End If
If Sheets("temperature").Cells(22, 60) = "����" Then
    �����������.Range("BH19:BH21").Interior.ColorIndex = 0
    �����������.Range("BH7:BH18").ClearContents
    �����������.Range("BH7:BH18").Interior.ColorIndex = 35
    �����������.Range("BH20").FormulaLocal = "=����($N$19="""";"""";����������(BH7:BH18;$N$7:$N$18)/$N$19)"
    �����������.Range("BH21").FormulaLocal = "=����($O$19="""";"""";����������(BH7:BH18;$O$7:$O$18)/$O$19)"
End If
'������ 8
If Sheets("temperature").Cells(22, 61) = "���" Then
    �����������.Range("BI20") = 5#
    �����������.Range("BI21") = 15#
    �����������.Range("BI20:BI21").Interior.ColorIndex = 35
    �����������.Range("BI7:BI18").Interior.ColorIndex = 0
    For i = 7 To 18
        �����������.Cells(i, 61).FormulaLocal = "=����(AC" + CStr(i) + "="""";"""";($BI$20*P" + CStr(i) + "+$BI$21*Q" + CStr(i) + ")/AC" + CStr(i) + ")"
    Next i
End If
If Sheets("temperature").Cells(22, 61) = "����" Then
    �����������.Range("BI19:BI21").Interior.ColorIndex = 0
    �����������.Range("BI7:BI18").ClearContents
    �����������.Range("BI7:BI18").Interior.ColorIndex = 35
    �����������.Range("BI20").FormulaLocal = "=����($P$19="""";"""";����������(BI7:BI18;$P$7:$P$18)/$P$19)"
    �����������.Range("BI21").FormulaLocal = "=����($Q$19="""";"""";����������(BI7:BI18;$Q$7:$Q$18)/$Q$19)"
End If
'������ 9
If Sheets("temperature").Cells(22, 62) = "���" Then
    �����������.Range("BJ20") = 5#
    �����������.Range("BJ21") = 15#
    �����������.Range("BJ20:BJ21").Interior.ColorIndex = 35
    �����������.Range("BJ7:BJ18").Interior.ColorIndex = 0
    For i = 7 To 18
        �����������.Cells(i, 62).FormulaLocal = "=����(AD" + CStr(i) + "="""";"""";($BJ$20*R" + CStr(i) + "+$BJ$21*S" + CStr(i) + ")/AD" + CStr(i) + ")"
    Next i
End If
If Sheets("temperature").Cells(22, 62) = "����" Then
    �����������.Range("BJ19:BJ21").Interior.ColorIndex = 0
    �����������.Range("BJ7:BJ18").ClearContents
    �����������.Range("BJ7:BJ18").Interior.ColorIndex = 35
    �����������.Range("BJ20").FormulaLocal = "=����($R$19="""";"""";����������(BJ7:BJ18;$R$7:$R$18)/$R$19)"
    �����������.Range("BJ21").FormulaLocal = "=����($S$19="""";"""";����������(BJ7:BJ18;$S$7:$S$18)/$S$19)"
End If
'������ 10
If Sheets("temperature").Cells(22, 63) = "���" Then
    �����������.Range("BK20") = 5#
    �����������.Range("BK21") = 15#
    �����������.Range("BK20:BK21").Interior.ColorIndex = 35
    �����������.Range("BK7:BK18").Interior.ColorIndex = 0
    For i = 7 To 18
        �����������.Cells(i, 63).FormulaLocal = "=����(AE" + CStr(i) + "="""";"""";($BK$20*T" + CStr(i) + "+$BK$21*U" + CStr(i) + ")/AE" + CStr(i) + ")"
    Next i
End If
If Sheets("temperature").Cells(22, 63) = "����" Then
    �����������.Range("BK19:BK21").Interior.ColorIndex = 0
    �����������.Range("BK7:BK18").ClearContents
    �����������.Range("BK7:BK18").Interior.ColorIndex = 35
    �����������.Range("BK20").FormulaLocal = "=����($T$19="""";"""";����������(BK7:BK18;$T$7:$T$18)/$T$19)"
    �����������.Range("BK21").FormulaLocal = "=����($U$19="""";"""";����������(BK7:BK18;$U$7:$U$18)/$U$19)"
End If
End Sub

'����������� ��������� �������� ������
Function efficiencyElectricMotor(power, load, turnovers)
If turnovers >= 1500 And turnovers <= 3000 Then
    Set rangeNormsEfficiency1500 = Sheets("1989N").Range("T8:AB21")
    Set rangeNormsEfficiency3000 = Sheets("1989N").Range("AC8:AK21")
        b = interpolationNorms(power, load, rangeNormsEfficiency1500)
        b2 = interpolationNorms(power, load, rangeNormsEfficiency3000)
            efficiencyElectricMotor = b + ((b2 - b) * (turnovers - 1500) / (3000 - 1500))
End If
If turnovers < 1500 Then
    Set rangeNormsEfficiency1500 = Sheets("1989N").Range("AL8:AT21")
    Set rangeNormsEfficiency3000 = Sheets("1989N").Range("T8:AB21")
        b = interpolationNorms(power, load, rangeNormsEfficiency1500)
        b2 = interpolationNorms(power, load, rangeNormsEfficiency3000)
            efficiencyElectricMotor = b + ((b2 - b) * (turnovers - 0) / (1500 - 0))
End If
If turnovers > 3000 Then
    Set rangeNormsEfficiency1500 = Sheets("1989N").Range("AC8:AK21")
    Set rangeNormsEfficiency3000 = Sheets("1989N").Range("AU8:BC21")
        b = interpolationNorms(power, load, rangeNormsEfficiency1500)
        b2 = interpolationNorms(power, load, rangeNormsEfficiency3000)
            efficiencyElectricMotor = b + ((b2 - b) * (turnovers - 3000) / (4500 - 3000))
End If
End Function

'������ ���������� ���������
Sub selectionTestParameters()
Dim accountRange As Range
Dim numberRings As Long
Dim rangeNormsResistance As Range
Set accountRange = Sheets("jumper").Range("A3:A1000")                               '�������� ����� �������������� �����
Set rangeNormsResistance = Sheets("1989N").Range("T49:W63")                         '�������� ������������� �������������� ��������� (�� 34.20.519-97)
    numberRings = Application.WorksheetFunction.Max(accountRange)                   '���������� �������������� ����
    For s = 3 To numberRings + 2 Step 1                                             '������� �� �������������� �������
        reductionTemperature = Sheets("jumper").Cells(s, 5)                         '����������� ��������� ����������� �� ���� ������ �� �������������� ������
        Dysl = Sheets("jumper").Cells(s, 3)
        Gr = Sheets("jumper").Cells(s, 4)
        dPper = Sheets("jumper").Cells(s, 9)
        If Dysl = Empty Then
            Dysl = 0.015
            Sheets("jumper").Cells(s, 3).Value = Dysl
        End If
        If Gr = Empty Then
            Gr = 5
            Sheets("jumper").Cells(s, 4).Value = Gr
        End If
        If dPper = Empty Then
            dPper = 2
            Sheets("jumper").Cells(s, 9).Value = dPper
        End If
        '�������� ������ �����
        Dim wks As Worksheet
        Application.Calculation = xlManual
        For Each wks In ActiveWorkbook.Worksheets
            wks.Calculate
        Next
        Set wks = Nothing
        Range("J" + CStr(s) + "").GoalSeek Goal:=reductionTemperature, ChangingCell:=Range("D" + CStr(s) + "") '������ ����������� � ����������� �� �������
    Next s
    For q = 3 To numberRings + 2 Step 1                                             '������� �� �������������� �������
        Set expense = Sheets("jumper").Range("D" + CStr(q) + "")
            If expense = Empty Then
                Exit For
            End If
            For i = 6 To 28 Step 1                                                  '������ �������� ��������� � ����������� �� ������ ������ �� ���������
                Set x = Sheets("1989N").Cells(i, 2)
                pressureLoss = interpolationNorms(x / 1000, 4, rangeNormsResistance) * (expense ^ 2)
                If pressureLoss <= 2 And pressureLoss > 0.0000001 Then
                    Exit For
                End If
            Next i
        Sheets("jumper").Range("C" + CStr(q) + "").Value = x / 1000
        Sheets("jumper").Range("I" + CStr(q) + "").Value = pressureLoss 'pressureLoss
    Next q
End Sub

'���������� �����������
Function limitingFactor(sourceTotal, regimeTotal, chartTotal, Typ)
    For i = 6 To 88 Step 2 '����������� ������� ���� ������������ �������������� �� ������� Ratio
        sourceRatio = Sheets("ratio").Cells(3, i)
        ChartRatio = Sheets("ratio").Cells(4, i)
        RegimeRatio = Sheets("ratio").Cells(5, i)
        If sourceRatio = sourceTotal And chartTotal = ChartRatio And regimeTotal = RegimeRatio Then
            columnFraction = i + 1
            Exit For
        End If
    Next i
    If Typ = "��������� ������������" Or Typ = "��������� ���������" Or Typ = "���������" Or Typ = "�������" Then
        shareUnderground = Sheets("ratio").Cells(26, columnFraction) '���� ���������
        limitingFactor = aq(shareUnderground)
    ElseIf Typ = "���������" Then
        shareGround = Sheets("ratio").Cells(27, columnFraction) '���� ���������
        limitingFactor = aq_1(shareGround)
    End If
End Function

'������ ������������� �� ������� total �� ������� coefficient
Sub entryCoefficients() 'source As String, Chart As String, Regime As String, Typ As String, Year As Integer, Direction As String
Dim source As Long
Dim Regime As String
Dim Chart As String
Dim Direction As String
Dim Year As Integer
Dim Typ As String
    ����20.Range("A1:AI263").Interior.Color = xlNone
    For j = 4 To 303 Step 1
        For l = 6 To 35 Step 1
            Sheets("coefficient").Cells(j, l).Value = 1
        Next l
    Next j
    For i = 4 To 1000 Step 1
        source = Sheets("total").Cells(i, 1)
        Regime = Sheets("total").Cells(i, 2)
        Chart = Sheets("total").Cells(i, 3)
        Direction = Sheets("total").Cells(i, 4)
        Year = Sheets("total").Cells(i, 5)
        Typ = Sheets("total").Cells(i, 6)
            If source = 1 And Year <= 1989 And Typ = "���������" And Direction = "������" Then Line = 4 '+ (source * 28)
            If source = 1 And Year <= 1989 And Typ = "���������" And Direction = "�������" Then Line = 5
            If source = 1 And Year <= 1989 And Typ = "�������" And Direction = "������" Then Line = 6 '+ (source * 28)
            If source = 1 And Year <= 1989 And Typ = "�������" And Direction = "�������" Then Line = 7
            If source = 1 And Year <= 1989 And Typ = "���������" And Direction = "������" Then Line = 8
            If source = 1 And Year <= 1989 And Typ = "���������" And Direction = "�������" Then Line = 9
            If source = 1 And Year <= 1989 And Typ = "��������� ���������" Then Line = 10
            If source = 1 And Year <= 1989 And Typ = "��������� ������������" Then Line = 10
                If source = 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "������" Then Line = 11
                If source = 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "�������" Then Line = 12
                If source = 1 And Year > 1989 And Year < 1998 And Typ = "�������" And Direction = "������" Then Line = 13
                If source = 1 And Year > 1989 And Year < 1998 And Typ = "�������" And Direction = "�������" Then Line = 14
                If source = 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "������" Then Line = 15
                If source = 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "�������" Then Line = 16
                If source = 1 And Year > 1989 And Year < 1998 And Typ = "��������� ������������" Then Line = 17
                If source = 1 And Year > 1989 And Year < 1998 And Typ = "��������� ���������" Then Line = 18
                    If source = 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "������" Then Line = 19
                    If source = 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "�������" Then Line = 20
                    If source = 1 And Year > 1997 And Year < 2004 And Typ = "�������" And Direction = "������" Then Line = 21
                    If source = 1 And Year > 1997 And Year < 2004 And Typ = "�������" And Direction = "�������" Then Line = 22
                    If source = 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "������" Then Line = 23
                    If source = 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "�������" Then Line = 24
                    If source = 1 And Year > 1997 And Year < 2004 And Typ = "��������� ���������" Then Line = 25
                    If source = 1 And Year > 1997 And Year < 2004 And Typ = "��������� ������������" Then Line = 25
                        If source = 1 And Year > 2003 And Typ = "���������" And Direction = "������" Then Line = 26
                        If source = 1 And Year > 2003 And Typ = "���������" And Direction = "�������" Then Line = 27
                        If source = 1 And Year > 2003 And Typ = "�������" And Direction = "������" Then Line = 28
                        If source = 1 And Year > 2003 And Typ = "�������" And Direction = "�������" Then Line = 29
                        If source = 1 And Year > 2003 And Typ = "���������" And Direction = "������" Then Line = 30
                        If source = 1 And Year > 2003 And Typ = "���������" And Direction = "�������" Then Line = 31
                        If source = 1 And Year > 2003 And Typ = "��������� ������������" Then Line = 32
                        If source = 1 And Year > 2003 And Typ = "��������� ���������" Then Line = 33
                            If source <> 1 And Year <= 1989 And Typ = "���������" And Direction = "������" Then Line = 4 + ((source - 1) * 30)
                            If source <> 1 And Year <= 1989 And Typ = "���������" And Direction = "�������" Then Line = 5 + ((source - 1) * 30)
                            If source <> 1 And Year <= 1989 And Typ = "�������" And Direction = "������" Then Line = 6 + ((source - 1) * 30)
                            If source <> 1 And Year <= 1989 And Typ = "�������" And Direction = "�������" Then Line = 7 + ((source - 1) * 30)
                            If source <> 1 And Year <= 1989 And Typ = "���������" And Direction = "������" Then Line = 8 + ((source - 1) * 30)
                            If source <> 1 And Year <= 1989 And Typ = "���������" And Direction = "�������" Then Line = 9 + ((source - 1) * 30)
                            If source <> 1 And Year <= 1989 And Typ = "��������� ���������" Then Line = 10 + ((source - 1) * 30)
                            If source <> 1 And Year <= 1989 And Typ = "��������� ������������" Then Line = 10 + ((source - 1) * 30)
                                If source <> 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "������" Then Line = 11 + ((source - 1) * 30)
                                If source <> 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "�������" Then Line = 12 + ((source - 1) * 30)
                                If source <> 1 And Year > 1989 And Year < 1998 And Typ = "�������" And Direction = "������" Then Line = 13 + ((source - 1) * 30)
                                If source <> 1 And Year > 1989 And Year < 1998 And Typ = "�������" And Direction = "�������" Then Line = 14 + ((source - 1) * 30)
                                If source <> 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "������" Then Line = 15 + ((source - 1) * 30)
                                If source <> 1 And Year > 1989 And Year < 1998 And Typ = "���������" And Direction = "�������" Then Line = 16 + ((source - 1) * 30)
                                If source <> 1 And Year > 1989 And Year < 1998 And Typ = "��������� ������������" Then Line = 17 + ((source - 1) * 30)
                                If source <> 1 And Year > 1989 And Year < 1998 And Typ = "��������� ���������" Then Line = 18 + ((source - 1) * 30)
                                    If source <> 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "������" Then Line = 19 + ((source - 1) * 30)
                                    If source <> 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "�������" Then Line = 20 + ((source - 1) * 30)
                                    If source <> 1 And Year > 1997 And Year < 2004 And Typ = "�������" And Direction = "������" Then Line = 21 + ((source - 1) * 30)
                                    If source <> 1 And Year > 1997 And Year < 2004 And Typ = "�������" And Direction = "�������" Then Line = 22 + ((source - 1) * 30)
                                    If source <> 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "������" Then Line = 23 + ((source - 1) * 30)
                                    If source <> 1 And Year > 1997 And Year < 2004 And Typ = "���������" And Direction = "�������" Then Line = 24 + ((source - 1) * 30)
                                    If source <> 1 And Year > 1997 And Year < 2004 And Typ = "��������� ���������" Then Line = 25 + ((source - 1) * 30)
                                    If source <> 1 And Year > 1997 And Year < 2004 And Typ = "��������� ������������" Then Line = 25 + ((source - 1) * 30)
                                        If source <> 1 And Year > 2003 And Typ = "���������" And Direction = "������" Then Line = 26 + ((source - 1) * 30)
                                        If source <> 1 And Year > 2003 And Typ = "���������" And Direction = "�������" Then Line = 27 + ((source - 1) * 30)
                                        If source <> 1 And Year > 2003 And Typ = "�������" And Direction = "������" Then Line = 28 + ((source - 1) * 30)
                                        If source <> 1 And Year > 2003 And Typ = "�������" And Direction = "�������" Then Line = 29 + ((source - 1) * 30)
                                        If source <> 1 And Year > 2003 And Typ = "���������" And Direction = "������" Then Line = 30 + ((source - 1) * 30)
                                        If source <> 1 And Year > 2003 And Typ = "���������" And Direction = "�������" Then Line = 31 + ((source - 1) * 30)
                                        If source <> 1 And Year > 2003 And Typ = "��������� ������������" Then Line = 32 + ((source - 1) * 30)
                                        If source <> 1 And Year > 2003 And Typ = "��������� ���������" Then Line = 33 + ((source - 1) * 30)
                                            If Chart = "������ 1" And Regime = "���" Then Column = 6
                                            If Chart = "������ 2" And Regime = "���" Then Column = 9
                                            If Chart = "������ 3" And Regime = "���" Then Column = 12
                                            If Chart = "������ 4" And Regime = "���" Then Column = 15
                                            If Chart = "������ 5" And Regime = "���" Then Column = 18
                                            If Chart = "������ 6" And Regime = "���" Then Column = 21
                                            If Chart = "������ 7" And Regime = "���" Then Column = 24
                                            If Chart = "������ 8" And Regime = "���" Then Column = 27
                                            If Chart = "������ 9" And Regime = "���" Then Column = 30
                                            If Chart = "������ 10" And Regime = "���" Then Column = 33
                                                If Chart = "������ 1" And Regime = "��" Then Column = 7
                                                If Chart = "������ 2" And Regime = "��" Then Column = 10
                                                If Chart = "������ 3" And Regime = "��" Then Column = 13
                                                If Chart = "������ 4" And Regime = "��" Then Column = 16
                                                If Chart = "������ 5" And Regime = "��" Then Column = 19
                                                If Chart = "������ 6" And Regime = "��" Then Column = 22
                                                If Chart = "������ 7" And Regime = "��" Then Column = 25
                                                If Chart = "������ 8" And Regime = "��" Then Column = 28
                                                If Chart = "������ 9" And Regime = "��" Then Column = 31
                                                If Chart = "������ 10" And Regime = "��" Then Column = 34
                                                    If Chart = "������ 1" And Regime = "����" Then Column = 8
                                                    If Chart = "������ 2" And Regime = "����" Then Column = 11
                                                    If Chart = "������ 3" And Regime = "����" Then Column = 14
                                                    If Chart = "������ 4" And Regime = "����" Then Column = 17
                                                    If Chart = "������ 5" And Regime = "����" Then Column = 20
                                                    If Chart = "������ 6" And Regime = "����" Then Column = 23
                                                    If Chart = "������ 7" And Regime = "����" Then Column = 26
                                                    If Chart = "������ 8" And Regime = "����" Then Column = 29
                                                    If Chart = "������ 9" And Regime = "����" Then Column = 32
                                                    If Chart = "������ 10" And Regime = "����" Then Column = 35
                            If Regime = "End" Then
                                Exit For
                            Else: Sheets("coefficient").Cells(Line, Column).Value = Sheets("total").Cells(i, 11)
                                ����20.Cells(Line, Column).Interior.ColorIndex = 3 'RGB(204, 255, 204)
                            End If
    Next i
End Sub

'�������� ������ �������� ����� (���������� ����� � �������� ����)

Sub bookTableFormatting()
    '���������� � ������������ �������������� ������� chart
    Call recalculationBook
    ����08.Range("A1:BB283").Interior.Color = xlNone
    ����08.Range("B1:U2").Interior.ColorIndex = 35
    ����08.Range("B6:U77").Interior.ColorIndex = 35
    '���������� � ������������ �������������� ������� temperature
    �����������.Range("A1:CC283").Interior.Color = xlNone
    �����������.Range("B7:U18").Interior.ColorIndex = 35
    �����������.Range("AF7:AG18").Interior.ColorIndex = 35
    �����������.Range("AF22:AF23").Interior.ColorIndex = 35
    �����������.Range("AH37:AH39").Interior.ColorIndex = 35
    �����������.Range("BB22:BK22").Interior.ColorIndex = 35
    '���������� � ������������ �������������� ������� isptemp
    Set accountRangeIsptemp = Sheets("isptemp").Range("B3:B1000")                               '�������� ����� �������������� �����
    numberRings = Application.WorksheetFunction.CountA(accountRangeIsptemp)                     '���������� �������������� ����
    ����5.Range("A1:BB1000").Interior.Color = xlNone
        ����5.Range(Sheets("isptemp").Cells(4, 2), Sheets("isptemp").Cells(numberRings + 4, 2)).Interior.ColorIndex = 35
        ����5.Range(Sheets("isptemp").Cells(4, 5), Sheets("isptemp").Cells(numberRings + 4, 5)).Interior.ColorIndex = 35
        ����5.Range(Sheets("isptemp").Cells(4, 10), Sheets("isptemp").Cells(numberRings + 4, 14)).Interior.ColorIndex = 35
    Range(Sheets("isptemp").Cells(numberRings + 4, 2), Sheets("isptemp").Cells(1000, 2)).ClearContents
    Range(Sheets("isptemp").Cells(numberRings + 4, 5), Sheets("isptemp").Cells(1000, 5)).ClearContents
    Range(Sheets("isptemp").Cells(numberRings + 4, 10), Sheets("isptemp").Cells(1000, 14)).ClearContents
    '���������� � ������������ �������������� ������� jumper
    ����14.Range("A1:BB1000").Interior.Color = xlNone
        For i = 3 To numberRings + 2 Step 1
            ����14.Cells(i, 3).Interior.ColorIndex = 35
            ����14.Cells(i, 4).Interior.ColorIndex = 35
            ����14.Cells(i, 9).Interior.ColorIndex = 35
        Next i
    Range(Sheets("jumper").Cells(numberRings + 3, 3), Sheets("jumper").Cells(1000, 3)).ClearContents
    Range(Sheets("jumper").Cells(numberRings + 3, 4), Sheets("jumper").Cells(1000, 4)).ClearContents
    Range(Sheets("jumper").Cells(numberRings + 3, 9), Sheets("jumper").Cells(1000, 9)).ClearContents
        '���������� � ������������ �������������� ������� Ratio
    Call numberOperatingModes
    numberModes = numberOperatingModes 'Sheets("ratio").Cells(28, 5) '- ���������� �������
        ����11.Range("A1:BB283").Interior.Color = xlNone '������� ��������� �� ������� �� ������� Ratio
        For i = 6 To numberModes * 2 + 4 Step 2
            For j = 3 To 5 Step 1
                ����11.Cells(j, i).Interior.ColorIndex = 35
            Next j
        Next i
        For Line = 7 To 24 Step 1 '��������� ����� ������������ �������������� ����� 20 % �� ������� Ratio
            For Column = 7 To numberModes * 2 + 5 Step 2
                more20 = Sheets("ratio").Cells(Line, Column)
                If more20 > 20 Then
                    ����11.Cells(Line, Column).Interior.ColorIndex = 4 'https://docs.microsoft.com/ru-ru/office/vba/api/excel.colorindex - ���� �������
                    'Else: Exit For
                End If
            Next Column
        Next Line
End Sub

'���������� ������� Ratio
Sub fillingRatioTab()
    Range(Sheets("ratio").Cells(3, 8), Sheets("ratio").Cells(6, 350)).ClearContents '�������� ��������� �� �����������
    For i = 7 To 20000 Step 1
        If i = 7 Then
            For p = 7 To 20000 Step 1
                chartCalculation = Sheets("calculation").Cells(p, 11)
                sourceCalculation = Sheets("calculation").Cells(p, 12)
                regimeCalculation = Sheets("calculation").Cells(p, 14)
                If chartCalculation <> Empty Or sourceCalculation <> Empty Or regimeCalculation <> Empty Then
                    Exit For
                End If
            Next p
        End If
        If i = 7 And chartCalculation <> Empty And sourceCalculation <> Empty And regimeCalculation <> Empty Then
            Sheets("ratio").Cells(3, 6).Value2 = sourceCalculation
            Sheets("ratio").Cells(4, 6).Value2 = chartCalculation
            Sheets("ratio").Cells(5, 6).Value2 = regimeCalculation
        End If
            chartCalculation1 = Sheets("calculation").Cells(i, 11)
            sourceCalculation1 = Sheets("calculation").Cells(i, 12)
            regimeCalculation1 = Sheets("calculation").Cells(i, 14)
        If chartCalculation1 <> Empty And sourceCalculation1 <> Empty And regimeCalculation1 <> Empty Then
            For j = 6 To 28 Step 2
                sourceRatio = Sheets("ratio").Cells(3, j)
                ChartRatio = Sheets("ratio").Cells(4, j)
                RegimeRatio = Sheets("ratio").Cells(5, j)

'                If sourceCalculation1 <> sourceRatio And chartCalculation1 <> ChartRatio And regimeCalculation1 <> RegimeRatio _
'                And sourceRatio <> Empty And ChartRatio <> Empty And RegimeRatio <> Empty Then
                If sourceCalculation1 = sourceRatio And chartCalculation1 = ChartRatio And regimeCalculation1 = RegimeRatio Then
                    Exit For
                Else
                    For ij = 6 To 28 Step 2
                        sourceRatio1 = Sheets("ratio").Cells(3, ij)
                        ChartRatio1 = Sheets("ratio").Cells(4, ij)
                        RegimeRatio1 = Sheets("ratio").Cells(5, ij)
                        
                        If sourceRatio1 = sourceRatio And ChartRatio1 = ChartRatio And RegimeRatio1 = RegimeRatio Then
                            Exit For
                        Else
                        For b = 6 To 28 Step 2
                            sourceRatio1 = Sheets("ratio").Cells(3, b)
                            
                            If sourceRatio1 = Empty Then
                                Sheets("ratio").Cells(3, b).Value2 = sourceCalculation1
                                Sheets("ratio").Cells(4, b).Value2 = chartCalculation1
                                Sheets("ratio").Cells(5, b).Value2 = regimeCalculation1
                                Exit For
                            End If
                        Next b
                        End If
                    Next ij
                End If
                'Exit For
            Next j
        End If
    Next i
End Sub

'����������� ���������� ����� ��� ��������� �� ������������
Sub temperatureClimatology()
Dim Month As String
Dim monthClimatologyAir As Double
Dim monthClimatologySoil As Double
Set accountRange = Sheets("isptemp").Range("A4:A1000")                              '�������� ����� �������������� �����
    numberRings = Application.WorksheetFunction.Max(accountRange)                   '���������� �������������� ����
    For j = 4 To numberRings + 3 Step 1
        Month = Sheets("isptemp").Cells(j, 12)
        If Month = "������" Then i = 7
        If Month = "�������" Then i = 8
        If Month = "����" Then i = 9
        If Month = "������" Then i = 10
        If Month = "���" Then i = 11
        If Month = "����" Then i = 12
        If Month = "����" Then i = 13
        If Month = "������" Then i = 14
        If Month = "��������" Then i = 15
        If Month = "�������" Then i = 16
        If Month = "������" Then i = 17
        If Month = "�������" Then i = 18
        monthClimatologyAir = Sheets("temperature").Cells(i, 33)
        monthClimatologySoil = Sheets("temperature").Cells(i, 32)
        Sheets("isptemp").Cells(j, 13).Value2 = monthClimatologyAir
        Sheets("isptemp").Cells(j, 14).Value2 = monthClimatologySoil
    Next j
End Sub
'�� �� ������� ��� �� excel �� �����
'Sub DeleteAllTextBox()
'    Dim oSh As Shape
'    For Each oSh In ActiveSheet.Shapes
'        oSh.Delete
'    Next oSh
'End Sub
'������������ ������� �������� ������������� �� �������� ������ �� ������� table_5
Sub table_5Calculation()
Dim accountRange As Range
Dim numberRings As Long
Dim Typ As String
    For j = 5 To 1500 Step 1
        endCalculation = Sheets("table_4").Cells(j, 2)
        If endCalculation = "End" Then
            Exit For
        Else: ����13.Cells(j - 1, 1).FormulaLocal = "=����(table_4!A" + CStr(j) + "="""";"""";table_4!A" + CStr(j) + ")"
              ����13.Cells(j - 1, 2).FormulaLocal = "=����(table_4!B" + CStr(j) + "="""";"""";table_4!B" + CStr(j) + ")"
              ����13.Cells(j - 1, 3).FormulaLocal = "=����(table_4!C" + CStr(j) + "="""";"""";table_4!C" + CStr(j) + ")"
              ����13.Cells(j - 1, 4).FormulaLocal = "=����(table_4!D" + CStr(j) + "="""";"""";table_4!D" + CStr(j) + ")"
              ����13.Cells(j - 1, 5).FormulaLocal = "=����(table_4!E" + CStr(j) + "="""";"""";table_4!E" + CStr(j) + ")"
              ����13.Cells(j - 1, 6).FormulaLocal = "=����(table_4!F" + CStr(j) + "="""";"""";table_4!F" + CStr(j) + ")"
              ����13.Cells(j - 1, 7).FormulaLocal = "=����(table_4!G" + CStr(j) + "="""";"""";table_4!G" + CStr(j) + ")"
        '����18.Range(Cells(5, 1), Cells(j, 11)).Interior.ColorIndex = 35
        End If
    Next j
Set accountRange = Sheets("table_5").Range("A4:A1000")                               '�������� ����� �������������� �����
    numberRings = Application.WorksheetFunction.CountA(accountRange)                 '���������� �������������� ����
    For i = 4 To numberRings + 3 Step 1
        Typ = Sheets("table_5").Cells(i, 5)
        If Typ = "��������� ���������" Or Typ = "��������� ������������" Then
            Range(Sheets("table_5").Cells(i, 8), Sheets("table_5").Cells(i + 1, 8)).Merge
            Range(Sheets("table_5").Cells(i, 9), Sheets("table_5").Cells(i + 1, 9)).Merge
            Range(Sheets("table_5").Cells(i, 10), Sheets("table_5").Cells(i + 1, 10)).Merge
            ����13.Cells(i, 8).FormulaLocal = "=Qpodz(table_4!M" + CStr(i + 1) + ";table_4!M" + CStr(i + 2) + ";table_4!J" + CStr(i + 1) + ";table_4!K" + CStr(i + 1) + ";table_4!J" + CStr(i + 2) + ";table_4!K" + CStr(i + 2) + ";A" + CStr(i) + ";B" + CStr(i) + ";E" + CStr(i) + ")"
            ����13.Cells(i, 9).FormulaLocal = "=SummaQpodzem(A" + CStr(i) + ";B" + CStr(i) + ")"
            ����13.Cells(i, 10).FormulaLocal = "=H" + CStr(i) + "/I" + CStr(i) + ""
            i = i + 1
        End If
        If Typ = "���������" Then
            ����13.Cells(i, 8).FormulaLocal = "=Qnadpod(table_4!M" + CStr(i + 1) + ";table_4!J" + CStr(i + 1) + ";table_4!K" + CStr(i + 1) + ";A" + CStr(i) + ";B" + CStr(i) + ";E" + CStr(i) + ")"
            ����13.Cells(i, 9).FormulaLocal = "=SummaQnadzem(A" + CStr(i) + ";B" + CStr(i) + ";D" + CStr(i) + ")"
            ����13.Cells(i, 10).FormulaLocal = "=H" + CStr(i) + "/I" + CStr(i) + ""
        End If
    Next i
End Sub

'������������ ������� �������� ������������� �� �������� ������ �� ������� table_4
Sub table_4Calculation()
Dim accountRange As Range
Dim numberRings As Long
Dim Typ As String
Set accountRange = Sheets("table_4").Range("A5:A1000")                               '�������� ����� �������������� �����
    numberRings = Application.WorksheetFunction.CountA(accountRange)                 '���������� �������������� ����
    For i = 5 To numberRings + 4 Step 1
        ����18.Cells(i, 12).FormulaLocal = "=����(E" + CStr(i) + "="""";"""";tokrsre(A" + CStr(i) + ";E" + CStr(i) + "))"
        ����18.Cells(i, 13).FormulaLocal = "=����(J" + CStr(i) + "="""";"""";����(D" + CStr(i) + "=""������"";(H" + CStr(i) + "-I" + CStr(i) + "/4)*(J" + CStr(i) + "-K" + CStr(i) + ")*1000;(H" + CStr(i) + "-3*I" + CStr(i) + "/4)*(J" + CStr(i) + "-K" + CStr(i) + ")*1000))"
    Next i
    For j = 5 To 1500 Step 1
        endCalculation = Sheets("table_4").Cells(j, 2)
        If endCalculation = "End" Then
            Exit For
        Else: ����18.Range(Cells(5, 1), Cells(j, 11)).Interior.ColorIndex = 35
        End If
    Next j
End Sub

' ������ ���������� ��������� � ����������� ������ ��������������� ������
Sub podborTisp()
c = Range("C1")
d = Range("D1")
e = Range("E1")
F = Range("F1")
x = Range("G1")
h = Range("H1")
i = Range("I1")
j = Range("j1")
k = Range("K1")
l = Range("L1")
m = Range("M1")
    If c <> 0 Then Range("C2").GoalSeek Goal:=c, ChangingCell:=Range("C3") Else: Range("C3") = 0
    If d <> 0 Then Range("D2").GoalSeek Goal:=d, ChangingCell:=Range("D3") Else: Range("D3") = 0
    If e <> 0 Then Range("E2").GoalSeek Goal:=e, ChangingCell:=Range("E3") Else: Range("E3") = 0
    If F <> 0 Then Range("F2").GoalSeek Goal:=F, ChangingCell:=Range("F3") Else: Range("F3") = 0
    If x <> 0 Then Range("G2").GoalSeek Goal:=x, ChangingCell:=Range("G3") Else: Range("G3") = 0
    If h <> 0 Then Range("H2").GoalSeek Goal:=h, ChangingCell:=Range("H3") Else: Range("H3") = 0
    If i <> 0 Then Range("I2").GoalSeek Goal:=i, ChangingCell:=Range("I3") Else: Range("I3") = 0
    If j <> 0 Then Range("J2").GoalSeek Goal:=j, ChangingCell:=Range("J3") Else: Range("J3") = 0
    If k <> 0 Then Range("K2").GoalSeek Goal:=k, ChangingCell:=Range("K3") Else: Range("K3") = 0
    If l <> 0 Then Range("L2").GoalSeek Goal:=l, ChangingCell:=Range("L3") Else: Range("L3") = 0
    If m <> 0 Then Range("M2").GoalSeek Goal:=m, ChangingCell:=Range("M3") Else: Range("M3") = 0
    Call recalculationBook
End Sub

'���������� ���������� � ����������� ������ �� ������� ring
Sub fillChartTemperatures()
    Range(Cells(1, 3), Cells(1, 13)).ClearContents
    numberRings = Cells(1, 1)
    numberPlots = Cells(1, 2)
    For i = 5 To 1000 Step 1
        numberRingsTable_4 = Sheets("table_4").Cells(i, 1)
        Direction = Sheets("table_4").Cells(i, 4)
        If Direction = "������" And numberRingsTable_4 = numberRings Then ' ���������� ����� ������
            For j = i To i + (numberPlots * 2 - 1) Step 2
                phaseNumberSupply = Sheets("table_4").Cells(j, 2)
                Direction = Sheets("table_4").Cells(j, 4)
                If phaseNumberSupply = 1 Then
                    Cells(1, 3).Value2 = Sheets("table_4").Cells(j, 10).Value2
                    plotReturn1 = Sheets("table_4").Cells(j + 1, 11).Value2
                ElseIf phaseNumberSupply = 2 Then
                    Cells(1, 4).Value2 = Sheets("table_4").Cells(j, 10).Value2
                    plotReturn2 = Sheets("table_4").Cells(j + 1, 11).Value2
                ElseIf phaseNumberSupply = 3 Then
                    Cells(1, 5).Value2 = Sheets("table_4").Cells(j, 10).Value2
                    plotReturn3 = Sheets("table_4").Cells(j + 1, 11).Value2
                ElseIf phaseNumberSupply = 4 Then
                    Cells(1, 6).Value2 = Sheets("table_4").Cells(j, 10).Value2
                    plotReturn4 = Sheets("table_4").Cells(j + 1, 11).Value2
                ElseIf phaseNumberSupply = 5 Then
                    Cells(1, 7).Value2 = Sheets("table_4").Cells(j, 10).Value2
                    plotReturn5 = Sheets("table_4").Cells(j + 1, 11).Value2
                End If
            Next j
        End If
            If phaseNumberSupply <> Empty And numberRingsTable_4 = numberRings Then
                Cells(1, 3 + phaseNumberSupply).Value2 = Sheets("table_4").Cells(j - 1, 10).Value2
            End If
        If numberPlots = 5 And phaseNumberSupply <> Empty Then
            Cells(1, 4 + phaseNumberSupply).Value2 = plotReturn5
            Cells(1, 5 + phaseNumberSupply).Value2 = plotReturn4
            Cells(1, 6 + phaseNumberSupply).Value2 = plotReturn3
            Cells(1, 7 + phaseNumberSupply).Value2 = plotReturn2
            Cells(1, 8 + phaseNumberSupply).Value2 = plotReturn1
        Exit For
        ElseIf numberPlots = 4 And phaseNumberSupply <> Empty Then
            Cells(1, 4 + phaseNumberSupply).Value2 = plotReturn4
            Cells(1, 5 + phaseNumberSupply).Value2 = plotReturn3
            Cells(1, 6 + phaseNumberSupply).Value2 = plotReturn2
            Cells(1, 7 + phaseNumberSupply).Value2 = plotReturn1
        Exit For
        ElseIf numberPlots = 3 And phaseNumberSupply <> Empty Then
            Cells(1, 4 + phaseNumberSupply).Value2 = plotReturn3
            Cells(1, 5 + phaseNumberSupply).Value2 = plotReturn2
            Cells(1, 6 + phaseNumberSupply).Value2 = plotReturn1
        Exit For
        ElseIf numberPlots = 2 And phaseNumberSupply <> Empty Then
            Cells(1, 4 + phaseNumberSupply).Value2 = plotReturn2
            Cells(1, 5 + phaseNumberSupply).Value2 = plotReturn1
        Exit For
        ElseIf numberPlots = 1 And phaseNumberSupply <> Empty Then
            Cells(1, 4 + phaseNumberSupply).Value2 = plotReturn1
        Exit For
        End If
    Next i
End Sub

'�������� �������� ������
Sub verificationSourseData()
    If Regime = "" Then
        MsgBox "�� ������ ����� �� ������� isptemp ��� ��������������� ������ � " & ring
        Exit Sub
    ElseIf Chart = "" Then
        MsgBox "�� ������ ������������� ������ �� ������� isptemp ��� ��������������� ������ � " & ring
        Exit Sub
    End If
End Sub

'�������� �������� �����
Sub recalculationBook()
    Dim wks As Worksheet
    Application.Calculation = xlManual
    For Each wks In ActiveWorkbook.Worksheets
        wks.Calculate
    Next
    Set wks = Nothing
End Sub

'������� �������� �� ������� Ratio �� ������� PSV
Function carryoverRatioPSV(sourceRatio, ChartRatio, RegimeRatio, numberModes)
    For l = 6 To 5 + numberModes Step 1
        sourcePSV = Sheets("PSV").Cells(3, l)
        If sourcePSV = Empty Then
            Sheets("PSV").Cells(3, l).Value = sourceRatio
            Sheets("PSV").Cells(4, l).Value = ChartRatio
            Sheets("PSV").Cells(6, l).Value = RegimeRatio
            Exit Function
        End If
    Next l
End Function
'���������� ������� ������ �� ������� ratio
Function numberOperatingModes()
    k = Sheets("ratio").Cells(25, 7)
    For i = 7 To 250 Step 2
        If Sheets("ratio").Cells(25, i) <> 100 Then
            Exit For
        Else
            p = p + Sheets("ratio").Cells(25, i)
        End If
    Next i
    If p = Empty Then
        numberOperatingModes = k / 100
    Else
        numberOperatingModes = p / 100
    End If
    'MsgBox numberOperatingModes
End Function

'�������������� ������ ������� �� �����
Function Num2ABC(ByVal x As Long) As String
x = x - 1
Do
Num2ABC = Chr$(65 + x Mod 26) & Num2ABC
x = x \ 26 - 1
Loop While x >= 0
End Function

'������������ ������ �� ������� PSV
Sub calculationFormationPSV()
    Application.ScreenUpdating = False                              '��������� ���������� ������ ��� ���������
    Columns.Hidden = False   '�������� ��� ������� ����� � ��������
    Rows.Hidden = False

    '������������ ������ �� ������� PSV
    Range(Sheets("PSV").Cells(3, 6), Sheets("PSV").Cells(59, 17)).ClearContents '�������� ��������� �� �����������
    Call numberOperatingModes
    numberModes = numberOperatingModes 'Sheets("ratio").Cells(28, 5) '- ���������� �������
    ����21.Range("A1:AI263").Interior.Color = xlNone '�������� ��������� �� �������
    For c = 6 To 5 + numberModes * 2 Step 2
        sourceRatio = Sheets("ratio").Cells(3, c) '������� ���������
        ChartRatio = Sheets("ratio").Cells(4, c) '������� �������
        RegimeRatio = Sheets("ratio").Cells(5, c) '������� ������
        Call carryoverRatioPSV(sourceRatio, ChartRatio, RegimeRatio, numberModes)
    Next c
    For i = 6 To 5 + numberModes Step 1
        source = Sheets("PSV").Cells(3, i)
        Chart = Sheets("PSV").Cells(4, i) 'workingHours = Sheets("PSV").Cells(5, i)
        Regime = Sheets("PSV").Cells(6, i)
        If source <> Empty And Chart <> Empty And Regime <> Empty Then
            For ascii = 64 + i To 69 + numberModes
                j = Num2ABC(i)
                ����21.Cells(5, ascii - 64).FormulaLocal = "=periodWork(" + j + "4;" + j + "6)"
                ����21.Cells(7, ascii - 64).FormulaLocal = "=length(" + j + "3;" + j + "4;" + j + "5)"
                ����21.Cells(8, ascii - 64).FormulaLocal = "=MX(" + j + "3;" + j + "4;" + j + "5)"
                ����21.Cells(9, ascii - 64).FormulaLocal = "=����(" + j + "7=0; 0; " + j + "8/" + j + "7)"
                ����21.Cells(10, ascii - 64).FormulaLocal = "=volume(" + j + "3;" + j + "4;" + j + "6)"
                ����21.Cells(11, ascii - 64).FormulaLocal = "=" + j + "5"
                ����21.Cells(12, ascii - 64).FormulaLocal = "=0,0025*" + j + "10*" + j + "11"
                ����21.Cells(13, ascii - 64).FormulaLocal = "=1,5*" + j + "10"
                ����21.Cells(14, ascii - 64).FormulaLocal = "=0,5*" + j + "10"
                ����21.Cells(15, ascii - 64).FormulaLocal = "=PSVSARZ(" + j + "$3;" + j + "$4;$E15;$D15)"
                ����21.Cells(16, ascii - 64).FormulaLocal = "=PSVSARZ(" + j + "$3;" + j + "$4;$E16;$D15)"
                ����21.Cells(17, ascii - 64).FormulaLocal = "=PSVSARZ(" + j + "$3;" + j + "$4;$E17;$D15)"
                ����21.Cells(18, ascii - 64).FormulaLocal = "=PSVSARZ(" + j + "$3;" + j + "$4;$E18;$D18)"
                ����21.Cells(19, ascii - 64).FormulaLocal = "=PSVSARZ(" + j + "$3;" + j + "$4;$E19;$D18)"
                ����21.Cells(20, ascii - 64).FormulaLocal = "=PSVSARZ(" + j + "$3;" + j + "$4;$E20;$D18)"
                ����21.Cells(21, ascii - 64).FormulaLocal = "=����(����(" + j + "15:" + j + "20)=0;""-"";����(" + j + "15:" + j + "20))"
                ����21.Cells(22, ascii - 64).FormulaLocal = "=����(" + j + "12:" + j + "14;" + j + "21)"
                ����21.Cells(23, ascii - 64).FormulaLocal = "=tyn(" + j + "4;" + j + "6)"
                ����21.Cells(24, ascii - 64).FormulaLocal = "=tzp(" + j + "4;" + j + "6)"
                ����21.Cells(25, ascii - 64).FormulaLocal = "=tisp(" + j + "4;" + j + "6)"
                ����21.Cells(26, ascii - 64).FormulaLocal = "=temperaturePSV(" + j + "$3;" + j + "$4;$E26;$D26;F$6)"
                ����21.Cells(27, ascii - 64).FormulaLocal = "=temperaturePSV(" + j + "$3;" + j + "$4;$E27;$D26;F$6)"
                ����21.Cells(28, ascii - 64).FormulaLocal = "=temperaturePSV(" + j + "$3;" + j + "$4;$E28;$D26;F$6)"
                ����21.Cells(29, ascii - 64).FormulaLocal = "=temperaturePSV(" + j + "$3;" + j + "$4;$E29;$D29;F$6)"
                ����21.Cells(30, ascii - 64).FormulaLocal = "=temperaturePSV(" + j + "$3;" + j + "$4;$E30;$D29;F$6)"
                ����21.Cells(31, ascii - 64).FormulaLocal = "=temperaturePSV(" + j + "$3;" + j + "$4;$E31;$D29;F$6)"
                ����21.Cells(32, ascii - 64).FormulaLocal = "=coolantDensity(" + j + "23)"
                ����21.Cells(33, ascii - 64).FormulaLocal = "=coolantDensity(" + j + "24)"
                ����21.Cells(34, ascii - 64).FormulaLocal = "=coolantDensity(" + j + "25)"
                ����21.Cells(35, ascii - 64).FormulaLocal = "=coolantDensity(" + j + "26)"
                ����21.Cells(36, ascii - 64).FormulaLocal = "=coolantDensity(" + j + "27)"
                ����21.Cells(37, ascii - 64).FormulaLocal = "=coolantDensity(" + j + "28)"
                ����21.Cells(38, ascii - 64).FormulaLocal = "=coolantDensity(" + j + "29)"
                ����21.Cells(39, ascii - 64).FormulaLocal = "=coolantDensity(" + j + "30)"
                ����21.Cells(40, ascii - 64).FormulaLocal = "=coolantDensity(" + j + "31)"
                ����21.Cells(41, ascii - 64).FormulaLocal = "=heatLossPSV(" + j + "12;" + j + "23;" + j + "32)"
                ����21.Cells(42, ascii - 64).FormulaLocal = "=heatLossPSV(" + j + "13;" + j + "24;" + j + "33)"
                ����21.Cells(43, ascii - 64).FormulaLocal = "=heatLossPSV(" + j + "14;" + j + "25;" + j + "34)"
                ����21.Cells(44, ascii - 64).FormulaLocal = "=heatLossPSV(" + j + "15;" + j + "26;" + j + "35)"
                ����21.Cells(45, ascii - 64).FormulaLocal = "=heatLossPSV(" + j + "16;" + j + "27;" + j + "36)"
                ����21.Cells(46, ascii - 64).FormulaLocal = "=heatLossPSV(" + j + "17;" + j + "28;" + j + "37)"
                ����21.Cells(47, ascii - 64).FormulaLocal = "=heatLossPSV(" + j + "18;" + j + "29;" + j + "38)"
                ����21.Cells(48, ascii - 64).FormulaLocal = "=heatLossPSV(" + j + "19;" + j + "30;" + j + "39)"
                ����21.Cells(49, ascii - 64).FormulaLocal = "=heatLossPSV(" + j + "20;" + j + "31;" + j + "40)"
                ����21.Cells(50, ascii - 64).FormulaLocal = "=����(����(" + j + "44:" + j + "49)=0;""-"";����(" + j + "44:" + j + "49))"
                ����21.Cells(51, ascii - 64).FormulaLocal = "=����(" + j + "41:" + j + "43;" + j + "50)"
                ����21.Cells(52, ascii - 64).FormulaLocal = "=Qizol(" + j + "3;" + j + "4;" + j + "6;E52)"
                ����21.Cells(53, ascii - 64).FormulaLocal = "=Qizol(" + j + "3;" + j + "4;" + j + "6;E53)"
                ����21.Cells(54, ascii - 64).FormulaLocal = "=Qizol(" + j + "3;" + j + "4;" + j + "6;E54)"
                ����21.Cells(55, ascii - 64).FormulaLocal = "=Qizol(" + j + "3;" + j + "4;" + j + "6;E55)"
                ����21.Cells(56, ascii - 64).FormulaLocal = "=Qizol(" + j + "3;" + j + "4;" + j + "6;E56)"
                ����21.Cells(57, ascii - 64).FormulaLocal = "=" + j + "52+" + j + "54+" + j + "55+" + j + "56+" + j + "53"
                ����21.Cells(58, ascii - 64).FormulaLocal = "=" + j + "57*" + j + "11*10^(-6)"
                ����21.Cells(59, ascii - 64).FormulaLocal = "=" + j + "51+" + j + "58"
                Exit For
            Next
'            Call recalculationBook
            Call hidingCellPSV
            Application.ScreenUpdating = True
        Else: Exit For
        End If
    Next i
    Application.ScreenUpdating = True
End Sub
'Sub dxghd()
'k = PSVSARZ(2, "������ 3", "������", "�� ������� �� ���� � ������������ ������")
'End Sub
'������ ������� ���� � ���� �� ������� PSV
Function PSVSARZ(source, Chart, Direction, season)
    For i = 6 To 10 Step 4
        sourceRegulators = Sheets("chart").Cells(i, 32)
        chartRegulators = Sheets("chart").Cells(i, 33)
        If sourceRegulators = source And chartRegulators = Chart And season = "�� ������� �� ���� � ������������ ������" Then
            For j = 35 To 37 Step 1
                directionRegulators = Sheets("chart").Cells(5, j)
                If directionRegulators = Direction Then
                    numberRegulators = Sheets("chart").Cells(i, j)
                    leakRate = Sheets("chart").Cells(i + 1, j)
                    workingHours = Sheets("chart").Cells(i + 2, j)
                    regulatorLeakageTemperature = Sheets("chart").Cells(i + 3, j)
                    If numberRegulators <> Empty And leakRate <> Empty And workingHours <> Empty Then
                        PSVSARZ = numberRegulators * leakRate * workingHours
                        Exit Function
                        Else: PSVSARZ = " - "
                    End If
                End If
            Next j
        End If
        If sourceRegulators = source And chartRegulators = Chart And season = "�� ������� �� ���� � ������ ������" Then
            For j = 38 To 40 Step 1
                directionRegulators = Sheets("chart").Cells(5, j)
                If directionRegulators = Direction Then
                    numberRegulators = Sheets("chart").Cells(i, j)
                    leakRate = Sheets("chart").Cells(i + 1, j)
                    workingHours = Sheets("chart").Cells(i + 2, j)
                    regulatorLeakageTemperature = Sheets("chart").Cells(i + 3, j)
                    If numberRegulators <> Empty And leakRate <> Empty And workingHours <> Empty Then
                        PSVSARZ = numberRegulators * leakRate * workingHours
                        Exit Function
                        Else: PSVSARZ = " - "
                    End If
                End If
            Next j
        End If
    Next i
    If PSVSARZ = Empty Then
        PSVSARZ = " - "
    End If
End Function
'Sub er()
'k = temperaturePSV(1, "������ 1", "�������", "�� ������� �� ���� � ������������ ������")
'End Sub
'����������� ������� ���� � ���� �� ������� PSV
Function temperaturePSV(source, Chart, Direction, season, Regime)
    For i = 6 To 10 Step 4
        sourceRegulators = Sheets("chart").Cells(i, 32)
        chartRegulators = Sheets("chart").Cells(i, 33)
        If sourceRegulators = source And chartRegulators = Chart And season = "�� ������� �� ���� � ������������ ������" Then
            For j = 35 To 37 Step 1
                directionRegulators = Sheets("chart").Cells(5, j)
                If directionRegulators = Direction Then
                    temperaturePSV = Sheets("chart").Cells(i + 3, j)
                    If temperaturePSV <> "-" And temperaturePSV <> "" Then
                        temperaturePSV = temperaturePSV - coldWaterTemperature(Chart, Regime)
                    End If
                    If temperaturePSV = Empty Then
                        temperaturePSV = " - "
                    End If
                    Exit Function
                    Else: temperaturePSV = " - "
                End If
            Next j
        End If
        If sourceRegulators = source And chartRegulators = Chart And season = "�� ������� �� ���� � ������ ������" Then
            For j = 38 To 40 Step 1
                directionRegulators = Sheets("chart").Cells(5, j)
                If directionRegulators = Direction Then
                    temperaturePSV = Sheets("chart").Cells(i + 3, j)
                    If temperaturePSV <> "-" And temperaturePSV <> "" Then
                        temperaturePSV = temperaturePSV - coldWaterTemperature(Chart, Regime)
                    End If
                    If temperaturePSV = Empty Then
                        temperaturePSV = " - "
                    End If
                    Exit Function
                    Else: temperaturePSV = " - "
                End If
            Next j
        End If
    Next i
    If temperaturePSV = Empty Then
        temperaturePSV = " - "
    End If
End Function
'Sub dfghjd()
'k = coolantDensity("-")
'End Sub

'���������� �������������
Function coolantDensity(temperaturePSV)
    If temperaturePSV <> Empty And temperaturePSV > 0 And temperaturePSV <> " - " And temperaturePSV <> "-" Then
        coolantDensity = wskDSWT(temperaturePSV)
    Else: coolantDensity = " - "
    End If
End Function

'������ ����� � �������������
Function heatLossPSV(PSVSARZ, temperaturePSV1, coolantDensity1)
    If PSVSARZ <> Empty And PSVSARZ > 0 And PSVSARZ <> " - " And PSVSARZ <> "-" Then
        heatLossPSV = PSVSARZ * temperaturePSV1 * coolantDensity1 / 1000000
    Else: heatLossPSV = " - "
    End If
End Function

'����������� �������� ����
Function coldWaterTemperature(Chart, Regime)
    If Chart = "������ 1" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 54)
        ElseIf Chart = "������ 2" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 55)
        ElseIf Chart = "������ 3" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 56)
        ElseIf Chart = "������ 4" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 57)
        ElseIf Chart = "������ 5" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 58)
        ElseIf Chart = "������ 6" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 59)
        ElseIf Chart = "������ 7" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 60)
        ElseIf Chart = "������ 8" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 61)
        ElseIf Chart = "������ 9" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 62)
        ElseIf Chart = "������ 10" And Regime = "���" Then
        tx = Sheets("temperature").Cells(19, 63)
        ElseIf Chart = "������ 1" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 54)
        ElseIf Chart = "������ 2" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 55)
        ElseIf Chart = "������ 3" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 56)
        ElseIf Chart = "������ 4" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 57)
        ElseIf Chart = "������ 5" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 58)
        ElseIf Chart = "������ 6" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 59)
        ElseIf Chart = "������ 7" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 60)
        ElseIf Chart = "������ 8" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 61)
        ElseIf Chart = "������ 9" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 62)
        ElseIf Chart = "������ 10" And Regime = "��" Then
        tx = Sheets("temperature").Cells(20, 63)
        ElseIf Chart = "������ 1" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 54)
        ElseIf Chart = "������ 2" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 55)
        ElseIf Chart = "������ 3" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 56)
        ElseIf Chart = "������ 4" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 57)
        ElseIf Chart = "������ 5" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 58)
        ElseIf Chart = "������ 6" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 59)
        ElseIf Chart = "������ 7" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 60)
        ElseIf Chart = "������ 8" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 61)
        ElseIf Chart = "������ 9" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 62)
        ElseIf Chart = "������ 10" And Regime = "����" Then
        tx = Sheets("temperature").Cells(21, 63)
    End If
    coldWaterTemperature = tx
End Function
'������� ����� �� ������� PSV
Sub hidingCellPSV()
    'Application.ScreenUpdating = False                              '��������� ���������� ������ ��� ���������
    Call numberOperatingModes
    numberModes = numberOperatingModes 'Sheets("ratio").Cells(28, 5) '- ���������� �������
    Columns.Hidden = False   '�������� ��� ������� ����� � ��������
    Rows.Hidden = False
    For k = 7 To 59 Step 1
        For b = 6 To 5 + numberModes Step 1
                variableCell = Sheets("PSV").Cells(k, b)
                rowCell = Sheets("PSV").Cells(k, b)
            If variableCell <> " - " And variableCell <> "-" Then
                Exit For
            End If
                Rows(k).EntireRow.Hidden = True
                Exit For
        Next b
    Next k
    'Application.ScreenUpdating = True
End Sub

'������� ������� ������ �� ������� temperature �� ������� summary
Sub perranceWorkTime()
    Application.ScreenUpdating = False
    Call numberOperatingModes
    numberModes = numberOperatingModes 'Sheets("ratio").Cells(28, 5) '- ���������� �������
    Range(Sheets("summary").Cells(23, 1), Sheets("summary").Cells(442, 30)).ClearContents '�������� ��������� �� �����������
    ����22.Range("A120:AI363").Interior.Color = xlNone '�������� ��������� �� �������
    
    For c = 6 To 5 + numberModes Step 1
        ds = ds + 1
        sourcePSV = Sheets("PSV").Cells(3, c) '������� ���������
        ChartPSV = Sheets("PSV").Cells(4, c) '������� �������
        RegimePSV = Sheets("PSV").Cells(6, c) '������� ������
        WorkTime = Sheets("PSV").Cells(5, c) '������� ������� ������
        Sheets("summary").Cells(23 + 16 * (ds - 1), 1) = "��������: " & sourcePSV & ","
        Sheets("summary").Cells(24 + 16 * (ds - 1), 1) = ChartPSV & ","
        Sheets("summary").Cells(25 + 16 * (ds - 1), 1) = "���� ������: " & WorkTime & " �"
            For ij = 4 To 9 Step 1
                sourceConsumers = Sheets("summary").Cells(1, ij)
                chartConsumers = Sheets("summary").Cells(2, ij)
                consumersChart = Sheets("summary").Cells(3, ij)
                If sourcePSV = sourceConsumers And ChartPSV = chartConsumers And RegimePSV = "��" Then
                    Exit For
                End If
            Next ij
            For s = 2 To 21 Step 1
                ChartTemperature = Sheets("temperature").Cells(5, s)
                If ChartPSV = ChartTemperature Then
                    For l = 7 To 18 Step 1
                        p = (l + 19) + 16 * (ds - 1) '������ ����� �����������
                        Month1 = Sheets("temperature").Cells(l, 1)
                        heatingPeriod = Sheets("temperature").Cells(l, s)
                        summerPeriod = Sheets("temperature").Cells(l, s + 1)
                        repairPeriodMonth = Sheets("temperature").Cells(l, 68)
                        repairPeriodYear = Sheets("temperature").Cells(19, 68)
                            If RegimePSV = "��" Then summerPeriod = 0
                            If RegimePSV = "����" Then heatingPeriod = 0
                        Sheets("summary").Cells(p, 1).Value = Month1
                        Sheets("summary").Cells(p, 2).Value = heatingPeriod
                            If heatingPeriod = Empty Then heatingPeriod = 0
                        Sheets("summary").Cells(p, 3).Value = summerPeriod
                            If summerPeriod = Empty Then summerPeriod = 0
                        Sheets("summary").Cells(p, 4).Value = repairPeriodMonth
                        Sheets("summary").Cells(p, 5).Value = volume(sourcePSV, ChartPSV, RegimePSV) * 0.0025 * (heatingPeriod + summerPeriod)
                            If repairPeriodMonth > 0 Then
                                Sheets("summary").Cells(p, 6).Value = volume(sourcePSV, ChartPSV, RegimePSV) * 1.5 * (repairPeriodMonth / repairPeriodYear)
                                Sheets("summary").Cells(p, 7).Value = volume(sourcePSV, ChartPSV, RegimePSV) * 0.5 * (repairPeriodMonth / repairPeriodYear)
                            End If
                        Sheets("summary").Cells(p, 8).Value = SARZsummaryPSV(sourcePSV, ChartPSV, heatingPeriod, summerPeriod)
                        ����22.Cells(p, 9).FormulaLocal = "=����(E" + CStr(p) + ":H" + CStr(p) + ")"
                            If sourcePSV = sourceConsumers And ChartPSV = chartConsumers And RegimePSV = "��" Then
                                Sheets("summary").Cells(p, 10).Value = Sheets("summary").Cells(13, ij) * 0.0025 * (heatingPeriod + summerPeriod) ' + Sheets("summary").Cells(p, 5)
                                Sheets("summary").Cells(p, 11).Value = Sheets("summary").Cells(13, ij) * 1.5 * (repairPeriodMonth / repairPeriodYear) ' + Sheets("summary").Cells(p, 6)
                                Sheets("summary").Cells(p, 12).Value = Sheets("summary").Cells(13, ij) * 0.5 * (repairPeriodMonth / repairPeriodYear) ' + Sheets("summary").Cells(p, 7)
                                Sheets("summary").Cells(p, 13).Value = 0
                            Else
                                Sheets("summary").Cells(p, 10).Value = 0 'Sheets("summary").Cells(p, 5)
                                Sheets("summary").Cells(p, 11).Value = 0 'Sheets("summary").Cells(p, 6)
                                Sheets("summary").Cells(p, 12).Value = 0 'Sheets("summary").Cells(p, 7)
                                Sheets("summary").Cells(p, 13).Value = 0 'Sheets("summary").Cells(p, 8)
                            End If
                        ����22.Cells(p, 14).FormulaLocal = "=����(J" + CStr(p) + ":M" + CStr(p) + ")"
                        gyn = Sheets("summary").Cells(p, 5)
                            If gyn <> 0 Then
                                Sheets("summary").Cells(p, 15).Value = Qynm(gyn, ChartPSV, RegimePSV, Month1)
                            Else: Sheets("summary").Cells(p, 15).Value = 0
                            End If
                            gzp = Sheets("summary").Cells(p, 6)
                        Sheets("summary").Cells(p, 16).Value = Qzpm(gzp, ChartPSV, RegimePSV, Month1)
                            gisp = Sheets("summary").Cells(p, 7)
                        Sheets("summary").Cells(p, 17).Value = Qispm(gisp, ChartPSV, RegimePSV, Month1)
                        Sheets("summary").Cells(p, 18).Value = SARZsummaryGcal(sourcePSV, ChartPSV, RegimePSV, Month1, heatingPeriod, summerPeriod)
                                gyn = Sheets("summary").Cells(p, 10)
                                gzp = Sheets("summary").Cells(p, 11)
                                gisp = Sheets("summary").Cells(p, 12)
                            If gyn <> 0 Then
                                gyn1 = Qynm(gyn, consumersChart, "��", Month1)
                            Else: gyn1 = 0
                            End If
                            If gzp <> 0 Then
                                gzp1 = Qzpm(gzp, consumersChart, "��", Month1)
                            Else: gzp1 = 0
                            End If
                            If gisp <> 0 Then
                                gisp1 = Qispm(gisp, consumersChart, "��", Month1)
                            Else: gisp1 = 0
                            End If
                            Sheets("summary").Cells(p, 19).Value = gyn1 + gzp1 + gisp1
                        ����22.Cells(p, 20).FormulaLocal = "=����(O" + CStr(p) + ":R" + CStr(p) + ")"
                            If heatingPeriod > 0 Or summerPeriod > 0 Then
                                Sheets("summary").Cells(p, 21).Value = Qizolsr(sourcePSV, ChartPSV, RegimePSV, "��������� ���������", Month1, "������")
                                Sheets("summary").Cells(p, 22).Value = Qizolsr(sourcePSV, ChartPSV, RegimePSV, "���������", Month1, "������")
                                Sheets("summary").Cells(p, 23).Value = Qizolsr(sourcePSV, ChartPSV, RegimePSV, "���������", Month1, "�������")
                            End If
                        ����22.Cells(p, 24).FormulaLocal = "=����(U" + CStr(p) + ":W" + CStr(p) + ")"
                        ����22.Cells(p, 25).FormulaLocal = "=U" + CStr(p) + "*($B" + CStr(p) + "+$C" + CStr(p) + ")*10^(-6)"
                        ����22.Cells(p, 26).FormulaLocal = "=(V" + CStr(p) + "+W" + CStr(p) + ")*($B" + CStr(p) + "+$C" + CStr(p) + ")*10^(-6)"
                        ����22.Cells(p, 27).FormulaLocal = "=����(Y" + CStr(p) + ":Z" + CStr(p) + ")"
                        ����22.Cells(p, 28).FormulaLocal = "=T" + CStr(p) + ""
                        ����22.Cells(p, 29).FormulaLocal = "=S" + CStr(p) + ""
                        ����22.Cells(p, 30).FormulaLocal = "=AA" + CStr(p) + "+AB" + CStr(p) + ""
                    Next l
                        Sheets("summary").Cells(p + 1, 1).Value = "���"
                        ����22.Cells(p + 1, 2).FormulaLocal = "=����(B" + CStr(p - 11) + ":B" + CStr(p) + ")"
                        ����22.Cells(p + 1, 3).FormulaLocal = "=����(C" + CStr(p - 11) + ":C" + CStr(p) + ")"
                        ����22.Cells(p + 1, 4).FormulaLocal = "=����(D" + CStr(p - 11) + ":D" + CStr(p) + ")"
                        ����22.Cells(p + 1, 5).FormulaLocal = "=����(E" + CStr(p - 11) + ":E" + CStr(p) + ")"
                        ����22.Cells(p + 1, 6).FormulaLocal = "=����(F" + CStr(p - 11) + ":F" + CStr(p) + ")"
                        ����22.Cells(p + 1, 7).FormulaLocal = "=����(G" + CStr(p - 11) + ":G" + CStr(p) + ")"
                        ����22.Cells(p + 1, 8).FormulaLocal = "=����(H" + CStr(p - 11) + ":H" + CStr(p) + ")"
                        ����22.Cells(p + 1, 9).FormulaLocal = "=����(I" + CStr(p - 11) + ":I" + CStr(p) + ")"
                        ����22.Cells(p + 1, 10).FormulaLocal = "=����(J" + CStr(p - 11) + ":J" + CStr(p) + ")"
                        ����22.Cells(p + 1, 11).FormulaLocal = "=����(K" + CStr(p - 11) + ":K" + CStr(p) + ")"
                        ����22.Cells(p + 1, 12).FormulaLocal = "=����(L" + CStr(p - 11) + ":L" + CStr(p) + ")"
                        ����22.Cells(p + 1, 13).FormulaLocal = "=����(M" + CStr(p - 11) + ":M" + CStr(p) + ")"
                        ����22.Cells(p + 1, 14).FormulaLocal = "=����(N" + CStr(p - 11) + ":N" + CStr(p) + ")"
                        ����22.Cells(p + 1, 15).FormulaLocal = "=����(O" + CStr(p - 11) + ":O" + CStr(p) + ")"
                        ����22.Cells(p + 1, 16).FormulaLocal = "=����(P" + CStr(p - 11) + ":P" + CStr(p) + ")"
                        ����22.Cells(p + 1, 17).FormulaLocal = "=����(Q" + CStr(p - 11) + ":Q" + CStr(p) + ")"
                        ����22.Cells(p + 1, 18).FormulaLocal = "=����(R" + CStr(p - 11) + ":R" + CStr(p) + ")"
                        ����22.Cells(p + 1, 19).FormulaLocal = "=����(S" + CStr(p - 11) + ":S" + CStr(p) + ")"
                        ����22.Cells(p + 1, 20).FormulaLocal = "=����(T" + CStr(p - 11) + ":T" + CStr(p) + ")"
                            Sheets("summary").Cells(p + 1, 21).Value = Qizol(sourcePSV, ChartPSV, RegimePSV, "��������� ���������") + Qizol(sourcePSV, ChartPSV, RegimePSV, "��������� ������������") + Qizol(sourcePSV, ChartPSV, RegimePSV, "���������") + Qizol(sourcePSV, ChartPSV, RegimePSV, "�������")
                            Sheets("summary").Cells(p + 1, 22).Value = Qizolnadz(sourcePSV, ChartPSV, RegimePSV, "���������", "������")
                            Sheets("summary").Cells(p + 1, 23).Value = Qizolnadz(sourcePSV, ChartPSV, RegimePSV, "���������", "�������")
                        ����22.Cells(p + 1, 24).FormulaLocal = "=����(U" + CStr(p) + ":W" + CStr(p) + ")"
                        ����22.Cells(p + 1, 25).FormulaLocal = "=����(Y" + CStr(p - 11) + ":Y" + CStr(p) + ")"
                        ����22.Cells(p + 1, 26).FormulaLocal = "=����(Z" + CStr(p - 11) + ":Z" + CStr(p) + ")"
                        ����22.Cells(p + 1, 27).FormulaLocal = "=����(AA" + CStr(p - 11) + ":AA" + CStr(p) + ")"
                        ����22.Cells(p + 1, 28).FormulaLocal = "=����(AB" + CStr(p - 11) + ":AB" + CStr(p) + ")"
                        ����22.Cells(p + 1, 29).FormulaLocal = "=����(AC" + CStr(p - 11) + ":AC" + CStr(p) + ")"
                        ����22.Cells(p + 1, 30).FormulaLocal = "=����(AD" + CStr(p - 11) + ":AD" + CStr(p) + ")"
                End If
                If ChartPSV = ChartTemperature Then
                    Exit For '����� �� ����� ����������� ������� ����������
                End If
            Next s
    Next c
'    Call recalculationBook
    Application.ScreenUpdating = True
End Sub
'Sub uijgjnfgb()
'lksdfg = SARZsummaryPSV(1, "������ 1", 744, 0)
'End Sub
'������ ������� ���� � ���� ���������
Function SARZsummaryPSV(source, Chart, heatingPeriod, summerPeriod) ', pPSV, WorkTime) '(source, Chart, cPSV, pPSV)
    SARZsummaryPSV1 = PSVSARZ(source, Chart, "������", "�� ������� �� ���� � ������������ ������")
    If SARZsummaryPSV1 = " - " Then
        SARZsummaryPSV1 = 0
    End If
    For ik = 2 To 21 Step 1
        temperatureChart = Sheets("temperature").Cells(5, ik)
        If temperatureChart = Chart Then
            workingYear1 = Sheets("temperature").Cells(19, ik)
            If workingYear1 = "" Or workingYear1 = Empty Then
                workingYear1 = 0
                Exit For
            End If
            If SARZsummaryPSV1 > 0 Or heatingPeriod > 0 Then
                SARZsummaryPSV1 = SARZsummaryPSV1 * (heatingPeriod / workingYear1)
                Exit For
            Else: Exit For
            End If
        End If
    Next ik
    SARZsummaryPSV2 = PSVSARZ(source, Chart, "�������", "�� ������� �� ���� � ������������ ������")
    If SARZsummaryPSV2 = " - " Then
        SARZsummaryPSV2 = 0
    End If
    For ik1 = 2 To 21 Step 1
        temperatureChart1 = Sheets("temperature").Cells(5, ik1)
        If temperatureChart1 = Chart Then
            workingYear2 = Sheets("temperature").Cells(19, ik1)
            If workingYear2 = "" Or workingYear1 = Empty Then
                workingYear2 = 0
                Exit For
            End If
            If SARZsummaryPSV2 > 0 Or heatingPeriod > 0 Then
                SARZsummaryPSV2 = SARZsummaryPSV2 * (heatingPeriod / workingYear1)
                Exit For
            Else: Exit For
            End If
        End If
    Next ik1
    SARZsummaryPSV3 = PSVSARZ(source, Chart, "��� ������", "�� ������� �� ���� � ������������ ������")
    If SARZsummaryPSV3 = " - " Then
        SARZsummaryPSV3 = 0
    End If
    For ik2 = 2 To 21 Step 1
        temperatureChart2 = Sheets("temperature").Cells(5, ik2)
        If temperatureChart2 = Chart Then
            workingYear3 = Sheets("temperature").Cells(19, ik2)
            If workingYear3 = "" Or workingYear1 = Empty Then
                workingYear3 = 0
                Exit For
            End If
            If SARZsummaryPSV3 > 0 Or heatingPeriod > 0 Then
                SARZsummaryPSV3 = SARZsummaryPSV3 * (heatingPeriod / workingYear1)
                Exit For
            Else: Exit For
            End If
        End If
    Next ik2
    SARZsummaryPSV4 = PSVSARZ(source, Chart, "������", "�� ������� �� ���� � ������ ������")
    If SARZsummaryPSV4 = " - " Then
        SARZsummaryPSV4 = 0
    End If
    For ik3 = 2 To 21 Step 1
        temperatureChart3 = Sheets("temperature").Cells(5, ik3)
        If temperatureChart3 = Chart Then
            workingYear4 = Sheets("temperature").Cells(19, ik3)
            If workingYear4 = "" Or workingYear1 = Empty Then
                workingYear4 = 0
                Exit For
            End If
            If SARZsummaryPSV4 > 0 Or heatingPeriod > 0 Then
                SARZsummaryPSV4 = SARZsummaryPSV4 * (summerPeriod / workingYear1)
                Exit For
            Else: Exit For
            End If
        End If
    Next ik3
    SARZsummaryPSV5 = PSVSARZ(source, Chart, "�������", "�� ������� �� ���� � ������ ������")
    If SARZsummaryPSV5 = " - " Then
        SARZsummaryPSV5 = 0
    End If
    For ik4 = 2 To 21 Step 1
        temperatureChart4 = Sheets("temperature").Cells(5, ik4)
        If temperatureChart4 = Chart Then
            workingYear5 = Sheets("temperature").Cells(19, ik4)
            If workingYear5 = "" Or workingYear1 = Empty Then
                workingYear5 = 0
                Exit For
            End If
            If SARZsummaryPSV5 > 0 Or heatingPeriod > 0 Then
                SARZsummaryPSV5 = SARZsummaryPSV5 * (summerPeriod / workingYear1)
                Exit For
            Else: Exit For
            End If
        End If
    Next ik4
    SARZsummaryPSV6 = PSVSARZ(source, Chart, "��� ������", "�� ������� �� ���� � ������ ������")
    If SARZsummaryPSV6 = " - " Then
        SARZsummaryPSV6 = 0
    End If
    For ik5 = 2 To 21 Step 1
        temperatureChart5 = Sheets("temperature").Cells(5, ik5)
        If temperatureChart5 = Chart Then
            workingYear6 = Sheets("temperature").Cells(19, ik5)
            If workingYear6 = "" Or workingYear1 = Empty Then
                workingYear6 = 0
                Exit For
            End If
            If SARZsummaryPSV6 > 0 Or heatingPeriod > 0 Then
                SARZsummaryPSV6 = SARZsummaryPSV6 * (summerPeriod / workingYear1)
                Exit For
            Else: Exit For
            End If
        End If
    Next ik5
    SARZsummaryPSV = SARZsummaryPSV1 + SARZsummaryPSV2 + SARZsummaryPSV3 + SARZsummaryPSV4 + SARZsummaryPSV5 + SARZsummaryPSV6
End Function

'������ ����� � �������� ������� ���� � ���� ���������
Function temperatureSARZMonth(Chart, Regime, Month, Direction)
If Month = "������" Then i = 7
If Month = "�������" Then i = 8
If Month = "����" Then i = 9
If Month = "������" Then i = 10
If Month = "���" Then i = 11
If Month = "����" Then i = 12
If Month = "����" Then i = 13
If Month = "������" Then i = 14
If Month = "��������" Then i = 15
If Month = "�������" Then i = 16
If Month = "������" Then i = 17
If Month = "�������" Then i = 18
    If Chart = "������ 1" Then k = 34
    If Chart = "������ 2" Then k = 36
    If Chart = "������ 3" Then k = 38
    If Chart = "������ 4" Then k = 40
    If Chart = "������ 5" Then k = 42
    If Chart = "������ 6" Then k = 44
    If Chart = "������ 7" Then k = 46
    If Chart = "������ 8" Then k = 48
    If Chart = "������ 9" Then k = 50
    If Chart = "������ 10" Then k = 52
        If Chart = "������ 1" Then j = 54
        If Chart = "������ 2" Then j = 55
        If Chart = "������ 3" Then j = 56
        If Chart = "������ 4" Then j = 57
        If Chart = "������ 5" Then j = 58
        If Chart = "������ 6" Then j = 59
        If Chart = "������ 7" Then j = 60
        If Chart = "������ 8" Then j = 61
        If Chart = "������ 9" Then j = 62
        If Chart = "������ 10" Then j = 63
            temperature1SARZ = Sheets("temperature").Cells(i, k)
            temperature2SARZ = Sheets("temperature").Cells(i, k + 1)
            temperatureSARZcold = Sheets("temperature").Cells(i, j)
        If Direction = "������" And temperature1SARZ <> "" Then
            temperatureSARZMonth = temperature1SARZ - temperatureSARZcold '����������� � ������ ��������������
        End If
        If Direction = "�������" And temperature2SARZ <> "" Then
            temperatureSARZMonth = temperature2SARZ - temperatureSARZcold '����������� � ������� ��������������
        End If
        If Direction = "��� ������" And temperature1SARZ <> "" And temperature2SARZ <> "" Then
            temperatureSARZMonth = (temperature1SARZ + temperatureSARZMonth) / 2 - temperatureSARZcold  '����������� � ������� ��������������
        End If
End Function
           
'������ ����� �� ������� �� ������� summary �� ������� �� ����
Function SARZsummaryGcal(source, Chart, Regime, Month, heatingPeriod, summerPeriod)
    SARZsummaryPSV1 = PSVSARZ(source, Chart, "������", "�� ������� �� ���� � ������������ ������")
    If SARZsummaryPSV1 = " - " Then
        SARZsummaryPSV1 = 0
    End If
    For ik = 2 To 21 Step 1
        temperatureChart = Sheets("temperature").Cells(5, ik)
        If temperatureChart = Chart Then
            workingYear1 = Sheets("temperature").Cells(19, ik)
            If workingYear1 = "" Or workingYear1 = Empty Then
                workingYear1 = 0
                Exit For
            End If
            If SARZsummaryPSV1 > 0 And heatingPeriod > 0 Then
                SARZsummaryGcal1 = (SARZsummaryPSV1 * (heatingPeriod / workingYear1) * _
                (temperatureSARZMonth(Chart, Regime, Month, "������") * wskDSWT(temperatureSARZMonth(Chart, Regime, Month, "������")))) / 1000000
                Exit For
            Else: Exit For
            End If
        End If
    Next ik
    SARZsummaryPSV2 = PSVSARZ(source, Chart, "�������", "�� ������� �� ���� � ������������ ������")
    If SARZsummaryPSV2 = " - " Then
        SARZsummaryPSV2 = 0
    End If
    For ik1 = 2 To 21 Step 1
        temperatureChart1 = Sheets("temperature").Cells(5, ik1)
        If temperatureChart1 = Chart Then
            workingYear2 = Sheets("temperature").Cells(19, ik1)
            If workingYear2 = "" Or workingYear1 = Empty Then
                workingYear2 = 0
                Exit For
            End If
            If SARZsummaryPSV2 > 0 And heatingPeriod > 0 Then
                SARZsummaryGcal2 = (SARZsummaryPSV2 * (heatingPeriod / workingYear1) * _
                (temperatureSARZMonth(Chart, Regime, Month, "�������") * wskDSWT(temperatureSARZMonth(Chart, Regime, Month, "�������")))) / 1000000
                Exit For
            Else: Exit For
            End If
        End If
    Next ik1
    SARZsummaryPSV3 = PSVSARZ(source, Chart, "��� ������", "�� ������� �� ���� � ������������ ������")
    If SARZsummaryPSV3 = " - " Then
        SARZsummaryPSV3 = 0
    End If
    For ik2 = 2 To 21 Step 1
        temperatureChart2 = Sheets("temperature").Cells(5, ik2)
        If temperatureChart2 = Chart Then
            workingYear3 = Sheets("temperature").Cells(19, ik2)
            If workingYear3 = "" Or workingYear1 = Empty Then
                workingYear3 = 0
                Exit For
            End If
            If SARZsummaryPSV3 > 0 And heatingPeriod > 0 Then
                SARZsummaryGcal3 = (SARZsummaryPSV3 * (heatingPeriod / workingYear1) * _
                (temperatureSARZMonth(Chart, Regime, Month, "��� ������") * wskDSWT(temperatureSARZMonth(Chart, Regime, Month, "��� ������")))) / 1000000
                Exit For
            Else: Exit For
            End If
        End If
    Next ik2
    SARZsummaryPSV4 = PSVSARZ(source, Chart, "������", "�� ������� �� ���� � ������ ������")
    If SARZsummaryPSV4 = " - " Then
        SARZsummaryPSV4 = 0
    End If
    For ik3 = 2 To 21 Step 1
        temperatureChart3 = Sheets("temperature").Cells(5, ik3)
        If temperatureChart3 = Chart Then
            workingYear4 = Sheets("temperature").Cells(19, ik3)
            If workingYear4 = "" Or workingYear1 = Empty Then
                workingYear4 = 0
                Exit For
            End If
            If SARZsummaryPSV4 > 0 And heatingPeriod > 0 Then
                SARZsummaryGcal4 = (SARZsummaryPSV4 * (heatingPeriod / workingYear1) * _
                (temperatureSARZMonth(Chart, Regime, Month, "������") * wskDSWT(temperatureSARZMonth(Chart, Regime, Month, "������")))) / 1000000
                Exit For
            Else: Exit For
            End If
        End If
    Next ik3
    SARZsummaryPSV5 = PSVSARZ(source, Chart, "�������", "�� ������� �� ���� � ������ ������")
    If SARZsummaryPSV5 = " - " Then
        SARZsummaryPSV5 = 0
    End If
    For ik4 = 2 To 21 Step 1
        temperatureChart4 = Sheets("temperature").Cells(5, ik4)
        If temperatureChart4 = Chart Then
            workingYear5 = Sheets("temperature").Cells(19, ik4)
            If workingYear5 = "" Or workingYear1 = Empty Then
                workingYear5 = 0
                Exit For
            End If
            If SARZsummaryPSV5 > 0 And heatingPeriod > 0 Then
                SARZsummaryGcal5 = (SARZsummaryPSV5 * (heatingPeriod / workingYear1) * _
                (temperatureSARZMonth(Chart, Regime, Month, "�������") * wskDSWT(temperatureSARZMonth(Chart, Regime, Month, "�������")))) / 1000000
                Exit For
            Else: Exit For
            End If
        End If
    Next ik4
    SARZsummaryPSV6 = PSVSARZ(source, Chart, "��� ������", "�� ������� �� ���� � ������ ������")
    If SARZsummaryPSV6 = " - " Then
        SARZsummaryPSV6 = 0
    End If
    For ik5 = 2 To 21 Step 1
        temperatureChart5 = Sheets("temperature").Cells(5, ik5)
        If temperatureChart5 = Chart Then
            workingYear6 = Sheets("temperature").Cells(19, ik5)
            If workingYear6 = "" Or workingYear1 = Empty Then
                workingYear6 = 0
                Exit For
            End If
            If SARZsummaryPSV6 > 0 And heatingPeriod > 0 Then
                SARZsummaryGcal6 = (SARZsummaryPSV6 * (heatingPeriod / workingYear1) * _
                (temperatureSARZMonth(Chart, Regime, Month, "��� ������") * wskDSWT(temperatureSARZMonth(Chart, Regime, Month, "��� ������")))) / 1000000
                Exit For
            Else: Exit For
            End If
        End If
    Next ik5
    SARZsummaryGcal = SARZsummaryGcal1 + SARZsummaryGcal2 + SARZsummaryGcal3 + SARZsummaryGcal4 + SARZsummaryGcal5 + SARZsummaryGcal6
End Function

'���������� ������ ��� ���������
Sub fillingDiagram()
    Application.ScreenUpdating = False
    Range(Sheets("diagram").Cells(40, 3), Sheets("diagram").Cells(52, 8)).ClearContents '�������� ��������� �� �����������
    Range(Sheets("diagram").Cells(40, 13), Sheets("diagram").Cells(52, 16)).ClearContents '�������� ��������� �� �����������
    For i = 40 To 52 Step 1
        Month1 = Sheets("diagram").Cells(i, 2)
        k = 27
        Sheets("diagram").Cells(i, 3).Value = tabAmountSummary(Month1, k)
        k = 28
        Sheets("diagram").Cells(i, 4).Value = tabAmountSummary(Month1, k)
        k = 30
        Sheets("diagram").Cells(i, 5).Value = tabAmountSummary(Month1, k)
        k = 21
        Sheets("diagram").Cells(i, 13).Value = tabAmountSummary(Month1, k)
        k = 22
        Sheets("diagram").Cells(i, 14).Value = tabAmountSummary(Month1, k)
        k = 23
        Sheets("diagram").Cells(i, 15).Value = tabAmountSummary(Month1, k)
        k = 24
        Sheets("diagram").Cells(i, 16).Value = tabAmountSummary(Month1, k)
    Next i
        lossThroughIsolation = Sheets("diagram").Cells(52, 3)
        networkWaterLoss = Sheets("diagram").Cells(52, 4)
        totalosses = Sheets("diagram").Cells(52, 5)
    For j = 40 To 51 Step 1
            Sheets("diagram").Cells(j, 6).Value = lossThroughIsolation
            Sheets("diagram").Cells(j, 7).Value = networkWaterLoss
            Sheets("diagram").Cells(j, 8).Value = totalosses
    Next j
    
    Application.ScreenUpdating = True
End Sub
'��������� ������ � ������� summary
Function tabAmountSummary(Month, k)
    Set list0 = Sheets("summary").Range(Sheets("summary").Cells(23, k), Sheets("summary").Cells(25000, k))
    Set list1 = Sheets("summary").Range(Sheets("summary").Cells(23, 1), Sheets("summary").Cells(25000, 1))
    tabAmountSummary = Application.SumIfs(list0, list1, Month)
End Function








