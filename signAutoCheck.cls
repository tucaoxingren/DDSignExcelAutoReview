VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ��ʼ������
' ����ǩ����ʼʱ��
Const AmStart As String = "08:00:00"
' ����ǩ������ʱ��
Const AmEnd As String = "09:00:00"
' ����ǩ����ʼʱ��
Const noonStart As String = "13:00:00"
' ����ǩ������ʱ��
Const noonEnd As String = "14:00:00"
' ����ǩ����ʼʱ��
Const PmStart As String = "18:00:00"
' ����ǩ������ʱ��
Const PmEnd As String = "21:00:00"
'Y-��˱�ע  Z-��Чǩ������  AA-�����涨����
Const reviewCol As String = "Y"
Const invalidCheckCol As String = "Z"
Const excessCountCol As String = "AA"
'������
Const nameCol As String = "A"
'������
Const dateCol As String = "F"
Sub ���ǩ������()
' ˼·:
' ���������� У��ͼƬ�в�Ϊ�� ����ݷö���ͱ�ע��ͬʱΪ��
' ����ÿ������ʱ���Ƿ��ڹ涨ʱ���� ��ʱ�����ݼ�Ϊ��Ч���� ������ ʱ���� ÿ��������ǩ������

' ����ÿ�¿�����9����Чǩ������δǩ����    ������������12����Чǩ������δǩ����
' ����ְλ�趨��Чǩ����ֵ ������ֵʱ�ڳ����涨��������� ��������
' ����˱�ע�� ������Чǩ��˵�� ����:
'   2018-10-10 ȫ��ǩ���ްݷö���
'   2018-10-11 ����ǩ���ްݷö���
'   2018-10-13 ����δǩ��
'   2018-10-14  ����δǩ��
'   ����δǩ��
'   7 ����Чǩ��

' A-����    C-����  D-��������  E-ְλ  F-����  G-ʱ��  H-�ص�  I-��ϸ��ַ  L-�ݷö���  M-��ע  P-ͼƬ
' Y-��˱�ע  Z-��Чǩ������  AA-�����涨����
'Application.ScreenUpdate = False

MsgBox "���������⽴��������"
' ѭ������
Dim i As Integer

' ��˱�ע
Dim reviewNote As String
' ��˱��
'Dim reviewFlag As Boolean
reviewFlag = True
' ��Чǩ������
Dim invalidCheckNum As Integer
' ��Чǩ������
Dim validCheckNum As Integer
' �����涨����
Dim excessCount As Integer

' ��˱��
'   ��λ:ʱ����˱��
'       0-��������У��Ϊ��ǩ��ʱ���û�з��ϵ�ʱ���
'       1-��������У��Ϊ����ǩ��ʱ��
'       2-��������У��Ϊ����ǩ��ʱ��
'       3-��������У��Ϊ����ǩ��ʱ��
'   ʮλ:ͼƬ��У����
'       1-У��ͨ��
'       2-У�鲻ͨ��
'   ��λ:�ݷö���ͱ�ע��ͬʱΪ��У����
'       1-У��ͨ��
'       2-У�鲻ͨ��
Dim dateCheckFlag As Integer
dateCheckFlag = 0

' ������ʼ�к�
i = 4
' ����ȫ������
Do While Cells(i, dateCol) <> ""
    ' ��ʼ��Чǩ�������볬���涨����
    invalidCheckNum = 0
    excessCount = 0
    ' ���� ĳ�˵�ȫ��ǩ������
    Dim personCow As Integer
    personCow = i
    ' ��ȡĳ�˼�¼����ֹ�к�
    Do While Cells(personCow, nameCol) = Cells(personCow + 1, nameCol)
        personCow = personCow + 1
    Loop
    Dim personCount As Integer
    personCount = personCow + 1 - i
    Do While i <= personCow
        dateCheckFlag = 0
        Dim dayCow As Integer
        dayCow = i
        ' ��ȡĳ��ĳ�յ���ֹ�к�
        Do While Cells(dayCow, "F") = Cells(dayCow + 1, dateCol) And dayCow < personCow
            dayCow = dayCow + 1
        Loop
        ' ��ʼ��¼����
        Dim k As Integer
        k = 1
        Dim dayCount As Integer
        dayCount = dayCow + 1 - i
        Dim arr()
        ReDim arr(1 To dayCount)
        ' ���� ĳ��ĳ�յ�ȫ��ǩ������
        Do While i <= dayCow
            dateCheckFlag = check(i)
            'Sheet1.Cells(i, "AB").Value = dateCheckFlag
            arr(k) = dateCheckFlag
            k = k + 1
            i = i + 1
        Loop
        reviewNote = Cells(i - 1, dateCol) & Chr(10)
        validCheckNum = 0
        ' ����ͨ������ 0-δͨ�� 1-ͨ��
        Dim arrSign(1 To 3)
        arrSign(1) = 0
        arrSign(2) = 0
        arrSign(3) = 0
        For j = 1 To dayCount
            If arr(j) = 111 Then
                arrSign(1) = 1
            ElseIf arr(j) = 112 Then
                arrSign(2) = 1
            ElseIf arr(j) = 113 Then
                arrSign(3) = 1
            End If
        Next j
        If (arrSign(1) + arrSign(2) + arrSign(3)) = 3 Then
            reviewNote = ""
        Else
            If arrSign(1) = 0 And arrSign(2) = 0 Then
                reviewNote = reviewNote + " ����δǩ�� ����δǩ��"
                invalidCheckNum = invalidCheckNum + 2
            ElseIf arrSign(1) = 0 And arrSign(3) = 0 Then
                reviewNote = reviewNote + " ����δǩ�� ����δǩ��"
                invalidCheckNum = invalidCheckNum + 2
            ElseIf arrSign(2) = 0 And arrSign(3) = 0 Then
                reviewNote = reviewNote + " ����δǩ�� ����δǩ"
                invalidCheckNum = invalidCheckNum + 2
            ElseIf arrSign(1) = 0 Then
                reviewNote = reviewNote + " ����δǩ��"
                invalidCheckNum = invalidCheckNum + 1
            ElseIf arrSign(2) = 0 Then
                reviewNote = reviewNote + " ����δǩ��"
                invalidCheckNum = invalidCheckNum + 1
            ElseIf arrSign(3) = 0 Then
                reviewNote = reviewNote + " ����δǩ��"
                invalidCheckNum = invalidCheckNum + 1
            Else
                reviewNote = reviewNote + " ȫ��δǩ��"
                invalidCheckNum = invalidCheckNum + 3
            End If
        End If
         ' ������Ԫ�����
        Dim reviewNoteRange As Range
        ' �����ϲ�����
        Dim reviewNoteStr As String
        reviewNoteStr = reviewCol & CStr(i - dayCount) & ":" & reviewCol & CStr(i - 1)
        ' ����Ԫ�����ֵΪָ������,����ָ�����ⵥԪ��
        Set reviewNoteRange = Range(reviewNoteStr)
        ' �ϲ��õ�Ԫ������
        reviewNoteRange.Merge
        ' �ͷŵ�Ԫ�����
        Set reviewNoteRange = Nothing
        ' ������˴�����˱�ע
        Sheet1.Cells(i - dayCount, reviewCol).Value = reviewNote
    Loop
    ' ������Ԫ�����
    Dim invalidCheckRange As Range
    ' �����ϲ�����
    Dim invalidCheckStr As String
    invalidCheckStr = invalidCheckCol & CStr(i - personCount) & ":" & invalidCheckCol & CStr(i - 1)
    ' ����Ԫ�����ֵΪָ������,����ָ�����ⵥԪ��
    Set invalidCheckRange = Range(invalidCheckStr)
    ' �ϲ��õ�Ԫ������
    invalidCheckRange.Merge
    ' �ͷŵ�Ԫ�����
    Set invalidCheckRange = Nothing
    ' ���������Чǩ������
    Sheet1.Cells(i - personCount, invalidCheckCol).Value = "��Чǩ������:" & CStr(invalidCheckNum) & "��"
    ' ��������涨���� ����ְλ������ֵ
    If Sheet1.Cells(i - 1, "E").Value = "������" Then
        excessCount = invalidCheckNum - 12
    ElseIf Sheet1.Cells(i - 1, "E").Value = "����" Then
        excessCount = invalidCheckNum - 9
    End If
    ' ������Ԫ�����
    Dim excessCountRange As Range
    ' �����ϲ�����
    Dim excessCountStr As String
    excessCountStr = excessCountCol & CStr(i - personCount) & ":" & excessCountCol & CStr(i - 1)
    ' ����Ԫ�����ֵΪָ������,����ָ�����ⵥԪ��
    Set excessCountRange = Range(excessCountStr)
    ' �ϲ��õ�Ԫ������
    excessCountRange.Merge
    ' �ͷŵ�Ԫ�����
    Set invalidCheckRange = Nothing
    If excessCount > 0 Then
        Sheet1.Cells(i - personCount, excessCountCol).Value = "�����涨����:" & CStr(excessCount) & "��"
    End If
Loop
End Sub
' ����ÿһ������ i-����ѭ������
Function check(i As Integer)
    ' ��˱��
    Dim dateCheckFlag As Integer
    dateCheckFlag = 0
    ' ÿ�ν���ѭ��������˱��Ϊ��Ч
    'reviewFlag = True
    ' У��ʱ��
    If Cells(i, "G") = "" Then
        dateCheckFlag = 0
    Else
       Dim dateValue As Date
       dateValue = TimeValue(Cells(i, "G"))
       dateCheckFlag = checkDate(dateValue)
    End If
    ' У��ͼƬ��
    If Cells(i, "P") = "" Then
        dateCheckFlag = 20 + dateCheckFlag
    Else
        dateCheckFlag = 10 + dateCheckFlag
    End If
     ' У��ݷö���ͱ�ע��ͬʱΪ��
    If Cells(i, "L") = "" And Cells(i, "M") = "" Then
        dateCheckFlag = 200 + dateCheckFlag
    Else
        dateCheckFlag = 100 + dateCheckFlag
    End If
    check = dateCheckFlag
End Function
' У��ʱ���Ƿ���ָ������ checkTime-��ҪУ���ʱ��
Function checkDate(checkTime As Date)
    Dim dateCheckFlag As Integer
    dateCheckFlag = 0
    Dim start As Date
    start = TimeValue(AmStart)
    If checkTime >= TimeValue(AmStart) And checkTime <= TimeValue(AmEnd) Then
        dateCheckFlag = 1
    ElseIf checkTime >= TimeValue(noonStart) And checkTime <= TimeValue(noonEnd) Then
        dateCheckFlag = 2
    ElseIf checkTime >= TimeValue(PmStart) And checkTime <= TimeValue(PmEnd) Then
        dateCheckFlag = 3
    End If
    checkDate = dateCheckFlag
End Function
