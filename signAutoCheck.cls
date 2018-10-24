VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 初始化常量
' 上午签到开始时间
Const AmStart As String = "08:00:00"
' 上午签到结束时间
Const AmEnd As String = "09:00:00"
' 中午签到开始时间
Const noonStart As String = "13:00:00"
' 中午签到结束时间
Const noonEnd As String = "14:00:00"
' 下午签到开始时间
Const PmStart As String = "18:00:00"
' 下午签到结束时间
Const PmEnd As String = "21:00:00"
'Y-审核备注  Z-无效签到次数  AA-超出规定次数
Const reviewCol As String = "Y"
Const invalidCheckCol As String = "Z"
Const excessCountCol As String = "AA"
'姓名列
Const nameCol As String = "A"
'日期列
Const dateCol As String = "F"
Sub 检查签到数据()
' 思路:
' 按姓名分组 校验图片列不为空 检验拜访对象和备注不同时为空
' 检验每天三次时间是否在规定时间内 无时间数据即为无效数据 日期列 时间列 每日有三次签到数据

' 督导每月可以有9次无效签到（含未签到）    区域经理可以有12次无效签到（含未签到）
' 根据职位设定无效签到阈值 超过阈值时在超出规定次数列输出 超出次数
' 在审核备注列 设置无效签到说明 如下:
'   2018-10-10 全天签到无拜访对象
'   2018-10-11 早上签到无拜访对象
'   2018-10-13 早上未签到
'   2018-10-14  中午未签到
'   下午未签到
'   7 次无效签到

' A-姓名    C-部门  D-完整部门  E-职位  F-日期  G-时间  H-地点  I-详细地址  L-拜访对象  M-备注  P-图片
' Y-审核备注  Z-无效签到次数  AA-超出规定次数
'Application.ScreenUpdate = False

MsgBox "美丽的柒柒酱工作快乐"
' 循环变量
Dim i As Integer

' 审核备注
Dim reviewNote As String
' 审核标记
'Dim reviewFlag As Boolean
reviewFlag = True
' 无效签到次数
Dim invalidCheckNum As Integer
' 有效签到次数
Dim validCheckNum As Integer
' 超出规定次数
Dim excessCount As Integer

' 审核标记
'   个位:时间审核标记
'       0-该行数据校验为无签到时间或没有符合的时间段
'       1-该行数据校验为上午签到时间
'       2-该行数据校验为上午签到时间
'       3-该行数据校验为上午签到时间
'   十位:图片列校验标记
'       1-校验通过
'       2-校验不通过
'   百位:拜访对象和备注不同时为空校验标记
'       1-校验通过
'       2-校验不通过
Dim dateCheckFlag As Integer
dateCheckFlag = 0

' 数据起始行号
i = 4
' 遍历全部数据
Do While Cells(i, dateCol) <> ""
    ' 初始无效签到次数与超出规定次数
    invalidCheckNum = 0
    excessCount = 0
    ' 遍历 某人的全部签到数据
    Dim personCow As Integer
    personCow = i
    ' 获取某人记录的终止行号
    Do While Cells(personCow, nameCol) = Cells(personCow + 1, nameCol)
        personCow = personCow + 1
    Loop
    Dim personCount As Integer
    personCount = personCow + 1 - i
    Do While i <= personCow
        dateCheckFlag = 0
        Dim dayCow As Integer
        dayCow = i
        ' 获取某人某日的终止行号
        Do While Cells(dayCow, "F") = Cells(dayCow + 1, dateCol) And dayCow < personCow
            dayCow = dayCow + 1
        Loop
        ' 初始记录数组
        Dim k As Integer
        k = 1
        Dim dayCount As Integer
        dayCount = dayCow + 1 - i
        Dim arr()
        ReDim arr(1 To dayCount)
        ' 遍历 某人某日的全部签到数据
        Do While i <= dayCow
            dateCheckFlag = check(i)
            'Sheet1.Cells(i, "AB").Value = dateCheckFlag
            arr(k) = dateCheckFlag
            k = k + 1
            i = i + 1
        Loop
        reviewNote = Cells(i - 1, dateCol) & Chr(10)
        validCheckNum = 0
        ' 检验通过数组 0-未通过 1-通过
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
                reviewNote = reviewNote + " 上午未签到 中午未签到"
                invalidCheckNum = invalidCheckNum + 2
            ElseIf arrSign(1) = 0 And arrSign(3) = 0 Then
                reviewNote = reviewNote + " 上午未签到 下午未签到"
                invalidCheckNum = invalidCheckNum + 2
            ElseIf arrSign(2) = 0 And arrSign(3) = 0 Then
                reviewNote = reviewNote + " 中午未签到 下午未签"
                invalidCheckNum = invalidCheckNum + 2
            ElseIf arrSign(1) = 0 Then
                reviewNote = reviewNote + " 上午未签到"
                invalidCheckNum = invalidCheckNum + 1
            ElseIf arrSign(2) = 0 Then
                reviewNote = reviewNote + " 中午未签到"
                invalidCheckNum = invalidCheckNum + 1
            ElseIf arrSign(3) = 0 Then
                reviewNote = reviewNote + " 下午未签到"
                invalidCheckNum = invalidCheckNum + 1
            Else
                reviewNote = reviewNote + " 全天未签到"
                invalidCheckNum = invalidCheckNum + 3
            End If
        End If
         ' 声明单元格对象
        Dim reviewNoteRange As Range
        ' 声明合并区域
        Dim reviewNoteStr As String
        reviewNoteStr = reviewCol & CStr(i - dayCount) & ":" & reviewCol & CStr(i - 1)
        ' 给单元格对象赋值为指定区域,可以指定任意单元格
        Set reviewNoteRange = Range(reviewNoteStr)
        ' 合并该单元格区域
        reviewNoteRange.Merge
        ' 释放单元格变量
        Set reviewNoteRange = Nothing
        ' 输出此人此日审核备注
        Sheet1.Cells(i - dayCount, reviewCol).Value = reviewNote
    Loop
    ' 声明单元格对象
    Dim invalidCheckRange As Range
    ' 声明合并区域
    Dim invalidCheckStr As String
    invalidCheckStr = invalidCheckCol & CStr(i - personCount) & ":" & invalidCheckCol & CStr(i - 1)
    ' 给单元格对象赋值为指定区域,可以指定任意单元格
    Set invalidCheckRange = Range(invalidCheckStr)
    ' 合并该单元格区域
    invalidCheckRange.Merge
    ' 释放单元格变量
    Set invalidCheckRange = Nothing
    ' 输出此人无效签到次数
    Sheet1.Cells(i - personCount, invalidCheckCol).Value = "无效签到次数:" & CStr(invalidCheckNum) & "次"
    ' 输出超过规定次数 根据职位设置阈值
    If Sheet1.Cells(i - 1, "E").Value = "区域经理" Then
        excessCount = invalidCheckNum - 12
    ElseIf Sheet1.Cells(i - 1, "E").Value = "督导" Then
        excessCount = invalidCheckNum - 9
    End If
    ' 声明单元格对象
    Dim excessCountRange As Range
    ' 声明合并区域
    Dim excessCountStr As String
    excessCountStr = excessCountCol & CStr(i - personCount) & ":" & excessCountCol & CStr(i - 1)
    ' 给单元格对象赋值为指定区域,可以指定任意单元格
    Set excessCountRange = Range(excessCountStr)
    ' 合并该单元格区域
    excessCountRange.Merge
    ' 释放单元格变量
    Set invalidCheckRange = Nothing
    If excessCount > 0 Then
        Sheet1.Cells(i - personCount, excessCountCol).Value = "超出规定次数:" & CStr(excessCount) & "次"
    End If
Loop
End Sub
' 检验每一行数据 i-整体循环变量
Function check(i As Integer)
    ' 审核标记
    Dim dateCheckFlag As Integer
    dateCheckFlag = 0
    ' 每次进入循环设置审核标记为有效
    'reviewFlag = True
    ' 校验时间
    If Cells(i, "G") = "" Then
        dateCheckFlag = 0
    Else
       Dim dateValue As Date
       dateValue = TimeValue(Cells(i, "G"))
       dateCheckFlag = checkDate(dateValue)
    End If
    ' 校验图片列
    If Cells(i, "P") = "" Then
        dateCheckFlag = 20 + dateCheckFlag
    Else
        dateCheckFlag = 10 + dateCheckFlag
    End If
     ' 校验拜访对象和备注不同时为空
    If Cells(i, "L") = "" And Cells(i, "M") = "" Then
        dateCheckFlag = 200 + dateCheckFlag
    Else
        dateCheckFlag = 100 + dateCheckFlag
    End If
    check = dateCheckFlag
End Function
' 校验时间是否在指定区间 checkTime-需要校验的时间
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

