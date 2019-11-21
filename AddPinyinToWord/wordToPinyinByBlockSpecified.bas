Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public g_manualCharSet As Object
Public g_specifiedCharSet As Object
Public g_exceptionCharSet As Object

Function isNeedZhuyin(ByVal tmpChar) As Boolean
            
            ' 非汉字，跳出
            If LenB(StrConv(tmpChar, vbFromUnicode)) <> 2 Then
                    isNeedZhuyin = False
                    Exit Function
            End If
            
            ' 手动指定字符 和 例外字符 两种模式，有点像黑白名单，没有做成优先级的方式，简单点，二选一就可以了
            If True Then
                    ' 特殊符号不需要加注音，跳出
                    If "" <> g_exceptionCharSet(tmpChar) Then
                            isNeedZhuyin = False
                            Exit Function
                    End If
                    
                    isNeedZhuyin = True
                    Exit Function
                    
            Else
                    '是指定的字才加注音
                    If "" <> g_specifiedCharSet(tmpChar) Then
                            isNeedZhuyin = True
                            Exit Function
                    End If
                    
                    isNeedZhuyin = False
                    Exit Function
                
            End If
            
End Function

Function initExceptionCharSet()

        Set g_exceptionCharSet = CreateObject("Scripting.Dictionary")

        '不需要加注音的字符，加到这个集合里
        
        g_exceptionCharSet.Add "。", "ok"
        g_exceptionCharSet.Add "，", "ok"
        g_exceptionCharSet.Add "、", "ok"
        g_exceptionCharSet.Add "（", "ok"
        g_exceptionCharSet.Add "）", "ok"
        g_exceptionCharSet.Add "(", "ok"
        g_exceptionCharSet.Add ")", "ok"
        g_exceptionCharSet.Add "：", "ok"
        g_exceptionCharSet.Add ": ", "ok"
        g_exceptionCharSet.Add "　", "ok"
        g_exceptionCharSet.Add "[", "ok"
        g_exceptionCharSet.Add "]", "ok"
        g_exceptionCharSet.Add "-", "ok"
        g_exceptionCharSet.Add "+", "ok"
        g_exceptionCharSet.Add "*", "ok"
        g_exceptionCharSet.Add "《", "ok"
        g_exceptionCharSet.Add "》", "ok"
        g_exceptionCharSet.Add "【", "ok"
        g_exceptionCharSet.Add "】", "ok"
        g_exceptionCharSet.Add "“", "ok"
        g_exceptionCharSet.Add Chr$(9), "ok"
        g_exceptionCharSet.Add Chr$(10), "ok"
        g_exceptionCharSet.Add Chr$(13), "ok"
    
    
End Function

Function initManualCharSet()

        Set g_manualCharSet = CreateObject("Scripting.Dictionary")
        
        '如果一个字，不是所有出现的地方都相同发音，那就需要后期手动修改
    
        g_manualCharSet.Add "南", "ná"
        g_manualCharSet.Add "无", "mó"
        g_manualCharSet.Add "唵", "ong"
        g_manualCharSet.Add "尽", "jìn"
        g_manualCharSet.Add "晒", "lì"
    
End Function

Function initSpecifiedCharSet()

        Set g_specifiedCharSet = CreateObject("Scripting.Dictionary")

        g_specifiedCharSet.Add "伽", "ok"
        g_specifiedCharSet.Add "梵", "ok"
        'g_specifiedCharSet.Add "乐", "ok"
        g_specifiedCharSet.Add "苾", "ok"
        g_specifiedCharSet.Add "刍", "ok"

End Function

Function getPinyin(ByVal text As String) As String

        Dim Result As String
        Dim cmd As String
        
        Dim wExec As Object
        Dim shellObj As Object
        Set shellObj = CreateObject("WScript.Shell")
        
        cmd = "pythonw C:\wordToPinyin.py " & text
        
        Set wExec = shellObj.Exec(cmd)
        Result = wExec.StdOut.ReadAll
    
        '如果运行时同时做其它操作，这里有时候会中段，可能是什么信号，懒得处理了，F5继续运行就可以了
        
        '去掉空格和后面的换行，否则注音显示会歪向左边
    '    spaceIdx = InStr(Result, " ")
    '    If 0 <> spaceIdx Then
    '        Result = Mid(Result, 1, spaceIdx - 1)
    '    End If
        
        getPinyin = Result
    
        Set wExec = Nothing
        Set shellObj = Nothing
    
End Function

Function isExceptionChar(ByVal char As String) As Boolean
    '有些字符不需要加注音
    
    If "" <> g_exceptionCharSet(tmpChar) Then
            isExceptionChar = True
    Else
            isExceptionChar = False
    End If
    
End Function

Function getManualPinyin(ByRef tmpChar As String, ByRef charPinyin As String)

        Dim tmpPinyin As String
        tmpPinyin = g_manualCharSet(tmpChar)
        If "" <> tmpPinyin Then
            charPinyin = tmpPinyin
        End If
        
End Function

Function countChar(areaStart As Long, areaEnd As Long)
    '手动算字数
    
    Dim charNum As Long
            
            '全选时要修正一下结束位置，否则全选时会死循环结束不了
            Selection.EndKey unit:=wdStory
            docEnd = Selection.End
            If docEnd < areaEnd Then
                    areaEnd = docEnd
            End If
    
            '恢复光标到起始位置
            Selection.Start = areaStart
            Selection.End = areaStart
            
            charNum = 0
            Do While Selection.End < areaEnd
        
                    If charNum Mod 256 = 0 Then
                            ' 防止假死
                            DoEvents
                            'Debug.Print "正在计算字数: " & Selection.End & "/" & areaEnd & "    百分比:" & Format((Selection.End - areaStart) / (areaEnd - areaStart), "Percent") & "%"
                    End If
            
                    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdMove
                    charNum = charNum + 1
        
            Loop
            
            countChar = charNum
            
End Function

Function showDebugLog(ByVal prompt As String, _
        ByVal allBeginTimeSec As Double, ByVal allDoneCount As Long, ByVal allNum As Long, _
        ByVal phaseBeginTimeSec As Double, ByVal phaseDoneCount As Long, ByVal phaseNum As Long)

        Dim nowTimeSec As Double
        
        Dim allAvrCharCount As Double
        Dim allCostTimeSec As Double
        
        Dim phaseAvrCharCount As Double
        Dim phaseCostTimeSec As Double
        
        Dim tmpMS As Long
        'tmpMS = timeGetTime()
        
        If (0 = allDoneCount) Or (0 = phaseDoneCount) Then
                Exit Function
        End If
        
        nowTimeSec = Timer()
        
        allCostTimeSec = nowTimeSec - allBeginTimeSec
        If 0 <> allCostTimeSec Then
                allAvrCharCount = allDoneCount / allCostTimeSec
        Else
                allAvrCharCount = 1
        End If
        
        phaseCostTimeSec = nowTimeSec - phaseBeginTimeSec
        If 0 <> phaseCostTimeSec Then
                phaseAvrCharCount = phaseDoneCount / (nowTimeSec - phaseBeginTimeSec)
        Else
                phaseAvrCharCount = 1
        End If
        
        Debug.Print prompt & "    all：" & Format(allCostTimeSec / 24 / 3600, "hh:mm:ss") & _
                "    " & Format((allNum - allDoneCount) / allAvrCharCount / 24 / 3600, "hh:mm:ss") & _
                "    " & allDoneCount & "/" & allNum & _
                "    " & Format(allAvrCharCount, "Standard") & "字/秒" & _
                "    " & Format(allDoneCount / allNum, "Percent") & _
                "    phase：" & Format(phaseCostTimeSec / 24 / 3600, "hh:mm:ss") & _
                "    " & Format((phaseNum - phaseDoneCount) / phaseAvrCharCount / 24 / 3600, "hh:mm:ss") & _
                "    " & phaseDoneCount & "/" & phaseNum & _
                "    " & Format(phaseAvrCharCount, "Standard") & "字/秒" & _
                "    " & Format(phaseDoneCount / phaseNum, "Percent")

End Function

Sub addPinyin()
        '给选中区域加注音
        
        '避免被系统信号打断，好像有点效果
        Application.EnableCancelKey = False
        
        Dim blockCharLimit As Long
        '可自行修改每次处理的字符数量的大小
        blockCharLimit = 100
    
        Dim textForPinyin As String
        Dim tmpChar As String
        Dim pinyin As String
        Dim blockText As String
        
        Dim cursor1 As Long
        Dim cursor2 As Long
        Dim areaStart As Long
        Dim areaEnd As Long
        Dim docEnd As Long
        Dim lastPos As Long
        Dim blockAddedCharCount As Long
        Dim charPinyin As String
        Dim pinyinNum As Long
        
        Dim blockStart As Long
        Dim blockEnd As Long
        Dim blockReadCharCount As Long
        Dim blockReadMovedCharCount As Long
        Dim blockAddMovedCharCount As Long
        
        Dim allReadMovedCharCount As Long
        Dim charNum As Long
        Dim allAddedMovedCharCount As Long
        
        Dim costTimeSec As Double
        Dim avrCharCount As Double
        
        Dim mainLoopCount As Long
        
        Dim logContent As String
        
        Dim appBeginTimeSec As Double
        Dim endTimeSec As Double
        Dim tmpTimeSec As Double
        Dim blockBeginTimeSec As Double
        Dim addBeginTimeSec As Double
        
        'charNum =countChar(areaStart, areaEnd)
        charNum = Selection.Characters.Count
        areaStart = Selection.Start
        areaEnd = Selection.End
        
        appBeginTimeSec = Timer()
        
        initManualCharSet       '手动指定的读音
        initSpecifiedCharSet     '需要加读音的字
        initExceptionCharSet    ' 例外字符
        
        Dim pinyinArr() As String
        
        Debug.Print ""
        
        
        ' test
        'textForPinyin = getPinyin("】")
        

        ' 收集需要加注音的字
        textForPinyin = ""
        blockText = ""
        allReadMovedCharCount = 0
        
        mainLoopCount = 0
        'blockStart = areaStart
        allAddedMovedCharCount = 0
        blockEnd = areaStart
        Selection.Start = areaStart
        Selection.End = areaStart
        Do While (allReadMovedCharCount < charNum)
        
                blockBeginTimeSec = Timer()
                
                '防止假死
                DoEvents
                
                mainLoopCount = mainLoopCount + 1
        
                Debug.Print "开始收集字"
                
                blockStart = Selection.End
                Selection.Start = blockStart
                
                'blockLoopCount = 0
                
                blockText = ""
                
                blockReadCharCount = 0
                blockReadMovedCharCount = 0
                
                Do While (allReadMovedCharCount < charNum)
                
                        '防止假死
                        DoEvents
                        
                        '进度显示
                        If allReadMovedCharCount Mod 64 = 0 Then
                                showDebugLog "阶段1/2    收集字", appBeginTimeSec, allReadMovedCharCount + 1, charNum, blockBeginTimeSec, blockReadCharCount + 1, blockCharLimit
                        End If
                        
                        ' 找句末，非汉字，或者不需要加注音的符号
                        'Do While (allReadMovedCharCount < charNum) And (blockReadCharCount < blockCharLimit)
                        
                                '防止假死
                                'DoEvents
                                
                                lastPos = Selection.End
                                Selection.Start = Selection.End
                                '光标往下移一个字
                                Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdMove
                                Selection.Start = lastPos
                                
                                allReadMovedCharCount = allReadMovedCharCount + 1
                                blockReadMovedCharCount = blockReadMovedCharCount + 1
                                
                                tmpChar = Selection.text
                                
                                If True = isNeedZhuyin(tmpChar) Then
                                        blockText = blockText + tmpChar
                                        blockReadCharCount = blockReadCharCount + 1
                                ElseIf (blockCharLimit <= blockReadCharCount) Then
                                        ' 达到限制后，遇到句末才结束。python汉字转语音模块，多音字会根据词自动选择相应的读音
                                        Exit Do
                                End If
                        
                        'Loop

                Loop
                
                blockEnd = Selection.End
                
                pinyin = getPinyin(blockText)
                pinyinArr = Split(pinyin, " ")
                pinyinNum = UBound(pinyinArr) + 1
                
                Debug.Print "开始加注音"
                
                addBeginTimeSec = Timer()
                
                ' 逐字加注音
                charPinyin = ""
                Selection.Start = blockStart
                Selection.End = blockStart
                blockAddedCharCount = 0
                blockAddMovedCharCount = 0
                Do While (blockAddMovedCharCount < blockReadMovedCharCount)
                
                        '防止假死
                        DoEvents
                        
                        '进度显示
                        If blockAddedCharCount Mod 16 = 0 Then
                                showDebugLog "阶段2/2    加注音", appBeginTimeSec, allAddedMovedCharCount + 1, charNum, addBeginTimeSec, blockAddMovedCharCount + 1, blockReadMovedCharCount
                        End If
                        
                        allAddedMovedCharCount = allAddedMovedCharCount + 1
                        blockAddMovedCharCount = blockAddMovedCharCount + 1
                        
                        Selection.Start = Selection.End
                        lastPos = Selection.End
                        Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdMove
                        Selection.Start = lastPos
                        tmpChar = Selection.text
                        
                        'Debug.Print "加注音:" & blockAddedCharCount & "/" & blockReadCharCount & "    " & tmpChar
                        
                        If True = isNeedZhuyin(tmpChar) Then
                        
                                charPinyin = pinyinArr(blockAddedCharCount)
                                
                                '如果是手动指定读音的字，就用手动指定的读音
                                getManualPinyin tmpChar, charPinyin
                                
                                
                                 ' 不需要加注音，跳出
                                'If False = isExceptionChar(tmpChar) Then
                                    
                                        '注音操作后，字的大小会增长，当前Selection.End等属性会自动变化
                                        cursor1 = lastPos
                                        cursor2 = Selection.End
                                        With Selection
                                                .SetRange Start:=cursor1, End:=cursor2
                                                .Range.PhoneticGuide text:=charPinyin, Alignment:=wdPhoneticGuideAlignmentCenter, Raise:=13, FontSize:=7, FontName:="微软雅黑"
                                                .SetRange Start:=cursor1, End:=cursor2
                                        End With
                                        
                                'End If
                                
                                blockAddedCharCount = blockAddedCharCount + 1
                                
                        End If
                        
                Loop
                
                
'                '如果add中移动光标的数量与read中的不一致，就需要修正光标位置
'                If (blockReadMovedCharCount <> blockAddMovedCharCount) Then
'                        Selection.Start = Selection.End
'                        Selection.MoveRight unit:=wdCharacter, Count:=blockReadMovedCharCount - blockAddMovedCharCount, Extend:=wdMove
'                End If
                
        Loop
        
        Set g_manualCharSet = Nothing
        Set g_specifiedCharSet = Nothing
        Set g_exceptionCharSet = Nothing
        
        ' 恢复选中区域
        Selection.Start = areaStart
        
        endTimeSec = Timer()
        
        logContent = "总耗时：" & Format((endTimeSec - appBeginTimeSec) / 24 / 3600, "hh:mm:ss") & "    总字数：" & charNum & _
            "    平均：" & Format(charNum / (endTimeSec - appBeginTimeSec), "Standard") & "字/秒"
        
        Debug.Print logContent
        MsgBox logContent
        

End Sub



