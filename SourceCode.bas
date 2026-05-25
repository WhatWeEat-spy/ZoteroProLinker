Attribute VB_Name = "模块1"
' ==========================================
' Zotero 顶刊级排版超链接引擎 (支持内容模式/年份模式)
' ==========================================
Const DEFAULT_COLOR As Long = 16711680 ' 默认学术蓝
Const SHOW_UNDERLINE As Boolean = False

' ==========================================
' 核心 1：生成超链引擎
' ==========================================
Sub ZoteroLink_GenerateAll()
    Dim doc As Document
    Dim fld As Field, bibFld As Field, para As Paragraph
    Dim jsonText As String, bkmkName As String
    Dim i As Integer, matchCount As Integer
    Dim newLink As Hyperlink, userColor As Long, linkMode As Integer
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    ' 读取配置：颜色和排版模式
    On Error Resume Next
    userColor = CLng(doc.Variables("ZoteroLinkColor").Value)
    If Err.Number <> 0 Then userColor = DEFAULT_COLOR
    linkMode = CInt(doc.Variables("ZoteroLinkMode").Value)
    If Err.Number <> 0 Or linkMode = 0 Then linkMode = 1 ' 1: 内容剥离括号, 2: 仅限年份
    
    doc.Styles(wdStyleHyperlink).Font.Color = userColor
    doc.Styles(wdStyleHyperlink).Font.Underline = wdUnderlineNone
    doc.Styles(wdStyleHyperlinkFollowed).Font.Color = userColor
    doc.Styles(wdStyleHyperlinkFollowed).Font.Underline = wdUnderlineNone
    On Error GoTo 0
    
    ' 1. 定位参考文献表并打书签
    For Each fld In doc.Fields
        If InStr(1, fld.Code.text, "ADDIN ZOTERO_BIBL") > 0 Then Set bibFld = fld: Exit For
    Next fld
    If bibFld Is Nothing Then
        MsgBox "未找到参考文献表，请先生成 References。", vbExclamation: Exit Sub
    End If
    
    Dim bibIdx As Integer: bibIdx = 1
    For Each para In bibFld.Result.Paragraphs
        If Len(Trim(Replace(para.Range.text, vbCr, ""))) > 5 Then
            bkmkName = "ZoteroBib_" & bibIdx
            If Not doc.Bookmarks.Exists(bkmkName) Then doc.Bookmarks.Add Name:=bkmkName, Range:=para.Range.Characters(1)
            bibIdx = bibIdx + 1
        End If
    Next para
    
    ' 2. 遍历正文，精准执行模式切片
    matchCount = 0
    For Each fld In doc.Fields
        If InStr(1, fld.Code.text, "ADDIN ZOTERO_ITEM") > 0 Then
            jsonText = fld.Code.text
            Dim itemAuthors(1 To 50) As String, itemTitles(1 To 50) As String
            Dim itemCount As Integer: itemCount = 0
            Dim pos As Long: pos = 1
            
            Do
                pos = InStr(pos, jsonText, """itemData"":")
                If pos = 0 Then Exit Do
                itemCount = itemCount + 1
                itemAuthors(itemCount) = ExtractJSON(jsonText, """family"":""", pos)
                itemTitles(itemCount) = ExtractJSON(jsonText, """title"":""", pos)
                pos = pos + 10
            Loop
            
            If itemCount > 0 Then
                Dim txt As String: txt = fld.Result.text
                Dim fStart As Long: fStart = fld.Result.Start
                Dim tStart(1 To 50) As Long, tEnd(1 To 50) As Long
                
                ' 基础切割（多引注分离）
                If itemCount = 1 Then
                    tStart(1) = 1: tEnd(1) = Len(txt)
                Else
                    Dim tempArr() As String, curPos As Long
                    tempArr = Split(txt, ";")
                    If UBound(tempArr) + 1 = itemCount Then
                        curPos = 1
                        For rIdx = 1 To itemCount
                            tStart(rIdx) = curPos
                            tEnd(rIdx) = curPos + Len(tempArr(rIdx - 1)) - 1
                            curPos = tEnd(rIdx) + 2
                        Next rIdx
                    Else
                        tempArr = Split(txt, ",")
                        If UBound(tempArr) + 1 = itemCount Then
                            curPos = 1
                            For rIdx = 1 To itemCount
                                tStart(rIdx) = curPos
                                tEnd(rIdx) = curPos + Len(tempArr(rIdx - 1)) - 1
                                curPos = tEnd(rIdx) + 2
                            Next rIdx
                        Else
                            Dim segLen As Long: segLen = Len(txt) \ itemCount
                            If segLen < 1 Then segLen = 1
                            For rIdx = 1 To itemCount
                                tStart(rIdx) = (rIdx - 1) * segLen + 1
                                If rIdx = itemCount Then tEnd(rIdx) = Len(txt) Else tEnd(rIdx) = rIdx * segLen
                            Next rIdx
                        End If
                    End If
                End If
                
                ' ??【核心算法】：根据模式剥离括号或提取年份
                Dim subRanges(1 To 50) As Range
                For rIdx = 1 To itemCount
                    Dim segStr As String
                    If tEnd(rIdx) > Len(txt) Then tEnd(rIdx) = Len(txt)
                    segStr = Mid(txt, tStart(rIdx), tEnd(rIdx) - tStart(rIdx) + 1)
                    
                    If linkMode = 2 Then
                        ' 模式 2：精准定位年份
                        Dim yStart As Long, yEnd As Long
                        If FindYear(segStr, yStart, yEnd) Then
                            tEnd(rIdx) = tStart(rIdx) + yEnd - 1
                            tStart(rIdx) = tStart(rIdx) + yStart - 1
                        Else
                            ShrinkBrackets segStr, tStart(rIdx), tEnd(rIdx) ' 找不到年份则退化为模式1
                        End If
                    Else
                        ' 模式 1：仅剥离括号和空白
                        ShrinkBrackets segStr, tStart(rIdx), tEnd(rIdx)
                    End If
                    
                    Set subRanges(rIdx) = doc.Range(fStart + tStart(rIdx) - 1, fStart + tEnd(rIdx))
                Next rIdx
                
                ' 倒序打桩防坐标崩塌
                For itemIdx = itemCount To 1 Step -1
                    Dim aStr As String: aStr = itemAuthors(itemIdx)
                    Dim tStr As String: tStr = itemTitles(itemIdx)
                    If Len(tStr) > 15 Then tStr = Left(tStr, 15)
                    Dim matched As String: matched = ""
                    Dim chkIdx As Integer: chkIdx = 1
                    
                    If aStr <> "" Or tStr <> "" Then
                        For Each para In bibFld.Result.Paragraphs
                            If Len(Trim(Replace(para.Range.text, vbCr, ""))) > 5 Then
                                If (aStr = "" Or InStr(1, para.Range.text, aStr, vbTextCompare) > 0) And _
                                   (tStr = "" Or InStr(1, para.Range.text, tStr, vbTextCompare) > 0) Then
                                    matched = "ZoteroBib_" & chkIdx
                                    Exit For
                                End If
                                chkIdx = chkIdx + 1
                            End If
                        Next para
                    End If
                    
                    If matched <> "" Then
                        Set newLink = doc.Hyperlinks.Add(Anchor:=subRanges(itemIdx), Address:="", SubAddress:=matched)
                        newLink.Range.Font.Color = userColor
                        newLink.Range.Font.Underline = wdUnderlineNone
                        matchCount = matchCount + 1
                    End If
                Next itemIdx
            End If
        End If
    Next fld
    
    Application.ScreenUpdating = True
    MsgBox "超链生成完毕！已启用顶刊模式，成功连接 " & matchCount & " 处。", vbInformation
End Sub

' ==========================================
' 核心 2：卸载引擎 (洗色)
' ==========================================
Sub ZoteroLink_RemoveAll()
    Dim doc As Document, hl As Hyperlink, fld As Field
    Dim i As Long, removeCount As Integer, trackState As Boolean
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    trackState = doc.TrackRevisions
    If trackState = True Then doc.TrackRevisions = False
    removeCount = 0
    
    ' 第一波：剥离系统超链接
    For i = doc.Hyperlinks.Count To 1 Step -1
        Set hl = doc.Hyperlinks(i)
        If InStr(1, hl.SubAddress, "ZoteroBib_") > 0 Then
            On Error Resume Next
            hl.Range.Select
            Application.CommandBars.ExecuteMso "HyperlinkRemove"
            removeCount = removeCount + 1
            On Error GoTo 0
        End If
    Next i
    
    ' ??【核心修复】：第二波，强制清洗所有 Zotero 域内部格式，彻底解决残留蓝色
    For Each fld In doc.Fields
        If InStr(1, fld.Code.text, "ADDIN ZOTERO_ITEM") > 0 Then
            fld.Result.Font.ColorIndex = wdAuto
            fld.Result.Font.Underline = wdUnderlineNone
        End If
    Next fld
    
    ' 清理垃圾书签
    For i = doc.Bookmarks.Count To 1 Step -1
        If InStr(1, doc.Bookmarks(i).Name, "ZoteroBib_") > 0 Then doc.Bookmarks(i).Delete
    Next i
    
    doc.Range(0, 0).Select
    doc.TrackRevisions = trackState
    Application.ScreenUpdating = True
    MsgBox "已移除 " & removeCount & " 处链接并重置字体颜色。", vbInformation
End Sub

' ==========================================
' 核心 3：综合设置控制台 (包含颜色与模式)
' ==========================================
Sub ZoteroLink_Settings()
    Dim doc As Document: Set doc = ActiveDocument
    Dim cChoice As String, mChoice As String, finalColor As Long, finalMode As Integer
    
    cChoice = InputBox("【步骤 1/2】请选择颜色：" & vbCrLf & "1 - 学术蓝 (默认)" & vbCrLf & "2 - 纯黑色" & vbCrLf & "3 - 经典红", "排版设置", "1")
    Select Case cChoice
        Case "1": finalColor = 16711680
        Case "2": finalColor = 0
        Case "3": finalColor = 255
        Case Else: Exit Sub
    End Select
    
    mChoice = InputBox("【步骤 2/2】请选择排版模式：" & vbCrLf & vbCrLf & _
                       "1 - 内容模式 (默认)：链接全部文字。" & vbCrLf & _
                       "2 - 年份模式 (AER/QJE)：仅年份(如2024a)变超链。", "排版设置", "1")
    Select Case mChoice
        Case "1": finalMode = 1
        Case "2": finalMode = 2
        Case Else: Exit Sub
    End Select
    
    doc.Variables("ZoteroLinkColor").Value = finalColor
    doc.Variables("ZoteroLinkMode").Value = finalMode
    MsgBox "设置已保存！点击【生成超链接】立即生效。", vbInformation
End Sub

' ==========================================
' 辅助算法区：数据切片与模式剥离 (勿删)
' ==========================================
Function ExtractJSON(source As String, key As String, startFrom As Long) As String
    Dim startPos As Long, endPos As Long, nextItem As Long
    nextItem = InStr(startFrom + 1, source, """itemData"":")
    startPos = InStr(startFrom, source, key)
    If startPos > 0 And (nextItem = 0 Or startPos < nextItem) Then
        startPos = startPos + Len(key)
        endPos = InStr(startPos, source, """")
        If endPos > startPos Then ExtractJSON = Mid(source, startPos, endPos - startPos)
    Else
        ExtractJSON = ""
    End If
End Function

Sub ShrinkBrackets(ByVal segText As String, ByRef tStart As Long, ByRef tEnd As Long)
    ' 剥离首尾的括号和多余空格
    Dim trStart As Long, trEnd As Long, c As String
    trStart = 1: trEnd = Len(segText)
    Do While trStart <= trEnd
        c = Mid(segText, trStart, 1)
        If c = "(" Or c = "[" Or c = "）" Or c = "（" Or c = " " Then trStart = trStart + 1 Else Exit Do
    Loop
    Do While trEnd >= trStart
        c = Mid(segText, trEnd, 1)
        If c = ")" Or c = "]" Or c = "）" Or c = "（" Or c = " " Then trEnd = trEnd - 1 Else Exit Do
    Loop
    tEnd = tStart + trEnd - 1
    tStart = tStart + trStart - 1
End Sub

Function FindYear(ByVal text As String, ByRef outStart As Long, ByRef outEnd As Long) As Boolean
    ' 跨平台年份 (精准捕捉 19xx, 20xx 以及带字母后缀的 2024a)
    Dim i As Long
    For i = 1 To Len(text) - 3
        If Mid(text, i, 1) Like "[1-2]" And Mid(text, i + 1, 1) Like "[0-9]" And _
           Mid(text, i + 2, 1) Like "[0-9]" And Mid(text, i + 3, 1) Like "[0-9]" Then
            outStart = i: outEnd = i + 3
            If outEnd < Len(text) Then
                If Mid(text, outEnd + 1, 1) Like "[a-zA-Z]" Then outEnd = outEnd + 1
            End If
            FindYear = True
            Exit Function
        End If
    Next i
    FindYear = False
End Function

