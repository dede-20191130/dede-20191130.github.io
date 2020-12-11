---
title: "[VBA] ブック内の図形内のテキストを検索・置換するマクロ（Qiitaの記事の拡張）"
author: dede-20191130
date: 2020-12-11T16:27:07+09:00
slug: Search-Shape-String
draft: false
toc: true
featured: false
tags: ['Excel', 'VBA']
categories: ['プログラミング']
archives:
    - 2020
    - 2020-12
---

## この記事について

Qiitaで見たこちらの記事を参考に、  
ブック内のすべての図形で検索できるように拡張したマクロです。  
[RelaxTools Addin](https://software.opensquare.net/relaxtools/about/)を利用できない環境や、検索機能だけほしい場合に、  
自分で使えたら便利かなと思い作成しました。  

[参考元：[Excel]図形内のテキストを検索・置換したい](https://qiita.com/s-hchika/items/dda585fa0bdb829e9713)

[<span id="srcURL"><u>説明のために作成したExcelファイルとソースコードはこちらでダウンロードできます。</u></span>](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2020/12/Search-Shape-String)



## コード

ほぼそのまんまです。  
違いといえば例外処理のいくつかの流れの変更と、  
未発見フラグと終了フラグの管理を自分的にわかりやすくしたところくらい。

### モジュール内定数・変数

```vb
Option Explicit
'**************************
'*図形の文字列検索・置換
'*
'*referencing https://qiita.com/s-hchika/items/dda585fa0bdb829e9713
'**************************

'定数
'ポップアップの名前
Private Const TITLE_SEARCH_SHAPE_TEXT As String = "オートシェイプ検索"

'変数
'無し


```

### searchShapeText

```vb
'******************************************************************************************
'*関数名    ：文字検索関数
'*機能      ：
'*引数      ：
'******************************************************************************************
Public Sub searchShapeText()

    
    '定数
    Const FUNC_NAME As String = "searchShapeText"
    
    '変数
    Dim mySheets As Variant                     'ワークシートの集合体
    Dim sheet As Variant
    Dim searchWord As String                     '検索ワード
    Dim flgTerminate As Boolean
    Dim flgFound As Boolean
    
    On Error GoTo ErrorHandler
    
    'ブック内検索orシート検索
    If MsgBox("ブック全体を検索場所としますか。", vbYesNo, TITLE_SEARCH_SHAPE_TEXT) = vbYes Then
        '対象のワークシートを現在開いているブックの全てのシートとする
        Set mySheets = ActiveWorkbook.Worksheets
    Else
        '対象のワークシートを現在開いているシートのみとする
        mySheets = Array(ActiveSheet)
    End If
    
    '検索ワード入力ポップアップを表示する
    searchWord = Trim(InputBox("検索したいワードを入力して下さい。", TITLE_SEARCH_SHAPE_TEXT))

    If searchWord = "" Then GoTo ExitHandler
    
    '検索
    For Each sheet In mySheets
        sheet.Activate
        If Not searchReplaceShapeText(sheet.Shapes, searchWord, flgTerminate, flgFound) Then GoTo ExitHandler
        '終了フラグTrueの場合
        If flgTerminate Then GoTo ExitHandler
    Next sheet
    
    'すべての検索範囲で未発見の場合
    If Not flgFound Then MsgBox "「" & searchWord & "」が見つかりません。", vbExclamation, TITLE_SEARCH_SHAPE_TEXT
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TITLE_SEARCH_SHAPE_TEXT
        
    GoTo ExitHandler
        
End Sub

```

### searchReplaceShapeText

```vb
'******************************************************************************************
'*関数名    ：searchReplaceShapeText
'*機能      ：図形内検索置換関数
'*引数      ：worksheetShapes Worksheetの図形コレクション
'*引数      ：searchWord      検索文字
'*引数      ：flgTerminate      探索終了フラグ
'*引数      ：flgFound      文字列発見フラグ
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Private Function searchReplaceShapeText(ByVal worksheetShapes As Object, ByVal searchWord As String, _
                                        ByRef flgTerminate As Boolean, ByRef flgFound As Boolean) As Boolean

    
    '定数
    Const FUNC_NAME As String = "searchReplaceShapeText"
    
    '変数
    Dim targetShape  As Excel.Shape              'ワークシート内の図形
    Dim shapeText   As String                    '図形内の文字
    Dim discoveryWord As Long                    '検索ワード発見位置
    Dim replaceWord As String                    '置換後の文字
    Dim replacePopupMsg As String                '置換ポップアップメッセージ
    Dim searchWordCnt As Long: searchWordCnt = 1 '図形内検索ワード数
    
    On Error GoTo ErrorHandler


    'ワークシートに図形が存在する間ループ
    For Each targetShape In worksheetShapes
        Do

            'クループ化された図形の時
            If (targetShape.Type = msoGroup) Then
    
                If Not (searchReplaceShapeText(targetShape.GroupItems, searchWord, flgTerminate, flgFound)) Then GoTo ExitHandler
                '終了フラグTrueの場合
                If flgTerminate Then GoTo TruePoint
    
                'コメントの時
            ElseIf (targetShape.Type = msoComment) Then
                Exit Do
            Else
                '指定したテキストフレームにテキストがあるかどうかを返す
                If (targetShape.TextFrame2.HasText) Then
    
                    '図形内のテキストを取得
                    shapeText = targetShape.TextFrame2.TextRange.Text
    
                    '図形内の文字列から検索
                    discoveryWord = InStr(shapeText, searchWord)
    
                    '検索ワードが見つかったとき、置換の処理を行う
                    If (discoveryWord > 0&) Then
                        
                        '文字列発見フラグTrue
                        flgFound = True
                        
                        'ウィンドウを図形の位置にスクロール
                        ActiveWindow.ScrollRow = targetShape.TopLeftCell.Row
                        ActiveWindow.ScrollColumn = targetShape.TopLeftCell.Column
    
                        Do While (discoveryWord > 0&)
                            
                            'テキスト範囲選択を解除するため、カレントセルを選択する
                            targetShape.TopLeftCell.Select
    
                            targetShape.TextFrame2.TextRange.Characters(discoveryWord, Len(searchWord)).Select
    
                            replacePopupMsg = "置換する場合、入力してください。" & vbNewLine & vbNewLine & "置換前 : " & searchWord & vbNewLine & "置換後"
    
                            ' 置換入力メッセージを出力する
                            replaceWord = InputBox(replacePopupMsg, "置換")
    
                            If Not replaceWord = "" Then
                            
                                '図形内の文字列を一箇所置換する
                                targetShape.TextFrame2.TextRange.Text = Replace(shapeText, searchWord, replaceWord, 1, searchWordCnt)
                                targetShape.TopLeftCell.Select
    
                            End If
    
                            '検索を継続するかどうか
                            If (MsgBox("continue?", vbQuestion Or vbOKCancel, TITLE_SEARCH_SHAPE_TEXT) <> vbOK) Then
                                flgTerminate = True
                                GoTo TruePoint
    
                                '同じ図形内で文字検索
                            Else
                                discoveryWord = InStr(discoveryWord + 1&, shapeText, searchWord)
                            End If
    
                        Loop
    
                    End If
                End If
            End If
        Loop While False
    Next
    

TruePoint:

    searchReplaceShapeText = True

ExitHandler:
    
    
    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TITLE_SEARCH_SHAPE_TEXT
        
    GoTo ExitHandler
        
End Function
```