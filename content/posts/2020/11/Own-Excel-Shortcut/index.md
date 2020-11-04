---
title: "[Excel VBA] 個人的に作業がはかどった自作Excelショートカット"
author: dede-20191130
date: 2020-11-05T00:48:58+09:00
slug: Own-Excel-Shortcut
draft: false
toc: true
tags: ['Excel', 'VBA','自作' ,'ツール']
categories: ['アプリケーション', 'プログラミング']
archives:
    - 2020
    - 2020-11
---

## この記事について

Excelから文章を別のファイルや文書に移したり、  
Excelのシート上で四角形や矢印などの図形の切り貼りや、図形でフローチャート、模式図を作成する機会はわりと多いかと思う。  

そうした際に、  
VBAを用いて、Excelにもともと備わっていないショートカットをいくつか登録しておくと、作業スピードが上がったため、  
紹介したい。

ちなみに、[クイックアクセスツールバー](https://www.relief.jp/docs/003103.html)を用いて、  
[Altキー　+　数字] でいくつかのショートカットキーを拡張することはできるが、  
個人的にAltキーからの数字キー押下はあまり押しやすいキーではないため、最小限でしか登録はしていない。  
また、クイックアクセスツールバーで対応が難しいショートカットも存在する。

## 一覧

このようなショートカットを作成した。

|機能|どんなときに使う？|割り当てたキー|
|--|--|--|
|対象のセルから<br>テキストのみコピーする|文字列両端の<br>ダブルクォーテーション無しで<br>セルの文章を取得したいとき|Ctrl + Shift + K|
|選択したオブジェクト（画像、図形等）の<br>最前面化・最背面化|入り組んだ模式図などを<br>作成するとき|Ctrl + Shift + B|
|Excelのイベント一時停止|マクロ有ブックで、<br>開くと自動的にフォームなどが開かれるExcelブックを、<br>それらの処理なしで開きたいとき|Ctrl + Shift + M|

## 自作ショートカットのためのマクロ登録方法

こちらに分かりやすい記事がありました。  
["エクセル初心者も出来る！マクロでショートカットキーを作成する方法"](https://fastclassinfo.com/entry/marco_shortcuts/)

①Alt + F11キーでVBEを開く。  
②標準モジュールにプロシージャを記入する。  
③ブックの画面に戻り、Alt + F8を押して、マクロの設定画面を開く。  
④プロシージャを選択し、オプションボタンからマクロ呼出のショートカットキーを登録する。

## どのキーにマクロを登録すべきか？

下記の記事に既存のショートカットが列挙されている。  
["【2020年最新】Excelのショートカットキー全230個の一覧表！ - Electrical Information"](https://detail-infomation.com/excel-shortcut-key/)

このうち、まだ何も確保されていないキーを登録するとキーが競合せずにすむため都合が良い。  
Ctrl + Shift + K、M、Nなどが良いかもしれない。

## 個別説明

### 対象のセルからテキストのみコピーするショートカット

エクセルでセルをコピーしてほかのアプリケーション（メモ帳とか）に張り付けしようとすると、
セル内容のほかに改行コード（LF）がついてくる。  
これが結構タスクの際に邪魔になることがある（消すためにいちいちBSキーを打つのはめんどくさい）。

エクセルに、「値の貼り付け」機能は存在するのに「値のコピー」機能が存在しないことがこういう悩みを生むんだろう。

これについて対策をするショートカット。

#### コード

```vb
'******************************************************************************************
'*関数名    ：copyCellValueToCB
'*機能      ：アクティブセル内容をクリップボードにコピー
'*引数(1)   ：無し
'******************************************************************************************
Public Sub copyCellValueToCB()

    '定数
    Const FUNC_NAME As String = "copyCellValueToCB"

    '変数

    On Error GoTo ErrorHandler
    '---以下に処理を記述---


    'クリップボードに文字列を格納
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = CStr(ActiveCell.Value)
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With

ExitHandler:

    Exit Sub

ErrorHandler:

        MsgBox "エラーが発生しましたので終了します" & _
                vbLf & _
                "関数名：" & FUNC_NAME & _
                vbLf & _
                "エラー番号" & Err.Number & Chr(13) & Err.Description, vbCritical

        GoTo ExitHandler

End Sub
```

これでキー押下だけで改行コード無しの値のコピーが可能になった。


### 選択したオブジェクト（画像、図形等）の最前面化・最背面化

フローチャートや模式図、組織図などをExcelで作成するときに、  
矢印線をテキストボックスや画像の背面に（あるいは画像などを他の図形の前面に）移動させて  
被っている部分を調整したい機会があるかと思う。

その際に最前面に移動の処理を右クリックから呼び出すのは遅いので、  
作業時間を大幅に削減する。

#### コード（最前面の場合）

```vb
'******************************************************************************************
'*関数名    ：ZOrderToFront
'*機能      ：
'*引数(1)   ：
'******************************************************************************************
Public Sub ZOrderToFront()
    
    '定数
    Const FUNC_NAME As String = "ZOrderToFront"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    Selection.ShapeRange.ZOrder msoBringToFront

ExitHandler:

    Exit Sub
    
ErrorHandler:
    
    If Err.Number = 438 Then
        MsgBox "オブジェクトを選択してから実行してください。", vbExclamation, "警告"
    Else
        MsgBox "エラーが発生したため、マクロを終了します。" & _
               vbLf & _
               "関数名：" & FUNC_NAME & _
               vbLf & _
               "エラー番号：" & Err.Number & vbNewLine & _
               Err.Description, vbCritical, "エラー"
    End If
    GoTo ExitHandler
        
End Sub

```

最背面の移動のショートカットの場合は、  
次のように置き換える。

```vb
Selection.ShapeRange.ZOrder msoSendToBack
```

### Excelのイベント一時停止	

こんなケースを想定。  
《マクロ付エクセルブックを編集するとき、起動時に実行するイベントプロシージャを働かせないで起動したい》  
《その他、シートアクティブ時のイベント等の「実行してほしくないプロシージャ」を一旦中断したい》

#### コード

下記のQiita記事で紹介したとおり。  
["【Excel VBA】起動時に実行するマクロが鬱陶しいブックを編集したいときのツールの作成 - Qiita"](https://qiita.com/dede-20191130/items/845fd382fe00ce18f767)

## 終わりに

何か他に便利そうなショートカットを見つけたら順次更新したい。