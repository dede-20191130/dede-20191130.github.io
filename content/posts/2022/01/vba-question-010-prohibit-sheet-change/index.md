---
title: "[教えて！VBA] 第10回 Excelのシートの移動・削除を禁止する方法 ＆ マクロ実行時だけ許可するにはどうすればいいの？？"
author: dede-20191130
date: 2022-01-11T13:47:44+09:00
slug: vba-question-010-prohibit-sheet-change
draft: false
toc: true
featured: false
tags: ["VBA","Excel"]
categories: ["プログラミング"]
archives:
    - 2022
    - 2022-01
---


![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641887344/learnerBlog/vba-question-010-prohibit-sheet-change/prohibit-sheet-change_levful.png)

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    ブックの保護機能を利用し、シートの移動・削除を禁止できます。<br/>
    マクロからシート構成をいじりたい場合は、一時的に解除して再度復旧するようなコードの組み方をすることが必要です。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
VBAマクロに関する質問のうち、  
皆が疑問に思っているトピックについて解説いたします。

今回は、  
{{< colored-span color="#fb9700" >}}
Excelのシートの移動・削除を禁止する方法と、  
それを一時的に許可してマクロの処理を実行する方法
{{< /colored-span >}}  
を解説いたします。

※この記事は、Office VBAマクロのうち  
Excel VBAマクロに関するトピックです。

レベル：<b>初級者向け</b>

## やりたいこと

マクロを組んだExcelブックを他の人にツールとして配るとき、  
シートの構成をいじられたら困る場合があるでしょう。

例えば、  
非表示シートに商品データシートを置いておいてマクロから参照する場合、  
商品データシートの名前を変えられたら参照できなくなる可能性があり、  
また、商品データシートを削除されたら、データが失われます。

![ユーザ操作により不具合が発生する例](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641887344/learnerBlog/vba-question-010-prohibit-sheet-change/prohibit-sheet-change2_bwacab.png)

よって、  
シート構成をいじることができないように  
プロテクト（保護）を掛ける必要があります。

また、シート構成を操作するようなマクロ実行時には、  
一時的にプロテクトを解除するような手続きが必要になるでしょう。

以下の手順セクションでは、  
シート構成の保護のやり方、  
およびマクロのコードについて説明します。
## 環境

以下は、  
Office 2016のExcel環境での説明です。  

※2022/1時点の最新バージョンのExcelでも内容は変わりません。

## 保護機能について

まずExcelの保護機能について簡単に説明します。

ユーザーが誤ってデータの一部を毀損することを防ぐために、  
Excelに備わっている動作制限機能のことを保護機能と呼びます。

校閲タブから操作できる保護機能は、  
{{< colored-span color="#fb9700" >}}シートの保護{{< /colored-span >}}と
{{< colored-span color="#fb9700" >}}ブックの保護{{< /colored-span >}}の
二種類に分類されます。

|保護の種類|対象|内容|
|-|-|-|
|ブックの保護|シート構成<br/>（ワークシートの数、名前など）|非表示のワークシートの表示、ワークシートの追加、<br/>移動、削除、非表示、ワークシートの名前変更を<br/>行うことができないようになる|
|シートの保護|シート上のオブジェクト|セル書き込み、セル追加・削除、図形の位置変更など<br/>（細かく設定が可能）|

## 手順

### レベル1. シートの移動・削除を禁止する

シートの移動・削除を禁止する、すなわちシート構成の保護を実施するためには、  
校閲タブの「ブックの保護」ボタンを押下します。

![校閲タブの「ブックの保護」](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641887343/learnerBlog/vba-question-010-prohibit-sheet-change/prohibit-sheet-change3_sm34zl.png)

保護のためのパスワードを入力し、実行します。

![ブックの保護設定画面](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641887343/learnerBlog/vba-question-010-prohibit-sheet-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-11_153805_vyzs7i.png)

シート構成が保護され、非表示メニューや名前変更メニューがグレーアウトされて  
選択できなくなります。

![シート構成保護](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641887343/learnerBlog/vba-question-010-prohibit-sheet-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-11_153835_lqhtfo.png)

### レベル2. 禁止を一時的に許可する

マクロからブックの保護を解除することができます。

```vb
ThisWorkbook.Unprotect PassWord
```

解除したいワークブックの`Unprotect`メソッドを呼び出して、  
引数にパスワードを指定します。

実行後、  
グレーアウトされていた非表示メニューや名前変更メニューが再度選択可能になっています。

![シート構成保護の解除](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641887343/learnerBlog/vba-question-010-prohibit-sheet-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-11_155647_vgukim.png)

### レベル3. マクロ実行時だけ許可し、復旧する

次のように、  
GOTOの機構（[例外処理の解説ページ](https://dede-20191130.github.io/learnerBlog/posts/2020/12/05/vba-exception-handle/)で詳細に説明しています）を使って、  
必ず保護の復旧の処理が実行されるようにします。

このようにしないと、  
エラーが起きた時に復旧されずにマクロが終了する可能性があるためです。

```vb
Public Sub マクロ実行時だけ許可()
    
    On Error GoTo ErrorHandler
    
    '//一時的に解除
    ThisWorkbook.Unprotect "test"
    
    '//シート構成を操作する処理……
    
ExitHandler:
    
    '//保護の復旧
    ThisWorkbook.Protect "test"
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。", vbCritical
        
    GoTo ExitHandler
        
End Sub

```
## デモ
### ABOUT
シートを追加したり消去したりするボタンを動作させるデモです。

実際にシートをイチから作成しているわけではなく、  
非表示にしていたシートを追加したり、表示中のシートを非表示化するマクロです。

### 前提
まず、  
シートが4種類あります。  
そのうち、接頭辞が「追加シート」であるシートが３つあります。

![デモ シート構成](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641887344/learnerBlog/vba-question-010-prohibit-sheet-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-11_162249_bz1cem.png)

シート1のシート上に、  
シート追加ボタン、シート消去ボタンを作成します。

![デモ用ボタン配置](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641887344/learnerBlog/vba-question-010-prohibit-sheet-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-11_162343_epq8y2.png)

それぞれのボタンにマクロを登録します。

```vb
Public Sub シート追加()
    
    Dim additinalSheetNumber As Long
    Dim vSheet As Object
    
    On Error GoTo ErrorHandler
    
    '//一時的に解除
    ThisWorkbook.Unprotect "test"
    
    '//表示中の追加シートの数を算出
    additinalSheetNumber = 0
    For Each vSheet In ThisWorkbook.Worksheets
        If vSheet.Visible = True And InStr(vSheet.Name, "追加シート") > 0 Then
            additinalSheetNumber = additinalSheetNumber + 1
        End If
    Next vSheet
    
    '//追加シートがすべて表示されているならば処理を終了
    If additinalSheetNumber = 3 Then GoTo ExitHandler
    
    '//次の追加シートを表示
    ThisWorkbook.Worksheets("追加シート" & (additinalSheetNumber + 1)).Visible = True
    
ExitHandler:
    
    '//保護の復旧
    ThisWorkbook.Protect "test"
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。", vbCritical
        
    GoTo ExitHandler
        
End Sub
```

```vb

Public Sub シート消去()
    
    Dim additinalSheetNumber As Long
    Dim vSheet As Object
    
    On Error GoTo ErrorHandler
    
    '//一時的に解除
    ThisWorkbook.Unprotect "test"
    
    '//表示中の追加シートの数を算出
    additinalSheetNumber = 0
    For Each vSheet In ThisWorkbook.Worksheets
        If vSheet.Visible = True And InStr(vSheet.Name, "追加シート") > 0 Then
            additinalSheetNumber = additinalSheetNumber + 1
        End If
    Next vSheet
    
    '//追加シートがすべて非表示ならば処理を終了
    If additinalSheetNumber = 0 Then GoTo ExitHandler
    
    '//最も連番の大きい追加シートを非表示
    ThisWorkbook.Worksheets("追加シート" & additinalSheetNumber).Visible = False
    
ExitHandler:
    
    '//保護の復旧
    ThisWorkbook.Protect "test"
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。", vbCritical
        
    GoTo ExitHandler
        
End Sub

```

３つの追加シートを、  
ボタンを押すたびに次々と追加（消去）する処理を持っています。

それぞれの処理の最後に、  
かならず保護を復旧するように組んでいます。

### デモ

次の動画のように、  
シートを追加したり消去したりできます。

また、➕マークがグレーアウトされていることからわかるように、  
ブックの保護は継続されています。

{{< video src="https://res.cloudinary.com/ddxhi1rnh/video/upload/v1641891076/learnerBlog/vba-question-010-prohibit-sheet-change/demo_auaenj.webm" max_width=600px is_bundle=false >}}

## デモで使用したファイルについて

[こちら](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2022/01/prohibit-sheet-change)からダウンロードできます。


## 終わりに

保護機能を使いこなせば、  
ユーザーフレンドリーで安全なツールを作成することができます。

機会があればシートの保護機能のほうについても  
解説する記事を作成したいと思っています。
