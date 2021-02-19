---
title: "[Excel VBA] PageSetup.PrintAreaによるシートの印刷範囲の変更を行うときにエラーになる場合＆その回避方法"
author: dede-20191130
date: 2021-01-15T23:38:43+09:00
slug: Change-PageSetup-PrintAria
draft: false
toc: true
featured: false
tags: ['VBA','Excel']
categories: ['プログラミング','トラブルシューティング']
vba_taxo: specification
archives:
    - 2021
    - 2021-01
---

## この記事について

ワークシートオブジェクトのPageSetup.PrintAreaプロパティを用いて  
条件に従ってシートの印刷範囲を変更するような処理を実装したい場合があるかもしれない。

そのとき、{{< colored-span color="#fb9700" >}}セルの参照形式{{< /colored-span >}}に気をつけないと、思わぬエラーになる可能性がある。

この記事で、エラーの発生ケースとその二通りの回避方法について記したい。

[<span id="srcURL"><u>説明のために作成したExcelファイルとソースコード、テスト用データはこちらでダウンロードできます。</u></span>](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2021/01/Change-PageSetup-PrintAria)

## 検証環境

Windows 10 Home(64bit)  
MSOffice 2016

## 事例

### About

ブックのシートに印刷範囲が設定されている。

その印刷範囲を、列は変えずに印刷範囲の下限をひとつ下の行に変更する処理をVBAで記述したい。  
e.g. 印刷範囲が$A$1:$E$5ならば、関数実行後に印刷範囲が$A$1:$E$6となるようにしたい。

### コード

```vb {hl_lines=[22]}
'******************************************************************************************
'*関数名    ：changePrintAreaBeforeRevised
'*機能      ：PrintAreaをひとつ下の行に変更する 修正前
'*引数      ：
'******************************************************************************************
Public Sub changePrintAreaBeforeRevised()
    
    '定数
    Const FUNC_NAME As String = "changePrintAreaBeforeRevised"
    
    '変数
    Dim prePrintAreaAddress As String
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
    
        '現在の印刷範囲アドレス
        prePrintAreaAddress = .PageSetup.PrintArea
        
        '印刷範囲をひとつ下の行に変更する
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address '★01
        
        Debug.Print .PageSetup.PrintArea
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub
```

## R1C1参照形式だとエラー発生

### 内容

上記は、  
参照形式がデフォルト（A1参照形式）の場合はうまく動作する。

しかし、  
{{< colored-span color="#fb9700" >}}R1C1参照形式{{< /colored-span >}}を使用している場合、  
★01の箇所でエラーとなる。

![エラー](./image01.png)

### 原因

PageSetupオブジェクトのPrintAreaプロパティは、  
コード実行時点の参照形式によって取得文字列が異なる。

- R1C1参照形式の場合はR1C1形式の文字列
- A1参照形式の場合はA1形式の文字列

また、Rangeオブジェクトに指定する文字列は  
A1参照形式のみ想定され、R1C1参照形式を許容していない。

したがって、  
prePrintAreaAddressにはR1C1参照形式のアドレス文字列が格納され、  
Range(prePrintAreaAddress)としてRangeオブジェクトに格納する時点でエラーとなる。

## 回避方法

### 参照形式自体を切り替える

```vb {hl_lines=["20-21"]}
'******************************************************************************************
'*関数名    ：changePrintAreaBeforeRevised
'*機能      ：PrintAreaをひとつ下の行に変更する 修正01
'*引数      ：
'******************************************************************************************
Public Sub changePrintAreaRevised01()
    
    '定数
    Const FUNC_NAME As String = "changePrintAreaRevised01"
    
    '変数
    Dim prePrintAreaAddress As String
    Dim currentStyle As XlReferenceStyle

    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
        
        'セルの参照形式をA1形式に変更
        currentStyle = Application.ReferenceStyle
        Application.ReferenceStyle = xlA1
        
        '現在の印刷範囲アドレス
        prePrintAreaAddress = .PageSetup.PrintArea
        
        '印刷範囲をひとつ下の行に変更する
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
        'セルの参照形式を復旧する
        Application.ReferenceStyle = currentStyle
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub

```

PrintArea取得・設定前後で  
参照形式を強制的にA1参照形式に切り替える。

欠点は、  
切り替え・復旧の間に時間がかかる処理がある場合や、  
この関数を何度も呼び出す場合、  
ユーザ側の視点から、シートの参照形式の部分が交互に変わってチラつくように見えるかもしれない。

### アドレス文字列を別の参照形式に変更する

Application.ConvertFormulaを用いて  
文字列だけを変更する。

```vb {hl_lines=["22"]}
'******************************************************************************************
'*関数名    ：changePrintAreaBeforeRevised
'*機能      ：PrintAreaをひとつ下の行に変更する 修正02
'*引数      ：
'******************************************************************************************
Public Sub changePrintAreaRevised02()
    
    '定数
    Const FUNC_NAME As String = "changePrintAreaRevised02"
    
    '変数
    Dim prePrintAreaAddress As String
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
    
        '現在の印刷範囲アドレス
        prePrintAreaAddress = .PageSetup.PrintArea
        
        'アドレスをxlA1参照形式のものに修正
        If Application.ReferenceStyle = xlR1C1 Then prePrintAreaAddress = Application.ConvertFormula(prePrintAreaAddress, xlR1C1, xlA1)
        
        '印刷範囲をひとつ下の行に変更する
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub
```

## 終わりに

アドレス文字列を別の参照形式に変更する方法が最も自然で応用性も高いかと思う。
