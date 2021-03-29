---
title: "[Excel VBA] 年度に関わらず営業日数を算出する方法について紹介"
author: dede-20191130
date: 2021-03-29T16:04:41+09:00
slug: wsfunc-networkday
draft: false
toc: true
featured: false
tags: ['Excel', 'VBA']
categories: ['プログラミング']
archives:
    - 2021
    - 2021-03
---

## この記事について

### About

Excel VBAで営業日数を計算する場合、  
自作関数で細かく営業日条件を設定して計算する以外に、  
組み込み関数の`NetworkDays`を使用する方法がある。

{{< box-with-title title="書式：" >}} 
    <code>WorksheetFunction.NetworkDays(開始日, 終了日, [祝日])</code><br>
    (<a href="https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheetfunction.networkdays" target="_blank" rel="noopener">NetworkDaysの仕様</a>)
{{< /box-with-title >}}



### 使用例

使用例としては、通常、次のようになる。
1. 祝日（土日以外）のリストをいずれかのシート上に作成する。
2. VBAコード上で、上記リストのセル範囲を祝日引数として指定し、営業日数を算出する。

![祝日リスト](./img01.png)

```vb {hl_lines=["21"]}
'******************************************************************************************
'*機能      ：営業日の計算　祝日としてセル範囲使用
'*引数      ：
'******************************************************************************************
Public Sub CalcWorkDayUseRange()
    
    '定数
    Const FUNC_NAME As String = "CalcWorkDayUseRange"
    
    '変数
    
    '営業日：祝日を除く月～金曜日とする
    'ex.1)2019/12/29から2020/1/6までの営業日を計算する
    Debug.Print WorksheetFunction.NetworkDays(#12/29/2019#, #1/6/2020#, ThisWorkbook.Worksheets("祝日").Range("B2:B9")) '=3
    'ex.2)2020/6/14から2020/6/16までの営業日を計算する
    Debug.Print WorksheetFunction.NetworkDays(#6/14/2020#, #6/16/2020#, ThisWorkbook.Worksheets("祝日").Range("B2:B9")) '=1
    
    'note)祝日シートには2020年の祝日カレンダーしかないため、
    '   他の年の祝日をまたぐ営業日の計算するためには、対象年のカレンダーを追加する必要がある
    '   ex)2021/6/14から2021/6/16までの営業日を計算すると、6/15（会社設立記念日）が祝日判定されず、3が返る
    Debug.Print WorksheetFunction.NetworkDays(#6/14/2021#, #6/16/2021#, ThisWorkbook.Worksheets("祝日").Range("B2:B9")) '=3
    
    
ExitHandler:

    Exit Sub
    
End Sub

```

### 課題点

上のコードのハイライトで示されているように、  
上記の方法だと、  
祝日シートに記述されていない年（2019年以前、2021年以降）の祝日またぎの計算は苦手である。

各年ごとの祝日をシートに準備しなければならないが、  
それが作業量とミスの可能性を増やすため、  
この記事では年度に関わらず営業日数を算出する方法を紹介したい。

## 方法

### NetworkDays関数の[祝日]引数はセル範囲だけではない

NetworkDays関数の第三引数（[祝日]）として  
セル範囲の他に、配列引数や{{< colored-span color="#fb9700" >}}Date型の配列変数{{< /colored-span >}}をとることができる。

VBAにおいては、  
{{< colored-span color="#fb9700" >}}Date型の配列変数{{< /colored-span >}}を利用すると  、
年に依らない日数の計算が容易になるかと考えられる。

よって、以下では、  
シートに祝日を記述するのではなく、  
配列としてVBA上にハードコードすることを基本方針としている。

### 配列変数取得用のClassを作成

まず、それぞれの祝日を格納している配列変数を取得するClassを作成した。

```vb {hl_lines=["49-56"]}
'@Folder("Class")
Option Explicit

'**************************
'クラス名：ClsSpecialHoliday
'*祝祭日の設定・取得
'**************************

'定数欄
Private Const SOURCE_NAME As String = ""

'変数欄
Private lArrHoliday() As String '祝祭日（月日のみ）格納先


'******************************************************************************************
'*getter/setter欄
'******************************************************************************************

'******************************************************************************************
'*引数      ：対象年
'******************************************************************************************
Public Property Get arrHoliday(ByVal yy As String) As Date()
    Dim dateArr() As Date
    Dim i As Long
    ReDim dateArr(0 To UBound(lArrHoliday))
    
    For i = 0 To UBound(lArrHoliday)
        dateArr(i) = CDate(yy & "/" & lArrHoliday(i))
    Next i
    arrHoliday = dateArr
End Property



'******************************************************************************************
'*機能      ：Class_Initialize
'*引数      ：
'******************************************************************************************
Private Sub Class_Initialize()

    '定数
    Const FUNC_NAME As String = "Class_Initialize"
    
    '変数
    
    '***ここで通年の祝祭日を設定します（○月/○日）***
    ReDim lArrHoliday(0 To 7)
    lArrHoliday(0) = "1/1"
    lArrHoliday(1) = "1/2"
    lArrHoliday(2) = "1/3"
    lArrHoliday(3) = "4/29"
    lArrHoliday(4) = "5/3"
    lArrHoliday(5) = "5/4"
    lArrHoliday(6) = "5/5"
    lArrHoliday(7) = "6/15"
    '*************************************
    

ExitHandler:

    Exit Sub
    
        
End Sub


```

クラスの利用時は、  
インスタンスを生成した後に  
`配列変数 = object.arrHoliday(string: 対象年)`というように利用する。

もし祝日を追加/削除したければ、  
上のコードでハイライトした配列の設定部分を追加したり削除したりすればいい。


### 営業日の計算

営業日の計算用の関数を書き直すと、  
以下のようになる。

```vb {hl_lines=[""]}
'******************************************************************************************
'*機能      ：営業日の計算　祝日としてハードコードされた値を使用
'*引数      ：
'******************************************************************************************
Public Sub CalcWorkDay()
    
    '定数
    Const FUNC_NAME As String = "CalcWorkDay"
    
    '変数
    Dim objSpecialHoliday As New ClsSpecialHoliday
    
    '営業日：祝日を除く月～金曜日とする
    'ex.1)2019/12/29から2020/1/6までの営業日を計算する
    Debug.Print WorksheetFunction.NetworkDays(#12/29/2019#, #1/6/2020#, objSpecialHoliday.arrHoliday("2020")) '=3
    'ex.2)2020/6/14から2020/6/16までの営業日を計算する
    Debug.Print WorksheetFunction.NetworkDays(#6/14/2020#, #6/16/2020#, objSpecialHoliday.arrHoliday("2020")) '=1
    
    'ex.3)2021/6/14から2021/6/16までの営業日を計算する
    '   引数として2021を渡せば2021年の祝日として配列を取得できるため、
    '   6/15（会社設立記念日）が祝日判定され、2が返る
    Debug.Print WorksheetFunction.NetworkDays(#6/14/2021#, #6/16/2021#, objSpecialHoliday.arrHoliday("2021")) '=2
    
    
ExitHandler:

    Exit Sub
    
End Sub

```

書き直し前とは異なり、  
2021/6/14から2021/6/16までの営業日がかんたんに計算できている。

## Furthermore

今回ClsSpecialHolidayクラスにハードコードした祝日配列の内容を  
ファイルから取得したりデータベースを利用したりすれば、  
さらにメンテナンス性は向上するかと思う。

