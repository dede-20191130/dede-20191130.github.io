---
title: "[Access VBA] 数式エラーの発生しているExcelシートをテーブルに取り込むときの注意点とサンプルコード"
author: dede-20191130
date: 2022-01-22T12:15:03+09:00
slug: import-calc-error-value
draft: false
toc: true
featured: false
tags: ["VBA","Access"]
categories: ["プログラミング"]
vba_taxo: vba_coding_sample
archives:
    - 2022
    - 2022-01
---

![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642835775/learnerBlog/import-calc-error-value/import-calc-error-value1_ozm6gj.png)

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    数式エラーが発生しているセルに対しては、<br/>
    エラーをキャッチし、代わりの値を設定するなどの対処が必要です。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
AccessからExcelシートのデータを取り込む際に、  
数式エラーが発生しているセルがあった場合を取り扱います。

その場合の注意点と、  
エラーの回避方法について、  
サンプルコード付きで解説します。

## 環境

以下は、  
Office 2019のAccess環境で検証済みです。  

※2022/1時点の最新バージョンのOfficeでも内容は変わりません。

## 現象
### 数式エラーとは

数式エラーについては、  
Excelでテーブルの作成や集計、分析などを行った経験がある人であれば  
見たことがあるかと思います。

セルに値を入力した後、  
その値が不正なデータであった場合には、  
#（シャープ）から始まるエラー値が表示されます。

![数式エラーの一例](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642835775/learnerBlog/import-calc-error-value/import-calc-error-value2_jkw7l0.png)

それぞれのエラー値には、  
別々の原因が存在し、  
VBAで参照した際には個別の「エラー番号」を含む特殊な値を返すため、  
戻り値を受け取る変数は、Variant 型である必要があります。

|表示されるエラー|エラー番号|
|:----|:----|
|#DIV/0!|2007|
|#VALUE!|2015|
|#NAME?|2029|
|#REF!|2023|



### 数式エラーを持つセルを取り込む場合

数式エラーは、  
シートの作業時になるべく排除しておくべきではあるでしょう。

しかしながら、  
別のシステムで自動作成されたシートに含まれる数式エラーなどは、  
あらかじめ排除しておくことが困難でしょう。

そのようなときに、  
Accessでシートのテーブルデータを取り込むと、  
予期しないエラーが発生し、  
悩まさせるかもしれません。

例えば、  
次のように、ある列のデータに数式エラーが発生している売上管理テーブルを取り込むことを考えます。

![Excelデータサンプル](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642835775/learnerBlog/import-calc-error-value/import-calc-error-value3_l0e4pn.png)

Access側のテーブル構造もシートに寄せます。

|フィールド名|データ型|
|:----|:----|
|管理番号|数値型|
|店舗|短いテキスト|
|売上金|数値型|
|調整後売上金|数値型|

#### TransferSpreadsheetメソッド使用デモ

まず、  
Accessの`DoCmd.TransferSpreadsheet`メソッドを利用して  
取り込むデモを行います。

```vb
Sub Excelインポート_Docmd使用_エラー発生デモ()
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, _
                              "売上管理テーブル", Application.CurrentProject.Path & "\" & "数式エラーのある表.xlsx", True, "売上管理テーブル!"
End Sub

```

実行すると、  
「キー違反」のアラートが表示されてしまいます。

![「キー違反」のアラート](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642835776/learnerBlog/import-calc-error-value/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-22_140304_zlktsq.png)

「はい」をクリックすると、  
数式エラーの発生したセルの値だけがからっぽのデータが挿入されます。

![アラート後のテーブル](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642835776/learnerBlog/import-calc-error-value/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-22_140552_jrb8rd.png)

エラーメッセージが発生することや、  
からっぽのデータが挿入されることは、  
おそらく望んでいた挙動に沿う処理ではないでしょう。

#### DAOで手続き的に取り込むデモ

`DAO`を用いて、  
手続き的に取り込む場合は、  
挙動が少々異なります。

{{% notice tip 手続き的とは %}}
「手続き的（Imperative ）」とは、  
処理の記述の方法の種類の一つで、  
行いたい処理を細かく一行ずつ書いていくことを指します。

対義語は宣言的（Declarative ）で、  
処理の記述の際に、  
行いたい処理の内容を端的に示す名前を持つ構文や関数名だけを書けば処理が行われることを指します（SQLなど）。

宣言的な関数を使うことは便利ですが、VBAにおいては自由度は手続き的な方が高いように感じます。
{{% /notice %}}

```vb

Sub Excelインポート_DAO使用_エラー発生デモ()
    Const xlUp = -4162
    Dim excelPath As String
    Dim exApp As Object
    Dim wb As Object
    Dim sheetValues() As Variant
    Dim rs As Recordset
    Dim i As Long

    On Error GoTo Err
    
    excelPath = Application.CurrentProject.Path & "\" & "数式エラーのある表.xlsx"
    
    '//Excelアプリを立ち上げる
    Set exApp = CreateObject("Excel.Application")
    Set wb = exApp.Workbooks.Open(excelPath, , True)
    '//売上管理シートを参照する
    With wb.Worksheets(1)
        '//売上管理シート上のテーブルデータを、行数のぶんだけ参照し、
        '//その内容を二次元配列にパースする
        sheetValues = .Range( _
                      .Cells(2, 1), _
                      .Cells( _
                      .Cells(.Rows.Count, 1).End(xlUp).Row, _
                      4 _
                      ) _
        ).Value
    End With
    
    '//売上管理テーブルのレコードセットを開く
    Set rs = CurrentDb.OpenRecordset("売上管理テーブル", dbOpenDynaset)
    
    With rs
        For i = LBound(sheetValues, 1) To UBound(sheetValues, 1)
            '//対応するフィールドに、データを入れていく
            .AddNew
        
            .Fields("管理番号").Value = sheetValues(i, 1)
            .Fields("店舗").Value = sheetValues(i, 2)
            .Fields("売上金").Value = sheetValues(i, 3)
            '//※※　エラーの発生
            .Fields("調整後売上金").Value = sheetValues(i, 4)
        
            .Update
        Next i
        
    End With

Exits:

    rs.Close
    wb.Close
    exApp.Quit
    
    Exit Sub

Err:

    MsgBox Err.Description, vbExclamation, Err.Number
    GoTo Exits
    
End Sub


```

これを実行すると、  
「Fields2オブジェクトのエラー」という文言で、  
調整後売上金を挿入する際にエラーが発生します。

![DAO使用処理のエラー](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642835776/learnerBlog/import-calc-error-value/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-22_142007_nbnxbq.png)


### どうすればいいのか？

では、どうすればエラーを回避する、  
あるいは利用者にエラーについてアラートしてあげることができるのか？

それについて、次のセクションで見ていきます。

## 回避策


### ABOUT

発生したエラーを検知するには、  
`TransferSpreadsheet`メソッドではなく、  
`DAO`を用いてコードを記述する必要があります。

エラー発生箇所において、  
`VBA.Information`のメンバーである`IsError`メソッドを利用して、  
エラー発生有無を検知します。

`IsError(value)`は、valueがエラーの場合のみTrueを返します。

```vb
'//※※　例：エラーの発生箇所で条件分岐を行う
.Fields("調整後売上金").Value = IIF(IsError(sheetValues(i, 4)),エラー発生した場合の値,エラー発生なしの場合の値)
```

なお、  
説明で使用したファイルについて、  
[こちら](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2022/01/import-calc-error-value)からダウンロードできます。

### 1. エラーの場所を教えてあげる

エラー検知した際にエラーの発生箇所を教えてあげるためには、  
次のようにエラー情報を変数に退避し、メッセージなどに渡す必要があります。

なお、このケースでは、エラー発生時はテーブルにはデータを格納しないものとします。

![エラーの場所を教えてあげる処理フロー](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642835775/learnerBlog/import-calc-error-value/import-calc-error-value4_xslcux.png)

```vb

Sub Excelインポート_エラー発生箇所を通知()
    Const xlUp = -4162
    Dim excelPath As String
    Dim exApp As Object
    Dim wb As Object
    Dim sheetValues() As Variant
    Dim myWorkspase As Workspace
    Dim myDB As DAO.Database
    Dim rs As Recordset
    Dim i As Long
    Dim errAlertText As String
    Dim isUpdatable As Boolean
    Dim isCommit As Boolean

    On Error GoTo Err
    
    excelPath = Application.CurrentProject.Path & "\" & "数式エラーのある表.xlsx"
    
    '//Excelアプリを立ち上げる
    Set exApp = CreateObject("Excel.Application")
    Set wb = exApp.Workbooks.Open(excelPath, , True)
    '//売上管理シートを参照する
    With wb.Worksheets(1)
        '//売上管理シート上のテーブルデータを、行数のぶんだけ参照し、
        '//その内容を二次元配列にパースする
        sheetValues = .Range( _
                      .Cells(2, 1), _
                      .Cells( _
                      .Cells(.Rows.Count, 1).End(xlUp).Row, _
                      4 _
                      ) _
        ).Value
    End With
    
    '//トランザクションの開始
    Set myWorkspase = DBEngine.Workspaces(0)
    myWorkspase.BeginTrans
    isCommit = False
    
    '//DBの取得
    Set myDB = myWorkspase.Databases(0)
    
    '//売上管理テーブルのレコードセットを開く
    Set rs = myDB.OpenRecordset("売上管理テーブル", dbOpenDynaset)
    
    With rs
        errAlertText = ""
        For i = LBound(sheetValues, 1) To UBound(sheetValues, 1)
            '//対応するフィールドに、データを入れていく
            .AddNew
            
            '//サブ関数を呼び出し、エラーが発生しなかった場合のみフィールドに値を代入する
            If Excelインポート_エラー発生箇所を通知_サブ関数_エラー検知(i, sheetValues(i, 1), "管理番号", errAlertText) Then
                .Fields("管理番号").Value = sheetValues(i, 1)
            Else
                isUpdatable = True
            End If
            
            '//管理番号と同様
            If Excelインポート_エラー発生箇所を通知_サブ関数_エラー検知(i, sheetValues(i, 2), "店舗", errAlertText) Then
                .Fields("店舗").Value = sheetValues(i, 2)
            Else
                isUpdatable = True
            End If
            
            '//管理番号と同様
            If Excelインポート_エラー発生箇所を通知_サブ関数_エラー検知(i, sheetValues(i, 3), "売上金", errAlertText) Then
                .Fields("売上金").Value = sheetValues(i, 3)
            Else
                isUpdatable = True
            End If
            
            '//管理番号と同様
            If Excelインポート_エラー発生箇所を通知_サブ関数_エラー検知(i, sheetValues(i, 4), "調整後売上金", errAlertText) Then
                .Fields("調整後売上金").Value = sheetValues(i, 4)
            Else
                isUpdatable = True
            End If
            
            '//エラー発生時は新規行の挿入をキャンセル
            If isUpdatable Then
                .Update
            Else
                .CancelUpdate
            End If
            
        Next i
        
    End With
    
    '//変更をコミットする
    '//エラー発生時はコミットせず、エラーメッセージを表示する
    If errAlertText = "" Then
        myWorkspase.CommitTrans
        isCommit = True
    Else
        errAlertText = "下記の行・項目において数式エラーが発生しました。" & vbLf & vbLf & errAlertText
        MsgBox errAlertText, vbExclamation, "数式エラー"
    End If
    

Exits:
    
    '//コミットしていなければロールバックする
    If Not isCommit Then myWorkspase.Rollback
    
    rs.Close
    wb.Close
    exApp.Quit
    
    Exit Sub

Err:

    MsgBox Err.Description, vbExclamation, Err.Number
    GoTo Exits
    
End Sub


'******************************************************************************************
'*機能      ：指定した値のエラーを検知する
'*引数      ：行番号
'*引数      ：調査するセル値
'*引数      ：項目名
'*引数      ：エラー情報格納用テキスト
'*戻り値    ：True > エラーなし、False > エラー発生
'******************************************************************************************
Function Excelインポート_エラー発生箇所を通知_サブ関数_エラー検知(ByVal rowNumber As Long, ByVal sheetValue As Variant, _
ByVal tgtItem As String, ByRef errAlertText As String) As Boolean
    
    If IsError(sheetValue) Then
        '//エラー発生時、エラーテキストに情報を追加し、戻り地をFalseとする
        errAlertText = errAlertText & rowNumber & "行目：" & tgtItem & vbLf
        Excelインポート_エラー発生箇所を通知_サブ関数_エラー検知 = False
    Else
        Excelインポート_エラー発生箇所を通知_サブ関数_エラー検知 = True
    End If

End Function

```

こちらを実行すると、  
数式エラーの発生箇所を記録し、  
次のようにエラーメッセージが表示されます。

![記録されたエラー情報のメッセージ](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642835776/learnerBlog/import-calc-error-value/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-22_154356_zgz3gv.png)

ただAccess標準のエラーが出るよりも、  
こちらのほうがユーザフレンドリーですね。

### 2. エラー時に代わりに特定の値を入れる

エラー発生時に逐一メッセージを出すよりも、  
あらかじめ決められた規則で代打の値を代入したい場合もあるでしょう。

そのような場合のサンプルコードを記しました。

![エラー時に代わりに特定の値を入れる処理フロー](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642835775/learnerBlog/import-calc-error-value/import-calc-error-value5_xvi0cn.png)

```vb

Sub Excelインポート_エラー発生時に代わりの値を代入()
    Const xlUp = -4162
    Dim excelPath As String
    Dim exApp As Object
    Dim wb As Object
    Dim sheetValues() As Variant
    Dim rs As Recordset
    Dim i As Long

    On Error GoTo Err
    
    excelPath = Application.CurrentProject.Path & "\" & "数式エラーのある表.xlsx"
    
    '//Excelアプリを立ち上げる
    Set exApp = CreateObject("Excel.Application")
    Set wb = exApp.Workbooks.Open(excelPath, , True)
    '//売上管理シートを参照する
    With wb.Worksheets(1)
        '//売上管理シート上のテーブルデータを、行数のぶんだけ参照し、
        '//その内容を二次元配列にパースする
        sheetValues = .Range( _
                      .Cells(2, 1), _
                      .Cells( _
                      .Cells(.Rows.Count, 1).End(xlUp).Row, _
                      4 _
                      ) _
        ).Value
    End With
    
    '//売上管理テーブルのレコードセットを開く
    Set rs = CurrentDb.OpenRecordset("売上管理テーブル", dbOpenDynaset)
    
    With rs
        For i = LBound(sheetValues, 1) To UBound(sheetValues, 1)
            '//対応するフィールドに、データを入れていく
            .AddNew
            
            '//エラー発生時には、代わりに現在のテーブルの最大の管理番号よりも1だけ大きい管理番号を挿入する
            .Fields("管理番号").Value = IIf(IsError(sheetValues(i, 1)), Nz(DMax("管理番号", "売上管理テーブル"), 1) + 1, sheetValues(i, 1))
            '//エラー発生時には、代わりに不正な店舗名であることを示す
            .Fields("店舗").Value = IIf(IsError(sheetValues(i, 2)), "※不正な店舗名です", sheetValues(i, 2))
            '//エラー発生時には、売上金はゼロとする
            .Fields("売上金").Value = IIf(IsError(sheetValues(i, 3)), 0, sheetValues(i, 3))
            '//エラー発生時には、代わりに売上金の90%の金額を設定
            .Fields("調整後売上金").Value = IIf(IsError(sheetValues(i, 4)), .Fields("売上金").Value * 0.9, sheetValues(i, 4))
            
            .Update
        Next i
        
    End With

Exits:

    rs.Close
    wb.Close
    exApp.Quit
    
    Exit Sub

Err:

    MsgBox Err.Description, vbExclamation, Err.Number
    GoTo Exits
    
End Sub


```

こちらを実行すると、  
調整後売上金として、  
数式エラーの行は売上金の90%の金額を設定するようになります。

![代わりの値が設定されたテーブル](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642835776/learnerBlog/import-calc-error-value/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-22_160455_o9b65k.png)

## サンプルファイル

上にも記載しましたが、  
説明で使用したファイルについて、  
[こちら](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2022/01/import-calc-error-value)からダウンロードできます。

## 終わりに

エラー発生時の分岐処理については、  
VBAマクロを扱う限り避けては通れない問題となるでしょう。

様々なケースを想定し、  
なるべくユーザに優しいマクロを作ることができるようになることが、  
マクロの性能向上にとって重要となります。





