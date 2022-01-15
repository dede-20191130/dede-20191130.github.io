---
title: "[教えて！VBA] 第11回 マクロからファイルを新しく名付けて保存する際の注意点とは？？"
author: dede-20191130
date: 2022-01-12T09:40:37+09:00
slug: vba-question-011-save-named-file
draft: false
toc: true
featured: false
tags: ["VBA","Excel"]
categories: ["プログラミング"]
vba_taxo: vbaq
archives:
    - 2022
    - 2022-01
---



![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641975714/learnerBlog/vba-question-011-save-named-file/save-named-file_z2hxc1.png)

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    フォルダに作成するファイルのパスを組み立てる際に、<br/>
    エラーが起きないようにするチェックの仕組みを記載しました。<br/><br/>
    また、既存のファイルが有った場合には、<br/>
    上書きする、別ファイルとして出力するなどの回避策があります。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
VBAマクロに関する質問のうち、  
皆が疑問に思っているトピックについて解説いたします。

今回は、  
{{< colored-span color="#fb9700" >}}
マクロからファイルを新しく名付けて保存する際に、エラーや想定外の結果にならないための注意点
{{< /colored-span >}}  
を解説いたします。

※この記事は、Office VBAマクロのうち  
Excel VBAマクロに関するトピックです。

レベル：<b>初級者向け</b>

## 環境

以下は、  
Office 2016のExcel環境での説明です。  

※2022/1時点の最新バージョンのExcelでも内容は変わりません。

また、WindowsOSのファイルシステムについての説明です（Macなどには当てはまらない箇所もあるかと思います）。  
もっとも、Officeアプリを使用するのは主にWindowsユーザのため問題ないかとは思いますが。
## ファイル保存について

VBAマクロでは、  
取り扱うデータを収集・入力・加工したのちに、  
別ファイルとしてデータを吐き出すような処理を書くことができます。

![マクロによるファイルの出力](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641975714/learnerBlog/vba-question-011-save-named-file/save-named-file2_he4uef.png)

例えば、新しくブックを作成し、  
マクロ登録済みのブックのセル内容を転機するマクロは次のようになります。

```vb

Sub ブックとして出力するサンプル()
    Dim newWorkBook As Workbook
    
    '//ブックの新規追加
    Set newWorkBook = Workbooks.Add
    
    '//マクロ登録しているブックのB2セルの内容を転記
    newWorkBook.Worksheets(1).Range("A1").Value = ThisWorkbook.Worksheets(1).Range("B2").Value
    
    '//追加したブックをtempフォルダに保存して閉じる
    newWorkBook.SaveAs "C:\temp\ブックとして出力するサンプル.xlsx"
    newWorkBook.Close SaveChanges:=False

    
End Sub

```

また、テキストファイルを作成し、  
そちらに内容を書き込むマクロは次のようになります。

```vb

Sub テキストファイルとして出力するサンプル()
    Open "C:\temp\テキストファイルとして出力するサンプル.txt" For Append As #1
    '//マクロ登録しているブックのB2セルの内容をテキストファイルに記入
    Print #1, ThisWorkbook.Worksheets(1).Range("B2").Value
    Close #1
End Sub

```

これらに付帯する注意点と、  
その対策についてを、  
以下のセクションで見ていきます。

## ファイル保存の注意点とチェック機構

### ABOUT

ファイルを保存するためには、  
必ずファイルパス（ファイルのアドレス。例：`C:\temp\サンプル.txt`）を指定します。

その際に、  
いくつかのチェックを設けることによって、  
予期せぬエラーや結果を回避することが、  
マクロの性能向上にとって重要となります。

### 注意点1. フォルダやファイル名に空欄が無いようにする

データを挿入してファイルパスを動的に生成する場合、  
挿入するデータの有効性をチェックする必要があります。

例えば、  
セルの入力内容によってファイルパスを生成するような処理の場合、  
セルの入力内容が空欄にならないようにチェックが必要です。

次のコードは  
場合によってはエラーが発生するサンプルです。

エラー発生時、  
新規作成したブックは開かれたままになってしまいます。

```vb

Sub セルの入力内容によってファイルパスを生成_NGサンプル()
    Dim newWorkBook As Workbook
    Dim filePath As String
    
    Set newWorkBook = Workbooks.Add
    newWorkBook.Worksheets(1).Range("A1").Value = ThisWorkbook.ActiveSheet.Range("B2").Value
    
    '//ファイルパスを動的に生成
    '////※　B3セルが空欄だったら1004エラー発生
    filePath = "C:\temp\" & ThisWorkbook.ActiveSheet.Range("B3").Value & ".xlsx"
    newWorkBook.SaveAs filePath
    newWorkBook.Close SaveChanges:=False
End Sub

```

これを改善するために、  
次のチェック機構2点を導入します。  
- B3セルが空欄でないことを確かめる
- 空欄であった場合に処理を終了し、ブックを保存せず閉じる

```vb

Sub セルの入力内容によってファイルパスを生成_チェック機構追加サンプル()
    Dim newWorkBook As Workbook
    Dim filePath As String
    
    On Error GoTo ErrorHandler
    
    Set newWorkBook = Workbooks.Add
    newWorkBook.Worksheets(1).Range("A1").Value = ThisWorkbook.ActiveSheet.Range("B2").Value
    
    '//ファイルパスを動的に生成
    '////B3セルが空欄でないことを確かめる
    If ThisWorkbook.ActiveSheet.Range("B3").Value = "" Then
        MsgBox "B3セルにファイル名を入力して下さい", vbExclamation
        GoTo ExitHandler
    End If
    filePath = "C:\temp\" & ThisWorkbook.ActiveSheet.Range("B3").Value & ".xlsx"
    
    newWorkBook.SaveAs filePath

ExitHandler:
    '//ブックを閉じる
    newWorkBook.Close SaveChanges:=False
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラー発生" & vbLf & Err.Description, vbCritical, Err.number
    
    GoTo ExitHandler
    
End Sub

```

空欄のエラーが起きた場合でも、  
後始末の処理を導入しやすい形になりました。

### 注意点2. ファイル名に禁止文字が存在してはならない

Windows OS環境のファイル名には次の文字が使えません。  
（いずれも半角です）

<table>
<thead>
    <tr>
        <th>文字</th>
        <th>意味</th>
    </tr>
</thead>
<tbody>
    <tr>
        <td>"</td>
        <td>ダブルクォーテーション</td>
    </tr>
    <tr>
        <td>&lt;</td>
        <td>小なり</td>
    </tr>
    <tr>
        <td>&gt;</td>
        <td>大なり</td>
    </tr>
    <tr>
        <td>｜</td>
        <td>バーティカルバー</td>
    </tr>
    <tr>
        <td>:</td>
        <td>コロン</td>
    </tr>
    <tr>
        <td>*</td>
        <td>アスタリスク</td>
    </tr>
    <tr>
        <td>?</td>
        <td>クエスチョンマーク</td>
    </tr>
    <tr>
        <td>¥</td>
        <td>円記号</td>
    </tr>
    <tr>
        <td>/</td>
        <td>スラッシュ</td>
    </tr>
</tbody>
</table>

もしいずれかの文字が含まれたパスでファイルを保存しようとする場合、  
保存メソッドがエラーとなります。

![保存メソッドのエラー](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641975713/learnerBlog/vba-question-011-save-named-file/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-12_134404_vpizxp.png)

ファイル名にこれらが検知された場合、  
主な対処法としては次の2点があるでしょう。  
- 警告メッセージを表示し、保存をキャンセル
- 禁止文字をすべてなにかの文字（例えばアンダースコア「_」）に置き換えて出力

{{< colored-span color="#fb9700" >}}前者{{< /colored-span >}}
についてのサンプルはつぎのようになります。

```vb

Private Const FILE_FORBIDDEN_CHARACTORS_STR = "\_/_:_*_?_""_<_>_|"

Sub ファイル名禁止文字を検知_アラートを出す場合()
    Dim newWorkBook As Workbook
    Dim filePath As String
    Dim myFilename  As String
    Dim forbiddenChar As Variant
    Dim cancel As Boolean
    
    On Error GoTo ErrorHandler
    
    Set newWorkBook = Workbooks.Add
    newWorkBook.Worksheets(1).Range("A1").Value = ThisWorkbook.ActiveSheet.Range("B2").Value
    
    '//ファイル名をセルから取得
    myFilename = ThisWorkbook.ActiveSheet.Range("B3").Value
    filePath = "C:\temp\" & myFilename & ".xlsx"
    
    
    '//ファイル名禁止文字が含まれていないかどうかをチェック
    cancel = False
    For Each forbiddenChar In Split(FILE_FORBIDDEN_CHARACTORS_STR, "_")
        If InStr(myFilename, forbiddenChar) > 0 Then
            '//含まれている場合
            cancel = True
            Exit For
        End If
    Next forbiddenChar
    If cancel Then
        '//キャンセルする場合
        MsgBox "ファイル名として使えない文字が含まれています", vbExclamation
        GoTo ExitHandler
    End If
    
    newWorkBook.SaveAs filePath

ExitHandler:
    '//ブックを閉じる
    newWorkBook.Close SaveChanges:=False
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラー発生" & vbLf & Err.Description, vbCritical, Err.number
    
    GoTo ExitHandler
    
End Sub

```

もし禁止文字9文字のいずれかがセルに書かれたファイル名に存在している場合、  
ブックは保存されません。

また、
{{< colored-span color="#fb9700" >}}後者{{< /colored-span >}}
（禁止文字を置換する）についてのサンプルは次のようになります。

```vb

Private Const FILE_FORBIDDEN_CHARACTORS_STR = "\_/_:_*_?_""_<_>_|"

Sub ファイル名禁止文字を検知_文字を置き換えて出力する場合()
    Dim newWorkBook As Workbook
    Dim filePath As String
    Dim myFilename  As String
    Dim forbiddenChar As Variant
    
    On Error GoTo ErrorHandler
    
    Set newWorkBook = Workbooks.Add
    newWorkBook.Worksheets(1).Range("A1").Value = ThisWorkbook.ActiveSheet.Range("B2").Value
    
    '//ファイル名をセルから取得
    myFilename = ThisWorkbook.ActiveSheet.Range("B3").Value
    
    '//ファイル名禁止文字が含まれていないかどうかをチェック
    For Each forbiddenChar In Split(FILE_FORBIDDEN_CHARACTORS_STR, "_")
        If InStr(myFilename, forbiddenChar) > 0 Then
            '//含まれている場合、すべてアンダースコアに置換する
            myFilename = Replace(myFilename, forbiddenChar, "_")
        End If
    Next forbiddenChar
    
    '//ファイルパスを作成
    filePath = "C:\temp\" & myFilename & ".xlsx"
    
    newWorkBook.SaveAs filePath

ExitHandler:
    '//ブックを閉じる
    newWorkBook.Close SaveChanges:=False
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラー発生" & vbLf & Err.Description, vbCritical, Err.number
    
    GoTo ExitHandler
    
End Sub

```

禁止文字のひとつひとつをファイル名に存在するか調べ、  
該当すれば、すべて置換します。

この方法ならば、  
どのようなファイル名（パス名全体で255文字を超えない限り）でも出力することができます。

### +α 厳密にチェックするには

上記の注意点1, 2を総括すると、  
ファイル名の有効性をチェックするというタスクになります。

ところで、  
文字列のチェックには{{< colored-span color="#fb9700" >}}正規表現{{< /colored-span >}}を使用すればより厳密なチェックが可能です。

多くの場合は注意点1, 2で十分に対応可能かと思いますが、  
より厳密にチェックが必要であれば、正規表現を使用しましょう。

正規表現の使い方については
[こちら](https://excel-ubara.com/excelvba4/EXCEL232.html)
に素晴らしい記事があります。

正規表現のパターン指定は使用環境や目的によってまちまちと思いますが、  
`VBA`のパターンは`Javascript`(ES2015以降)などの他のモダン言語に比べてやや貧弱であることは注意しなければならないでしょう。

### 注意点3. 指定パスにファイルが存在する場合

指定パスに作成したいファイルと同名ファイルが既に存在していた場合、  
状況はやや違ってきます。

![同名ファイルが既に存在](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641975714/learnerBlog/vba-question-011-save-named-file/save-named-file3_uhlvv3.png)

ファイルシステムの都合上、  
ファイルはユニークなパス（唯一無二のパス）を持つ必要があるため、  
フォルダに同名ファイルは設置できません。

回避策として、  
1. 既存ファイルを上書きする
2. 警告を表示して保存をキャンセルする
3. ファイル名にランダム文字列を付け、別ファイルとして出力する
が挙げられるでしょう。

それらについて見ていきます。

#### 既存ファイルを上書き

この方法のユースケースとしては、  
指定フォルダに同名ファイルがあっても気にしない（すでに古くなったファイルとみなす）場合や、  
出力するファイルの内容が変化しないことを想定している場合に用いられるでしょう。

上書きする際には、  
通常はアラートメッセージが表示されます。

![アラートメッセージ](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641975714/learnerBlog/vba-question-011-save-named-file/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-12_155251_nfnbfq.png)

それを回避するために、  
`DisplayAlerts`プロパティをいったん`False`に設定します。

コードは次のようになります。

```vb

Sub 指定パスにファイルが存在する_既存ファイルを上書きする方法()
    Dim newWorkBook As Workbook
    Dim filePath As String
    
    On Error GoTo ErrorHandler
    
    '//いったん警告メッセージを非表示化
    Application.DisplayAlerts = False
    
    Set newWorkBook = Workbooks.Add
    newWorkBook.Worksheets(1).Range("A1").Value = ThisWorkbook.ActiveSheet.Range("B2").Value
    
    '//ファイルパスを動的に生成
    filePath = "C:\temp\" & ThisWorkbook.ActiveSheet.Range("B3").Value & ".xlsx"
    
    newWorkBook.SaveAs filePath

ExitHandler:
    '//ブックを閉じる
    newWorkBook.Close SaveChanges:=False
    
    '//警告メッセージの設定を戻す
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラー発生" & vbLf & Err.Description, vbCritical, Err.number
    
    GoTo ExitHandler
    
End Sub

```

注意しなければならないのは、  
`Application.DisplayAlerts`の値を必ず`True`に戻すことです。

`DisplayAlerts`を無効化している間は、  
すべての警告メッセージが出ないようになるので、  
普段遣いのExcelでの作業がやりにくくなる可能性があるためです。

#### 警告を表示しキャンセル

この方法のユースケースとしては、  
既存ファイルを削除したくない場合や、  
そもそも同名のファイルを出力するようなオペレーションが、業務フローに対して本質的に間違っているので  
ユーザにやりなおしをさせたい場合などが該当するでしょう。

ファイル存在有無の検知には、  
`Dir`関数を利用します。  
（FSOの`FileExists`メソッドを利用しても可能です）

```vb

Sub 指定パスにファイルが存在する_警告を表示しキャンセルする方法()
    Dim newWorkBook As Workbook
    Dim filePath As String
    
    On Error GoTo ErrorHandler
    
    Set newWorkBook = Workbooks.Add
    newWorkBook.Worksheets(1).Range("A1").Value = ThisWorkbook.ActiveSheet.Range("B2").Value
    
    '//ファイルパスを動的に生成
    filePath = "C:\temp\" & ThisWorkbook.ActiveSheet.Range("B3").Value & ".xlsx"
    
    '//フォルダに同名ファイルが存在するかどうかをチェック
    If Dir(filePath) <> "" Then
        '//キャンセルする
        MsgBox "出力先のフォルダに、すでに同じ名前のファイルが存在します。", vbExclamation
        GoTo ExitHandler
    End If
    
    newWorkBook.SaveAs filePath

ExitHandler:
    '//ブックを閉じる
    newWorkBook.Close SaveChanges:=False
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラー発生" & vbLf & Err.Description, vbCritical, Err.number
    
    GoTo ExitHandler
    
End Sub

```
#### 別ファイルとして出力

この方法のユースケースとしては、  
既存・新規それぞれのファイルは維持したいが、  
ユーザにもう一度ファイル名の設定をやり直させたくない場合、  
および、マクロで定期的に自動でファイルを吐き出すような処理  
（VBAでそのようなプログラムを走らせるのはあまり現実的ではないかもしれませんが）を実行したい場合が挙げられるでしょう。


別ファイルとして出力するために、  
ファイル名の最後にランダム文字列を付与します。

ランダム文字列の実装には、  
SHA-256ハッシュ値を使用します。

[こちら](https://blog.nekonium.com/vba-hash/)
でご紹介されていたハッシュ関数を利用しました。

リンク先でもご紹介されているように、  
.NET Frameworkの`System.Security.Cryptography`ライブラリを利用することで、  
VBAの環境でもSHA-256を利用することが可能になります。

ハッシュの引数として、  
現在時刻（`Now`）を採用すれば、  
それぞれのファイルで決して被ることがない文字列が生成できます。

```vb

'// 引用:https://blog.nekonium.com/vba-hash/
Public Function SHA256_HEX(str As String) As String
    Dim sha256m As Object
    Dim utf8 As Object
    Dim bytes() As Byte
    Dim hash() As Byte
    Dim i As Integer
    Dim res As String

    Set utf8 = CreateObject("System.Text.UTF8Encoding")
    bytes = utf8.GetBytes_4(str)
    Debug.Print bytes

    Set sha256m = CreateObject("System.Security.Cryptography.SHA256Managed")
    hash = sha256m.ComputeHash_2((bytes))
    Debug.Print hash
    
    For i = LBound(hash) To UBound(hash)
        res = res & LCase(Right("0" & Hex(hash(i)), 2))
    Next i

    SHA256_HEX = LCase(res)
End Function



Sub 指定パスにファイルが存在する_別ファイルとして出力する方法()
    Dim newWorkBook As Workbook
    Dim filePath As String
    Dim myFileName As String
    
    On Error GoTo ErrorHandler
    
    Set newWorkBook = Workbooks.Add
    newWorkBook.Worksheets(1).Range("A1").Value = ThisWorkbook.ActiveSheet.Range("B2").Value
    
    '//ファイルパスを動的に生成
    myFileName = ThisWorkbook.ActiveSheet.Range("B3").Value
    filePath = "C:\temp\" & myFileName & ".xlsx"
    
    '//フォルダに同名ファイルが存在するかどうかをチェック
    If Dir(filePath) <> "" Then
        '//ファイル名にハッシュ値を付与
        myFileName = myFileName & "_" & SHA256_HEX(Now())
        '//ファイルパスを再設定
        filePath = "C:\temp\" & myFileName & ".xlsx"
    End If
    
    newWorkBook.SaveAs filePath

ExitHandler:
    '//ブックを閉じる
    newWorkBook.Close SaveChanges:=False
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラー発生" & vbLf & Err.Description, vbCritical, Err.number
    
    GoTo ExitHandler
    
End Sub

```

【以下、デモです】

フォルダに「テストファイル」が存在する状態でマクロを実行すると……

![マクロ実行前](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641975714/learnerBlog/vba-question-011-save-named-file/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-12_164325_ndtbsb.png)

ハッシュ値が付与されたファイルが出力されます。

![マクロ実行後](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641975714/learnerBlog/vba-question-011-save-named-file/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-12_164341_hvhvbr.png)

## 終わりに

以上、ファイルを新規生成してフォルダに出力する際の注意点について説明しました。  

ケースバイケースでそれぞれの回避策を組み合わせれば、  
既存のマクロをさらに使い勝手の良いマクロに成長させることができるでしょう。