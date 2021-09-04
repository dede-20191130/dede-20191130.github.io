---
title: "[Access VBA] フォームのコントロール操作系の関数をどのようにユニットテストするかの方法メモ"
author: dede-20191130
date: 2021-09-04T10:17:36+09:00
slug: access-change-control-test
draft: true
toc: true
featured: false
tags: ["VBA","Access"]
categories: ["プログラミング"]
archives:
    - 2021
    - 2021-09
---

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    動的なフォーム生成、および動的なフォーム上のコントロール生成によって<br>
    実際に使用するフォームなどの環境から切り離して<br>
    コントロール操作系の関数のユニットテストができる。
{{< /box-with-title >}}

MSAccessのVBAで安定性のあるアプリケーションを作成する場合、  
関数を機能単位で分割してユニットテストすると安定性が高まる。

[純粋な関数](https://qiita.com/oedkty/items/f5fb807390a87359da0f)の場合はテストは簡単だが、  
関数の外部のグローバル変数やコンポーネント  
（フォームのテキストボックスの値や背景色など）を変更する場合はユニットテストが難しくなり、  
また、使用するフォームやクエリに何かしらの影響を与えて、それらを汚す（想定外の挙動を付与する）かもしれない。

フォームのコントロール操作系の関数にしぼって考えると、  
VBAによってフォームやその上のコントロールを動的に生成し、  
後始末をちゃんとすることで、  
コード上でテスト環境の作成まですべて完結、かつ環境をなるべく汚さない方法でユニットテストできるだろう。

その方法を記したい。

[<span id="srcURL"><u>説明のために作成したサンプルを含むツール（Accessファイル）とソースコードはこちらでダウンロードできます。</u></span>](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2021/09/access-change-control-test)

## 作成環境

- windows10
- MSOffice2019

## 想定するケース

### About

次のように、  

1. ボタン内容でプルダウンのリストを切り替える関数
2. ボタン内容で入力欄の使用可能/不可能を切り替える関数

の機能をそれぞれもつフレームがあるMainフォームを考える。

![サンプル](./01.png)

それぞれのフレームでは、  
内部のラジオボタンの切り替えにより、  
コンボボックスのプルダウンのリストを食べ物から飲み物のリストに切り替えたり、  
使用可能な入力欄を18歳未満専用のものから18歳以上用のものに切り替えたりする（コードは後述）。  

フレームのイベントから呼び出される関数の処理は、  
操作するコントロールに強く結びついているため、  
Mainフォーム上でテストするとMainフォーム自体の何かを変更する可能性がある。

### コード

【サンプル①】

```vb

'******************************************************************************************
'*機能      ：コンボボックスの項目リストを変更
'*引数      ：
'******************************************************************************************
Public Sub changeCmbBoxItems(ByVal selectedNumber As Long, ByVal cmbBox As ComboBox)
    
    '定数
    Const FUNC_NAME As String = "changeCmbBoxItems"
    
    '変数
    
    On Error GoTo ErrorHandler
    
    '//項目のクリア
    cmbBox.RowSource = ""
    
    Select Case selectedNumber
    '//食べ物
    Case 1
        cmbBox.AddItem "ピザ"
        cmbBox.AddItem "そば"
        cmbBox.AddItem "焼き肉"
    '//飲み物
    Case 2
        cmbBox.AddItem "コーラ"
        cmbBox.AddItem "緑茶"
        cmbBox.AddItem "水"
    End Select

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "クラス名：" & SOURCE_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical
        
    GoTo ExitHandler
        
End Sub

```

【サンプル②】

```vb
'******************************************************************************************
'*機能      ：テキストボックスの使用可能状態を変更
'*引数      ：
'******************************************************************************************
Public Sub changeTextBoxesEnabled(ByVal selectedNumber As Long, ByRef textboxes() As textbox)
    
    '定数
    Const FUNC_NAME As String = "changeTextBoxesEnabled"
    
    '変数
    Dim canUnder18Enable As Boolean '//18歳未満のためのテキストボックスが有効かどうか
    Dim textbox As Variant
    
    On Error GoTo ErrorHandler
    
    '//18歳未満を選択時はTrue、それ以外の場合はFalse
    canUnder18Enable = (selectedNumber = 1)
    
    '//タグがunder18かover18かによって
    '//使用可能状態を切り替える
    For Each textbox In textboxes
        If InStr(textbox.Tag, "under18") <> 0 Then
            textbox.Enabled = canUnder18Enable
        Else
            textbox.Enabled = Not canUnder18Enable
        End If
    Next textbox

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "クラス名：" & SOURCE_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical
        
    GoTo ExitHandler
        
End Sub
```


## ユニットテストの方法

### About

[動的なフォーム、コントロール作成](https://www.feedsoft.net/access/tips/tips81.html)により、  
テスタブルな環境を即時作成し、すぐに削除するようにする。

その際に、環境構築の順序を気をつけないとエラーが頻出したため、  
[気をつけなければならなかったこと](#気をつけなければならなかったこと)としてそれを記した。

### テストコード

【サンプル①のテスト】

```vb
'******************************************************************************************
'*機能      ：テスト　コンボボックスの項目リストを変更関数
'******************************************************************************************
Public Sub テスト_changeCmbBoxItems()
    
    '定数
    Const FUNC_NAME As String = "テスト_changeCmbBoxItems"
    
    '変数
    Dim tForm As Form
    Dim fName As String
    Dim cmb As ComboBox
    
    On Error GoTo ErrorHandler

    '//フォームの動的作成
    Set tForm = CreateForm()
    fName = tForm.Name
    
    '//デザインビューで開く
    DoCmd.OpenForm fName, acDesign
    
    '//コンボボックスの動的作成
    Set cmb = CreateControl(fName, _
                            AcControlType.acComboBox)
    Dim mycmb As String
    mycmb = "mycmb"
    cmb.Name = mycmb
    cmb.RowSourceType = "Value List"
    
    '//デザインビューを閉じる
    DoCmd.Close acForm, fName, acSaveYes
    
    '//フォームビューで開く
    DoCmd.OpenForm fName, acNormal
    
    '//上記で作成したコンボボックスを再度参照
    Set cmb = Forms(fName).Controls(mycmb)
    
    '//■テスト01：食べ物のリスト設定
    '////関数呼び出し
    Call changeCmbBoxItems(1, cmb)
    '////アサーション
    Debug.Assert cmb.ListCount = 3
    Debug.Assert cmb.Column(0, 0) = "ピザ"
    Debug.Assert cmb.Column(0, 1) = "そば"
    Debug.Assert cmb.Column(0, 2) = "焼き肉"
    Debug.Print cmb.ListCount
        
    '//■テスト02：飲み物のリスト設定
    '////関数呼び出し
    Call changeCmbBoxItems(2, cmb)
    '////アサーション
    Debug.Assert cmb.ListCount = 3
    Debug.Assert cmb.Column(0, 0) = "コーラ"
    Debug.Assert cmb.Column(0, 1) = "緑茶"
    Debug.Assert cmb.Column(0, 2) = "水"
    Debug.Print cmb.ListCount
    
    '//フォームビューを閉じる
    DoCmd.Close , , acSaveNo
    
    '//動的生成したフォームを削除
    DoCmd.DeleteObject acForm, fName
    
ExitHandler:
    
    '//テスト完了
    Debug.Print Now & ":Finish " & FUNC_NAME
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "クラス名：" & SOURCE_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical
        
    GoTo ExitHandler
        
End Sub
```


【サンプル②のテスト】

```vb
'******************************************************************************************
'*機能      ：テスト　テキストボックスの使用可能状態の変更関数
'******************************************************************************************
Public Sub テスト_changeTextBoxesEnabled()
    
    '定数
    Const FUNC_NAME As String = "テスト_changeTextBoxesEnabled"
    
    '変数
    Dim tForm As Form
    Dim fName As String
    Dim textboxes(0 To 3) As textbox
    Dim i As Long
    
    On Error GoTo ErrorHandler

    '//フォームの動的作成
    Set tForm = CreateForm()
    fName = tForm.Name
    
    '//デザインビューで開く
    DoCmd.OpenForm fName, acDesign
    
    '//テキストボックス配列の動的作成
    For i = 0 To 3
        Set textboxes(i) = CreateControl(fName, _
                            AcControlType.acTextBox)
                            
        textboxes(i).Name = "mytext_" & i
        
        '//一部のみunder18、それ以外はover18のタグを付与
        If i < 2 Then
            textboxes(i).Tag = "under18"
        Else
            textboxes(i).Tag = "over18"
        End If
    Next i
    
    '//デザインビューを閉じる
    DoCmd.Close acForm, fName, acSaveYes
    
    '//フォームビューで開く
    DoCmd.OpenForm fName, acNormal
    
    '//上記で作成したテキストボックス配列を再度参照
    For i = 0 To 3
        Set textboxes(i) = Forms(fName).Controls("mytext_" & i)
    Next i
    
    '//■テスト01：18歳未満専用のテキストボックスの有効化
    '////関数呼び出し
    Call changeTextBoxesEnabled(1, textboxes)
    '////アサーション
    Debug.Assert textboxes(0).Tag = "under18"
    Debug.Assert textboxes(0).Enabled = True
    Debug.Assert textboxes(1).Tag = "under18"
    Debug.Assert textboxes(1).Enabled = True
    Debug.Assert textboxes(2).Tag <> "under18"
    Debug.Assert textboxes(2).Enabled = False
    Debug.Assert textboxes(3).Tag <> "under18"
    Debug.Assert textboxes(3).Enabled = False
    
    '//■テスト02：18歳以上専用のテキストボックスの有効化
    '////関数呼び出し
    Call changeTextBoxesEnabled(2, textboxes)
    '////アサーション
    Debug.Assert textboxes(0).Tag = "under18"
    Debug.Assert textboxes(0).Enabled = False
    Debug.Assert textboxes(1).Tag = "under18"
    Debug.Assert textboxes(1).Enabled = False
    Debug.Assert textboxes(2).Tag <> "under18"
    Debug.Assert textboxes(2).Enabled = True
    Debug.Assert textboxes(3).Tag <> "under18"
    Debug.Assert textboxes(3).Enabled = True
        
    '//フォームビューを閉じる
    DoCmd.Close , , acSaveNo
    
    '//動的生成したフォームを削除
    DoCmd.DeleteObject acForm, fName
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "クラス名：" & SOURCE_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical
        
    GoTo ExitHandler
        
End Sub


```


### 気をつけなければならなかったこと

#### ビューによって設定可能な部分が異なる

私自身まだあまりAccessのビューごとの性質の違いについて  
十分に把握していないため、  
エラーのトラブルシューティングに見舞われることになった。

```vb {hl_lines=[2,10,13]}
'//デザインビューで開く
DoCmd.OpenForm fName, acDesign

'//コンボボックスの動的作成
Set cmb = CreateControl(fName, _
                        AcControlType.acComboBox)
Dim mycmb As String
mycmb = "mycmb"
cmb.Name = mycmb
cmb.RowSourceType = "Value List"

'//デザインビューを閉じる
DoCmd.Close acForm, fName, acSaveYes
```

`Name`や`RowSourceType`の指定は  
デザインビューでないと機能しない  
（本当は何か回避策があるかもしれないが、私のコードだとそうなった）ため、  
このようにしてデザインビューにおいて開閉することでプロパティをテスタブルに設定した。

同様に、  
次のようにコンボボックスのリスト項目を変更して参照する場合も  
フォームビューで開いておかないとエラーとなるため、  
次のようにする。

```vb {hl_lines=[2,9,19,28]}
'//フォームビューで開く
DoCmd.OpenForm fName, acNormal

'//上記で作成したコンボボックスを再度参照
Set cmb = Forms(fName).Controls(mycmb)

'//■テスト01：食べ物のリスト設定
'////関数呼び出し
Call changeCmbBoxItems(1, cmb)
'////アサーション
Debug.Assert cmb.ListCount = 3
Debug.Assert cmb.Column(0, 0) = "ピザ"
Debug.Assert cmb.Column(0, 1) = "そば"
Debug.Assert cmb.Column(0, 2) = "焼き肉"
Debug.Print cmb.ListCount
    
'//■テスト02：飲み物のリスト設定
'////関数呼び出し
Call changeCmbBoxItems(2, cmb)
'////アサーション
Debug.Assert cmb.ListCount = 3
Debug.Assert cmb.Column(0, 0) = "コーラ"
Debug.Assert cmb.Column(0, 1) = "緑茶"
Debug.Assert cmb.Column(0, 2) = "水"
Debug.Print cmb.ListCount

'//フォームビューを閉じる
DoCmd.Close , , acSaveNo
```

#### ビュー変更の際に参照がリセットされるため、再度参照を設定し直す

次のように、  
フォームビューを開いた後に変数`cmb`のコンボボックスに対しての参照を  
復旧させないといけない。

16行目を怠ると、  
リセットによりcmbはNull参照をしているためエラーが発生する。

```vb {hl_lines=[2,5,13,16]}
'//デザインビューで開く
DoCmd.OpenForm fName, acDesign

'//コンボボックスの動的作成
Set cmb = CreateControl(fName, _
                        AcControlType.acComboBox)
......

'//デザインビューを閉じる
DoCmd.Close acForm, fName, acSaveYes

'//フォームビューで開く
DoCmd.OpenForm fName, acNormal

'//上記で作成したコンボボックスを再度参照
Set cmb = Forms(fName).Controls(mycmb)
```

### 実行

上記テストコードを実行すると、  
自動的に新規フォームが作成され、  
フォーム上にテストに必要なコンポーネント（コンボボックス、テキストボックス数個）が整えられる。

作成した関数が実行され、適切な結果かどうかを`Debug.Assert`メソッドで評価。  
もし想定通りならば実行は一時停止せずにそのまま処理される。

最後にフォームが削除され、  
環境を汚さずにテストが終了する。

## まとめ

コントロールを動的に設定する際に  
いくつか気をつけないとエラーを吐くのは  
対処法を知っておかないと思わぬボトムネックになりかねない。

それさえクリアすれば、  
コントロール操作系の関数を安全にユニットテストする方法として適していると思う。
