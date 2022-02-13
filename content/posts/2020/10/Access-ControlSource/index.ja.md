---
title: "[Access VBA] コントロールソースにユーザ定義関数を用いると列幅の自動調整が想定通りに機能しない問題の解決"
author: dede-20191130
date: 2020-10-13T09:37:47+09:00
slug: Access-ControlSource
draft: false
toc: true
tags: ['Access', 'VBA']
categories: ['トラブルシューティング', 'プログラミング']
vba_taxo: specification
archives:
    - 2020
    - 2020-10
---

## この記事について

Accessでフォーム（データシート型）を使用する際、
テキストボックスのコントロールソースとして文字列を返すユーザ定義関数を指定すると、
フォームの列幅の自動調整の動作が想定通りにならない問題がある。

例えば、ユーザ定義関数で長い文字列を取得すると、文字列がフィールドに収まりきらず見切れてしまう。

その解決策を模索した。

## 要約

ユーザ定義関数の呼び出しタイミングが
FormのResizeプロシージャよりも後であるため、
対象のテキストボックスに関しては、
```vb
Private Sub Form_Resize()
    
    '...code

    ctl.ColumnWidth = -2
    
    '...code

End Sub
```
による自動調整が機能しない。

そのため、下記のいずれかの措置をとる。
- ユーザ定義関数の処理にControl.ColumnWidth = -2が追加された関数を作成する。
- Resizeする際に長さを明示的に指定。

## 本文

### 前提

このようなテーブルT_01がある。  
![T_01](./image01.png)

T_01をサブフォームに組み込んで、このようなフォームF_01（およびサブフォームSubF_01）を表示する。  
![F_01](./image02.png)

このとき、各列ID、_Name、sizeの列幅を自動で調整したい。  
また、あたらしくsize typeテキスト列を追加し、  
sizeの値ごとにS,M,Lのサイズ記号を示す文字列「This is ○ type for he/she's size.」を格納したい。  
こちらの列幅も自動で調整したい。


### 環境
Microsoft Access 2019


### 実装

サブフォームのResizeプロシージャで  
各々のテキストボックスに対して  
列幅を指定した。

コントロールのColumnWidthの値を-2と指定することで列幅が自動調整される。
```vb
'******************************************************************************************
'*関数名    ：リサイズ
'*機能      ：
'*引数(1)   ：
'******************************************************************************************
Private Sub Form_Resize()
    
    '定数
    Const FUNC_NAME As String = "Form_Resize"
    
    '変数
    Dim ctl As Access.Control
    
    On Error GoTo ErrorHandler
    '---以下に処理を記述---
    
    For Each ctl In Me.Controls
        If ctl.ControlType = acTextBox Then
            ctl.ColumnWidth = -2
        End If
    Next
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生しましたのでマクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Sub
```


また、size typeテキスト列については、  
sizeの値の範囲ごとにS,M,Lのサイズ記号を分けたい。

そのため、  
コントロールソースとして次の値を指定し、  
ユーザ定義関数getSizeTypeを作成した。

```
# コントロールソース式
=getSizeType([size])
```

```vb
'******************************************************************************************
'*関数名    ：サイズ・タイプ取得
'*機能      ：
'*引数(1)   ：サイズ数字
'*戻り値    ：サイズ・タイプ
'******************************************************************************************
Public Function getSizeType(ByVal sizeNum As Long) As String
    
    '定数
    Const FUNC_NAME As String = "getSizeType"
    
    '変数
    Dim rtn As String
    
    On Error GoTo ErrorHandler
    
    getSizeType = ""
    
    '～169      ：S
    '170～175   ：M
    '～176      ：L
    Select Case True
    Case sizeNum < 169
        rtn = "This is S type for he/she's size."
    Case 176 < sizeNum
        rtn = "This is L type for he/she's size."
    Case Else
        rtn = "This is M type for he/she's size."
    End Select

    getSizeType = rtn
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生しましたのでマクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function
```

### トラブル

#### フィールドの見切れ

テーブルに既存のフィールドであるID、_Name、sizeについては  
正常に列幅が調整されているが、  
size type列については見切れが発生している。  

文字列が長すぎるために、Form_Resize()で列幅が調整される想定であったが、  
うまく機能していない。

![見切れ](./image03.png)

#### 関数の呼び出し順序

各関数に次のようにデバッグ出力を設定して  
実行したところ、  
関数の呼び出し順序は  
①サブフォームForm_Resize → ②親フォームForm_Resize → ③getSizeType  
であった

```vb
Debug.Print FUNC_NAME
```

### 解決措置

#### ①ユーザ定義関数の修正

Form_Resize()では対応できないため、  
ユーザ定義関数の処理にControl.ColumnWidth = -2を追加する必要がある。  

そのため、getSizeType()をラッピングする新しい関数getSizeTypeForSubF01Tb()を作成し、
比較のために別のコントロールソースを持つ新しいテキストボックスsize type ver2を作成した。

```
# コントロールソース式
=getSizeTypeForSubF01Tb([size],"txtSizeTypeVer2")
```

```vb
'******************************************************************************************
'*関数名    ：サイズ・タイプ取得 SubF01用
'*機能      ：
'*引数(1)   ：サイズ数字
'*引数(2)   ：コントロール名
'*戻り値    ：サイズ・タイプ
'******************************************************************************************
Public Function getSizeTypeForSubF01Tb(ByVal sizeNum As Long, ByVal ctlName As String) As String
    
    '定数
    Const FUNC_NAME As String = "getSizeTypeForSubF01Tb"
    
    '変数
    Dim rtn As String
    
    On Error GoTo ErrorHandler
    
    getSizeTypeForSubF01Tb = ""
    
    rtn = Module_ManageFormControls.getSizeType(sizeNum)

    '列幅を再設定
    If SysCmd(acSysCmdGetObjectState, acForm, Form_F_01.Name) <> 0 Then Form_SubF_01.Controls(ctlName).ColumnWidth = -2

    getSizeTypeForSubF01Tb = rtn
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生しましたのでマクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function
```

結果、列幅の自動調整がsize type ver2についても機能するようになった。

![size type ver2の列幅の調整](./image04.png)


#### ②長さを明示的に指定

Form_Resize()について、  
ループの後に一部のコントロール幅の数値をハードコーディングした。

```vb
'******************************************************************************************
'*関数名    ：リサイズ
'*機能      ：
'*引数(1)   ：
'******************************************************************************************
Private Sub Form_Resize()
    
    '定数
    Const FUNC_NAME As String = "Form_Resize"
    
    '変数
    Dim ctl As Access.Control
    
    On Error GoTo ErrorHandler
    '---以下に処理を記述---
    
    'ループ
    For Each ctl In Me.Controls
        If ctl.ControlType = acTextBox Then
            ctl.ColumnWidth = -2
        End If
    Next

    'size type ver2列の幅を明示的に指定（下記例では最低8cmの長さとなること）
    if Me.txtSizeTypeVer2.ColumnWidth < 8 * 567 then Me.txtSizeTypeVer2.ColumnWidth = 8 * 567
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生しましたのでマクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Sub
```

こちらについても、  
列幅の自動調整がsize type ver2についても機能するようになった。

#### ①および②の比較

①の方法は関数の数は増えないが、  
長さをハードコードしているため、  
getSizeTypeで返す文字列が変更した場合、メンテナンスの必要がある。

②の方法はすべての列幅を自動で調整できるが、  
関数の数が増え、コントロールソースの処理が若干煩雑になる。

