---
title: "[Access VBA] デザインビューのテーブル定義を表形式でエクスポートする関数を作成した"
author: dede-20191130
date: 2020-10-25T22:01:43+09:00
slug: Access-tableDef
draft: False
toc: true
tags: ['Access', 'VBA','自作']
categories: ['課題解決', 'プログラミング']
vba_taxo: help_develop
archives:
    - 2020
    - 2020-10
---

## この記事について

Accessの各々のテーブルの内容をエクスポートする方法として、
- ナビゲーションバーのテーブルの右クリックメニューからエクスポート -> Excel、テキスト等形式を選択して出力
- データベースツールタブのデータベース構造の解析 -> テーブル指定 -> 印刷、またはExcel、PDF等形式を選択して出力

の選択肢がある。  
  
しかし、  
前者は、フィールドの型や主キーの有無など、テーブルの詳細な情報はエクスポートできず、  
後者は、詳細な情報をエクスポートできるが、各フィールドごとに単票形式でエクスポートデータが得られるため、デザインビューのように一覧としてテーブルの設定を把握することが難しい。  

![デザインビューによる一覧化されたテーブル情報](./image01.png)  
  
よって、既存の機能で対応できないため、  
デザインビューのテーブル定義を表形式でエクスポートする関数を作成した。  

{{< colored-span color="#fb9700" >}}※本記事の関数の内容をもとにツール化したものも作成しました。🔽🔽 {{< /colored-span >}}  
{{< page-titled-link page="access-tabledef-tool" >}}



## 作成環境
Microsoft Office 2019

## 機能
カレントプロジェクト（コードが貼り付けてあるAccessファイル）の  
すべてのテーブルのテーブル定義を表形式でエクスポートする。  
エクスポート先は新規エクセルブック。テーブルデータがシートごとに格納される。  



## 関数概要
  
|名前|種類|機能|
| ---- | ---- | ---- |
|exportTableDefTablesMain|Subプロシージャ|テーブル定義情報表形式エクスポートのMain関数|
|getTableDefArray|Functionプロシージャ|テーブルの定義情報を取得|
|getFieldTypeString|Functionプロシージャ|引数フィールドのデータ型文字列を取得|
|getPKs|Functionプロシージャ|テーブルの主キーであるフィールド名を辞書として取得|
|getFKs|Functionプロシージャ|テーブルの外部キーであるフィールド名を辞書として取得|
|setWSName|Functionプロシージャ|引数で指定されたエクセルシートの名前をセット|

## 呼び出し関係
  
- exportTableDefTablesMain
    - 呼出 -> getTableDefArray
        - 呼出 -> getFieldTypeString
        - 呼出 -> getPKs
        - 呼出 -> getFKs
    - 呼出 -> setWSName
  
## コード

### [exportTableDefTablesMain]
```vb
'******************************************************************************************
'*関数名    ：exportTableDefTablesMain
'*機能      ：テーブル定義情報テーブルを作成
'*引数(1)   ：
'******************************************************************************************
Public Sub exportTableDefTablesMain()
    
    '定数
    Const FUNC_NAME As String = "exportTableDefTablesMain"
    
    '変数
    Dim xlApp As Object
    Dim wb As Object
    Dim tdf As DAO.TableDef
    Dim defArr As Variant
    Dim fstWs As Object
    Dim ws As Object
    
    On Error GoTo ErrorHandler
    
    'エクセルブック開始
    Set xlApp = CreateObject("Excel.Application")
    With xlApp
        .Visible = False
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    Set wb = xlApp.Workbooks.Add
    
    '初期シート
    Set fstWs = wb.Worksheets(1)
    
    'テーブルごとに別シートにテーブル定義情報テーブルを作成
    For Each tdf In CurrentDb.TableDefs
        Do
            'システムテーブル等出力の必要のないテーブルの場合はcontinue
            If Left(tdf.Name, 4) = "Msys" Or Left(tdf.Name, 4) = "Usys" Or Left(tdf.Name, 1) = "~" Then Exit Do
            
            'テーブルの定義情報配列を取得
            defArr = getTableDefArray(tdf)
            If IsNull(defArr) Then GoTo ExitHandler
            
            'ブックで新規シートを作成
            Set ws = wb.Worksheets.Add
            If Not setWSName(ws, tdf.Name) Then Call Err.Raise(1000, "シート名指定エラー", "シート名指定の際にエラーが発生しました。")
            
            '定義情報配列を記入し、列幅調整
            With ws.Range(ws.cells(1, 1), ws.cells(UBound(defArr) - LBound(defArr) + 1, UBound(defArr, 2) - LBound(defArr, 2) + 1))
                .Value = defArr
                .EntireColumn.AutoFit
            End With
            
        Loop While False
    Next tdf
    
    '初期シートの削除
    If wb.Worksheets.Count > 1 Then fstWs.Delete
    
    'ブック保存
    wb.saveas Application.CurrentProject.Path & _
              "\" & _
              Left( _
              CurrentProject.Name, _
              InStrRev(CurrentProject.Name, ".") - 1 _
              ) & _
                "_テーブル定義一覧.xlsx"
    
    '完了
    MsgBox "エクスポート完了", , "通知"
    
    
ExitHandler:
    
    'クローズ
    If Not wb Is Nothing Then wb.Close: Set wb = Nothing
    If Not xlApp Is Nothing Then xlApp.Quit: Set xlApp = Nothing
    
    Set tdf = Nothing
    Set ws = Nothing
    Set fstWs = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.description, vbCritical, "エラー"
        
    GoTo ExitHandler
        
End Sub
```
新規インスタンスでエクセルを立ち上げる。  
初期シートはデータの格納後削除される。
  
    
  

### [getTableDefArray]
```vb
'******************************************************************************************
'*関数名    ：getTableDefArray
'*機能      ：テーブルの定義情報を取得
'*            項目：フィールド名
'*                  データ型
'*                  サイズ
'*                  必須項目かどうか
'*                  主キー（PK）
'*                  外部キー（FK）
'*                  説明
'*
'*引数(1)   ：テーブル定義
'*戻り値    ：定義情報配列
'******************************************************************************************
Public Function getTableDefArray( _
       ByVal pTdf As DAO.TableDef _
       ) As Variant
    
    '定数
    Const FUNC_NAME As String = "getTableDefArray"
    
    '変数
    Dim defArr() As Variant
    Dim fld As DAO.Field
    Dim i As Long
    Dim dicPKs As Object
    Dim dicFKs As Object
    Dim description As String
    
    On Error GoTo ErrorHandler

    getTableDefArray = Null
    
    '(テーブルのフィールド数 + 1)×7のサイズの配列
    ReDim defArr(0 To pTdf.Fields.Count, 0 To 6)
    
    'ヘッダ設定
    defArr(0, 0) = "フィールド名"
    defArr(0, 1) = "データ型"
    defArr(0, 2) = "サイズ"
    defArr(0, 3) = "必須"
    defArr(0, 4) = "PK"
    defArr(0, 5) = "FK"
    defArr(0, 6) = "説明"
    
    'テーブルのすべての主キーであるフィールド名を辞書として取得
    Set dicPKs = getPKs(pTdf)
    If dicPKs Is Nothing Then GoTo ExitHandler
    
    'テーブルのすべての外部キーであるフィールド名を辞書として取得
    Set dicFKs = getFKs(pTdf)
    If dicFKs Is Nothing Then GoTo ExitHandler
    
    'フィールドごとに探索
    For i = 1 To pTdf.Fields.Count
        Set fld = pTdf.Fields(i - 1)
        'フィールド名
        defArr(i, 0) = fld.Name
        'データ型
        defArr(i, 1) = getFieldTypeString(fld.Type)
        'サイズ
        If fld.Type = dbText Then
            defArr(i, 2) = fld.Size
        Else
            defArr(i, 2) = "-"
        End If
        '必須項目かどうか
        If fld.Required Then defArr(i, 3) = "○"
        '主キー（PK）かどうか ◆note1
        If dicPKs.Exists(fld.Name) Then defArr(i, 4) = "○"
        '外部キー（FK）かどうか ◆note1
        If dicFKs.Exists(fld.Name) Then defArr(i, 5) = "○"
        '説明
        On Error Resume Next
        description = fld.Properties("Description")
        On Error GoTo ErrorHandler
        defArr(i, 6) = description
    Next i


    getTableDefArray = defArr
    
ExitHandler:
    
    Set fld = Nothing
    Set dicFKs = Nothing
    Set dicPKs = Nothing
    
    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.description, vbCritical, "エラー"
        
    GoTo ExitHandler
        
End Function
```
(引数テーブルのフィールド数+1)×7の大きさの配列を戻り値とする。  
◆note1 - テーブルの主キー及び外部キーの辞書配列を先に取得しておいて、  
各フィールドがそれに含まれているかどうかをチェックする。

### [getFieldTypeString]
```vb
'******************************************************************************************
'*関数名    ：getFieldTypeString
'*機能      ：フィールドのデータ型文字列を取得
'*引数(1)   ：フィールドタイプ
'*戻り値    ：フィールドのデータ型文字列
'******************************************************************************************
Public Function getFieldTypeString(ByVal pFldTyepNum As Long) As String
    
    '定数
    Const FUNC_NAME As String = "getFieldTypeString"
    
    '変数
    Dim strType As String
    
    On Error GoTo ErrorHandler

    strType = ""
    

    Select Case pFldTyepNum
    Case dbBoolean
        strType = "ブール型"
    Case dbByte
        strType = "バイト型"
    Case dbInteger
        strType = "整数型"
    Case dbLong
        strType = "長整数型"
    Case dbSingle
        strType = "単精度浮動小数点型"
    Case dbDouble
        strType = "倍精度浮動小数点型"
    Case dbCurrency
        strType = "通貨型"
    Case dbDate
        strType = "日付/時刻型"
    Case dbText
        strType = "テキスト型"
    Case dbLongBinary
        strType = "OLEオブジェクト型"
    Case dbMemo
        strType = "メモ型"
    End Select

    getFieldTypeString = strType
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.description, vbCritical, "エラー"
        
    GoTo ExitHandler
        
End Function
```
引数で与えられるDAO.FieldオブジェクトのTypeプロパティは数値であるため、  
それを文字列に変換する。

### [getPKs]
```vb
'******************************************************************************************
'*関数名    ：getPKs
'*機能      ：テーブルの主キーであるフィールド名を辞書として取得
'*引数(1)   ：フィールドタイプ
'*戻り値    ：辞書
'******************************************************************************************
Public Function getPKs(ByVal pTdf As DAO.TableDef) As Object
    
    '定数
    Const FUNC_NAME As String = "getPKs"
    
    '変数
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim dic As Object
    
    On Error GoTo ErrorHandler

    Set getPKs = Nothing
    Set dic = CreateObject("Scripting.Dictionary")
    
    
    'インデックスより探索
    For Each idx In pTdf.Indexes
        If idx.Primary = True Then
            For Each fld In idx.Fields
                dic.Add fld.Name, True
            Next
        End If
    Next

    'Return
    Set getPKs = dic
    
ExitHandler:

    Set dic = Nothing

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.description, vbCritical, "エラー"
        
    GoTo ExitHandler
        
End Function
```

### [getFKs]
```vb
'******************************************************************************************
'*関数名    ：getFKs
'*機能      ：テーブルの外部キーであるフィールド名を辞書として取得
'*引数(1)   ：
'*戻り値    ：辞書
'******************************************************************************************
Public Function getFKs(ByVal pTdf As DAO.TableDef) As Object
    
    '定数
    Const FUNC_NAME As String = "getFKs"
    
    '変数
    Dim rsRelation As DAO.Recordset
    Dim dic As Object
    
    On Error GoTo ErrorHandler

    Set getFKs = Nothing
    Set dic = CreateObject("Scripting.Dictionary")
    
    'リレーションテーブルにアクセス
    Set rsRelation = CurrentDb.OpenRecordset( _
                     "SELECT szColumn FROM MSysRelationships WHERE szObject =" & _
                     " " & _
                     "'" & _
                     pTdf.Name & _
                     "'" & _
                     ";" _
                     )
    
    With rsRelation
        If .EOF Then Set getFKs = dic: GoTo ExitHandler
        .MoveFirst
        Do Until .EOF
            dic.Add .Fields("szColumn").Value, True
            .MoveNext
        Loop
    End With
    
    'Return
    Set getFKs = dic
    
ExitHandler:
    
    If Not rsRelation Is Nothing Then rsRelation.Close: Set rsRelation = Nothing
        
    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & Err.Number & vbNewLine & _
           Err.description, vbCritical, "エラー"

        
    GoTo ExitHandler
        
End Function
```

テーブルの外部キー情報は  
システムテーブルのMSysRelationshipsに格納されているため、
Recordsetを用いて情報を取得する。

### [setWSName]
```vb
'******************************************************************************************
'*関数名    ：setWSName
'*機能      ：エクセルシートの名前をセット
'*引数(1)   ：エクセルシート
'*引数(2)   ：代入する名前
'*戻り値    ：True > 正常終了、False > 異常終了
'******************************************************************************************
Public Function setWSName( _
       ByVal ws As Object, _
       ByVal newName As String _
       ) As Boolean
    
    '定数
    Const FUNC_NAME As String = "setWSName"
    
    '変数
    
    On Error GoTo ErrorHandler

    setWSName = False
    
    ws.Name = newName

    setWSName = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    'シート名に使用できない文字であった場合
    ws.Name = "テーブル_" & ws.Parent.Worksheets.Count & "_" & Format(Now, "yyyymmddhhnnss")

    setWSName = True
    GoTo ExitHandler
        
End Function
```

シート名とするテーブル名に  
シート名として禁止されている文字（*、?等）が含まれている場合、  
シート名重複防止にシート名は「テーブル_全シート数_yyyymmddhhnnss」としている。

## 使用方法
抽出したいAccessファイルのモジュールのどこかに  
上記関数群を記入し、  
exportTableDefTablesMainを実行する。

## イメージ
次のようにエクセルファイルが出力される。  
デザインビュー表示に似せた様式で各情報が一覧できる。

![エクスポートNo1](./image02.png)  
![エクスポートNo2](./image03.png)  