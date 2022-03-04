---
title: "VBA: THE BENEFIT OF USING CLASS AND THE WAY TO USE, WITH A SIMPLE SAMPLE PART 2"
author: dede-20191130
date: 2020-11-22T14:07:24+09:00
slug: Excel-Class-Sample-02
draft: false
toc: true
featured: false
tags: ['Excel', 'VBA','HOMEMADE', 'object-oriented']
categories: ['programming']
vba_taxo: class_how_to
archives:
    - 2020
    - 2020-11
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

Hi, I'm Dede.

This article is a a sequel of a article I posted before.

{{< page-titled-link page="Excel-Class-Sample-01" >}}

Here, I'm going to take anothier sample file and describe how to use class and the bbenefits in a different way from that in the previous article.  
The sample class `clsCreateNewExcel` will express the benefits from encapsulation and routine processing in initialization and termination processes, which are excluded from `TableCreater` in the previous article.


You can download Access file created for explanation and view its source code from [<span id="srcURL"><u>here</u></span>](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2020/11/TableCreater/en)!


## CREATION ENVIRONMENT
Windows10  
MSOffice 2019

## OVERVIEW OF SAMPLE

I created a Access file containing Item Data table.

A user can execute different processes with buttons in Main Form: 

![Main Form](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645882070/learnerBlog/Excel-Class-Sample-02/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-02-26_222208_jsn1ie.png)

All of these buttons have a common processing:  
- Create a new Excel book on user's Desktop.

The individual processing for each button is in the following list:  
- Extract Item data where its unit price more than 10000 and post them to new book's sheet.
- Extract Item data where its code starts with 'B' and post them to new book's sheet.
- Get Item data from Web API and post them to new book's sheet.

I created a class module `clsCreateNewExcel` for the sake of book creation and management of book-relational objects.


## clsCreateNewExcel CLASS

### ROLE

The class instantiates a new Excel Application Interface, creates new Excel book, and stores objects which can manipulate them in itself.  
And, it has a function which adds worksheet to the book every time it's called.


### CODE

```vb
Option Compare Database
Option Explicit

'**************************
'*Excel Book Creation Class
'**************************

'Consts
Private currentSheetNum As Long

'Vars
Private myXlApp As Object
Private myWorkBook As Object
Private dicWorkSheet As Dictionary 'store all sheet objects of the book


'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Get xlApplication() As Object
    Set xlApplication = myXlApp
End Property


Public Property Get Workbook() As Object
    Set Workbook = myWorkBook
End Property


Public Property Get WorkSheets(ByVal num As Long) As Object
    If Not dicWorkSheet.Exists(num) Then Call MsgBox("The Sheet does not exists.", vbExclamation, TOOL_NAME): Set WorkSheets = Nothing: Exit Property
    Set WorkSheets = dicWorkSheet.Item(num)
End Property


'******************************************************************************************
'******************************************************************************************
Private Sub Class_Initialize()
    
    'Consts
    Const FUNC_NAME As String = "Class_Initialize"
    
    'Vars
    
    On Error GoTo ErrorHandler
    
    'initial sheet number
    currentSheetNum = 1
    
    'instance of ExcelApp
    Set myXlApp = CreateObject("Excel.Application")
    With myXlApp
        'all processing are done n the background
        .Visible = False
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    Set myWorkBook = myXlApp.Workbooks.Add
    Set dicWorkSheet = New Dictionary
    dicWorkSheet.Add currentSheetNum, myWorkBook.WorkSheets(currentSheetNum)
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub


'******************************************************************************************
'******************************************************************************************
Private Sub Class_Terminate()
    
    'Consts
    Const FUNC_NAME As String = "Class_Terminate"
    
    'Vars
    
    On Error GoTo ErrorHandler
    
    'save the book to user's Desktop
    With CreateObject("WScript.Shell")
        myWorkBook.SaveAs .SpecialFolders("Desktop") & "\" & "Test-Excel-Class-" & Format(Now, "yyyymmddhhnnss") & ".xlsx"
    End With
    
    'restore ExcelApp settings
    With myXlApp
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
    'close
    myWorkBook.Close
    myXlApp.Quit

ExitHandler:
    
    Set dicWorkSheet = Nothing
    Set myWorkBook = Nothing
    Set myXlApp = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Sub



'******************************************************************************************
'*Function :add new sheet
'*Return   :added sheet
'******************************************************************************************
Public Function addNewSheet() As Object
    
    'Consts
    Const FUNC_NAME As String = "addNewSheet"
    
    'Vars
    Dim ws As Object
    
    On Error GoTo ErrorHandler

    Set addNewSheet = Nothing
    
    currentSheetNum = currentSheetNum + 1
    'add new sheet at the end
    Set ws = myWorkBook.WorkSheets.Add(After:=myWorkBook.WorkSheets(myWorkBook.WorkSheets.Count))
    dicWorkSheet.Add currentSheetNum, ws
    
    Set addNewSheet = ws
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

```

### ENCAPSULATION

As you can see above, variable inner the class such as `myXlApp` and `myWorkBook` are declared in Private scope. So they can be only derived through `Property Get` procedure, that indicates we can't refer and get them directly.

In this way, it's called **Encapsulation** that variables inner class are protected from being changed except for ways we allowed.  
By encapsulation, we can keep variables from being changed and removed wrongfully, remove potential bugs, enhance a perspective of code.



### INITIALIZAITON AND TERMINATION

As you can see above class code too, in `Class_Initialize`, we can do the minimum number of tasks we want to do at once.

'The minimum number of tasks' is:
- creating a new Excel Application interface
- creating a new Excel book
- storing a first worksheet object in the book to Dictionary object
- hiding Excel App's behavior
- stopping screen updating of Excel
- deterring displaying any alert

on the other hand, `Class_Terminate` does:  
- restoring Excel App's settings
- saving the book to user's Desktop 
- closing the book and Excel App

In this way, we are free from describing processing code of these minimum tasks outer the class and free from forgetting it.  
Especially, in terms of Excel Application object, if you stored it to global variable and didn't set it `Nothing` after using it, Excel App instance will continue to run in the background. Routine processing frees you from this unexpected result.


## OTHER CODE

### getTableHeader

```vb
'******************************************************************************************
'*Function :get header data of target table
'*Arg      :table name
'*Arg      :array for gotten data
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function getTableHeader(ByVal tblName As String, ByRef pArrHeader() As String) As Boolean
    
    'Consts
    Const FUNC_NAME As String = "getTableHeader"
    
    'Vars
    Dim i As Long
    
    On Error GoTo ErrorHandler

    getTableHeader = False
    Erase pArrHeader
    
    With db.TableDefs(tblName)
        ReDim pArrHeader(0 To .Fields.Count - 1)
        For i = 0 To .Fields.Count - 1
            pArrHeader(i) = .Fields(i).Name
        Next
    End With
    
    getTableHeader = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

```

### getTableDataBySQL


```vb
'******************************************************************************************
'*Function :get recordset data as a 2-dimentional array
'*Arg      :sql string for target recordset
'*Arg      :array for gotten data
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function getTableDataBySQL(ByVal sql As String, ByRef arrData() As Variant) As Boolean
    
    'Consts
    Const FUNC_NAME As String = "getTableDataBySQL"
    
    'Vars
    Dim rs As DAO.Recordset
    Dim i As Long
    Dim j As Long
    
    On Error GoTo ErrorHandler

    getTableDataBySQL = False
    Erase arrData
    
    Set rs = db.OpenRecordset(sql)
    With rs
        If .EOF Then GoTo TruePoint
        .MoveLast
        ReDim arrData(0 To .RecordCount - 1, 0 To .Fields.Count - 1)
        .MoveFirst
        
        i = 0
        Do Until .EOF
            For j = 0 To .Fields.Count - 1
                arrData(i, j) = .Fields(j).Value
            Next j
            i = i + 1
            .MoveNext
        Loop
    End With

TruePoint:

    getTableDataBySQL = True
    
ExitHandler:
    
    If Not rs Is Nothing Then rs.Clone: Set rs = Nothing
    
    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

```

### postDataToSheet

```vb
'******************************************************************************************
'*Function :post data to sheet
'*Arg      :target sheet
'*Arg      :assigned sheet name
'*Arg      :aheader data array
'*Arg      :data array
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function postDataToSheet( _
    ByVal tgtSheet As Object, _
    ByVal sheetName As String, _
    ByVal pArrHeader As Variant, _
    ByVal pArrData As Variant _
) As Boolean
    
    'Consts
    Const FUNC_NAME As String = "postDataToSheet"
    
    'Vars
    
    On Error GoTo ErrorHandler

    postDataToSheet = False
    
    With tgtSheet
        .Name = sheetName
        .Range(.cells(1, 1), .cells(1, UBound(pArrHeader) - LBound(pArrHeader) + 1)).Value = pArrHeader
        .Range(.cells(2, 1), .cells(UBound(pArrData, 1) - LBound(pArrData, 1) + 2, UBound(pArrData, 2) - LBound(pArrData, 2) + 1)).Value = pArrData
        'lines
        .Range(.cells(1, 1), .cells(UBound(pArrData, 1) - LBound(pArrData, 1) + 2, UBound(pArrData, 2) - LBound(pArrData, 2) + 1)).Borders.LineStyle = xlContinuous
        'column widths adjustment
        .Range(.Columns(1), .Columns(UBound(pArrHeader) - LBound(pArrHeader) + 1)).AutoFit
    End With

    postDataToSheet = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

```

### getJsonFromAPI


```vb
'******************************************************************************************
'*Function :get Json string from specified URL
'*Arg      :URL
'*Return   :Json string
'******************************************************************************************
Public Function getJsonFromAPI(URL As String) As String

    'Consts
    Const FUNC_NAME As String = "getJsonFromAPI"
    
    'Vars
    Dim objXMLHttp As Object
    
    On Error GoTo ErrorHandler

    getJsonFromAPI = ""
    
    Set objXMLHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        objXMLHttp.Open "GET", URL, False
        objXMLHttp.Send


    getJsonFromAPI = objXMLHttp.responseText
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, TOOL_NAME
        
    GoTo ExitHandler
        
End Function

```


## DEMO

Push the 'Execute 1. + 2. + 3. + 4.' button, and you get a Excel book in your Desktop.

![Unit Price over 1000](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645881445/learnerBlog/Excel-Class-Sample-02/en/Excel-Class-Sample-02-1_duzxww.png)

![Item Code Starts With B](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645881445/learnerBlog/Excel-Class-Sample-02/en/Excel-Class-Sample-02-2_nckroj.png)

![Data From WebAPI](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645881445/learnerBlog/Excel-Class-Sample-02/en/Excel-Class-Sample-02-3_iypvoo.png)


## SAMPLE AND SOURCE CODE

Please refer <a href="#srcURL">Here!</a>



