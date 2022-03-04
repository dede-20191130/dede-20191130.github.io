---
title: "EXCEL VBA: TYPICAL PATTERNS FOR EXCEPTION HANDLING & SAMPLE OF HOW TO USE"
author: dede-20191130
date: 2020-12-05T18:02:12+09:00
slug: VBA-Exception-Handle
draft: false
toc: true
featured: true
tags: ['Excel', 'VBA','HOMEMADE']
categories: ['programming']
vba_taxo: vba_coding_sample
archives:
    - 2020
    - 2020-12
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE


In programming, it's a general implementaion to do special processing for handling a exception when a error within expectation or a custom error has occured.

VBA has a functionality of exception handling, but it has some complicated features than late languages,  
so I wrote templates for typical exception handling patterns and sample of how to use in this article.

You can download Excel file created for explanation, files for test, and view its source code from [<span id="srcURL"><u>here</u></span>](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2020/12/Exception-Handling/en)!



## WAHT IS EXCEPTION HANDLING

### THE MECHANISM

Let's consider what happens if the program without any exception handlings runs.

When a error occurs during running,   
it triggers behaviors determined by its running environment, such as outputting the error infomation to console window and displaying a error message,   
and sometimes the process itself is interrupted at the line where the error occurred.

in terms of tool or application, it's normally not so desirable situation behavior.   
Because you might want to recover the process so far so the user doesn't recognize it when the error has occured.

So exception handling is very useful. It can separate normal and abnormal processings.



### TYPICAL PATTERNS BY PYTHON CODES

#### TYR-EXCEPT

For exmaple, we describe it in Python.


```Python
def divide_each(a, b):
    try:
        print(a / b)
    except ZeroDivisionError as e:
        print('catch ZeroDivisionError:', e)
    except TypeError as e:
        print('catch TypeError:', e)

divide_each(1, 0)　# first calling
# catch ZeroDivisionError: division by zero

divide_each('a', 'b')　# second calling
# catch TypeError: unsupported operand type(s) for /: 'str' and 'str'
```

`divide_each` function calculates the division of `a/b` and catch each errors happening at that time and print each error information.

The errors captured are:   
- the error due to denominator being zero (ZeroDivisionError)
- the error due to the type of one of the arguments is not a number type (TypeError)

The former is printed to console (standard output) as `catch ZeroDivisionError: division by zero`,   
and the latter is printed as `catch ZeroDivisionError: division by zero`. 

This statement of Python is called {{< colored-span color="#fb9700" >}}try-except{{< /colored-span >}}.



#### TERMINATION PROCESSING BY FINALLY STATEMENT

Either in the case of normal termination or in the case that error has occured and captured in the except clause,   
when you want to do the process which must be executed finally, you can add {{< colored-span color="#fb9700" >}}finally clause{{< /colored-span >}} as follows:  




```python
def divide_each(a, b):
    try:
        print(a / b)
    except ZeroDivisionError as e:
        print('catch ZeroDivisionError:', e)
    except TypeError as e:
        print('catch TypeError:', e)
    finally:
        print('passed end processing')


divide_each(1, 0)
# catch ZeroDivisionError: division by zero

divide_each('a', 'b')
# catch TypeError: unsupported operand type(s) for /: 'str' and 'str'

```

Run this, and prints as follows:


```language
catch ZeroDivisionError: division by zero
passed end processing
catch TypeError: unsupported operand type(s) for /: 'str' and 'str'
passed end processing
```



## EXCEPTION HANDLING ON VBA

### FEATURES

Annoyingly, In VBA, there is no definition of exception handling as a fixed syntax like {{< colored-span color="#fb9700" >}}try-except statement{{< /colored-span >}} in Python.  
Thus, this means we have to define the handlings by ourselves which line the process jumps if error has occured, and which line for termination processing.

The statements for realizing above are `GOTO` and `Error`.



||||
|-|-|-|
|Goto Statement|moves to a specified line <br/>unconditionally|[Official Link](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/goto-statement)|
|On Error Statement|defines the program behavior on error.<br/><br/>move to a specified line on error in combination with `GOTO`<br/>disables branching processing itself on error|[Official Link](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement)|


### TEMPLATES

VBA has mainly two types of procedure: Sub Procedure and Function Procedure.

For each, I wrote a template for typical exception handling pattern.


#### SUB PROCEDURE

```vb
'******************************************************************************************
'*Function :template for sub-procedure
'******************************************************************************************
Public Sub subTemplate()
    
    'Consts
    Const FUNC_NAME As String = "subTemplate"
    
    'Vars
    
    On Error GoTo ErrorHandler

    '---write processing---
    

ExitHandler:
    
    '---write termination processing---
    
    Exit Sub
    
ErrorHandler:
    
    '---write processing for excetion---
    '   - show message
    '   - write the sysmte error infomation into a logfile
    '   - create a e-mail to notice the system error and send it
    
    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Sub

```

The declaration of `On Error GoTo ErrorHandler` set that the process moves to the line labelled as `ErrorHandler` on error.  
`ErrorHandler` corresponds to the `except` clause in Python.

After error handling such as writing log file or message displaying, the process moves to termination processing by order of `GoTo ExitHandler` snippet.  
It corresponds to the `finally` clause in Python.



#### FUNCTION PROCEDURE

In Function Procedure, two templates are possible, depending on the method used to inform the calling function that an error has occurred.

##### TEMPLATE 1

```vb
'******************************************************************************************
'*Function :template for function-procedure no1
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function functionTemplate01() As Boolean
    
    'Consts
    Const FUNC_NAME As String = "functionTemplate01"
    
    'Vars
    
    On Error GoTo ErrorHandler

    functionTemplate01 = False
    
    '---write processing---

TruePoint:
    
    '---write termination processing only when normal termination---
    
    functionTemplate01 = True

ExitHandler:
    
    '---write termination processing---
    
    Exit Function
    
ErrorHandler:

    '---write processing for excetion---
    '   - show message
    '   - write the sysmte error infomation into a logfile
    '   - create a e-mail to notice the system error and send it
    
    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Function

```

One big different between above Sub Procedure and Function Procedure is that the latter has a return value whose type is boolean, and it's true if normal termination and false if termination with error.

On error `functionTemplate01 = True` line isn't passed and the error is communicated to the calling function, which notices an abnormal termination.



```vb
if not functionTemplate01() then Call Msgbox("The calling of the functionTemplate01 is incorrect.")
```





##### TEMPLATE 2

```vb
'******************************************************************************************
'*Function :template for function-procedure no2
'*Return   :any type except for Null > normal termination; Null > abnormal termination
'******************************************************************************************
Public Function functionTemplate02() As Variant
    
    'Consts
    Const FUNC_NAME As String = "functionTemplate02"
    
    'Vars
    
    On Error GoTo ErrorHandler

    functionTemplate02 = Null
    
    '---write processing---

ExitHandler:
    
    '---write termination processing---
    
    Exit Function
    
ErrorHandler:

    '---write processing for excetion---
    '   - show message
    '   - write the sysmte error infomation into a logfile
    '   - create a e-mail to notice the system error and send it
    
    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Function

```

One big different against above Function Procedure Template 1 is that template 2 has a return value whose type is `Variant`, and the process terminates normally if the type is anything except for `Null` and abnormally if the type is `Null`.

FIrst of the process, the line `functionTemplate02 = Null` is run, and on the way `functionTemplate02` is assigned the value you want to return.

You use `isNull` function to detect the abnormal termination.



```vb
Dim returnValue as Variant
returnValue = functionTemplate02()
if isNull(returnValue) then Call Msgbox("The calling of the functionTemplate01 is incorrect.")
```

### SAMPLE OF HOW TO USE

#### BEHAVIORS OF SAMPLE

1. Main function calls *funcSample01*.
*funcSample01* retrieves a array of file paths from specified worksheet in Excel book.
2. Main function calls *funcSample02*.
*funcSample02* opens a Excel file whose path is given as a argument.   
It then write something into A1 Cell in first and second worksheet, and close the file.

I'd like to describe the error occurence and exception handling flow in them.


#### THE ENVIRONMENT IN WHICH I CREATED

Windows 10 Home(64bit)  
MSOffice 2016

#### SAMPLE EXCEL FILE WITH TEST DATA FILES

[The sample file](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2020/12/Exception-Handling/en) has a *FilePath* worksheet containg a total of three relative file paths of data files for test.


![FilePath Sheet](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646398776/learnerBlog/VBA-Exception-Handle/en/VBA-Exception-Handle1_cs2k2t.png)

|File Name|Value Of A1 Cell<br/>In First Sheet|Has Second Sheet|
|-|-|-|
|foo.xlsx|*Enpty*|False|
|mario.xlsx|'FireBall'|False|
|bar.xlsx|*Enpty*|True|




#### PROCESSING FLOW DIAGRAM

![Processing Flow Diaglam](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646394025/learnerBlog/VBA-Exception-Handle/en/Dependencies_Of_Functions_og3brh.svg)

#### CODE

##### CALLER SUB PROCEDURE

```vb
'******************************************************************************************
'*Function :exception handling sample main
'******************************************************************************************
Public Sub main()
    
    'Consts
    Const FUNC_NAME As String = "main"
    
    'Vars
    Dim filePathArr As Variant
    Dim filePath As Variant
    Dim sheetName As String
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    'call funcSample01 with a sheet name which doesn't exist as a argument.
    sheetName = "sheetNotExist"
    filePathArr = funcSample01(sheetName)
    'show message if Null value is returned
    If IsNull(filePathArr) Then MsgBox sheetName & "The '" & sheetName & "' sheet doesn't exist." & vbNewLine & "Failed to retrieve the file path array, but the process continues."
    
    'call funcSample01 with a sheet name which exists as a argument.
    sheetName = "FilePath"
    filePathArr = funcSample01(sheetName)
    'show message if Null value is returned
    If IsNull(filePathArr) Then MsgBox sheetName & "The '" & sheetName & "' sheet doesn't exist." & vbNewLine & "Failed to retrieve the file path array, but the process continues."
    
    'call funcSample02 with each excel file path
    For Each filePath In filePathArr
        'if there is already some text in A1 cell, output the path in which the process failed to write into Immediate Window
        If Not funcSample02(ThisWorkbook.Path & filePath) Then
            Debug.Print "The file path in which the process failed to write: " & filePath
        End If
    Next filePath
    
    'the other errors not caught by funcSamples are caught by ErrorHandler labeded line in this procedure

ExitHandler:
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Sub

```

##### funcSample01

```vb
'******************************************************************************************
'*Function :example of function procedure containing a exception handling no1
'*          get a array of file paths
'*Arg      :worksheet name
'*Return   :array > normal termination; Null > abnormal termination
'******************************************************************************************
Public Function funcSample01(ByVal wsName As String) As Variant
    
    'Consts
    Const FUNC_NAME As String = "funcSample01"
    
    
    On Error GoTo ErrorHandler

    funcSample01 = Null
    
    'get a array of the values from A1 cell to A3 cell
    With ThisWorkbook.Worksheets(wsName)
        funcSample01 = .Range("A1:A3").Value
    End With

ExitHandler:
    
    
    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Function
```

##### funcSample02

```vb
'******************************************************************************************
'*Function :example of function procedure containing a exception handling no1
'*          open a excel file whose path is given as a argument
'*          write current time in A1 cell of first sheet
'*          if second sheet exists, write 'Completed' in A1 cell of it
'*Arg      :the excel file path
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function funcSample02(ByVal filePath As String) As Boolean
    
    'Consts
    Const FUNC_NAME As String = "funcSample02"
    
    'Vars
    Dim wb As Workbook
    
    On Error GoTo ErrorHandler

    funcSample02 = False
    
    Set wb = Workbooks.Open(filePath)
    
    
    With wb
        'write current time
        'a error occurs if there is already a text in A1. This is an abnormal termination
        If Trim(.Worksheets(1).Range("A1").Value) <> "" Then Err.Raise 1000, , "There is already a text in A1 Cell."
        .Worksheets(1).Range("A1").Value = Now
        
        'this process terminates normally if second sheet doesn't exist
        If .Worksheets.Count < 2 Then GoTo TruePoint
        
        'write 'Completed'
        .Worksheets(2).Range("A1").Value = "Completed"
        
    End With
    

TruePoint:
    
    'save the book
    wb.Save
    
    funcSample02 = True

ExitHandler:
    
    'never fail to close the book whether if this process terminates normally or abnormally.
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    
    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Function

```

The following part prevents the book from remaining opening after the entire process is terminated.

```vb
ExitHandler:
    'never fail to close the book whether if this process terminates normally or abnormally.
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
```




#### DEMO

Run `main()`.

i. Call `funcSample01` and a system error message is displayed and the message of the failure to get file path follows.  

But the proper exception handling is put in, so the main process escapes being interrupted and continues with the next line.

![Error In funcSample01](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646398776/learnerBlog/VBA-Exception-Handle/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-03-04_215746_roy6wd.png)  

![Failure To Get File](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646398776/learnerBlog/VBA-Exception-Handle/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-03-04_215801_t2jkqd.png)

ii. In the loop of calling `funcSample02`, when the process opened mario.xlsx and is trying to write a text in A1 Cell,  
The position is already filled with 'FireBall', so a custom error message is displayed and error information output to Immediate Window follows.  


![Error In funcSample02](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646398776/learnerBlog/VBA-Exception-Handle/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-03-04_215814_vnnpyz.png)  

![Error Information Output](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646398777/learnerBlog/VBA-Exception-Handle/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-03-04_215839_jlqpxa.png)

But, the proper exception handling is put in too, interruption is escaped.



## SAMPLE FILE AND SOURCE CODE

Please refer <a href="#srcURL">Here!</a>




[^1]: Cited: [https://github.com/nkmk/python-snippets/blob/0bc3839319270c61ac37bd2112dd5996a4fe248b/notebook/exception_handling.py#L39-L51](https://github.com/nkmk/python-snippets/blob/0bc3839319270c61ac37bd2112dd5996a4fe248b/notebook/exception_handling.py#L39-L51)