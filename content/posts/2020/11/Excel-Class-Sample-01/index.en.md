---
title: "VBA: THE BENEFIT OF USING CLASS AND THE WAY TO USE, WITH A SIMPLE SAMPLE PART 1"
author: dede-20191130
date: 2020-11-20T23:47:25+09:00
slug: Excel-Class-Sample-01
draft: false
toc: true
featured: true
tags: ['Excel', 'VBA','HOMEMADE', 'object-oriented']
categories: ['programming']
vba_taxo: class_how_to
archives:
    - 2020
    - 2020-11
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

Class in VBA is kind of minor and somehow hard-to-use than any other programming languages.  
But actually there are some cases in which class enables us to code more safely, with more highly maintainability, and with less bugs.  
So I would like to describe the benefits and how to use them. Moreover, I would also like to describe a simple sample of using Class as well. 

You can download Excel file created for explanation and view its source code from [<span id="srcURL"><u>here</u></span>](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2020/11/TableCreater/en)!

## THE WAY TO USE CLASS

### WHAT IS CLASS?

Class is a combination of used data information(i.e. Variables and Constants) and information of processing details the class has (i.e. Functions) in a box (i.e. Class module).  
Having said that, it's not good to combine data and functions without any rules, but it must be a variable or function that defines one entity pointed to by the class and belongs to that entity.

I have to explain the term of 'object'.  
An object is an instance of a class. Conversely, class is a blueprint of an object. A class is a description of the information of properties of an object. By instantiation, you can treat it as an entity, not as information, for the first time.

### SAMPLE OF CLASS

For example, In case that one class expresses a human being, it has data (variables) such as eye, mouth, and body, and has processing details (functions) such as running, eating, and talking.

From a MSOffice point of view, the Range object in Excel is typical class object, having variables such as Address property for reference range, Row property for row number, and Value property for cell value, and having functions such as Select method for moving the cursor to range, AutoFit method for automatically adjusting row widths or column widths.

Also, What has a certain functionality can be class.  
The example is what performs a series of processes in which it imports Excel Range object, draws lines as a table, and writes a total value of subtotal column. Let's name it as **TableCreater**. I will explain **TableCreater** in detail later.


|Entity|Variables, Constants|Functions|
|-|-|-|
|Human Being|Eye<br>Mouse<br>Body|Run<br>Eat<br>Talk|
|Excel<br>Range Object|Address<br>Row<br>Value|Select<br>AutoFit|
|TableCreater|Target Range<br>Column Number for Subtotal<br>Header Color|Draw lines<br>Set Header's Style<br>Calculate Total Value|

### THE BENEFIT OF USING IT

#### ENABLES SAFER CODING

Variables and Constants belonging to class are basically declared in a scope of `Private` and are used by functions inner the class, so they are free from having it's value affected by alterring them accidentially (referring causes compiling error in the first place).

And, in the case that you want to refer them from outer the function, you can implement getter/setter procedure using `Property Get` statement and `Property Let` statement as a dedicated function, and interact with the outside through them.   
This mechanism is called **Encapsulation** (or data hiding).

The example is following short code. This percent storing variable and Property Let procedure only accept a value between 0 and 100 as a percent and otherwise the variable is `Null`. In this way, you can enhance safety of your code by filtering or checking when getting or setting a value.


```vb
Private percentVal As Variant

'**************************
'*Setter
'**************************
Public Property Let percent(ByVal v As Long)
    If v < 0 Or 100 < v Then percentVal = Null: Exit Property
    percentVal = v
End Property

```







#### ENABLES MORE HIGHLY MAINTAINABLE CODING

If you design functions of class use only its arguments or variables and constants declared inner the class as much as possible, this will reduce the frequency of using and being used by variables and functions from other classes and modules, and reduce the bondability of the codes to each other.  
This mechanism is called **loose coupling**.

In this way, most of the impact of changing a processing of function and adding new functions can be inside the class, and you are free from suffering from the risk of unexpected behavior when changing the specification.

Also, this approach enhances the perspective of code and you can write code everyone can read easily.





#### REDUCES CODING MISTAKES BY EXECUTING A ROUTINE PROCESS

Class in VBA has functions which run when creating and destroying objects of the class.  
When creating, `Class_Initialize()` runs, When destroying, `Class_Terminate()` runs.

If the routine processing which must be done when creating and destroying is described in these functions, this keep you from forgetting the processing or writing wrong processing.

Below is examples.
- If the class use Mail Item of Outlook, you may want to get Outlook Application as soon as the class object is created, or may want to have the Mail Imte visible as soon as the class object is destroyed. If you forget to change its visibility, the macro user won't be aware of the mail created background.
- If the class use Recordset of Access, you may want to connect Access Database as soon as the class object is created. Also, it's good to close RecordSet as soon as the class object is destroyed, in some cases, database itself too. If you forget to close the database, memory leaking or other unwelcome consequences may come.




### KEEP IN MIND: CLASS FUNCTIONALITY OF VBA IS WEAK RELATIVELY

You had better realize the weakness of VBA class.  

- `Class_Initialize` doesn't have any arguments. So you can't assign class variables some value as soon as the class object is created. 
- In VBA, there is not the concept of Class Inheritance.
- Class can have neither static variable or static function. That means you can't use these variables and functions unless you instantiate the class once.

I think these uncomfortabilities are weakness of VBA.






### ONE SAMPLE OF HOW TO USE VBA CLASS 

#### CREATION ENVIRONMENT
Windows10  
MSOffice 2016

#### i. CREATE A CLASS MODULE IN VBE

TableCreater mentioned above appears again.

In VBE (Development Environment of VBA), select class module in insertion tab and create it.

```vb
Option Explicit

'**************************
'*TableCreater
'**************************

```

#### ii. DECLARE ITS VARIABLES AND CONSTANTS

As a variable, prepare following.  

- Target Range object
- column number of subtotal column
- header cell's color

Concurrently, describe `Property Let` procedure and `Property Set` procedure.

```vb
Option Explicit

'**************************
'*TableCreater
'**************************


'Const
Private Const HEADER_COLOR = 15917529            'header cell color

'Vars
Private myRange As Range                         'range of target table
Private myColumnSubTotal As Long                 'column number of subtotal


'******************************************************************************************
'*getter/setter
'******************************************************************************************


Public Property Set Range(ByVal pRng As Range)
    Set myRange = pRng
End Property


Public Property Let ColumnSubTotal(ByVal num As Long)
    'prohibit being refered with the Range isn't set yet
    If myRange Is Nothing Then Err.Raise 1000, , "The range is not set."
    'error if argument number is out of range of 'range' object.
    If num < myRange.EntireColumn(1).Column Or myRange.EntireColumn(myRange.EntireColumn.Count).Column < num Then Err.Raise 1001, , "Invalid column number specification."
    'set
    myColumnSubTotal = num
End Property


```

#### iii. Class_Initialize、Class_Terminateを記述する

Normally, We describe `Class_Initialize` and `Class_Terminate`.  
But this itme there is nothing to do in them.



```vb
Option Explicit

'**************************
'*TableCreater
'**************************


'(omission)


'******************************************************************************************
'*Function ：
'*Arg      ：
'******************************************************************************************
Private Sub Class_Initialize()
    
    'Const
    Const FUNC_NAME As String = "Class_Initialize"
    
    'Vars
    
    On Error GoTo ErrorHandler
    
    'There's nothing special to do here this Class.

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Sub



```


#### iv. DESCRIBE EACH FUNCTION

Describe each processing detail as one function as a time.

- Draw lines
- Set header part's style
- Calculate total value and output



```vb
Option Explicit

'**************************
'*TableCreater
'**************************


'(omission)

'******************************************************************************************
'*Function ：draw lines
'*Arg      ：
'*Return   ：True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function drawLines() As Boolean
    
    'Const
    Const FUNC_NAME As String = "drawLines"
    
    
    On Error GoTo ErrorHandler

    drawLines = False
    
    'prohibit being called with the Range isn't set yet
    If myRange Is Nothing Then Err.Raise 1000, , "The range is not set."
    
    'draw lines
    myRange.Borders.LineStyle = xlContinuous

    drawLines = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Function


'******************************************************************************************
'*Function ：set header part's style
'            The header is cells of the first row of given range.
'*Return   ：True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function setStyleForHeader() As Boolean
    
    'Const
    Const FUNC_NAME As String = "setStyleForHeader"
    
    
    On Error GoTo ErrorHandler

    setStyleForHeader = False
    
    'prohibit being called with the Range isn't set yet
    If myRange Is Nothing Then Err.Raise 1000, , "The range is not set."
    
    'change styles
    With myRange.Rows(1)
        'change background color
        .Interior.color = HEADER_COLOR
        'change font weight to bold
        .Font.Bold = True
        'change text alignment to center
        .HorizontalAlignment = xlCenter
    End With
        
    
    setStyleForHeader = True
    
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Function




'******************************************************************************************
'*Function ：calculate total value from subtotal column and output it
'*Return   ：True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function calcTotalFromSubTotal() As Boolean
    
    'Const
    Const FUNC_NAME As String = "calcTotalFromSubTotal"
    
    'Vars
    Dim sumVal As Long
    Dim cell As Range
    Dim subTotalOrder As Long
    
    On Error GoTo ErrorHandler

    calcTotalFromSubTotal = False
    
    'prohibit being called with the Range isn't set yet
    If myRange Is Nothing Then Err.Raise 1000, , "The range is not set."
    
    'prohibit being called with the column number for subtotal isn't set yet
    If myColumnSubTotal = 0 Then Err.Raise 1002, , "The column number for subtotal is not set."
        
    'calculate the order of subtotal column
    subTotalOrder = myColumnSubTotal - myRange(1).Column + 1
        
    'calculate total value, except for header row
    For Each cell In myRange.Columns(subTotalOrder).Cells.Offset(1).Resize(myRange.Columns(subTotalOrder).Cells.Offset(1).Cells.Count - 1)
        'add only numeric value
        If IsNumeric(cell.Value) Then sumVal = sumVal + cell.Value
    Next cell
    If sumVal = 0 Then GoTo TruePoint
    
    'write the total value in the bottom cell of subtotal column
    With myRange.Columns(subTotalOrder).Rows(myRange.Columns(subTotalOrder).Cells.Count).Offset(1)
        .Value = sumVal
        'refer the label cell
        With .Offset(, -1)
            'write a label
            .Value = "Total"
            'draw lines to label cell and total cell
            .Resize(, .Columns.Count + 1).Borders.LineStyle = xlContinuous
        End With
        
    End With
       
TruePoint:
       
    calcTotalFromSubTotal = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Function

```

At this point coding for class is completed.  
You got the class expressing what creates table with data and behaviors.



#### v. USE THE CLASS FROM AN EXTERNAL FUNCTION

Following is 'base' sheet containing two table data.

![Base](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645855941/learnerBlog/Excel-Class-Sample-01/en/Excel-Class-Sample-01_xsxbih.png)

Based on each table data, I created functions which creates corresponging table in new sheet.  
Of course theses functions uses TableCreater.




```vb

'******************************************************************************************
'*Function ：create a table for template A in base sheet through TableCreater
'            creation location: new sheet
'******************************************************************************************
Public Sub TestTemplateA()
    
    'Const
    Const FUNC_NAME As String = "TestTemplateA"
    
    'Vars
    Dim ws As Worksheet
    Dim tableRange As Range
    Dim objTableCreater As TableCreater
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook
        'create new sheet
        Set ws = .Worksheets.Add(, .Worksheets(.Worksheets.Count))
        ws.Name = FUNC_NAME & "_" & Format(Now, "yyyymmddhhnnss")
        
        'copy template range from
        Set tableRange = ws.Range(ws.Cells(2, 2), ws.Cells(9, 4))
        tableRange.Value = .Worksheets(BASE_SHEET).Range(.Worksheets(BASE_SHEET).Cells(3, 2), .Worksheets(BASE_SHEET).Cells(10, 4)).Value
        
        'instanciate TableCreater
        Set objTableCreater = New TableCreater
        
        'set params
        Set objTableCreater.Range = tableRange
        objTableCreater.ColumnSubTotal = 4
        
        'draw lines: if error, shift to the exit process
        If Not objTableCreater.drawLines Then GoTo ExitHandler
         
        'set styles for header part for emphasis: if error, shift to the exit process
        If Not objTableCreater.setStyleForHeader Then GoTo ExitHandler
        
        'calc total: if error, shift to the exit process
        If Not objTableCreater.calcTotalFromSubTotal Then GoTo ExitHandler
        
        'adjust column widths
        tableRange.EntireColumn.AutoFit
        
    End With
    

ExitHandler:
    
    'release memory
    Set objTableCreater = Nothing
    Set ws = Nothing
    Set tableRange = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Sub






'******************************************************************************************
'*Function ：create a table for template B in base sheet through TableCreater
'            creation location: new sheet
'******************************************************************************************
Public Sub TestTemplateB()
    
    'Const
    Const FUNC_NAME As String = "TestTemplateB"
    
    'Vars
    Dim ws As Worksheet
    Dim tableRange As Range
    Dim objTableCreater As TableCreater
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook
        'create new sheet
        Set ws = .Worksheets.Add(, .Worksheets(.Worksheets.Count))
        ws.Name = FUNC_NAME & "_" & Format(Now, "yyyymmddhhnnss")
        
        'copy template range from
        Set tableRange = ws.Range(ws.Cells(2, 2), ws.Cells(8, 8))
        tableRange.Value = .Worksheets(BASE_SHEET).Range(.Worksheets(BASE_SHEET).Cells(13, 2), .Worksheets(BASE_SHEET).Cells(19, 8)).Value
        
        'instanciate TableCreater
        Set objTableCreater = New TableCreater
        
        'set params
        Set objTableCreater.Range = tableRange
        objTableCreater.ColumnSubTotal = 8
        
        'draw lines: if error, shift to the exit process
        If Not objTableCreater.drawLines Then GoTo ExitHandler
         
        'set styles for header part for emphasis: if error, shift to the exit process
        If Not objTableCreater.setStyleForHeader Then GoTo ExitHandler
        
        'calc total: if error, shift to the exit process
        If Not objTableCreater.calcTotalFromSubTotal Then GoTo ExitHandler
        
        'adjust column widths
        tableRange.EntireColumn.AutoFit
        
    End With
    

ExitHandler:
    
    'release memory
    Set objTableCreater = Nothing
    Set ws = Nothing
    Set tableRange = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "TableCreater"
        
    GoTo ExitHandler
        
End Sub


```

Running them results:


![Table A](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645855940/learnerBlog/Excel-Class-Sample-01/en/Excel-Class-Sample-01-2_qo4izn.png)
![Table B](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645855940/learnerBlog/Excel-Class-Sample-01/en/Excel-Class-Sample-01-3_ydb2sn.png)

#### vi. WHAT YOU GOT FROM USING THE CLASS

With TableCreater, I think the perspective of code is good.  
We'll realize that the sentence including `objTableCreater.something` is related to creation of table.

## AT THE END

### TABLE_CREATER SAMPLE AND SOURCE CODE

Please refer <a href="#srcURL">Here!</a>


### SEQUEL OF THIS ARTICLE

With a example of TableCreater, I regretted a little that the benefit of encapsulation and routine processing in initialization and termination is not expressed.  
So, I created a sequel of this article.

{{< page-titled-link page="Excel-Class-Sample-02" >}}
