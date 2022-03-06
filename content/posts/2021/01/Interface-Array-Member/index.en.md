---
title: "VBA: HOW TO USE ARRAYS AS A CLASS MEMBER WHEN USING INTERFACE INHERITANCE"
author: dede-20191130
date: 2021-01-13T19:49:08+09:00
slug: Interface-Array-Member
draft: false
toc: true
featured: false
tags: ['VBA','object-oriented']
categories: ['programming']
vba_taxo: inheritance
archives:
    - 2021
    - 2021-01
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

In VBA we can do object-oriented programming with *interface* feature.

But since there is one restriction against class variables as described below,   
we'll encounter syntax errors when we set an array  accessible from the outside as a class variable in a class implementing interface.

In this article I'd like to describe the case and how to avoid it.

You can download Excel file created for explanation and view its source code from [<span id="srcURL"><u>here</u></span>](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2021/01/Interface-Array-Member/en)!


## CREATION ENVIRONMENT

Windows 10 Home(64bit)  
MSOffice 2016

## WHAT THE CONSTRUCTION IS

### ABOUT

There are two team class `clsAnalysisTeam` and `clsNewTeam`, both implementing interface `clsAbsTeam`.

Each team class has an array to store team member's name and the method to get member's name.


### CLASS DIAGRAM

![Class Diaglam Without Its Variables And Functions](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646464425/learnerBlog/Interface-Array-Member/en/class_l3ufdu.svg)

## ERROR OCCURED

### THE CASE

I wrote each class as follows:


```vb
Option Explicit

'**************************
'*Team Class Interface
'**************************

'Consts

'Vars


'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Let arrayMenberName(ByVal idx As Long, ByVal name As String)

End Property

Public Property Get arrayMenberName(ByVal idx As Long) As String

End Property



'Functions
Public Function getMemberName(ByVal idx As Long) As String

End Function

```

```vb
Option Explicit

Implements clsAbsTeam

'**************************
'*Team Class:  Analysis Teram
'**************************

'Consts

'Vars



'******************************************************************************************
'*Function : get the menber name of index
'*Arg      : index number of target member
'*Return   : the name
'******************************************************************************************
Private Function clsAbsTeam_getMemberName(ByVal idx As Long) As String
    
    'Consts
    
    'Vars
    
    '*** here name-returning process is inserted ***
    
    
ExitHandler:

    Exit Function
        
End Function
```

```vb
Option Explicit

Implements clsAbsTeam

'**************************
'*Team Class:  New Team
'**************************

'Consts

'Vars



'******************************************************************************************
'*Function : get the menber name of index, but new team has no member so it returns 'no member'.
'*Arg      : index number of target member
'*Return   : the name
'******************************************************************************************
Private Function clsAbsTeam_getMemberName(ByVal idx As Long) As String
    
    'Consts
    
    'Vars
    
    '*** here name-returning process is inserted ***
    
    
ExitHandler:

    Exit Function
        
End Function

```

Then, in order to give the class an array to store the member's name,  
I added an array variable which scope is *Public*.


```vb

'**************************
'*Team Class Interface
'**************************

'Consts

'Vars
Public arrayMenberName(1 To 6) As String


'Functions
Public Function getMemberName(ByVal idx As Long) As String

End Function

```

Then, the compile error which says,   
{{< colored-span color="red" >}}Constants, fixed-length strings, arrays, user-defined types, and Declare statements not allowed as Public members of an object module{{< /colored-span >}}   
, apeears.


![Compile Error](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646464425/learnerBlog/Interface-Array-Member/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-03-05_151929_eeqzkd.png)

### CAUSE

The VB6 specification doesn't allow the variables of some types such as an array or user defined types to be set in class module.  
Thus we should prepare the mechanism which enables these variables to get accessed from outer modules.



### HOW TO AVOID

In interface module I defined only getter/setter functions,  
and gave the implementing class an array variable for outer modules to access the array through getter/setter.


```vb
Option Explicit

'**************************
'*Team Class Interface
'**************************

'Consts

'Vars


'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Let arrayMenberName(ByVal idx As Long, ByVal name As String)

End Property

Public Property Get arrayMenberName(ByVal idx As Long) As String

End Property



'Functions
Public Function getMemberName(ByVal idx As Long) As String

End Function

```

```vb
Option Explicit

Implements clsAbsTeam

'**************************
'*Team Class:  Analysis Teram
'**************************

'Consts

'Vars
Private myArrayMenberName(1 To 6) As String 'Max member number of 6



'******************************************************************************************
'*getter/setter
'******************************************************************************************
Private Property Let clsAbsTeam_arrayMenberName(ByVal idx As Long, ByVal name As String)
    myArrayMenberName(idx) = name
End Property

Private Property Get clsAbsTeam_arrayMenberName(ByVal idx As Long) As String
    clsAbsTeam_arrayMenberName = myArrayMenberName(idx)
End Property



'******************************************************************************************
'*Function : get the menber name of index
'*Arg      : index number of target member
'*Return   : the name
'******************************************************************************************
Private Function clsAbsTeam_getMemberName(ByVal idx As Long) As String
    
    'Consts
    
    'Vars
    
    clsAbsTeam_getMemberName = "The " & idx & "th team member is " & myArrayMenberName(idx)
    
    
ExitHandler:

    Exit Function
        
End Function
```

```vb
Option Explicit

Implements clsAbsTeam

'**************************
'*Team Class:  New Team
'**************************

'Consts

'Vars


'******************************************************************************************
'*getter/setter
'******************************************************************************************
Private Property Let clsAbsTeam_arrayMenberName(ByVal idx As Long, ByVal name As String)
    'nothing to do
End Property

Private Property Get clsAbsTeam_arrayMenberName(ByVal idx As Long) As String
    clsAbsTeam_arrayMenberName = "There is no member in this new team."
End Property

'******************************************************************************************
'*Function : get the menber name of index, but new team has no member so it returns 'no member'.
'*Arg      : index number of target member
'*Return   : the name
'******************************************************************************************
Private Function clsAbsTeam_getMemberName(ByVal idx As Long) As String
    
    'Consts
    
    'Vars
    
    clsAbsTeam_getMemberName = "There is no member in this new team."
    
    
ExitHandler:

    Exit Function
        
End Function
```

By writing above, we can avoid the compile error while setting an array to behave as expected in implementing class.

Incidentally, `clsNewTeam` has no member and the class doesn't have to have an array variable.


## SAMPLE

### CLASS DIAGRAM REVISIT

![Class Diaglam](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646464425/learnerBlog/Interface-Array-Member/en/classWithFunc_w2ujj4.svg)

### CALLING FUNCTION CODE

```vb
Option Explicit

'**************************
'*Calling Module
'**************************


'******************************************************************************************
'*FUnction:  operation testing function
'*Arg     :
'*Return  :  True > normal termination; False > abnormal termination
'******************************************************************************************
Public Sub testFunc()
    
    'Consts
    Const FUNC_NAME As String = "testFunc"
    
    'Vars
    Dim team As clsAbsTeam
    Dim coll As New Collection
    
    On Error GoTo ErrorHandler

    'set names for analysis team members
    Set team = New clsAnalysisTeam
    team.arrayMenberName(1) = "佐藤"
    team.arrayMenberName(3) = "Mike"
    team.arrayMenberName(5) = "Abdallah"
    
    'add analysis team
    'add new team
    coll.Add team
    coll.Add New clsNewTeam
    
    'output 3rd member's name for each team
    If Not outputSelectedMemberName(coll, 3) Then GoTo ExitHandler

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name: " & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Interface-Array-Member"
        
    GoTo ExitHandler
        
End Sub


'******************************************************************************************
'*FUnction: Outputs the name of the member whose number is given by index
'*Arg     : collection of the team. All of them implements clsAbsTeam.
'*Arg     : the index number
'*Return  : True > normal termination; False > abnormal termination
'******************************************************************************************
Private Function outputSelectedMemberName(ByVal collTeam As Collection, ByVal idx As Long) As Boolean
    
    'Consts
    Const FUNC_NAME As String = "outputSelectedMemberName"
    
    'Vars
    Dim cntTeam As clsAbsTeam
    
    On Error GoTo ErrorHandler

    outputSelectedMemberName = False
    
    For Each cntTeam In collTeam
        Debug.Print cntTeam.getMemberName(idx)
    Next cntTeam

TruePoint:

    outputSelectedMemberName = True

ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name: " & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Interface-Array-Member"
        
    GoTo ExitHandler
        
End Function

```

### DEMO

Run `testFunc`, and the following data will be output in Immediate Window.


```
The 3th team member is Mike
There is no member in this new team.
```




