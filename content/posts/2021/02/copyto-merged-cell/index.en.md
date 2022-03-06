---
title: "EXCEL VBA: THE MARCO TO COPY DATA TO MERGED CELLS"
author: dede-20191130
date: 2021-02-13T11:32:30+09:00
slug: copyto-merged-cell
draft: false
toc: true
featured: true
tags: ['Excel', 'VBA','HOMEMADE']
categories: ['programming']
vba_taxo: help_office
archives:
    - 2021
    - 2021-02
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

In Excel, when copying data from normal cells to merged ones, we'll encounter a warning that says 'We can't do that to merged cells' and are prohibited to do copy-paste.

For example, in image below we try to copy a green range of 'Valerie'-'Terry' to Name column in the table, and then we'll get a error.

![Copy Trial](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646488044/learnerBlog/copyto-merged-cell/en/copyto-merged-cell1_purokp.png)

![Get a Error](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646488044/learnerBlog/copyto-merged-cell/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-03-05_222233_krcqmi.png)

To avoid this, in advance the pasted cells must be th same merging-style as copied cells, or we must create a dedicated macro.  
This article describes the latter.



## CREATION ENVIRONMENT

- windows10
- MSOffice 2016

## REQUIREMENT

1. Before the macro runs, the user copy data of target cells to clipboard.  
And after running the cells the user is selecting at that time will get data from clipboard.  
2. It can work with any format-style of both copied or pasted cells.
3. In fact, the individual cell which make up a merged cell is empty except for top left cell.  
So it's desiable for the macro to ignore empty cells.



## PROCESSING FLOW

![Processing Flow](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646488044/learnerBlog/copyto-merged-cell/en/PROCESSING_FLOW_q5lyek.svg)

## CODES

### SUB PROCEDURE: COPY_TO_MERGED_CELLS

```vb
'******************************************************************************************
'*Function : copy from cells to cells
'            both source cells and destination cells can be normal or merged
'******************************************************************************************
Public Sub copyToMergedCells()
    
    'Consts
    Const FUNC_NAME As String = "copyToMergedCells"
    
    'Vars
    Dim arr() As Variant
    Dim row As Long: row = 0
    Dim col As Long: col = 1
    Dim dicIgnoreRow As Object: Set dicIgnoreRow = CreateObject("Scripting.Dictionary")
    Dim dicIgnoreCol As Object: Set dicIgnoreCol = CreateObject("Scripting.Dictionary")
    
    On Error GoTo ErrorHandler
        
    'check if the data is in text format.
    If Application.ClipboardFormats(1) <> xlClipboardFormatText Then MsgBox "The data in the clipboard is not in text format.", vbExclamation, FUNC_NAME: GoTo ExitHandler
        
    'step 1. move data from clipboard to 2-dimensional array
    '   delimiter of rows   : vbCrLf (Carriage return & Line feed)
    '   delimiter of columns: vbTab (Tab key)
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        
        .GetFromClipboard
        Dim c, d
        ReDim arr(1 To 1, 1 To 1)
        For Each c In Split(.GetText, vbCrLf)
            row = row + 1
            If row > 1 Then arr = redimPreserveFor1stDimension(arr, row)
            'currentCol: the index of the column currently being stored
            Dim currentCol As Long: currentCol = 0
            For Each d In Split(c, vbTab)
                currentCol = currentCol + 1
                If currentCol > col Then
                    col = col + 1
                    ReDim Preserve arr(1 To row, 1 To col)
                End If
                arr(row, currentCol) = Trim(d)
            Next d
        Next c
    End With
    
    'step 2. Record the rows and columnswhose cells is all emnpty
    '   Ignore them when pasting is done.
    
    'check rows
    Dim i, j As Long
    Dim isIgnore As Boolean
    For i = 1 To UBound(arr)
        isIgnore = True
        For j = 1 To UBound(arr, 2)
            'the row isn't ignored if target element isn't empty
            If Trim(arr(i, j)) <> "" Then isIgnore = False: Exit For
        Next j
        If isIgnore Then dicIgnoreRow.Add i, True
    Next i
    'check columns
    For j = 1 To UBound(arr, 2)
        isIgnore = True
        For i = 1 To UBound(arr)
            'the column isn't ignored if target element isn't empty
            If Trim(arr(i, j)) <> "" Then isIgnore = False: Exit For
        Next i
        If isIgnore Then dicIgnoreCol.Add j, True
    Next j
    
    'step 3. paste the array starting from the selected top left cell
    Dim k, l As Long
    Dim r, tmp As Range
    Set r = Selection(1)
    For k = 1 To UBound(arr): Do
        'continue if the row must be ignored
        If dicIgnoreRow.exists(k) Then Exit Do
        'store the current range object
        Set tmp = r
        
        For l = 1 To UBound(arr, 2): Do
            'continue if the column must be ignored
            If dicIgnoreCol.exists(l) Then Exit Do
            'paste a value
            r.Value = arr(k, l)
            'move the paste destination cell one column to the right
            Set r = r.Offset(, 1)
        Loop While False: Next l
        'move the paste destination cell one row to the below
        Set r = tmp.Offset(1)
    Loop While False: Next k
    
    Application.CutCopyMode = False
    
ExitHandler:

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


### FUNCTION PROCEDURE: REDIM_PRESERVE_FOR_1ST_DIM

```vb
'******************************************************************************************
'*Function: expantion of Redim Preserve statement
'           Redim Preserve statement can't change the length of first dimension of 2-dimensional array.
'           Thus the function works well in such case.
'*Arg     : target array
'*Arg     : upper limit on the number of elements in the first dimension 
'*Return  : True > normal termination; False > abnormal termination
'******************************************************************************************
Private Function redimPreserveFor1stDimension(ByVal arr As Variant, ByVal sLen As Long) As Variant
    
    'Consts
    Const FUNC_NAME As String = "redimPreserveFor1stDimension"
    
    'Vars
    Dim tspsedArr As Variant
        
    On Error Resume Next
    
    'transpose target array
    tspsedArr = WorksheetFunction.Transpose(arr)
    'step A. Redim the second dimension of the array to the length of argument number
    ReDim Preserve tspsedArr(1 To UBound(tspsedArr, 1), 1 To sLen)

    redimPreserveFor1stDimension = WorksheetFunction.Transpose(tspsedArr)
    
    'step B. if the array is '1 * N' style, by being transposed it turns 1-dimensional array and a error occurs at the step A.
    '   Instead, at that case, the error will be ignored and the following process replaces it.
    If Err.Number = 9 Then
        Dim newArr As Variant
        Dim i As Long
        Err.Clear
        On Error GoTo ErrorHandler
        ReDim newArr(1 To UBound(arr, 1) + 1, 1 To 1)
        'reflect a existing value to new array
        For i = 1 To UBound(arr, 1)
            newArr(i, 1) = arr(i, 1)
        Next i
        redimPreserveFor1stDimension = newArr
    End If
    
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


## HOW TO USE


1. Copy the range you want.
2. Select the top left cell of the range in which you want to paste.
3. Run the function.

## DEMO

### I. FORM NORMAL CELLS TO MERGED CELLS

As you can see below,  
after copying 'Valerie'-'Terry' by pressing ctrk + c and so on, then select the Name cell of No.5 row and run the macro.  
Then name cells will be filled with the data you copied.



![Before Running](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646488044/learnerBlog/copyto-merged-cell/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-03-05_224028_oncrje.png)

![After Running](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646488044/learnerBlog/copyto-merged-cell/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-03-05_224057_spahh9.png)

### II. FROM MERGED CELLS TO NORMAL CELLS


As you can see below,  
you can copy data from address column No1-4 and paste it to blue area.


![Before Running](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646488044/learnerBlog/copyto-merged-cell/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-03-05_224201_lssrwe.png)

![After Running](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646488045/learnerBlog/copyto-merged-cell/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-03-05_224221_qhw1o5.png)

### III. FROM MERGED CELLS TO MERGED CELLS

Even if both range consist of merged cells, it works.

Data of the orange color area can be pasted to the range of Address cells of No.5-7.


![Before Running](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646488045/learnerBlog/copyto-merged-cell/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-03-05_224439_wsv4vr.png)

![After Running](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646488045/learnerBlog/copyto-merged-cell/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-03-05_224453_ayfbeb.png)

