---
title: "EXCEL VBA: I CREATED A QUOTATION CREATION TOOL."
author: dede-20191130
date: 2020-07-16T08:05:13+00:00
slug: craete-estm
draft: false
categories:
  - Application
  - programming
tags:
  - Excel
  - VBA
  - tool
  - HOMEMADE
vba_taxo: help_office
toc: true
archives:
    - 2020
    - 2020-07
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

This article describes a Excel tool for creating a quotaion.  
You can download the tool from [here](https://github.com/dede-20191130/My_VBA_Tools/blob/master/T0001_01_%E8%A6%8B%E7%A9%8D%E6%9B%B8%E4%BD%9C%E6%88%90%E3%83%84%E3%83%BC%E3%83%AB_ExcelVer/en/T0001_01_Quotation_Creation_Tool_ExcelVer.xlsm).

## TOOL OVERVIEW

The tool creates one or more quotations.  
The schemes is:
1. A user inputs necessary data to each sheets.
2. The tool inserts the data to a template quotation sheet and creates a new book containing quotations.


## SCREEN IMAGES


![Setting Area][2]

![Creation Screen][3]

![Quotation Sample][4]


## HOW TO USE

### 1. INPUT

Input the data to insert into the template quotation.

#### INPUT BASIC DATA

Input common parameters between each quotation.

In the current tool, the following values can be set.
- Tax rate
- Unit of Quantity

![basic-data sheet](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644033526/learnerBlog/craete-estm/en/t0001-en03_grso9f.png)



#### REGISTER ITEM DATA

Register items to use some quotaions.  
If duplicated code is input in imte code column, alert message will be displayed and input code is removed.

![item-data sheet](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644033526/learnerBlog/craete-estm/en/t0001-en04_nol4av.png)


#### REGISTER ITEM SET DATA

Create data of item set that contains all items used in each quotation.

- Input any ID in the item set data No. field (duplicated values allowed).
- Select item code from a dropdown list which contains all codes input in item data sheet.
- Input other data.

Repeat the above steps for the number of quotations you want to create.

![item-set-data sheet](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644033526/learnerBlog/craete-estm/en/t0001-en05_p3jybw.png)



#### REGISTER QUOTATION DATA

Input QTN(quotaion) No. and other setting item, for example, the name of the destination company, delivery date, and so on.

Select item set data number from a dropdown list which contains all numbers input in item set data sheet.

The required fields are from the QTN No. to the item set data No.

![quotation-data sheet](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644033526/learnerBlog/craete-estm/en/t0001-en06_dfrq5q.png)


### 2. CREATE QUOTATIONS



#### SPECIFY THE TARGET DATA NUMBER RANGE

Push the Quotation Creation button,  
and specify the QTN No. of target quotation on the displayed screen.

You can specify the range of numbers in the range of 1 to the max number input in the QTN No. column.

![specify the target number range](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644033527/learnerBlog/craete-estm/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-02-05_113702_j8t63v.png)



#### CREATE

Push the create button,  
And new book containing the quotations is created in the folder the tool exists.  

A name of the book is 'Quotations _ current time'.

![new book for quotations](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644033526/learnerBlog/craete-estm/en/t0001-en02_cnkiyl.png)



#### CREATE PDF AS A OPTION

In case that you enable a option of pdf creation,  
PDF files corresponding with each quotation are created and stored in the same folder.


 ## P.S.: ACCESS VERSION OF THE TOOL

in writing


 [2]: https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644033527/learnerBlog/craete-estm/en/t0001-en01_absuqj.png
 [3]: https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644033527/learnerBlog/craete-estm/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-02-05_113702_j8t63v.png
 [4]: https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644033526/learnerBlog/craete-estm/en/t0001-en02_cnkiyl.png