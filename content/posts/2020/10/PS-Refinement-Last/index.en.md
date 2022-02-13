---
title: "POWERSHELL: THE WAY TO GET THE PATH LIST OF THE FILES WHICH HAVE BEEN UPDATED SINCE [SPECIFIED TIME]."
author: dede-20191130
date: 2020-10-30T00:00:04+09:00
slug: PS-Refinement-Last
draft: false 
toc: true
tags: ['Powershell']
categories: ['problem-solving']
archives:
    - 2020
    - 2020-10
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

I had a occation of working on a different PC (Windows) than usual.  
After work, I needed to sort the files which had been updated and were to be brought to my original PC,  
So I wanted the Powershell command which narrow down the files which had been updated after 9:00 today in a workspace folder.

If you can use version control apps such as Git in your environment, which manages change states automatically by commiting as appropriate, this articel is just a script practice of Powershell!




## CREATION ENVIRONMENT

|||
|-|-|
|PSVersion|5.1|


## COMMAND



Output the information to the console.
```Powershell
ls -r  -File | ?{$_.LastWriteTime -gt [Datetime]"10-11-2020 9:00:00"} | select FullName
```

Output to a file and make it easier to read.
```Powershell
ls -r  -File | ?{$_.LastWriteTime -gt [Datetime]"10-11-2020 9:00:00"} | select FullName | ft  -A   > "C:\temp\output.txt"
```

 No Aliases Version.
```Powershell
Get-ChildItem -Recurse  -File | Where-Object{$_.LastWriteTime -gt [Datetime]"2020/10/27 18:00:00"} | Select-Object FullName | Format-Table  -AutoSize   > "C:\temp\output.txt"
```






## EXPLANATION

```Powershell
Get-ChildItem -Recurse  -File
```

Recursively get the list of file information.

```Powershell
 $_.LastWriteTime -gt [Datetime]"10-11-2020 9:00:00"
```

Cast A date String to Datetime type value and compare it with the last write time of the file.

