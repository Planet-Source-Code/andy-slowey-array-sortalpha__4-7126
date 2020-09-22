<div align="center">

## Array SortAlpha


</div>

### Description

Sorts arrays alphabetically, ascending or descending.
 
### More Info
 
Array and Sort order

Okay so this isnt that complicated but it is a rather old bit of code I dug up while clearing out an old hard drive. I do have a versionfloating about someplace that does multidimensional arrays which will follow when I find it!

Array


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andy Slowey](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andy-slowey.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Sorting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/sorting__4-24.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andy-slowey-array-sortalpha__4-7126/archive/master.zip)

### API Declarations

Copyright... Dont be stupid!


### Source Code

```
<%
Function SortAlpha(ary, direction)
'###############################################################
'# USAGE: Call The function as any other, specify ASC or DESC #
'#    to change the ordering direction. This is meant to  #
'#    be similar to SQL style syntax            #
'###############################################################
  StopWork=False
  Do Until StopWork=True
    StopWork=True
    For i = 0 to UBound(ary)
      If i=UBound(ary) then Exit For
      If UCase(Direction) = "DESC" Then
        If ary(i) < ary(i+1) Then
          firstval = ary(i)
          secondval = ary(i+1)
          ary(i) = secondval
          ary(i+1) = firstval
          StopWork=False
        End If
      Else
        If ary(i) > ary(i+1) Then
          firstval = ary(i)
          secondval = ary(i+1)
          ary(i) = secondval
          ary(i+1) = firstval
          StopWork=False
        End If
      End If
    Next
  Loop
  SortAlpha=ary
End Function
Function Sort(ary)
  KeepChecking = TRUE
  Do Until KeepChecking = FALSE
    KeepChecking = FALSE
    For I = 0 to UBound(ary)
      If I = UBound(ary) Then Exit For
      If ary(I) > ary(I+1) Then
        FirstValue = ary(I)
        SecondValue = ary(I+1)
        ary(I) = SecondValue
        ary(I+1) = FirstValue
        KeepChecking = TRUE
      End If
    Next
  Loop
  Sort = ary
End Function
Function PrintArray(ary)
  For i=0 to UBound(ary)
    Response.Write(ary(i) & "<BR>" & vbCrLf)
  Next
End Function
Dim MyArray
MyArray = Array(1,5,"shawn","says","hello",123,12,98)
PrintArray(MyArray)
Response.Write("<HR><h1>Sorted Alpha ASC</h1><br><br>")
PrintArray(SortAlpha(MyArray, "ASC"))
Response.Write("<HR><h1>Sorted Alpha DESC</h1><br><br>")
PrintArray(SortAlpha(MyArray, "DESC"))
Response.Write "<HR>"
For I = 0 to Ubound(Sort(MyArray))
  Response.Write MyArray(I) & "<br>" & vbCRLF
Next
%>
```

