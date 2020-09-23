<div align="center">

## Image Directory Listing


</div>

### Description

Lists thumbnail images within the current directory. Very simple beginner stuff dealing with file scripting object. Handy for viewing your images quickly. Just copy the code and paste it a file called default.asp. Then view the directory on your website.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Graphics/ Sound](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics-sound__4-15.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-image-directory-listing__4-7522/archive/master.zip)





### Source Code

```
<H1>Image Directory Listing</H1>
<%
Dim FSO
Dim Files
Dim File
Dim Count
Const Columns = 3
Const ImageWidth = 100
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
Set Files = FSO.GetFolder(Server.MapPath("./")).Files
Set FSO = Nothing
Response.Write "<TABLE width=""100%"" border=""1"" cellspacing=""0"">"
Response.Write "<TR>"
Count = 0
For Each File In Files
	Select Case LCase(Right(File.Name, 3))
		Case "jpg", "gif", "bmp", "png"
			Count = Count + 1
			If Count Mod Columns = 1 Then Response.Write "</TR><TR>"
			Response.Write "<TD align=""center"" valign=""top"">"
			Response.Write "<A href=""" & File.Name & """>"
			Response.Write File.Name
			Response.Write "<BR><IMG src=""" & File.Name & """ border=""1"" width=""" & ImageWidth & """><BR>"
			Response.Write "</A>"
			Response.Write "</TD>"
	End Select
Next
Response.Write "</TR>"
Response.Write "</TABLE>"
Set File = Nothing
Set Files = Nothing
%>
```

