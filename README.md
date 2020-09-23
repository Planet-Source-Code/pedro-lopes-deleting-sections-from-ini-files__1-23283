<div align="center">

## Deleting Sections from \.INI Files


</div>

### Description

In this article, I explain how you can delete a specific entry from an .INI file.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Pedro Lopes](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pedro-lopes.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pedro-lopes-deleting-sections-from-ini-files__1-23283/archive/master.zip)





### Source Code

<p align=center style="font:15pt verdana;color:#F90000">Deleting Sections from .INI Files
<p style="font:10pt verdana">In this article, I explain how you can delete a specific entry from an .INI file.
<p style="font:10pt verdana">An initialization (.INI) file is an ASCII text file that follows a specific format. The file is divided into sections where the name of the section is enclosed in brackets. Directly below the section headings are one or more entries. Each entry (or key name) is the name you want to set a value for. This is followed by an equal sign. Next, the value to be assigned to the key name is specified.
<br>To modify an .INI file, you use the Windows WritePrivateProfileString() and WriteProfileString() functions. The WriteProfileString() function is used to modify the Windows WIN.INI initialization file, while all other .INI files are modified by calling the WritePrivateProfileString() function.
<br>The following is an example of an .INI file's contents:
<p style="font:10pt verdana">[progsetup]
<br>Date=10/10/95
<br>Datafile=c:\temp.dat
<p style="font:10pt verdana">In this example, the section name is "progsetup", the key names are Date and Datafile, and the values to be given to the key names are 10/10/95 and c:\temp.dat.
<br>To delete a specific entry from an initialization file, call the WritePrivateProfileString() function with the statement:
<p style="font:10pt verdana">x = WritePrivateProfileString(lpAppName, 0&, 0&, FileName)
<p style="font:10pt verdana">specifying the following parameters:
<p style="font:10pt verdana">lpAppName \The name of the section you want to remove from the INI file
<br>lpKeyName \The entry you want to delete. This must be set to a NULL string
   \to delete the entire section.
<br>lpString \The string to be written to the entry. When set to an empty string,
   \this causes the lpKeyName entry to be deleted.
<br>lpFileName \The name of the INI file to modify.
<p style="font:10pt verdana">In our example above, we would set lpAppName to "progsetup", lpFileName to "C:\DEMO.INI", and both lpKeyName and lpString to 0& (zero). After you call this function, the entire "progsetup" section of the DEMO.INI file will be deleted.
<p style="font:10pt verdana">The lpKeyName and lpString variables are of type Any. If you use the type String, the function may or may not work properly, so be sure to specify these as type Any when deleting entries from initialization files. The same rule applies when using the WriteProfileString() function.
<p style="font:12pt verdana"><font color="#F90000"><b>Example:</b></font>
<p style="font:10pt verdana">This is how you can aply this:
<p style="font:10pt verdana">1.Using the Windows Notepad application, create a new text file called DEMO.INI. Save the file to the root directory on drive C. Add the following lines to this text file:
<br>[progsetup]
<br>Date=10/10/95
<br>Datafile=c:\temp.dat
<br>[colors]
<br>Background=red
<br>Foreground=white
<p style="font:10pt verdana">2.Start a new project in Visual Basic. Form1 is created by default.
<p style="font:10pt verdana">3.Create a module, and type the following Declare statement (note that this should be typed as a single line of text):
<br>Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
<p style="font:10pt verdana">4.Add the following code to Form1_Load():
<br>Sub Form_Load()
<br> crlf$ = Chr(13) & Chr(10)
<br> Text1.Text = ""
<br> Open "c:\demo.ini" For Input As #1
<br> While Not EOF(1)
<br>  Line Input #1, file_data$
<br>  Text1.Text = Text1.Text & file_data$ & crlf$
<br> Wend
<br> Close #1
<br>End Sub
<p style="font:10pt verdana">5.Add a text box control to Form1. Set its MultiLine property to True and its ScrollBars property to 3-Both. Adjust the size of the text box so that the contents of the C:\DEMO.INI file can be displayed in it.
<p style="font:10pt verdana">6.Add a command button control to Form1. Command1 is created by default. Set its Caption property to "Modify DEMO.INI".
<p style="font:10pt verdana">7.Add the following code to the Click event of Command1:
<br>Sub Command1_Click()
<br> FileName = "c:\demo.ini"
<br> lpAppName = "progsetup"
<br> x = WritePrivateProfileString(lpAppName, 0&, 0&, FileName)
<br>End Sub
<p style="font:10pt verdana">When you execute this sample program, the current contents of the file C::\DEMO.INI are displayed in the text box. Click once on the "Modify DEMO.INI" command button. The program has now deleted the entire "progsetup" section from the DEMO.INI file. You can verify that the file's contents were changed by running the demonstration program a second time.
<p><font face="verdana">Come and see my site for more of my stuff at <a href="http://www.atomsoftware.cjb.net"><font color="#F90000">www.AtomSoftware.Cjb.Net</a></font>.

