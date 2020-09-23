<div align="center">

## File Associations and the Command Line


</div>

### Description

This tutorial will help you create a custom file extension that opens directly in your program when you open the file. Read on...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Triumph](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/triumph.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/triumph-file-associations-and-the-command-line__1-59217/archive/master.zip)





### Source Code

<P><FONT face=Arial>
 So,
you want a file to open in your program, just
by opening the file, not starting the program first?  Read
on...</FONT></P>
<P><FONT face=Arial>First, you need to come up with the extension(s) that your
program needs to be associated with.  For example if you have a program
that saves a profile, ".prf" might be an appropriate extension.  Just make
sure that your extension is not already used by something else.  To do
this, enter the registry editor ([click] Start -> [click] Run... -> [type]
regedit).  Go to HKEY_CLASSES_ROOT and search for your desired
extension.  If it's not there, it's yours!!!</FONT></P>
<P><FONT face=Arial>OK, now let's get to
some code.  To set the extension, use this handy snippet (compliments of
MSDN Knowledge Base)</FONT><br><br></P>
<P>
<pre>
Option Explicit
 Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
 "RegCreateKeyA" (ByVal hKey As Long, _
 ByVal lpSubKey As String, _
 phkResult As Long) As Long
 Private Declare Function RegSetValue Lib "advapi32.dll" Alias _
 "RegSetValueA" (ByVal hKey As Long, _
 ByVal lpSubKey As String, _
 ByVal dwType As Long, _
 ByVal lpData As String, _
 ByVal cbData As Long) As Long
 ' Return codes from Registration functions.
 Const ERROR_SUCCESS = 0&
 Const ERROR_BADDB = 1&
 Const ERROR_BADKEY = 2&
 Const ERROR_CANTOPEN = 3&
 Const ERROR_CANTREAD = 4&
 Const ERROR_CANTWRITE = 5&
 Const ERROR_OUTOFMEMORY = 6&
 Const ERROR_INVALID_PARAMETER = 7&
 Const ERROR_ACCESS_DENIED = 8&
 Private Const HKEY_CLASSES_ROOT = &H80000000
 Private Const MAX_PATH = 260&
 Private Const REG_SZ = 1
 Private Sub Form_Click()
 Dim sKeyName As String 'Holds Key Name in registry.
 Dim sKeyValue As String 'Holds Key Value in registry.
 Dim ret& 'Holds error status if any from API calls.
 Dim lphKey& 'Holds created key handle from RegCreateKey.
 'This creates a Root entry called "MyApp".
 sKeyName = "MyApp"
 sKeyValue = "My Application"
 ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
 ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
 'This creates a Root entry called .BAR associated with "MyApp".
 sKeyName = ".BAR"
 sKeyValue = "MyApp"
 ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
 ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
 'This sets the command line for "MyApp".
 sKeyName = "MyApp"
 sKeyValue = "c:\mydir\my.exe %1"
 ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
 ret& = RegSetValue&(lphKey&, "shell\open\command", REG_SZ, _
 sKeyValue, MAX_PATH)
 End Sub
</pre>
</P>
<P><FONT face=Arial><br>Then, change 'MyApp' to match your program's name (abbreviated) and 'My Application' to your program's name, and the file path to your program's (make sure to leave the %1). Also, change the extension (here, .BAR) to your own, which can be lowercase or uppercase.
<br><br>This will tell the computer that when a file of this
extension is to be opened, it should start your program.  It does not,
however, tell your program to open the file when it starts up.  To do that,
you need the next bit of code.</FONT></P>
<P><FONT face=Arial>The way your program knows what file to open is the Command
Line.  The command line contains any arguments your program needs when
it starts.  In this case, it contains the path of the file to open. 
The way to access the Command Line is through the keyword "Command".  It
contains the string that is the file's path.  So, in the Form_Load
event...</FONT></P>
<P>
<pre>
If Command <> "" Then
 LoadFile(Command)
End If
</pre>
<P><FONT face=Arial>One thing about the command line: it only works like this
for compiled EXEs.  To simulate this in VB's IDE, go to the Project menu,
click Project Properties, then the Make tab, and put in the file path in the
"Command Line Arguments" text box.  Compile and run.</FONT></P>
<P><FONT face=Arial>Now you should be set!  Start your program, create a
file, save it, exit the program, and double click on the file.  It should
open in your program.  If you have any problems, feel free to email
me.</FONT></P>
<P><FONT face=Arial>Hope this helped.</FONT></P>

