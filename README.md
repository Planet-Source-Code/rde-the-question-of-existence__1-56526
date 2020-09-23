<div align="center">

## The Question of Existence

<img src="PIC2004104124151557.gif">
</div>

### Description

I have come across several different solutions for testing for a file's existance, from Dir to opening the file, testing for error, then closing the file again. I remembered an article/comment by Bruce McKinney that influenced my solution to this problem and thought I would share it.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rde](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rde.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rde-the-question-of-existence__1-56526/archive/master.zip)





### Source Code


<font color="#000066">
<h2 align="center">The Question of Existence - by Bruce McKinney</h2>
<p>Testing for the existence of a file ought to be easy (and is in most
languages), but it turns out to be one of the most annoying problems
in Visual Basic. Don't count on simple solutions like this:</p>
<code><nobr>fExist = (Dir$(sFullPath) &lt;&gt; vbNullString)</nobr></code>
<p>Dir will return the first file found if you happen to pass sFullPath as
an empty string, and so will set fExist to True. You could use:</p>
<code><nobr>If sFullPath &lt;&gt; vbNullString Then fExist = (Dir$(sFullPath) &lt;&gt; vbNullString)</nobr></code>
<p>That statement works until you specify a file on an empty floppy or
on a CD-ROM drive. Then you're stuck in a message box.</p>
<p>Here's another common one:</p>
<code><nobr>fExist = FileLen(sFullPath)</nobr></code>
<p>It fails on 0-length files — uncommon but certainly not unheard of.</p>
<p>My theory is that the only reliable way to check for file existence
in Basic (without benefit of API calls) is to use error trapping.</p>
<p>I've challenged many Visual Basic programmers to give me an
alternative, but so far no joy. Here's the shortest way I know:</p>
<code><nobr>Function FileExists(sSpec As String) As Boolean</nobr><br />
 &#160; On Error Resume Next<br />
 &#160; Call FileLen(sSpec)<br />
 &#160; FileExists = (Err = 0)<br />
End Function</code>
<p>This can't be very efficient. Error trapping is designed to be fast
for the no fail case, but this function is as likely to hit errors
as not.</p>
<p>Perhaps you'll be the one to send me a Basic-only ExistFile function
with no error trapping that I can't break.</p>
<p>Until then, here's an API alternative:</p>
<code><nobr>Function ExistFileDir(sSpec As String) As Boolean</nobr><br />
 &#160; Dim af As Long<br />
 &#160; af = GetFileAttributes(sSpec)<br />
 &#160; ExistFileDir = (af &lt;&gt; -1)<br />
End Function</code>
<p>I didn't think there would be any way to break this one, but it turns
out that certain filenames containing control characters are legal on
Windows 95 but illegal on Windows NT. Or is it the other way around?</p>
<p>Anyway, I have seen this function fail in situations too obscure to
describe here.</p>
<p>Bruce McKinney</p>
<p></p>
<p>Please note that the VB6 File System Object's GetAttr function cannot be used in place of the GetFileAttributesA API function in this technique as it raises an error when the path is invalid.</p>
</font><font color="#660066">
<p><code><nobr>Private Declare Function GetFileAttributes Lib "kernel32" _<br />
 &#160; &#160; Alias "GetFileAttributesA" (ByVal lpSpec As String) As Long</nobr></p>
<p><nobr>Private Const INVALID_FILE_ATTRIBUTES As Long = -1</nobr></p>
<p><nobr>Function FileExists(sFileSpec As String) As Boolean<br />
 &#160; Dim Attribs As Long<br />
 &#160; Attribs = GetFileAttributes(sFileSpec)<br />
 &#160; If (Attribs &lt;&gt; INVALID_FILE_ATTRIBUTES) Then<br />
 &#160; &#160;  FileExists = ((Attribs And vbDirectory) &lt;&gt; vbDirectory)<br />
 &#160; End If<br />
End Function</nobr></p>
<p><nobr>Function DirExists(sPath As String) As Boolean<br />
 &#160; Dim Attribs As Long<br />
 &#160; Attribs = GetFileAttributes(sPath)<br />
 &#160; If (Attribs &lt;&gt; INVALID_FILE_ATTRIBUTES) Then<br />
 &#160; &#160;  DirExists = ((Attribs And vbDirectory) = vbDirectory)<br />
 &#160; End If<br />
End Function</nobr></code></p></font>

