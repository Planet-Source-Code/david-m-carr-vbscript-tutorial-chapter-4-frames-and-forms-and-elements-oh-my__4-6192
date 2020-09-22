<div align="center">

## VBScript Tutorial: Chapter 4\-\-Frames and Forms and Elements, Oh My\!


</div>

### Description

One of the most common uses of scripting is to respond to events from a form, such as typing text into a text box or checking data when the Submit button is pressed. There are two ways to reference form elements...this chapter tells you how.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David M\. Carr](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-m-carr.md)
**Level**          |Beginner
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VbScript \(browser/client side\)

**Category**       |[Controls/ Forms/ Dialogs/ Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/controls-forms-dialogs-menus__4-3.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-m-carr-vbscript-tutorial-chapter-4-frames-and-forms-and-elements-oh-my__4-6192/archive/master.zip)





### Source Code

```
<h4><font face="Verdana">Frames and Forms and Elements, Oh My!</font></h4>
<p><font face="Verdana">One of the most common uses of scripting is to respond to events from a form, such as
typing text into a text box or checking data when the Submit button is pressed. There are
two ways to reference form elements. First, you can reference them through the Object
Model. Document.Form(0).Elements(1).Value refers to the Value property of the second
element in the first form on the page. You can determine the number of forms in the page
using Document.Forms.Count, and can determine the number of elements in the first form
using Document.Forms(0).Elements.Count.</font></p>
<p><font face="Verdana">This is an awkward way of referring to elements. Instead, if you include a Name
attribute in the form and element's HTML tags, you can refer to it by that.</font></p>
<p><font face="Verdana">&lt;FORM NAME=&quot;aForm&quot;&gt;<br>
&lt;INPUT TYPE=&quot;TEXT&quot; NAME=&quot;txtGreet&quot;&gt;<br>
&lt;/FORM&gt;<br>
This element's Value could be referenced using aForm.txtGreet.Value</font></p>
<p><font face="Verdana">Unless you are write exclusively for IE4+, you should always put form elements in a
form, and use the NAME attribute for them rather than the ID attribute.</font></p>
<p><font face="Verdana">Form elements expose both properties and events. The following example shows this.</font></p>
<div class="vbscode">
<p><code><font face="Verdana">Sub txtFred_OnFocus()<br>
&nbsp;&nbsp;&nbsp; With aForm.txtGreet<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; If .Value = &quot;Hello&quot; Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; .Value =
&quot;Goodbye&quot;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Else<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; .Value =
&quot;Hello&quot;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End If<br>
&nbsp;&nbsp;&nbsp; End With<br>
End Sub</font></code></p>
</div>
<p><font face="Verdana">The onFocus event fires whenever the element becomes active for input, such as by
clicking it or tabbing to it. With is a keyword which allows the use of relative
references. Normally, it would be necessary to use aForm.txtGreet.Value all three times,
but it is much simpler to use with. Another way to do this would be to use the Set
keyword. It makes a variable into a shortcut to another object. Ex:</font></p>
<div class="vbscode">
<p><code><font face="Verdana">Sub txtFred_OnFocus()<br>
&nbsp;&nbsp;&nbsp; Dim theCtl<br>
&nbsp;&nbsp;&nbsp; Set theCtl = aForm.txtGreet<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; If theCtl.Value = &quot;Hello&quot; Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; theCtl.Value =
&quot;Goodbye&quot;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Else<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; theCtl.Value =
&quot;Hello&quot;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End If<br>
&nbsp;&nbsp;&nbsp; End With<br>
End Sub</font></code></p>
</div>
<p><font face="Verdana">Here is an example of using VBScript to check user form input.</font></p>
<hr>
<div class="vbscode">
<p><code><font face="Verdana">&lt;SCRIPT TYPE=&quot;text/vbscript&quot; LANGUAGE=&quot;VBScript&quot;&gt;<br>
&lt;!--<br>
Function lycosForm_OnSubmit()<br>
&nbsp;&nbsp;&nbsp; If Len(lycosForm.query.Value) = 0 Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; lycosForm_OnSubmit = False<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Alert &quot;You must enter a keyword.&quot;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Exit Function<br>
&nbsp;&nbsp;&nbsp; End If<br>
&nbsp;&nbsp;&nbsp; If lycosForm.cat.Value = &quot;null&quot; Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; lycosForm_OnSubmit = False<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Alert &quot;You must choose a search
type.&quot;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Exit Function<br>
&nbsp;&nbsp;&nbsp; End If<br>
End Function<br>
'--&gt;<br>
&lt;/SCRIPT&gt;</font></code></p>
</div>
<p><code><font face="Verdana">&lt;form action=&quot;http://www.lycos.com/cgi-bin/pursuit&quot; method=GET
NAME=&quot;lycosForm&quot;&gt;<br>
Search <br>
&lt;SELECT NAME=&quot;cat&quot;&gt;<br>
&lt;OPTION value=&quot;null&quot;&gt;Please Choose a Search Type&lt;/OPTION&gt;<br>
&lt;OPTION value=&quot;lycos&quot;&gt;The Web&lt;/OPTION&gt;<br>
&lt;OPTION value=&quot;sounds&quot;&gt;Sounds&lt;/OPTION&gt;<br>
&lt;OPTION value=&quot;graphics&quot;&gt;Pictures&lt;/OPTION&gt;<br>
&lt;OPTION value=&quot;point&quot;&gt;TOP 5%&lt;/OPTION&gt;<br>
&lt;/SELECT&gt; for:&lt;INPUT TYPE=&quot;text&quot; NAME=&quot;query&quot;
VALUE=&quot;&quot; SIZE=22&gt;<br>
&lt;input type=&quot;submit&quot; value=&quot;Go Get It&quot;&gt;&lt;/form&gt;</font></code></p>
<hr>
<p><font face="Verdana">And here is the result. A form which searches Lycos, but won't submit unless you've
chosen a search type, and typed in a keyword.</font></p>
<p><font face="Verdana"><script TYPE="text/vbscript" LANGUAGE="VBScript">
<!--
Function lycosForm_OnSubmit()
  If Len(lycosForm.query.Value) = 0 Then
    lycosForm_OnSubmit = False
    Alert "You must enter a keyword."
    Exit Function
  End If
  If lycosForm.cat.Value = "null" Then
    lycosForm_OnSubmit = False
    Alert "You must choose a search type."
    Exit Function
  End If
End Function
'-->
</script> </font> </p>
<form action="http://www.lycos.com/cgi-bin/pursuit" method="GET" NAME="lycosForm">
 <p><font face="Verdana">Search <select NAME="cat" size="1">
  <option value="null">Please Choose a Search Type</option>
  <option value="lycos">The Web</option>
  <option value="sounds">Sounds</option>
  <option value="graphics">Pictures</option>
  <option value="point">TOP 5%</option>
 </select> for:<input TYPE="text" NAME="query" VALUE SIZE="22"> <input type="submit" value="Go Get It"></font></p>
</form>
<p><font face="Verdana">Note: This is a variation of the <a href="http://www.lycos.com/linktolycos.html#searchform">search form</a> that Lycos makes
available for use on other people's pages.</font> </p>
<hr>
<p><font face="Verdana">Notice that here, I treated the event as a Function rather than a Sub. Is you want to
be able to cancel an event, that is how you do it, by making it a function and then
returning either True or False, False canceling the event.</font> </p>
<p><font face="Verdana">Window.Frames is like Document.Forms, in that it can contain multiple instances that
can be referred to by name or index, and has a Count property.&nbsp; The difference is
that while Forms contain Elements, Frames contain a Document object.</font> </p>
<p><font face="Verdana">Frames are much like windows. The contain a document, and can contain other frames. In
fact, in almost all ways, you can treat a frame as if it was a window.</font></p>
<p><font face="Verdana">You can refer from one frame to another by using the Window.Parent object. It points at
the window above it.</font></p>
<p><font face="Verdana">Frames are supported in NN2+ and IE3+, and are part of the HTML4 specification. For
scripting, though, floating frames are even more useful than regular frames. The IFRAME
tag allows you to make a frame which sits in the middle of a page, like an image. It is
supported by IE3+, and is not part of HTML4. They are referred to in the same way as
regular frames.</font></p>
<p><font face="Verdana">Below is an example of a scripted floating frame.</font></p>
<p align="center"><font face="Verdana"><!--webbot bot="HTMLMarkup" startspan --><IFRAME NAME="fraHello" HEIGHT="70" WIDTH="100"></IFRAME><!--webbot bot="HTMLMarkup" endspan --> <script type="text/vbscript" language="VBScript">
<!--
Dim SPre
Dim EPre
SPre1 = "<BODY BGCOLOR='#000080' TEXT='#00FF00'><PRE>"
SPre2 = "<BODY BGCOLOR='#000080' TEXT='#C0C0C0'><PRE>"
SPre3 = "<BODY BGCOLOR='#000080' TEXT='#008080'><PRE>"
SPre4 = "<BODY BGCOLOR='#000080' TEXT='#FFFF00'><PRE>"
SPre5 = "<BODY BGCOLOR='#000080' TEXT='#FF0000'><PRE>"
EPre = "<" & Chr(47) & "PRE><" & Chr(47) & "BODY>"
Sub AnimHelloFrame1
	Window.fraHello.Document.Write SPre1 & "H  " & EPre
	Window.fraHello.Document.Close
	SetTimeOut "AnimHelloFrame2 True", 250, "VBScript"
End Sub
Sub AnimHelloFrame2(forw)
	Window.fraHello.Document.Write SPre2 & "He  " & EPre
	Window.fraHello.Document.Close
	If forw Then
		SetTimeOut "AnimHelloFrame3 True", 250, "VBScript"
	Else
		SetTimeOut "AnimHelloFrame1", 250, "VBScript"
	End If
End Sub
Sub AnimHelloFrame3(forw)
	Window.fraHello.Document.Write SPre3 & "Hel " & EPre
	Window.fraHello.Document.Close
	If forw Then
		SetTimeOut "AnimHelloFrame4 True", 250, "VBScript"
	Else
		SetTimeOut "AnimHelloFrame2 False", 250, "VBScript"
	End If
End Sub
Sub AnimHelloFrame4(forw)
	Window.fraHello.Document.Write SPre4 & "Hell " & EPre
	Window.fraHello.Document.Close
	If forw Then
		SetTimeOut "AnimHelloFrame5", 250, "VBScript"
	Else
		SetTimeOut "AnimHelloFrame3 False", 250, "VBScript"
	End If
End Sub
Sub AnimHelloFrame5
	Window.fraHello.Document.Write SPre5 & "Hello" & EPre
	Window.fraHello.Document.Close
	SetTimeOut "AnimHelloFrame4 False", 250, "VBScript"
End Sub
Sub Window_OnLoad()
	AnimHelloFrame1
End Sub
-->
</script> </font> </p>
<h5 align="center"><font face="Verdana"><!--webbot bot="Navigation" S-Type="arrows" S-Orientation="horizontal" S-Rendering="text" B-Include-Home="FALSE" B-Include-Up="TRUE" U-Page S-Target --></font></h5>
<h6 align="left">&nbsp;</h6>
<SCRIPT SRC="http://library.thinkquest.org/tq-admin/tqtrailer.js"> </SCRIPT>
```

