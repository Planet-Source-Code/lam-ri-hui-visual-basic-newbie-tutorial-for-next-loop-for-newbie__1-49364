<div align="center">

## Visual Basic Newbie Tutorial : For\.\.\.Next Loop \(For Newbie\)


</div>

### Description

Hi VB newbie, do you know how to use for...next loop? If you don't, never mind. Read this tutorial to learn about that. You will learn how to use for...next loop by explanation and examples by me. Leave comments about this tutorial. Happy coding!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lam Ri Hui](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lam-ri-hui.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lam-ri-hui-visual-basic-newbie-tutorial-for-next-loop-for-newbie__1-49364/archive/master.zip)





### Source Code

<p><u><b><font face="Comic Sans MS" size="6" color="#000080">Visual Basic Newbie
Tutorial 1 : For...Next Loop </font></b></u></p>
<p><font face="Comic Sans MS"><b><font size="4">Loop</font></b><br>
A loop is used to run a block of statements over and over. It is simply a group
of commands that is repeated a specified number of times or for a specified
length of time. An example of looping action : imagine you are sticking mail
labels to several invitation cards. For each card, you check that the label
matches the name on it, then you tick it off from your guest list, before
pasting the mail label onto the card. You then proceed to the next card, repeat
the check, tick and paste actions, and so on. This is essentially how a loop
works.</font></p>
<p><font face="Comic Sans MS">There are two types of loops :<br>
<b>·</b> Counter Loop - where repetition is based on a specified number of times<br>
<b>·</b> Conditional Loop - where repetition is based on set conditions</font></p>
<p><font face="Comic Sans MS"><b><font size="4">Counter Loops</font></b><br>
You use a counter loop when you want the computer to perform a task for a
specific number of times. This is similar to running a track race of, say 12
laps. While running, you count the number of laps, and when you have completed
12, you stop.</font></p>
<p><font face="Comic Sans MS">A counter loop is also known as a <i><b>For loop</b></i>
or a <b><i>For/Next loop</i></b>. This is because the ends of the loop are
defined by the </font><font face="Courier New">For</font><font face="Comic Sans MS">
statement and the </font><font face="Courier New">Next</font><font face="Comic Sans MS">
statement. A <i><b>For/Next</b></i> loop requires two statements: the </font>
<font face="Courier New">For</font><font face="Comic Sans MS"> statement at the
beginning of the loop and the </font><font face="Courier New">Next</font><font face="Comic Sans MS">
statement at the end of loop.</font></p>
<p><font face="Comic Sans MS">At the beginning of a </font>
<font face="Courier New">For</font><font face="Comic Sans MS"> loop, you define
a counter variable as well as the start and end values for the variable. For
example, if you want the loop to repeat 12 times, you would set</font></p>
<p><font face="Comic Sans MS">&nbsp;&nbsp;&nbsp; </font>
<font face="Courier New">For X = 1 To 12</font></p>
<p><font face="Comic Sans MS">The syntax of the loop is as follows</font></p>
<blockquote>
 <p><font face="Courier New">For<i> countervariable </i>=<i> start</i> To <i>
 end</i><br>
&nbsp;&nbsp; <i>&nbsp;Statements to be executed</i><br>
 Next <i>countervariable</i></font></p>
</blockquote>
<p><font face="Comic Sans MS">The first time the loop is run, the counter
variable is set to the value of the starting point (usually 1). After the
statements are executed once, the counter variable reaches the Next statement
where a counter registers a count of 1. The counter variable increase by one for
each loop. Each time, the value in the counter is checked against the value of
the end point. It stops when this value is reached.</font></p>
<p><i><font face="Comic Sans MS" size="4">Example</font></i></p>
<p><font face="Comic Sans MS">The following program will cause the computer to
send out five beeps, one after another from the computer's speaker.</font></p>
<p><font face="Courier New">&nbsp;&nbsp;&nbsp; For b = 1 To 5<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Beep<br>
&nbsp;&nbsp;&nbsp; Next b</font></p>
<p><font face="Comic Sans MS">The variable used is b. It stands for the first
number in a For loop. Each time the loop is executed, the counter variable b
increases by 1, until it reaches a value of 5.</font></p>
<p><font face="Comic Sans MS">Thus, the above code is same as following :</font></p>
<p><font face="Courier New">&nbsp;&nbsp;&nbsp; Beep<br>
&nbsp;&nbsp;&nbsp; Beep<br>
&nbsp;&nbsp;&nbsp; Beep<br>
&nbsp;&nbsp;&nbsp; Beep<br>
&nbsp;&nbsp;&nbsp; Beep</font></p>
<p><font face="Comic Sans MS"><b><font size="4">Step Size</font></b><br>
The default change in the loop counter is one. You can specify a different value
for the change. This value is known as the <b><i>step size</i></b>. Referring to
our example of track, if the track is 1000 meters long, a race of 10,000 meters
will require 10 laps. On a track of 500 meters, it would require 20 laps. The
step size changes the number of required laps or runs.</font></p>
<p><font face="Comic Sans MS">To change the step size, include step in the For
loop. You can use any number, including decimals and negative numbers for the
step size.</font></p>
<p><i><font face="Comic Sans MS" size="4">Example</font></i></p>
<p><font face="Courier New">Dim Count</font></p>
<p><font face="Courier New">&nbsp;&nbsp;&nbsp; For Count = 0 To 100 Step 10<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Print Count<br>
&nbsp;&nbsp;&nbsp; Next Count</font></p>
<p><font face="Comic Sans MS">The output will as follow :</font></p>
<p><font face="Comic Sans MS">&nbsp;&nbsp;&nbsp; 0<br>
&nbsp;&nbsp;&nbsp; 10<br>
&nbsp;&nbsp;&nbsp; 20<br>
&nbsp;&nbsp;&nbsp; 30<br>
&nbsp;&nbsp;&nbsp; 40<br>
&nbsp;&nbsp;&nbsp; 50<br>
&nbsp;&nbsp;&nbsp; 60<br>
&nbsp;&nbsp;&nbsp; 70<br>
&nbsp;&nbsp;&nbsp; 80<br>
&nbsp;&nbsp;&nbsp; 90<br>
&nbsp;&nbsp;&nbsp; 100</font></p>
<p><font face="Comic Sans MS"><b><font size="4">Nested Loops</font></b><br>
You can nest two or more For loops inside one another. Whenever you program
needs to repeat a loop more than once, use a nested loop.</font></p>
<p><font face="Comic Sans MS">The following example shows two nested loops:</font></p>
<p><font face="Courier New">&nbsp;&nbsp;&nbsp; For x = 1 To 3
&lt;------------------------------<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
|<br>
&nbsp;&nbsp;&nbsp; For y = 1 To 5 &lt;---&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
|<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Print &quot;#&quot;;&nbsp;&nbsp;&nbsp; |&nbsp;&nbsp;
Inner loop&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
|&nbsp;&nbsp;&nbsp; Outer loop<br>
&nbsp;&nbsp;&nbsp; Next y ------------&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
|<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
|<br>
&nbsp;&nbsp;&nbsp; Print&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
|<br>
&nbsp;&nbsp;&nbsp; Next x ---------------------------------------</font></p>
<p><font face="Comic Sans MS">The inner loop is performed first. After it is
completed, then the outer loop carries on.</font></p>
<p><font face="Comic Sans MS"><b><font size="4">Inner Loop</font></b><br>
In the above example, the inner loop starts with the variable </font>
<font face="Courier New">y = 1</font><font face="Comic Sans MS">. It prints #.
The semicolon will cause the next statement to print on the same line. Hence the
second time the loop is carried out&nbsp; another # will be printed next to the
first, resulting in ##. After the inner loop is repeated 5 times, the result
will be #####.</font></p>
<p><font face="Comic Sans MS"><b><font size="4">Outer Loop</font></b><br>
The outer loop begins with the variable </font><font face="Courier New">x = 1</font><font face="Comic Sans MS">.
This will cause the inner loop to occur once, resulting in #####. The second
time the outer loop occurs, another ##### will be printed, and so on. Thus after
3 repetitions of the outer loop the result will be</font></p>
<p><font face="Comic Sans MS">&nbsp;&nbsp;&nbsp; #####<br>
&nbsp;&nbsp;&nbsp; #####<br>
&nbsp;&nbsp;&nbsp; #####</font></p>
<p><font face="Comic Sans MS">The Print statement by itself starts a new line
for each row of #####, then include another Print Statement.</font></p>
<p><font face="Courier New">&nbsp;&nbsp;&nbsp; Dim x As Integer, y As Integer</font></p>
<p><font face="Courier New">&nbsp;&nbsp;&nbsp; For x = 1 To 3</font></p>
<p><font face="Courier New">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; For y = 1
To 5<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Print &quot;#&quot;;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Next y</font></p>
<p><font face="Courier New">&nbsp;&nbsp;&nbsp; Print<br>
&nbsp;&nbsp;&nbsp; Print<br>
&nbsp;&nbsp;&nbsp; Next x</font></p>
<p><font face="Comic Sans MS">Nested loops need not be hard to use. Just
remember the rule : The inner loop must be completed before the Next statement
for the outer loop is encountered. Use indentation and blank lines between loops
to make them easy to read and debug.</font></p>
<p align="center"><b><font face="Comic Sans MS" size="6">
----------------------------------<br>
Visual Basic Newbie Tutorial 1 : For...Next Loop Ends Here<br>
----------------------------------</font></b></p>

