# Crash Course on VBScript

_The original document is on my [Github](https://github.com/filovirid/tutorials/blob/main/vbscript.md) for editing_

Lately, I was looking for a syntax of one function in Windows Script Host (WSH) and I realized that it's difficult these days to find a nice tutorial
about WSH or even VBScript since both of them are kind of dead! (VBS replaced by JavaScript and WSH replaced by Powershell). There are some websites
offering some tutorials but you need to watch lots of nasty advertise to be able to get something from the tutorial.

If you need to write some code in these languages or even more importantly, you want to 
read some code that is based on these scripting languages, here is a nice all-in-one tutorial covering both VBS 
and WSH (next post) mostly useful for reading codes and maybe writing some script (like automation scripts).

Let's start with different aspects of VBS first.


## Introduction to VBScript

VBScript or let's say Microsoft Visual Basic Scripting Edition is a scripting language that can be used both as a client-side language in HTML pages and server-side language on Windows operating system. However, these days the only browser which support VBS is Microsoft Internet Explorer (IE) and others like Firefox and Chrome only support JavaScript. So you can guess that no one is using VBS as a client-side language anymore.

However, regarding the server-side part, Windows OS still supports VBS and it's actually very useful to use it in writing some scripts for administrative stuff (removing files, scheduling, etc) and This is the bad part since it can also be used by attackers (e.g., malware authors) to use to this language as part of their attack vector to infect user's computer.

If you are already familiar with Visual Basic for application, this means that you already know VBS but if that's not the case, continue reading this tutorial.

### VBS variables (defining and using)

Like all other languages, we first need to know how to define variables in VBS. In VBS, we only have **one** variable type called **Variant**.
A Variant can contain different types of information (e.g., numbers, strings, booleans, etc). The VBS engine will convert the variant to the appropriate type when it's necessary based on the context. Here is the list of all sub-types for a variant.


| **Name**         |  **Value**         | **variable type** | 
|------------- | ------------- | -------------- |
| Empty  | Uninitialized (`0` for numbers and `""` for strings)  | 0
|  Null   | No valid data  | 1
|  Boolean | either `True` or `False`  | 11
|  Byte | from 0 to 255   | 17
|  Integer | from -32,768 to 32,767  | 2
|  Currency | -922,337,203,685,477.5808 to 922,337,203,685,477.5807  |6
|  Long | -2,147,483,648 to 2,147,483,647  | 3
| Single  |  -3.402823E38 to -1.401298E-45 (negative) & 1.401298E-45 to 3.402823E38 (positive) | 4
|  Double | -1.79769313486232E308 to -4.94065645841247E-324 (negative) & 4.94065645841247E-324 to 1.79769313486232E308 (positive)  | 5
| Date (Time)  | a date between January 1, 100 to December 31, 9999  | 7
| String  | variable-length string up to approximately 2 billion characters  | 8
| Object  | Contains an object  | 9
| Error  | Contains an error number  | 10

In the above table, the **variable type** is a number assigned to each variable type so that you can use it in your code if you need it (you never need it!). Let's take a look at a simple code which shows some variable types:

```basic
a = 12
b = 12.8
c = "Hi"
d = Null
msgbox VarType(a)
msgbox VarType(b)
msgbox VarType(c)
msgbox VarType(d)
```
Open up a file and change the extension to '.vbs', copy & paste the above code in to the file and double click on the file to execute it. You will see four message boxes showing the numbers 2, 5, 8 and, 1 respectively. `VarType` function returns the type of the variable and in this case: `a` is an *Integer*, `b` is a *Double*, `c` is a *String* and `d` is *Null* (`msgbox` is just a function to show a pop-up window).

**VBS is a case-insensitive languaage**. This means that all the keywords, functions and variables can be capital, small or the combination of both! All four lines in the following code generate the very same result, showing "Hi".

```basic
MsgBox "Hi"
msgbox "Hi"
MSGBOX "Hi"
MSGbox "Hi"
```

Let's see an example for variables:

```
Myname = "John"
myName = "Jack"
msgbox Myname
msgbox MYNAME
msgbox myName
```

All the three `msgbox` functions show "Jack" since `Myname`, `myName` and `MYNAME` are the same.

#### Declaring Variables

In order to declare a variable in VBS, you can just use the variable wherever you want without declaring it (as we did in the above examples). However, it's better to explicitly declare variables. To declare variables, you can use one of the three possible keywords:

1. **Dim**: Declares a variable
2. **Private**: Declares a private variable in a Class
3. **Public**: Declares a public variable in a Class

We don't care about **public** and **private** variables since we are not dealing with classes right now, but we use **dim** most of the time to declare variables.

#### Variable scopes

A variable scope is determined by where you declare it. If you declare it inside a function, then the variable exists only inside that function. If you declare it at script level, then it exists in all parts of the script.

#### Variable naming convention
1. Maximum of 255 characters
2. Must start with alphabet
3. Can not contain period
4. Can contain underline and letters and numbers

#### Assigning values to variables 

Just like all other languages, just use '=' (equal) sign to assign a value to a variable

Here is an example of declaring variables and assigning values:

```basic
Dim name, lastname, age
name = "John"
lastname = "Doe"
age = 100
msgbox "My name is " + name + " " + lastname + " and I am " + cstr(age)
```

Well, we have a new function in the example: `cstr`. This is one of the conversion function to convert different types to each other. Here is the list of conversion functions:

| **Name** | **Description** |
| -------- | --------------- |
| asc(string) | Returns the ANSI code for the first character of the string  |
| cbool(exp) |  Returns the boolean value from an expression |
| cbyte(exp) |  Returns the byte value of the expression |
| ccur(exp) |  Returns the currency type from the expression |
| cdate(exp) | Returns the date value from the expression  |
| cdbl(exp) |  Returns the double value from the expression |
| chr(code) |  Returns the character of the specified ANSI code |
| cint(exp) |  Returns the Integer value from the expression |
| clng(exp) |  Returns the long value from the expression |
| csng(exp) |  Returns the single value from the expression |
| cstr(exp) |  Converts expression exp to string |
| hex(number) |Returns the string representation of the hexadecimal value of the number  |
| oct(number) | Returns a string representing the octal value of a number |


These functions are very useful and used a lot in malware scripts. Let's see some examples:

```basic
rem A comment can start with 'rem' keyword
' Also we can have a comment with single quote
msgbox hex(65)		'41
msgbox hex(1223545)  '12AB79
msgbox cstr(12)		'12
msgbox cdate("2021-12-03 06:23:34 AM")   '03/12/2021 06:23:34
msgbox cbool(1)		'True
msgbox cbool(0)		'False
```

- We just learned two more things! A comment in VBS can begin with **`rem`** keyword or with a single quote (').

What if we are not sure if some functions can be converted to a specific data type or not? Well, there are some checking functions to do the job for us:

| **Name** | **Description** |
| -------- | --------------- |
| isdate(exp) | Returns True or False indicating whether the expression can be converted to date type. |
| isnumeric(exp) | Returns True or False indicating whether the expression can be evaluated to date type. |
| isempty(exp) | Returns True or False indicating whether the expression has been initialized |
| isnull(exp) | Returns True or False indicating whether the expression contains no valid data |
| isobject(exp) | Returns True or False indicating whether the expression is a valid reference to an object |
| isarray(variable) | Returns True or False indicating whether the variable is an array. |

**`IsArray()`** function is used to determine if a variable is an array or not, but we haven't told yet how to specify an array. Let's talk about arrays in VBS.

### Declaring and using arrays

Declaring an array variable is as easy as scalar variables. You just need to add parentheses and specify the length of the array (if it's a fixed-size array). To assign values to array cells, just use the array index and equal sign. Array index starts from 0 and include the number in parentheses.

```basic
dim myvar(10)
myvar(0) = "Hello"
myvar(1) = "World!"
myvar(3) = " I am John"
myvar(4) = " and I am "
myvar(5) = Cstr(26)
msgbox myvar(0) + myvar(1) + myvar(2) + myvar(3) + myvar(4) + myvar(5) 
```

In the above example, `myvar` array has 11 elements starts from 0 and end with the last one which is `myvar(10)`.

Wait...This is a fixed-size array. What if we want to declare a dynamic array? Well, it's easy:

1. First you declare an array without specifying the size
2. then you specify whatever size you want using the `redim` keyword.
3. If you want to increase the size again, you just use `redim preserve` and then the name of the array with the new size. Remember that using `preserve` is important to make sure the array keeps the previous values.

```basic
dim a()
redim a(1)
a(0) = "Hello "
a(1) = "World"
redim preserve a(2)
a(2) = "!"
msgbox a(0) + a(1) + a(2)
```

### Pre-defined and Custom Constants in VBScript

There is a list of constant keywords in VBS that are useful to know but not necessary since you can declare your constant whenever you want. However, here is the list of useful constants.

| **Name** | **Description** |   **Name** | **Description** |
| -------- | --------------- | -------- | --------------- |
| vbTab  |  Tab character |  vbVerticalTab | 0x0B |
| vbNullChar | 0x00 | vbNewLine | 0x0A or 0x0D0A |
| vbLf | 0x0A | vbFormFeed | 0x0C |
| VbCrLf | 0x0D0A | vbCr | 0x0D |
| vbOKOnly | 0 | vbOKCancel | 1 |
| vbAbortRetryIgnore | 2 | vbYesNoCancel | 3 |
| vbYesNo | 4 | vbRetryCancel | 5 |
| vbCritical | 16 | vbQuestion | 32 |
| vbExclamation | 48 | vbInformation | 64 |
| vbDefaultButton1 | 0 | vbApplicationModal | 0 |
| vbSystemModal | 4096 | vbOK | 1 |
| vbCancel | 2 | vbAbort | 3 |
| vbRetry | 4 | vbIgnore | 5 |
| vbYes | 6 | vbNo | 7 |
| vbTrue | -1 | vbFalse | 0 |


Let's see some examples for the constant values:

```basic
msgbox "Hello" + vbtab + "world!"
msgbox "Hello", vbokcancel +  vbExclamation , "Title"
msgbox "Hello", vbokcancel +  vbExclamation , ""
msgbox "Hello", vbokcancel +  vbCritical , "Title"
msgbox "Hello", vbAbortRetryIgnore +  vbApplicationModal , "Title"
msgbox vbtrue   'shows -1
msgbox vbtrue + vbFalse  'shows -1
```

And you can see from the example that **`vbTrue` constant is evaluated -1 and `vbFalse` constant is evaluated to 0**.

How to specify custom constants? Using **const** keywords:
```basic
const myname = "John"
const myage = 16
msgbox "My name is " + myname + " and I am " + cstr(myage)
```

Try to change the constants in the example to see the error.

### VBScript Operators

Here is the list of operators in VBScript:

| **Operator** | **Description** | **Operator** | **Description** |
| ------------ | --------------- | ------------ | --------------- |
| ^ | Exponentiation  |  -  |  Unary negation |
| * | Multiplication | / | Division |
| \ | Integer division | mod | Modulus arithmetic |
| + | Addition | - | Subtraction  |
| & | String concatenation|  |  |

| **Operator** | **Description** | **Operator** | **Description** |
| ------------ | --------------- | ------------ | --------------- |
| = | Equality | <> | Inequality |
| < | Less than | > | Greater than |
| <= | Less than or equal to | >= | Greater than or equal to |
| is | Object equivalence |  |  |


| **Operator** | **Description** | **Operator** | **Description** |
| ------------ | --------------- | ------------ | --------------- |
| not |  Logical negation | and | Logical conjunction |
| or | Logical disjunction | xor | Logical exclusion |
| eqv | Logical equivalence | imp | Logical implication |

Here is the definition of [Logical implication](https://calcworkshop.com/logic/logical-implication/) and here is the definition of [Logical equivalence](https://en.wikipedia.org/wiki/Logical_equivalence) in case you don't know.

Let's see some examples with outputs:

```basic
dim a
a = 12 ^ 3      ' 1728
a = 12 * 3      ' 36
a = 12 / 5      ' 2.4
a = 12 \ 5      ' 2
a = 12 + 5      ' 17
a = 12 mod 5    ' 2
a = "12" & "5"  ' "125"
a = -12         ' -12
```

**I am not covering operator precedence since it's boring! and as other programming/scripting languages, you can always use parentheses to override the the order of precedence**


### Conditional statements (IF...ELSE)

IF..ELSE statement in VBS is very simple and intuitive. There are two types of if..else: one-line statement and multiline statement. If you have only one statement, then use one-line if..else, otherwise, use multiline if..else.

#### One-line IF...ELSE

IF expression THEN statement ELSE statement

```basic
dim a
a = inputbox("Please enter a number")
a = cint(a)
if a < 10 then msgbox "a < 10" else msgbox "a > 10"
```

* The _ELSE_ part is arbitrary and can be removed.

We also learned about another function called `inputbox`. This function gets an input from the user.

#### Multi-line IF...ELSE

Here is the other forms of IF...ELSE statement.

```basic
IF condition THEN
    statement 1
    statement 2
ELSE
    statement 3
    statement 4
END IF
```
Or

```basic
IF condition1 THEN
    statement 1
    statement 2
ELSEif condition2 THEN
    statement 3
    statement 4
END IF
```

* There is no curly bracket like C/C++ or colon like Python.

### Conditional statements (Select...case)

Select...Case structure in VBS is similar to IF...ELSE but it's more efficient in a way that it only evaluates the condition one time. Here is the structure and example:

```basic
SELECT CASE condition
    CASE firstCase
        statement
        statement
    CASE secondCase
        statement
        statement
    CASE ELSE
        statement
        statement
END SELECT
```

```basic
dim a 
a = inputbox("Your favorite fruit?")
select case a
	case "apple"
		msgbox "I also like apple!"
	case "banana"
		msgbox "I don't like banana"
	case else
		msgbox "I don't know shit about " & a
end select
```

### Loops in VBScript

There are a few different forms of loops in VBScript. The first and the easiest one is **DO...LOOP** with the following syntax:


#### DO...LOOP statement

```basic
DO [<WHILE | UNTIL> condition]
    statement1
    statement2
    [EXIT DO]
    ....
LOOP
```

Everything inside the brackes ([]) are optional. This means that you are not forced to use **WHILE** or **UNTIL** or **EXIT DO**.

1. *DO WHILE condition* means that repeat the loop while the condition is True.
2. *DO UNTIL condition* means that repeat the loop until the condition become True.

**Note**: *EXIT DO* can only (and only) be used with DO...LOOP to exit the loop.

Here is a simple example: If you enter 20, I will print a string contains 20 A character.

```basic
dim a
a = inputbox("Enter a number between 1 and 100")
a = cint(a)
if a < 1 or a > 100 then
	msgbox cstr(a) & " is not between 1 and 100"
	wscript.quit	'consider it as a way to exit the code
end if
dim b, c
b = ""
c = 1
do while c < a
	b = b & "A"
	c =  c + 1
loop
msgbox b
```

#### FOR...NEXT statement

The syntax is as follows:

```basic
FOR counter = start TO end [STEP step]
    statement
    statement
    [EXIT FOR]
    statement
NEXT
```

Let's see the above example with FOR...LOOP statement.

```basic
dim a
a = inputbox("Enter a number between 1 and 100")
a = cint(a)
if a < 1 or a > 100 then
	msgbox cstr(a) & " is not between 1 and 100"
	wscript.quit	'consider it as a way to exit the code
end if
dim b, c
b = ""
c = 1
for c = 1 to a step 1
	b = b & "A"
next
msgbox b
```

#### FOR EACH...NEXT

Simply, whatever that is iterable, you can use *FOR EACH...NEXT* on it with the following syntax:

```basic
FOR EACH item IN iterable
    statements
    [Exit For]
    statements
Next
```

And let's see the same example a new style:

```basic
dim a(2), b
a(0) = "A"
a(1) = "B"
a(2) = "C"
b = ""
for each i in a
	b = b & i
next
msgbox b    ' it will show "ABC"
```

#### WHILE...WEND statement

WHILE...WEND is a simple version of DO...LOOP and just checks the condition and repeats the statement while the condition is True. Here is the syntax:

```basic
WHILE condition
    statement
    statement
    .
    .
    statement
WEND
````

And here is the same example using WHILE...WEND:

```basic
dim a
a = inputbox("Enter a number between 1 and 100")
a = cint(a)
if a < 1 or a > 100 then
	msgbox cstr(a) & " is not between 1 and 100"
	wscript.quit	'consider it as a way to exit the code
end if
dim b, c
b = ""
c = 1
while c < a
	b = b & "A"
	c =  c + 1
wend
msgbox b
```

### Procedures and Functions

While in almost all programming languages, there is only one way to define a routine (usually called function or subroutine), in VBScript, 
there are two ways to do that:

1. Define a function
2. Define a procedure

The major difference between function and procedure is that functions return values while
procedures only perform some actions without returning any value.

#### Defining a function in VBS

VBS uses the following syntax to define a function:

```basic
FUCNTION func_name (argument1, argument2, argumentn)
    statement 1
    statement 2
    func_name = expression
    [EXIT FUNCTION]
    func_name = expression
    [EXIT FUNCTION]
END FUNCTION
```

The funny part is that if you want to return a value from a function, you need to assign 
that value to the function name. Whenever you want to return from a function, you can 
use the term `EXIT FUNCTION` and just exit the function. Let's take a look at an example:

```basic
REM write a function to multiply/add two numbers
REM and return the value. On error, shows a msgbox
function operator(m1, m2, op)
	if op = "+" then
		operator = cdbl(m1) + cdbl(m2)
		exit function
	elseif op = "*" then
		operator = cdbl(m1) * cdbl(m2)
		exit function
	else 
		msgbox cstr(op) & " is not defined!"
	end if
end function

msgbox cstr(operator(2,3,"+"))    'shows 5
msgbox cstr(operator(2,3,"*"))	  'shows 6
```

There are other concepts related to functions like defining a public/private functions,
or passing arguments by value or by reference which we are not going to explore
(remember that this is a no deep shit tutorial!).

#### Defining a sub procedure in VBS

Here is the definition of the sub procedures in VBS:

```basic
SUB sub_name(arg1, arg2, arg3)
   statement 1
   [EXIT SUB]
   statement 2 
   statement 3
End Sub
```
**NOTES**:

1. Sub routines don't return any value. They can just perform some actions.
2. As soon as you call `EXIT SUB`, the program will exit the subroutine.

Her is an example of a subroutine:

```basic
sub print_error(msg)
	'we just print an error here
	msgbox msg
end sub

a = inputbox("Please enter a number")
if not isnumeric(a) then
	print_error("You must enter a valid number!")
	wscript.quit	'consider it as a way to exit the code
end if
a = cint(a)
msgbox "I am "& cstr(a) & " old"
```

**NOTE**: In order to call a subroutine (only subroutines not functions), you have three
options:

1. Using the name of the subroutine and putting the arguments not using parentheses.
2. Using the name of the subroutine and putting the arguments inside parentheses.
3. Using the keyword **CALL** and you MUST put the arguments inside parentheses.

Here is the example of three possible ways to call the example routine:

```basic
' Using call keyword and parentheses
call print_error("You must enter a valid number!")
' Without using the call keyword but with parentheses
print_error("You must enter a valid number!")
' Without using the call keyword and without parentheses
print_error "You must enter a valid number!"
```


### Other useful statements in VBS

We cover almost all basics of VBS. There are a few useful statements that fit nowhere
so we have to put it here.

#### ON ERROR Statement

Like other programming languages, VBScript also has a mechanism to handle
(run-time) errors. Either you can choose to stop the execution if an error 
occurred or you can continue the execution as some errors are not that critical.

##### Statement: On Error Resume Next

In order to tell VBS engine to continue the execution even if something is wrong, 
You can use the term `ON ERROR RESUME NEXT` either at the beginning of the script
or at the beginning of each routine/function. The example below shows the situation 
where the code execution will break because of the first `msgbox` (too many arguments provided).
However, using the error control, we can prevent that and make sure that the 
second msgbox will execute.

```basic
on error resume next
msgbox "This line raise an error", 2,2,2,2,2,2  ' this line should break the code
msgbox "If you see this, it means we handled the error"
```

##### Statement: On Error Goto 0

Using this statement, we can reset the error control to the default (throwing
error if something is wrong).

```basic
on error resume next
on error goto 0
msgbox "This line raise an error", 2,2,2,2,2,2
msgbox "If you see this, it means we handled the error"
```

#### OPTION EXPLICIT Statement

As you see, in VBS you can do almost whatever you want. A variable can be declared
first and then used but it can also be used without declaration. You can use uppercase
 and lowercase or a combination of both when you are using variables. This is somehow
 confusing and may introduce some errors in your code. To solve this problem, you can 
 use `OPTION EXPLICIT` statement at the beginning of your code. This makes sure that
 all the variables are declared and then used. The example below will work perfectly
 fine without using `option explicit`. However, utilizing this term, you will get
 an error (variable is undefined) for this code.

 ```basic
option explicit
name = "John"
msgbox name
 ```

#### Statement: SET

Assigns an **object** reference to a variable. We use the word "object" since it can 
be an instance of a class (defined by user) or any of the language built-in objects.
This is different from using "dim" in declaring variables. You should first, declare
the variable using "DIM" keyword and then assign the reference of the object to our 
variable. Check the FileSystemObject section for an example.


### Miscellaneous 

#### Several commands in one line

While in C language, you use semicolon (; ) to separate commands from each other,
in VBScript you must use colon ( : ) to separate commands in one line.

```basic
dim a:a=12:msgbox a
```

### Some useful functions in VBS

Here is a list of some useful functions that you may see from time to time in scripts. 
We write a one-line description for each function and a big example to cover their usage
at the end of this section.


|Function|Description|
|--------|-----------|
|abs()   |Returns the absolute value of a number| 
|array() |Returns a variant containing an array| 
|asc()   |Returns the ANSI character code corresponding to the first letter in a string|
|chr()   |Returns the character associated with the specified ANSI character code| 
|eval()  |Evaluates an expression and returns the result| 
|cos()   |Returns the cosine of an angle. | 
|filter()|Returns a zero-based array containing a subset of a string array based on a specified filter criteria| 
|exp()   |Returns e (the base of natural logarithms) raised to a power|
|hex()   |Returns a string representing the hexadecimal value of a number. |
|instrrev()   |Returns the position of an occurrence of one string within another, from the end of string| 
|instr()  |Returns the position of the first occurrence of one string within another(starts from 1)| 
|join()   |Returns a string created by joining a number of substrings contained in an array| 
|len()    |Returns the number of characters in a string or the number of bytes required to store a variable| 
|log()    |Returns the natural logarithm of a number. | 
|ltrim()  |Returns a copy of a string without leading spaces| 
|rtrim()  |Returns a copy of a string without trailing spaces| 
|trim()   |Returns a copy of a string without leading and trailing spaces| 
|mid()    |Returns a specified number of characters from a string| 
|replace()   |Returns a string in which a specified substring has been replaced with another substring a specified number of times| 
|oct()   |Returns a string representing the octal value of a number |
|left()  | Returns a specified number of characters from the left side of a string|
|lcase() | Returns a string that has been converted to lowercase|
|ucase() | Returns a string that has been converted to uppercase|
|now()   | Returns the current date and time according to the setting of your computer's system date and time|
|space() |Returns a string consisting of the specified number of spaces |
|split() |Returns a zero-based, one-dimensional array containing a specified number of substrings |
|strcomp() |Compares two strings and returns the result of the comparison |
|time()        |Returns a Variant of subtype Date indicating the current system time |
|DateDiff()  |Returns the difference between two dates|
|Lbound()  |Returns the smallest available subscript for the indicated dimension of an array |
|Ubound()  |Returns the largest available subscript for the indicated dimension of an array |

and Here is the example of all the functions:

```basic
dim a,b,c
a = -2
b = 4
abs(a)     ' return 2
abs(b)     ' return 4
c = Array("J", "o", "h", "n")
' print the word "John"
msgbox c(0) & c(1) & c(2) & c(3) 
msgbox asc("a")  ' print 97
' print "abc"
msgbox chr(97) & chr(98) & chr(99)
msgbox eval("5 + 2") ' eval() returns 7
a = array("This", "is", "a", "test","string")
' select all those have "is" 
b = filter(a, "is", true)
msgbox "length of b: " & cstr(ubound(b)+1)
msgbox b(0) & " " & b(1)        ' This is
' select all those don't have "is"
b = filter(a, "is", false)
msgbox b(0) & " " & b(1) & " " & b(2)  ' a test string
msgbox cstr(exp(2)) ' 7.38905609893065
msgbox(hex(10)) ' prints "A"
a = "This is a test string"
msgbox instr(a, "test")     ' prints 11
msgbox instr(a, "shit")     ' prints 0
msgbox instrrev(a, "i")     ' prints 19
a = Array("J", "o", "h", "n")
msgbox join(a, "")      ' prints "John"
msgbox join(a, "-")     ' prints "J-o-h-n"
b= "This is a test string"
msgbox len(b)       ' prints 21
b= "    string with spaces    "
msgbox "'" & ltrim(b) & "'"   ' prints "string with spaces    "
msgbox "'" & rtrim(b) & "'"   ' prints "   string with spaces"
msgbox "'" & trim(b) & "'"    ' prints "string with spaces"
b = "This is a test string"
msgbox mid(b, 10, 5)
a = "This is a test string"
msgbox "'" & mid(a, 11, 4) & "'"      ' prints 'test'
a = "what is this? this is a test"
' from beginning of the string, search for "is"
' and replace it with "aa" two times.
b = replace(a, "is", "aa", 1, 2)
msgbox b        ' what aa thaa? this is a test
a = "This is a test"
msgbox left(a, 4)   ' prints "This"
msgbox right(a, 4)  ' prints "test"
a = "THIS IS A TEST"
msgbox lcase(a)         ' prints "this is a test"
msgbox ucase(lcase(a))  ' prints "THIS IS A TEST"
msgbox now()            ' pirnts "24/01/2021 17:37:30"
msgbox "'" & space(5) & "'"     ' prints "     "
a = "aAbAcA"
b = split(a, "A")
msgbox join(b)      ' prints "a b c"
a = "This is a test"
b = "this is a test"
c = "THis is a test"
msgbox strcomp(a,c)     ' prints 1
msgbox strcomp(a,b)     ' prints -1
msgbox strcomp(a,a)     ' prints 0
msgbox time()           ' prints 17:37:30
' calculates epoch (from first of Jan 1970) - Unix timestamp
msgbox datediff("s", "1970/01/01 00:00:00", now())
```








