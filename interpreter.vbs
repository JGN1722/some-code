'Copyright (c) JGN1722
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

'//DOCUMENTATION
'LANGUAGE:
'
'this is a stack-oriented interpreted language written in vbscript
'it supports only integers
'it allows the user to manipulate four general purpose registers (a,b,c and d)
'the two other registers are ip, used to contain the current instruction number,
'and sp, used to point at the top of the stack
'the value of each registers can be accessed manually by the programmer
'the interpreter doesn't mind if there's garbage at the end of a line
'uninitialized stack items are empty, you can see it by incrementing sp and printing
'
'comments are signaled by a ' or the word rem
'
'to execute a file with it, pass it's path as an argument to this program
'
'big thanks to @Charles Ma on stackOverflow, who provided an important part of the language specification in the question 'How would I go about writing an interpreter in C? [closed]'
'
'INSTRUCTIONS:
' push <num> : push a number on to the stack
' pop [register] 	: pop off the first number on the stack. The 'register' operand is optional, you may use it if you want to pop the value into a register
' add			: pop off the top 2 items on the stack and push their sum on to the stack
' sub			: pop off the top 2 items on the stack and push their difference on to the stack
' ifeq <label>		: examine the top of the stack; if it's 0, continue, else, jump to the specified label
' jump <label>		: jump to the specified label
' print			: print the value at the top of the stack
' dup			: push a copy of what's at the top of the stack
' mov			: puts a value into a register
' end			: ends the program. The interpreter will automatically add one at the end of the source
'
'
'//EXAMPLES
'here is a valid script that can be executed with this interpreter:
'
'rem this puts the value 10 in the a register, and then pushes and prints it
'mov a 10
'push a
'print
'
'rem this will display 10
'push 10
'print 10
'
'push 0
'ifeq label
'
'rem this code will never be executed
'push 2
'print
'
'label:
'push 100
'print
'
'end

Option Explicit

'//objects
dim shell, fso, labelregex, intregex, labeladdresses
set shell = createobject("wscript.shell")
set fso = createobject("scripting.filesystemobject")
set labelregex = createobject("vbscript.regexp")
set intregex = createobject("vbscript.regexp")
set labeladdresses = createobject("scripting.dictionary")

'//this regex will be used to ensure that only integers are pushed
'//to the stack
intregex.pattern = "^[0-9\-]+$"

'//the regex object will be used to detect if a line is a label
labelregex.pattern = "^[A-Za-z0-9_]+:$"

'//loading the code in memory
dim filepath, file, textstream, source
'//the name of the file to be interpreted is the first command line argument
if WScript.Arguments.Count = 0 then
	msgbox "no script file name was passed as an argument",vbcritical,"error"
	wscript.quit
end if
filepath = WScript.Arguments.Item(0)
if not fso.fileexists(filepath) then
	wscript.quit
end if
set file = fso.getfile(filepath)
if file.size = 0 then
	'//textstream.readall will throw an error if the script
	'//tries to read an empty text file, so quit instead
	wscript.quit
end if
set textstream = file.openastextstream
source = textstream.readall

'//registers and stack
dim ip, sp, a, b, c, d
ip = 0
sp = 0
dim stack(65353)
a = 0
b = 0
c = 0
d = 0


'//remove blank lines, comments and get the addresses of the labels
'//I know, this block seems a bit messy
'//first, we declare some temporary variables
dim i,result,temparr,element,linecount,temptokens

'//then we put each line of the source into a temporary array
temparr = split(source,vbcrlf)

'//let's look at each line
for i=0 to ubound(temparr)
	'//first, let's put the line we're looking at in a variable, so it's
	'//clearer
	element = temparr(i)
	'//we trim the tabs and spaces to allow indentation
	element = trim(replace(element,"	"," "))
	
	'//if the line is blank, ignore it
	if not element = "" then
		'//ignore the comments as well
		if not left(element,1) = "'" and not left(element,3) = "rem" then
			if instr(element,"'") then
				element = left(element,instr(element,"'")-1)
			elseif instr(element,"rem") then
				element = left(element,instr(element,"rem")-1)
			end if
			'//if we're here, then the line is a line of code or a label
			'//add it to the result
			result = result + element + vbcrlf
			'//after having examined each line, we'll put the
			'//content of the result variable in the source
			'//variable
			
			'//increment the linecount variable
			'//it's used to keep track of the current line number,
			'//which will soon be useful because we need to
			'//associate the name of each label to the number
			'//of the line it is on
			linecount = linecount + 1
			
			'//here, we test if the line is a label
			if labelregex.test(element) then
				'//if it is, add a new entry to the dictionary
				'//the name is the key, and it corresponds
				'//to the current line number
				labeladdresses.add left(element,len(element)-1), linecount
			end if
		end if
		
	end if
next


'//add an 'end' instruction at the end of the program because else if the user forget to add it the interpreter
'//will encounter an error
result = result & vbcrlf & "end"

'//finally, let's put our result into the source variable
source = result


'//splitting the code into lines so it can be read line by line
dim code
code = split(source,vbcrlf)

'//main loop
'//it is an infinite loop used only to process line after line
'//it does never end, users need to use the 'end' instruction
'//note that the next instruction number is determined by the
'//return value of ExecuteLine()
dim line
do
	line = code(ip)
	ip = ExecuteLine(line,ip)
loop

function ExecuteLine(line,lineNumber)
	'//first, prepare the line by trimming spaces and tabs
	'//tabs count as spaces
	line = replace(line,"	"," ")
	line = trim(line)
	
	'//then split the words of the line
	dim tokens
	tokens = split(line," ")
	
	'//initialize the return value, it may be modified afterwards
	dim nextinstruction
	nextinstruction = lineNumber + 1
	
	'//if the line is empty, don't execute it and return immediately
	if ubound(tokens) = -1 then
		ExecuteLine = nextinstruction
		exit function
	end if
	
	'//if the line is a label, don't execute it and return immediately
	if labelregex.test(tokens(0)) then
		ExecuteLine = nextinstruction
		exit function
	end if
	
	'//Now, here's the main part of this function
	'//In this language, the first token is the name of the operation
	select case tokens(0)
		
		'//for explanations about this part, check the documentation
		case "push"
			sp = sp + 1
			stack(sp) = Eval(tokens(1))
		
		case "pop"
			if ubound(tokens) > 0 then
				select case tokens(1)
					case "a"
						a = stack(sp)
					case "b"
						b = stack(sp)
					case "c"
						c = stack(sp)
					case "d"
						d = stack(sp)
					case "sp"
						sp = stack(sp)
					case "ip"
						ip = stack(sp)
					case else
						throwerror lineNumber,"invalid register name: "+tokens(1)
				end select
			end if
			stack(sp) = 0
			sp = sp - 1
		
		case "add"
			stack(sp-1) = cstr(cint(stack(sp)) + cint(stack(sp-1)))
			stack(sp) = 0
			sp = sp-1
		
		case "sub"
			stack(sp-1) = cstr(cint(stack(sp)) - cint(stack(sp-1)))
			stack(sp) = 0
			sp = sp-1
		
		case "ifeq"
			if stack(sp) = 0 then
				if labeladdresses.exists(tokens(1)) then
					nextinstruction = labeladdresses(tokens(1))
				else
					'//when we throw an error, we indicate the line to help the user
					'//you can see we wrote lineNumber+1
					'//that's because in the program, we start counting instructions from 0
					'//but the line numbers in a text editor start from 1
					'//so we have to add 1 to make the line number and the
					'//indicated number match
					throwerror lineNumber+1, "the following label does not exist: " & tokens(1)
				end if
			end if
		
		case "jump"
			if labeladdresses.exists(tokens(1)) then
				nextinstruction = labeladdresses(tokens(1))
			else
				throwerror lineNumber+1, "the following label does not exist: " & tokens(1)
			end if
		
		case "print"
			msgbox stack(sp),vbinformation,"the program says:"
		
		case "dup"
			sp = sp + 1
			stack(sp) = stack(sp-1)
		
		case "mov"
			select case tokens(1)
				case "a"
					a = Eval(tokens(2))
				case "b"
					b = Eval(tokens(2))
				case "c"
					c = Eval(tokens(2))
				case "d"
					d = Eval(tokens(2))
				case "sp"
					sp = Eval(tokens(2))
				case "ip"
					ip = Eval(tokens(2))
			end select
		
		case "end"
			wscript.quit
		
		case else
			'//if the token isn't recognized, tell the user that
			'//there is a syntax error
			throwerror lineNumber+1, "syntax error"
	end select
	
	'//finally, return the next line number as the return value of the function
	ExecuteLine = nextinstruction
end function

function Eval(token)
	select case token
		case "a"
			Eval = a
		case "b"
			Eval = b
		case "c"
			Eval = c
		case "d"
			Eval = d
		case "sp"
			Eval = sp
		case "ip"
			Eval = ip
		case else
			if intregex.test(token) then
				Eval = token
			else
				throwerror ip+1,"incorrect data type"
			end if
	end select
end function

function throwerror(linenumber,message)
	msgbox 	"error on instruction: " & linenumber & vbcrlf &_
		 message,vbcritical,"error"
	wscript.quit
end function
