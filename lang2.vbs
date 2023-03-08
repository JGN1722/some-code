'//SETUP SECTION
'still the same language specification, with the 65353 cells stack, 6
'registers and 10 operations, and more to be implemented after I finally
'have a fucking working implementation
option explicit

'//classes
'In memory, the program will consist of an array of those classes, each
'representing an operation, and following the order of execution
'I first need a lexer to convert the text into objoperations and add them
'to the array, and then an interpreter which will read the array from
'bottom to top, jumping between instructions if needed with the help of
'array indexes

Class objtoken
	Public value
	Public tokentype
End Class

class objoperation
	public name
	public arguments
	
	private sub class_initialize()
		arguments = array()
	end sub
	
	public sub addargument(value)
		ReDim preserve arguments(ubound(arguments)+1)
		arguments(ubound(arguments)) = value
	end sub
end Class

'//variable declarations
'//retrieve the program texts
dim fso,position,text,current_char,current_token,token_count
set fso = createobject("scripting.filesystemobject")
token_count = -1

if not fso.getfile("code.txt").size = 0 then
	text = fso.opentextfile("code.txt").ReadAll
	text = LCase(text)
else
	wscript.quit
end if

position = 1
current_char = mid(text,position,1)

'//program variables
Dim stack(65353)
Dim a,b,c,d,sp,ip

'//error
Function Error(text)
	MsgBox text,vbCritical,"error"
	WScript.Quit
End Function

'//type functions
Dim r1
Set r1 = CreateObject("vbscript.regexp")
r1.Pattern = "^[0-9\-][0-9]*$"
Function isdigit(text)
	isdigit = r1.Test(text)
End Function

Dim r2
Set r2 = CreateObject("vbscript.regexp")
r2.Pattern = "^[a-zA-Z_][a-zA-Z0-9_]*$"
Function isalnum(text)
	isalnum = r2.Test(text)
End Function

'//label dictionary
Dim label_addresses
Set label_addresses = CreateObject("scripting.dictionary")

'//operation list
dim operation_list
operation_list = "|push|pop|add|sub|ifeq|jump|print|dup|mov|end|"












'//LEXER SECTION
'//this function returns the next token in the text to the parser
function get_next_token
	dim objreturn
	do
		if isnull(current_char) then
			set get_next_token = new objtoken
			get_next_token.tokentype = "EOF"
			get_next_token.value = "EOF"
			exit function
			
		elseif current_char =  " " or current_char = chr(9) then
			skip_whitespace
			
		'//that case checks for a newline, which is 2 characters
		elseif current_char = chr(13) and peek = chr(10) then
			skip_whitespace
			
		elseif current_char = "'" then
			skip_comments
			
		elseif isalnum(current_char) then
			set objreturn = get_whole_word
			if objreturn.value = "REM" then
				skip_comments
			else
				set get_next_token = objreturn
				exit function
			end if
			'//same as for the digit case, here we don't call
			'//advance
			
		elseif isdigit(current_char) then
			set get_next_token = get_whole_number
			'//in that case the pointer is already incremented
			'//by get_whole_number, so there is no need to
			'//call advance here
			exit function
			
		elseif current_char = ";" then
			set get_next_token = new objtoken
			get_next_token.tokentype = "SEMI"
			get_next_token.value = ";"
			advance
			exit function
			
		elseif current_char = "," then
			set get_next_token = new objtoken
			get_next_token.tokentype = "COMMA"
			get_next_token.value = ","
			advance
			exit function
			
		else
			error "unexpected symbol: "&current_char
		end if
	loop
end function

Function peek
	If position + 1 > Len(text) Then
		peek = Null
	Else
		peek = Mid(text,position+1,1)
	End If
End Function

'//this function advances by one the pointer in the text
'//the pointer needs to be incremented after the next token is returned
'//it alse sets the current_char variable
function advance
	position = position + 1
	if position > len(text) then
		current_char = null
	else
		current_char = mid(text,position,1)
	end if
end function

function get_whole_number
	dim result
	result = ""
	do while isdigit(current_char)
		result = result & current_char
		advance
		if isnull(current_char) then
			break
		end if
	loop
	set get_whole_number = new objtoken
	get_whole_number.value = result
	get_whole_number.tokentype = "INTEGER"
end function

function get_whole_word
	dim result
	result = ""
	do while isalnum(current_char)
		result = result & current_char
		advance
		if isnull(current_char) then
			exit do
		end if
	loop
	if current_char = ":" then
		set get_whole_word = new objtoken
		'//we have a special case here with labels
		'//we need to set their value as their addresses
		token_count = token_count + 1
		label_addresses(result) = token_count
		get_whole_word.value = result
		get_whole_word.tokentype = "LABEL"
		advance
		
	elseif result = "a" or result = "b" or result = "c" or result = "d" or result = "ip" or result = "sp" then
		set get_whole_word = new objtoken
		get_whole_word.value = result
		get_whole_word.tokentype = "REGISTER"
		
	elseif result = "rem" then
		set get_whole_word = new objtoken
		get_whole_word.value = "REM"
		get_whole_word.tokentype = "COMMENT"
		
	elseif instr(operation_list,"|" & result & "|") then
		set get_whole_word = new objtoken
		get_whole_word.value = result
		get_whole_word.tokentype = "OPERATION"
		token_count = token_count + 1
		
	else
		Set get_whole_word = new objtoken
		get_whole_word.value = result
		get_whole_word.tokentype = "LABEL_REFERENCE"
	end if
end function

function skip_whitespace
	do while current_char = " " or current_char = chr(13) or current_char = chr(10) or current_char = chr(9)
		advance
		if isnull(current_char) then
			exit do
		end if
	loop
end function

function skip_comments
	do until current_char = chr(13) and peek = chr(10)
		advance
		if isnull(current_char) then
			exit do
		end if
	loop
	advance
	advance
end function












'//PARSER SECTION
dim program_array,l
program_array = array()
'l acts as a counter so I can avoid overcomplicated syntax
l = 0

Sub eat(datatype)
	If current_token.tokentype = datatype Then
		Set current_token = get_next_token
		'MsgBox "eaten "&datatype&" and got "&current_token.tokentype
	Else
		Error "wrong token type:"&current_token.tokentype
	End If
End Sub

'//first, I'll get a token, then eat following the type, and add the values to the argument array
Set current_token = get_next_token

function tokenize_program
	do until isnull(current_char)
		'What I want to do here is to retrieve operation after
		'operation, and put them in an objoperation with the
		'appropriated properties

		'MsgBox "new tokenization iteration: "&current_token.tokentype
		ReDim preserve program_array(l)
		If Not current_token.tokentype = "LABEL" Then
			Select Case current_token.value
				Case "push"
					'MsgBox "push operation"
					eat "OPERATION"
					Set program_array(l) = New objoperation
					program_array(l).name = "push"
					program_array(l).addargument current_token.value
					If current_token.tokentype = "REGISTER" Then
						eat "REGISTER"
					Else
						eat "INTEGER"
					End If
					eat "SEMI"
					'MsgBox "got here"
					
				Case "pop"
					eat "OPERATION"
					Set program_array(l) = New objoperation
					program_array(l).name = "pop"
					If current_token.tokentype = "REGISTER" Then
						program_array(l).addargument current_token.value
						eat "REGISTER"
					End If
					eat "SEMI"
					
				Case "add"
					eat "OPERATION"
					Set program_array(l) = New objoperation
					program_array(l).name = "add"
					eat "SEMI"
					
				Case "sub"
					eat "OPERATION"
					Set program_array(l) = New objoperation
					program_array(l).name = "sub"
					eat "SEMI"
					
				Case "dup"
					eat "OPERATION"
					Set program_array(l) = New objoperation
					program_array(l).name = "dup"
					eat "SEMI"
					
				Case "ifeq"
					eat "OPERATION"
					Set program_array(l) = New objoperation
					program_array(l).name = "ifeq"
					program_array(l).addargument current_token.value
					eat "LABEL_REFERENCE"
					eat "SEMI"
					
				Case "jump"
					eat "OPERATION"
					Set program_array(l) = New objoperation
					program_array(l).name = "jump"
					program_array(l).addargument current_token.value
					eat "LABEL_REFERENCE"
					eat "SEMI"
					
				Case "print"
					eat "OPERATION"
					Set program_array(l) = New objoperation
					program_array(l).name = "print"
					eat "SEMI"
					
				Case "mov"
					eat "OPERATION"
					Set program_array(l) = New objoperation
					program_array(l).name = "mov"
					program_array(l).addargument(current_token.value)
					eat "REGISTER"
					eat "COMMA"
					If current_token.tokentype = "REGISTER" Then
						program_array(l).addargument(current_token.value)
						eat "REGISTER"
					Else
						program_array(l).addargument(current_token.value)
						eat "INTEGER"
					End If
					eat "SEMI"
					
				Case "end"
					eat "OPERATION"
					Set program_array(l) = New objoperation
					program_array(l).name = "end"
					eat "SEMI"
					
				Case Else
					Error "wrong token:"&current_token.value
					
			End Select
		Else
			Set program_array(l) = New objoperation
			program_array(l).name = "label"
			program_array(l).addargument(label_addresses.Item(current_token.value))
			eat "LABEL"
		End If
		l = l + 1
	loop
end function







'//smol interpreter
Function interpret
	tokenize_program
	
	Do
		if ip > ubound(program_array) then
			'//the program execution is finished
			wscript.quit
		end if
		
		Select Case program_array(ip).name
			Case "push"
				sp = sp + 1
				Select Case program_array(ip).arguments(0)
					Case "a"
						stack(sp) = a
					Case "b"
						stack(sp) = b
					Case "c"
						stack(sp) = c
					Case "d"
						stack(sp) = d
					Case "sp"
						stack(sp) = sp
					Case "ip"
						stack(sp) = ip
					Case Else
						stack(sp) = program_array(ip).arguments(0)
				End Select
				
				ip = ip + 1
			Case "pop"
				If UBound(program_array(ip).arguments) = -1 Then
					stack(sp) = Empty
					sp = sp - 1
				Else
					Select Case program_array(ip).arguments(0)
						Case "a"
							a = stack(sp)
						Case "b"
							b = stack(sp)
						Case "c"
							c = stack(sp)
						Case "d"
							d = stack(sp)
						Case "sp"
							sp = stack(sp)
						Case "ip"
							ip = stack(sp)
					End Select
				End If
				
				ip = ip + 1
			Case "add"
				stack(sp - 1) = CStr(CLng(stack(sp)) + CLng(stack(sp - 1)))
				stack(sp) = Empty
				sp = sp - 1
				
				ip = ip + 1
			Case "sub"
				stack(sp - 1) = stack(sp) - stack(sp - 1)
				stack(sp) = Empty
				sp = sp - 1
				
				ip = ip + 1
			Case "ifeq"
				If stack(sp) = 0 Then
					ip = label_addresses.Item(program_array(ip).arguments(0))
				Else
					ip = ip + 1
				End If
				
			Case "jump"
				ip = label_addresses.Item(program_array(ip).arguments(0))
				
			Case "print"
				MsgBox stack(sp),vbInformation,"the program says:"
				
				ip = ip + 1
			Case "dup"
				sp = sp + 1
				stack(sp) = stack(sp - 1)
				
				ip = ip + 1
			Case "mov"
				Select Case program_array(ip).arguments(0)
					Case "a"
						a = evaluate(program_array(ip).arguments(1))
					Case "b"
						b = evaluate(program_array(ip).arguments(1))
					Case "c"
						c = evaluate(program_array(ip).arguments(1))
					Case "d"
						d = evaluate(program_array(ip).arguments(1))
					Case "sp"
						sp = evaluate(program_array(ip).arguments(1))
					Case "ip"
						ip = evaluate(program_array(ip).arguments(1))
				End Select
				
				ip = ip + 1
			Case "end"
				WScript.Quit
			Case "label"
				'do nothing, it's only a label
				ip = ip + 1
			Case Else
				Error "invalid operation: "&program_array(ip).name
		End Select
	Loop
End Function

Function evaluate(expr)
	Select Case expr
		Case a
			evaluate = a
		Case b
			evaluate = b
		Case c
			evaluate = c
		Case d
			evaluate = d
		Case sp
			evaluate = sp
		Case ip
			evaluate = ip
		Case Else
			evaluate = expr
	End Select
End Function

interpret