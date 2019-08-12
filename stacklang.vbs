'stacklang
option explicit

class stack
	dim tos
	dim stack()
	dim stacksize
	
	private sub class_initialize
		stacksize = 1000
		redim stack( stacksize )
		tos = 0
	end sub

	public sub push( x )
		stack(tos) = x
		tos = tos + 1
	end sub
	
	public property get stackempty
		stackempty = ( tos = 0 )
	end property
	
	public property get stackfull
		stackfull = ( tos > stacksize )
	end property
	
	public property get stackroom
		stackroom = stacksize - tos
	end property
	
	public property get stackcount
		stackcount = tos + 1
	end property
	
	public function pop()
		if tos > 0 then
			pop = stack( tos - 1 )
			tos = tos - 1
		else
			wscript.echo "Error: 'pop' but not enough data on stack"
			'~ wscript.quit
		end if
	end function

	public sub resizestack( n )
		redim preserve stack( n )
		stacksize = n
		if tos > stacksize then
			tos = stacksize
		end if
	end sub
	
	public sub rotate
		dim last, i
		dim base
		base = tos - 1
		last = stack( base )
		for i = base to 1 step -1
			stack( i ) = stack( i - 1 )
		next
		stack( 0 ) = last
	end sub
	
	public sub show
		dim i
		wscript.stdout.write "["
		for i = 0 to tos - 1
			wscript.stdout.write stack( i )
			if i < tos - 1 then
				wscript.stdout.write ", "
			end if
		next
		wscript.echo "]"
	end sub 
end class

class machine
	private ip
	private script
	private finished
	private macros
	private subros
	
	sub class_initialize
		set macros = createobject("scripting.dictionary")
		set subros = createobject("scripting.dictionary")
	end sub
	
	public sub macro( key, data )
		if macros.exists( key ) then
			macros(key) = data
		else
			macros.add key, data
		end if
	end sub
	
	public function code( key )
		if macros.exists( key ) then
			code = macros(key)
		else
			code = vbnullstring
		end if
	end function
	
	public sub setscript( s )
		script = s
		ip = 1
		finished = false
	end sub
	
	public property get nextop
		wscript.stdout.write "(IP=" & ip & ") "
		if ip = 0 then 
			finished = true
			ip = len( script ) + 1
		end if
		nextop = mid( script, ip, 1 )
		ip = ip + 1
		if ip > len( script ) then
			finished = true
		end if
	end property
	
	public sub prevop
		ip = ip - 1
		if ip = 0 then ip = 1
	end sub
	
	public property get isfinished
		isfinished = finished
	end property
	
	public sub firstop
		ip = 1
	end sub
	
	public function evaluate( CS )
		dim c, tmp, macro, m2, tmp2
		dim macsub
		dim hit, ques
		do while not isfinished
			c = nextop
			'~ if isfinished then exit do
			wscript.echo left( script, ip - 2 ) & " (" & c & ") " & mid( script, ip  )
			select case c
			case "("
				do while true
					c = nextop
					if c = ")" then exit do
				loop
			case "+"
				apply CS, c, 2
			case "-"
				apply CS, c, 2
			'~ case "*"
				'~ apply CS, c, 2
			'~ case "/"
				'~ apply CS, c, 2
			'~ case "^"
				'~ apply CS, c, 2
			case ">"
				apply CS, c, 2
			case "<"
				apply CS, c, 2
			case "="
				apply CS, c, 2
			case "#"
				apply CS, "<>", 2
			case "&"
				apply CS, "and", 2
			case ":" 'dup
				if CS.stackcount > 0 then
					tmp = CS.pop
					CS.push tmp
					CS.push tmp
				else
					wscript.echo "Error: ':' but not enough data on stack"
				end if
			case "@" 'rotate top n elements of stack
				if CS.stackcount > 1 then
					CS.rotate
				else
					wscript.echo "Error: '@' but not enough data on stack"
				end if
			case "_" ' drop
				if CS.stackcount > 0 then
					CS.pop
				else
					wscript.echo "Error: '_' but not enough data on stack"
				end if
			case "?" 'test top of stack. Next op should be lowercase a..z being macro
				'~ wscript.echo " QUES "
				ques = cs.pop
				'~ wscript.echo "cs.pop",ques
				if ques then
					macsub = nextop
					hit = instr( "abcdefghijklmnopqrstuvwxyz", macsub ) 
					'~ wscript.echo "hit",hit
					if hit = 0 then
						hit = instr( "ABCDEFGHIJKLMNOPQRSTUVWXYZ", macsub )
						if hit = 0 then
							prevop
							wscript.echo "Error: '?' not followed by macro or subroutine name"
						else
							wscript.echo "[subr " & macsub & "]"
							RS.push ip
							ip = subros(macsub)
						end if
					else
						evalmacro CS, macsub
						'~ embed code(macro)
					end if
				end if
			case "!" 'Next op should be lowercase a..z being macro
				macsub = nextop
				hit = instr( "abcdefghijklmnopqrstuvwxyz", macsub ) 
				'~ wscript.echo "hit",hit
				if hit = 0 then
					hit = instr( "ABCDEFGHIJKLMNOPQRSTUVWXYZ", macsub )
					if hit = 0 then
						prevop
						wscript.echo "Error: '?' not followed by macro or subroutine name"
					else
						wscript.echo "[subr " & macsub & "]"
						RS.push ip
						ip = subros(macsub)
					end if
				else
					evalmacro CS, macsub
					'~ embed code(macro)
				end if
			case "\" 'Next op should be lowercase a..z being macro
				macro = nextop
				if instr( "abcdefghijklmnopqrstuvwxyz", macro ) = 0 then
					prevop
					wscript.echo "Error: '?' not followed by macro name"
				else
					wscript.echo "[clear " & macro & "]"
					m.macro macro, ""
				end if
			case "~" ' not
				if CS.stackcount > 0 then
					CS.push ( not CS.pop )
				else
					wscript.echo "Error: '~' but not enough data on stack"
				end if
			case "^" 'call
				wscript.echo "%%" & ip & "%%"
				RS.push ip
				ip = 1
			case "*" 'return
				if RS.stackcount > 0 then
					ip = RS.pop
				else
					ip = len( script ) + 1
				end if
			case "$" 'swap
				if CS.stackcount > 1 then
					tmp = CS.pop
					tmp2 = CS.pop
					CS.push tmp
					CS.push tmp2
				else
					wscript.echo "Error: '$' but not enough data on stack"
				end if
			case "}" 'pop from PS and push to RS
				IF PS.stackcount > 0 then
					RS.push PS.pop
				else
					wscript.echo "Error: '}' but not enough data on stack"
				end if
			case "{" 'pop from RS and push to PS
				IF RS.stackcount > 0 then
					PS.push RS.pop
				else
					wscript.echo "Error: '}' but not enough data on return-stack"
				end if
			case "1" 'push 1 on stack
				PS.Push 1
			case "0" 'push 0 on stack
				PS.push 0
			case "."
				wscript.echo CS.pop
			case "|"
				wscript.quit 'exit do
			case "`" 'ip = 1
				m.firstop
			end select
			CS.show
		loop	
		'~ evaluate = CS.pop
	end function
	
	private sub apply( S, op, count )
		if S.stackcount > count then
			'~ wscript.echo "[" & op & "]"
			 S.push Eval( "S.pop " & op & " S.pop" )
		else
			wscript.echo "Error in Apply: '" & op & "' but not enough data on stack"
		end if
	end sub
	
	private sub evalmacro( context, macsub )
		dim m2
		wscript.echo "[eval " & macsub & "]"
		set m2 = new machine
		m2.setscript code(macsub)
		m2.evaluate context
		set m2 = nothing
	end sub

	sub showscript
		wscript.echo "script: " & script
	end sub

	sub subroutine( sName, sCode )
		dim newip
		newip = len(script) + 1
		if subros.exists( sName ) then 
			wscript.echo "Error: '" & sName & "' already defined as a subroutine"
		else
			subros.add sName, newIp
			script = script & sCode
		end if
	end sub
	
end class 


dim PS
dim RS

set PS = new stack
set RS = new stack

'~ PS.push 5
'~ PS.push 1
PS.push 0 'n
PS.push 1 'm
PS.show

dim m
set m = new machine
'~ m.define "s", "$}:" 'swap, PS->RS, dup
'~ m.define "q", ".|" 'print return
'~ m.define "a", "1+:" '1 plus
'~ m.define "k", "{:}" 'PS->RS dup RS->PS
'~ m.setscript "!s\s!a!k<?q`|" 'execute s, clear s, execute a, RS->PS, dup, PS->RS, lessthan, test-exec q, jump to start
'~ m.define "a", "$1+.|"
'~ m.define "b", "$:0=?c"
'~ m.define "c", "$1-"
'~ m.setscript ":0=?A!B*"
'~ M.SUBROUTINE "A", ".*"
'~ M.SUBROUTINE "B", ":+.*"
'~ m.macro "a", ":}$:}$"
'~ m.macro "m", ":@:@$"
'~ m.setscript "!a$0=?A{{!a$:@0$>?B|"
'~ m.subroutine "B", "1$-1$^*"
'~ !m0$>?B|"
'~ m.subroutine "A", "1+^*"
'~ m.subroutine "B", "0$>?m"
'macros do testing
'subroutines do manipulations
'all macros assume n|m and leave f
m.macro "z", ":}$:}$"
m.macro "d", "_0="
m.macro "e", "_0$>"
'~ m.macro "c", "$_0="
'~ m.macro "d", "$!b"
'~ m.macro "e", "!z!b{{!c&"
'~ m.macro "f", "!z!b{{!d&"
m.setscript "!z!d?A{{!z!e?B{{!z!C*"
m.subroutine "A", "(restore from return stack){{(swap)$(drop)_(add one)1+(return)*"
m.subroutine "B", "(restore from resturn stack){{(drop)_(add one)1+(one)1(call start)^(return)*"
m.subroutine "C", "(swap)$(swap)$(dup):(rot)@(dup):(rot)@(swap)$1(swap)$(subtract)-(call start)^(swap)$1(swap)$(subtract)-(swap)$(call start)^(return)*"
'~ m.subroutine "Z", "(save to return stack):}$:}$*"
'~ m.subroutine "D", "(drop)_(zero)0(equals)=(return)*"
'~ m.subroutine "E", "_0$>*"
m.subroutine "F", "$_0=*"
m.subroutine "G", "$!E*"
m.subroutine "H", "!z!E{{!F&*"
m.subroutine "J", "!z!E{{!G&*"
m.evaluate PS

