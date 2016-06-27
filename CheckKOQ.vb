Option Compare Text
Option Explicit On

Module Module1
    ' NOTES
    ' Program CheckKOQ v.0.99
    ' Usage: "CheckKOQ filename"
    ' Checks Fortran 90 source-code, to verify the correctness of kind-of-quantity in scientific programs,
    ' and outputs the line numbers and variables which have wrong KOQ
    ' For definitions and examples see ISO 80000-1:2009 Quantities and units -- Part 1: General
    ' It is a PROOF-OF-CONCEPT extension to Camfort, a Fortran source-code units-of-measure checker described in:
    ' M. Contrastin, et al., "Units-of-Measure Correctness in Fortran Programs," Computing in Science & Engineering, vol. 18, pp. 102-107, 2016.
    ' The source code must be annotated with comments as follows:
    ' != KOQRelation :: KOQ1 = KOQ2 [*|/] KOQ3 [*|/] KOQ4...
    ' != KOQ <KOQ1 > :: variable1 [, variable2]
    ' This version requires Fortran source to have: no labels; no parenthesis in assignment expressions; full spacing
    ' TO DO:
    ' lots of error checking
    ' size arrays dynamically to hold the maximum number of tokens on a line (or make a more compact data structure)
    ' add parsing to deal with parenthesis in assignment expressions
    ' add parsing to ignore labels
    ' split the unary operator '**2' into a binary operator '** 2', to keep the number of tokens in Varmath/KOQMath even
    ' add space padding either side of operators +,-,/,*

    Sub Main(ByVal cmdArgs() As String)

        ' 0.1 setup the command line stuff
        Dim fileName As String
        fileName = ""
        Dim returnValue As Integer = 0
        ' See if there are any arguments.
        If cmdArgs.Length > 0 Then
            For argNum As Integer = 0 To UBound(cmdArgs, 1)
                ' Console.WriteLine("debug argnum " & argNum & " " & cmdArgs(argNum))
                ' Insert code to examine cmdArgs(argNum) and take appropriate action based on its value.
            Next argNum
            fileName = cmdArgs(0)
        End If
        ' fileName = "Consolidate.f90"  'for IDE debugging

        Dim Version As String
        Dim fileReader1 As String
        Dim NL As Integer
        Version = "CheckKOQ v.0.99"
        NL = 0
        fileReader1 = ""

        ' 0.2 write the initial output
        Console.WriteLine(Version)

        If fileName = "" Or fileName = "?" Or fileName = "/?" Or fileName = "\?" Or fileName = "-?" Then
            ' Console.WriteLine("debug filename= " & fileName)
            Console.WriteLine("Checks Fortran 90 source-code, to verify the correctness of kind-of-quantity in scientific programs")
            Console.WriteLine("For definitions and examples see ISO 80000-1:2009 Quantities and units -- Part 1: General")
            Console.WriteLine("Usage: CheckKOQ filename")
            Exit Sub
        Else
            Console.WriteLine("Checking " & fileName & "...")
        End If

        ' 1.0. count the # of lines in the source code
        FileSystem.FileOpen(1, fileName, OpenMode.Input)
        Do
            NL = NL + 1
            fileReader1 = Microsoft.VisualBasic.FileSystem.LineInput(1)
        Loop Until (EOF(1))
        FileSystem.FileClose(1)
        ' Console.WriteLine("debug Number of lines is: " & NL.ToString())

        '1.2 allocate array sizes and initialise stuff

        Dim LineFlag, KOQFlag, VarFlag, StateFlag As String 'placeholders
        Dim I, J, K, L, M, TC, FL, SubFL, SEC, S As Integer 'counters: I,J,K,L,M,O token count, error flags, sub error flags,sub expression#,?,LC line counter
        Dim Subflag As Integer
        Dim NKOQE, NVE, NVAR, NLS, NE, NLSE As Integer 'totals: KOQ expressions, variable expressions, KOQvariables, lines with subexp, expression lines, subexp line
        Dim ErrorFlag(2 * NL, 4) As String 'integer_line_number, KOQ, variable, error_state
        Dim SubErrorFlag(2 * NL, 4) As String 'decimal_line_number, KOQ, variable, error_state
        Dim tokens(NL, 30) As String 'token, [token]  (40 lines, 30 tokens per line (from INDEX=0 !!!))
        Dim meaning(NL, 3) As String 'line number, type of line, edited_token_count
        Dim variables(NL, 3) As String 'line number, varname, KOQ
        Dim KOQMath(NL, 22) As String 'line number, token count, KOQname, =, exponent, KOQname, [exponent, KOQname]
        Dim KOQExponents(NL, 22) As String 'line number, token count, KOQname, exponent, KOQname [exponent, KOQname]
        Dim VarExpr(NL, 22) As String 'line number, token count, varname, =, exponent, varname, [exponent, varname]
        Dim VarMath(NL, 22) As String 'line number, token count, varname, =, exponent, varname, [exponent, varname]
        Dim VarExponents(NL, 22) As String 'line number, token count, varname, exponent, varname [exponent, varname] {factor as 1,1}
        Dim SubExpr(NL, 22) As String 'line number, token count, varname, exponent, varname,  [ exponent, varname]
        Dim SubExponents(NL, 22) As String 'line number, token count, varname, exponent, varname [exponent, varname] {factor as 1,1}
        Dim Subs(NL, 2) As String 'line number
        Dim FnVarlookup, FnSublookup, FnLineLookup As String
        Dim A, B, C As Integer
        Dim FinalFlag As String
        Dim KOQannotations As Boolean

        For I = 1 To NL : For J = 1 To 30 : tokens(I, J) = "#" : Next J : Next I
        For I = 1 To NL : For J = 1 To 3 : meaning(I, J) = "#" : Next J : Next I
        For I = 1 To NL : For J = 1 To 22 : VarMath(I, J) = "#" : Next J : Next I
        For I = 1 To NL : For J = 1 To 22 : KOQMath(I, J) = "#" : Next J : Next I
        For I = 1 To NL : For J = 1 To 22 : KOQExponents(I, J) = "#" : Next J : Next I
        For I = 1 To NL : For J = 1 To 22 : VarExponents(I, J) = "#" : Next J : Next I
        For I = 1 To NL : For J = 1 To 22 : SubExponents(I, J) = "#" : Next J : Next I
        For I = 1 To NL : For J = 1 To 2 : Subs(I, J) = "#" : Next J : Next I
        For I = 1 To NL : For J = 1 To 22 : SubExpr(I, J) = "#" : Next J : Next I
        For I = 1 To 2 * NL : For J = 1 To 4 : ErrorFlag(I, J) = "#" : Next J : Next I
        For I = 1 To 2 * NL : For J = 1 To 4 : SubErrorFlag(I, J) = "#" : Next J : Next I
        FinalFlag = ""
        KOQannotations = False

        ' 1.5 DEBUG add spaces around = and binary operators  *,/,+,- and around unary operator ** to allow consistent tokenizing
        ' FileSystem.FileOpen(1, fileName, OpenMode.Input)
        '        For I = 1 To NL
        'fr2 = Microsoft.VisualBasic.FileSystem.LineInput(1)
        'Console.WriteLine(fr2)
        'Console.WriteLine("Line " & I & " length =" & Len(fr2))
        'If Len(fr2) < 3 Then GoTo 30
        'For J = 1  To Len(fR2) - 2  'start at second character, end at second last
        'J = 1
        '20:  If ((fr2(J) = "=") And (fr2(J - 1) = "!")) Then
        'J = J + 2
        'Console.WriteLine(fr2)
        'End If
        '       If ((fr2(J) = "/") Or (fr2(J) = "-") Or (fr2(J) = "+") Or (fr2(J) = "=")) Then
        'Console.WriteLine("Line " & I & " position of operators/assignment " & J & " is " & fr2(J))
        'fr2 = Left(fr2, J) & " " & fr2(J) & " " & Right(fr2, Len(fr2) - 1 - J)
        'J = J + 1
        'Console.WriteLine(fr2)
        'End If
        'If (fr2(J) = "*") Then
        'If (fr2(J + 1) = "*") Then
        'Console.WriteLine("Line " & I & " position of double star " & J & " is " & fr2(J))
        'fr2 = Left(fr2, J) & " " & fr2(J) & fr2(J + 1) & " " & Right(fr2, Len(fr2) - 2 - J)
        'J = J + 2 'skip  the next star
        'Console.WriteLine(fr2)
        'Else
        'Console.WriteLine("Line " & I & " position of lone star " & J & " is " & fr2(J))
        'fr2 = Left(fr2, J) & " " & fr2(J) & " " & Right(fr2, Len(fr2) - 1 - J)
        '   J = J + 1
        'Console.WriteLine(fr2)
        'End If
        'End If
        'Next J
        'If J < Len(fr2) - 2 Then
        'J = J + 1
        'GoTo 20
        'End If
        '30:     Next I
        'FileSystem.FileClose(1)


        ' 2. break the lines into tokens and store in the array tokens
        I = 1
        FileSystem.FileOpen(1, fileName, OpenMode.Input)
        Do
            fileReader1 = Microsoft.VisualBasic.FileSystem.LineInput(1)
            fileReader1 = LTrim(RTrim(fileReader1))
            ' Console.WriteLine(I.ToString() & " " & fileReader1)
            ' Sub GetTokens(fileReader1)
            '-------------------------------
            ' was Sub GetTokens(Line1 As String)
            Dim Firstblank As Integer
            Dim NextToken, Line1 As String
            Line1 = fileReader1
            J = 0
            Do
                J = J + 1
                Line1 = LTrim(RTrim(Line1))
                If Line1 = "" Then GoTo 110
                Firstblank = InStr(Line1, " ")
                If Firstblank = 0 Then
                    NextToken = Line1
                Else
                    NextToken = Mid(Line1, 1, Firstblank - 1)
                End If
                'Console.WriteLine("next token: " & NextToken)
                tokens(I, J) = NextToken
                'Console.WriteLine("length of that token: " & Len(NextToken).ToString)
                'Console.WriteLine("remaining length of string: " & Len(Line1).ToString)
                Line1 = Right(Line1, Len(Line1) - Len(NextToken))
                'Console.WriteLine(Line1)
            Loop Until Line1 = ""
            '-------------------------------
110:        I = I + 1
        Loop Until (EOF(1))
        FileSystem.FileClose(1)


        '3. trim the trailing commas off some tokens, and debug print the tokens array
        For I = 1 To NL
            For J = 1 To 30
                If Right(tokens(I, J), 1) Like "," Then
                    tokens(I, J) = Left(tokens(I, J), Len(tokens(I, J)) - 1)
                End If
            Next J

            '3.1 Trim the parameter tokens to remove the equals and arithmetic values (numbers)
            If tokens(I, 1) = "real" And tokens(I, 2) = "parameter" Then
                For J = 5 To 29
                    ' Console.WriteLine("I=" & I.ToString & " J=" & J.ToString & " tokenLeft=" & Left(tokens(I, J), 1) & " " & (Left(tokens(I, J), 1) Like "#")) 'debug
                    If tokens(I, J) = "=" Then
                        For K = J To 29 : tokens(I, K) = tokens(I, K + 1) : Next K
                    End If
                    If Left(tokens(I, J), 1) Like "#" Then
                        For K = J To 29 : tokens(I, K) = tokens(I, K + 1) : Next K
                    End If
                Next J
            End If
        Next I
        ' Console.WriteLine("")
        ' Console.WriteLine(" 3. debug print the tokens array")
        ' For I = 1 To NL
        ' For J = 1 To 30
        ' Console.Write(tokens(I, J) & " ")
        ' Next J
        ' Console.Write(vbCrLf)
        '  Next I

        '4. parse each line in the tokens array to find the meaning and number of tokens per line
        For I = 1 To NL
            meaning(I, 1) = I.ToString
            TC = 0
            If tokens(I, 1) = "program" Then : meaning(I, 2) = "program" : End If
            If tokens(I, 1) = "implicit" Then : meaning(I, 2) = "program" : End If
            If tokens(I, 1) = "#" Then : meaning(I, 2) = "blank" : End If
            If tokens(I, 1) = "!" Then : meaning(I, 2) = "comment" : End If
            If tokens(I, 1) = "!=" Then : meaning(I, 2) = tokens(I, 2) : End If
            If tokens(I, 1) = "end" Then : meaning(I, 2) = "program" : End If
            If tokens(I, 1) = "for" Then : meaning(I, 2) = "program" : End If
            If tokens(I, 1) = "do" Then : meaning(I, 2) = "program" : End If
            If tokens(I, 1) = "until" Then : meaning(I, 2) = "program" : End If
            If tokens(I, 1) = "next" Then : meaning(I, 2) = "program" : End If
            If tokens(I, 1) = "print" Then : meaning(I, 2) = "output" : End If
            If tokens(I, 2) = "=" Then : meaning(I, 2) = "VariableMath" : End If
            If tokens(I, 1) = "real" And tokens(I, 2) = "::" Then : meaning(I, 2) = "DeclareVariables" : End If
            If tokens(I, 1) = "real" And tokens(I, 2) = "parameter" Then : meaning(I, 2) = "DeclareParameters" : End If

            For J = 1 To 30
                If tokens(I, J) <> "#" Then : TC = TC + 1 : End If
            Next J
            meaning(I, 3) = TC.ToString
        Next I

        ' 4.1 debug print the meaning array
        '  Console.WriteLine("")
        '   Console.WriteLine("4. debug print the meaning array")
        '   For I = 1 To NL
        ' Console.WriteLine(meaning(I, 1) & " " & meaning(I, 2) & " " & meaning(I, 3))
        '   Next I

        '5.1 parse the relevant lines in the token array to find the KOQ of each math variable
        K = 0
        For I = 1 To NL
            If meaning(I, 2) = "KOQ" Then
                For J = 5 To CInt(meaning(I, 3))
                    K = K + 1
                    ' Console.WriteLine("debug I=" & I.ToString & " J=" & J.ToString & " K=" & K.ToString & " " & "tokens=" & meaning(I, 3))
                    variables(K, 1) = I.ToString
                    variables(K, 2) = tokens(I, J)
                    variables(K, 3) = tokens(I, 3)
                Next J
            End If
        Next I
        NVAR = K

        '  Console.WriteLine("") 'debug
        ' Console.WriteLine("5.1 debug print the KOQvariables array") 'debug
        '  Console.WriteLine("NVAR=" & NVAR)
        '  For I = 1 To NVAR : Console.WriteLine(variables(I, 1) & " " & variables(I, 2) & " " & variables(I, 3)) : Next I  'debug


        '5.2 parse the variable declarations in the token array to verify the math variables
        ' this is one of hundreds of error checks that could be introduced
        ' difficult because of variable (0+) number of blanks allowed in parameter declaration?
        ' but then Fortan suppodedly ignores blanks, they're just for humans, so how hard can it be?
        ' K = 0
        'For I = 1 To N
        'If meaning(I, 2) = "DeclareParameters" Then
        'For J = 3 To CInt(meaning(I, 3))
        'K = K + 1
        ' Console.WriteLine("debug I=" & I.ToString & " J=" & J.ToString & " K=" & K.ToString & " " & "tokens=" & meaning(I, 3))
        'variables(K, 1) = tokens(I, J)
        'variables(K, 2) = tokens(I, 3)
        'variables(K, 3) = I.ToString
        'Next J
        'End If
        'Next I
        'Console.WriteLine("") 'debug
        'Console.WriteLine("5.1 debug print the variables array") 'debug
        'For I = 1 To K : Console.WriteLine(variables(I, 1) & " " & variables(I, 2) & " " & variables(I, 3)) : Next I  'debug

        '5.3 parse the relevant lines in the token array to find the KOQ math...
        NKOQE = 0
        For I = 1 To NL
            If meaning(I, 2) = "KOQRelation" Then
                KOQannotations = True
                NKOQE = NKOQE + 1
                KOQMath(NKOQE, 1) = (meaning(I, 1))
                KOQMath(NKOQE, 2) = CInt((meaning(I, 3)) - 3).ToString
                For J = 1 To CInt(KOQMath(NKOQE, 2))
                    KOQMath(NKOQE, J + 2) = tokens(I, J + 3)
                Next J
            End If
        Next I

        ' If Not KOQannotations Then
        ' Console.WriteLine("No KOQ annotations found")
        ' Exit Sub
        ' End If

        ' Console.WriteLine("") 'debug
        'Console.WriteLine("5.3b debug print the KOQmath array") 'debug
        ' Console.WriteLine("NKOQE= " & NKOQE.ToString) 'debug
        '  For I = 1 To NKOQE : For J = 1 To CInt(KOQMath(I, 2)) + 2 : Console.Write(KOQMath(I, J) & " ") : Next J : Console.WriteLine() : Next I


        '5.4 parse the relevant lines in the token array to find the variable expressions
        NVE = 0
        For I = 1 To NL
            If meaning(I, 2) = "VariableMath" Then
                NVE = NVE + 1
                VarExpr(NVE, 1) = (meaning(I, 1))
                VarExpr(NVE, 2) = (meaning(I, 3))
                For J = 1 To CInt(VarExpr(NVE, 2))
                    VarExpr(NVE, J + 2) = tokens(I, J)
                    ' Console.WriteLine("5.4a debug VarExpr " & VarExpr(NV, 1) & "= " & tokens(I, J))
                Next J
            End If
        Next I
        ' Console.WriteLine("")
        '  Console.WriteLine("5.4b debug print the VarExpr array") 'debug
        '   Console.WriteLine("NVE= " & NVE.ToString) 'debug
        ' For I = 1 To NVE : For J = 1 To CInt(VarExpr(I, 2)) + 2 : Console.Write(VarExpr(I, J)) : Console.Write(" ") : Next J : Console.WriteLine() : Next I  'debug


        ' 5.5  make a copy of the VarExpr array for further processing
        For I = 1 To NL : For J = 1 To 22 : VarMath(I, J) = VarExpr(I, J) : Next J : Next I


        '6.1 substitute KOQ for variables in the VarMath array
        For I = 1 To NVE
            For J = 3 To CInt(VarMath(I, 2)) + 2 'why 3?
                For K = 1 To NVAR   'look through the Variables array to substitute KOQ
                    ' temp(I, J) = VarMath(I, J)
                    If VarMath(I, J) = variables(K, 2) Then
                        VarMath(I, J) = variables(K, 3) ' was temp
                        Exit For
                    End If
                Next K
                'Console.WriteLine("temp(I, J)= " & temp(I, J) & " variables(K, 2)= " & variables(K, 2) & " variables(K,3)= " & variables(K, 3))
                'Console.WriteLine(I.ToString & " " & J.ToString & " " & K.ToString & " " & temp(I, J))
            Next J
        Next I
        '  Console.WriteLine() 'debug
        '  Console.WriteLine("6.1 debug print the Varmath array") 'debug
        '  For I = 1 To NVE : For J = 1 To CInt(VarMath(I, 2)) + 2 : Console.Write(VarMath(I, J)) : Console.Write(" ") : Next J : Console.WriteLine() : Next I


        '6.3A construct the subs array (line numbers which have subexpressions bound by '+' and '-') (too complex to combine VarExpr and SubExpr?)
        NLS = 0
        For I = 1 To NVE 'check for subexpressions, build array of counters
            SEC = 1
            For J = 5 To CInt(VarMath(I, 2)) + 2
                If VarMath(I, J) = "+" Or VarMath(I, J) = "-" Then
                    SEC = SEC + 1
                End If
            Next J
            If SEC > 1 Then : NLS = NLS + 1 : Subs(NLS, 1) = VarMath(I, 1) : Subs(NLS, 2) = SEC.ToString : End If
        Next I
        ' Now count the number of lines in the Subs array
        NLS = 0 : For I = 1 To NL : If (Subs(I, 1) <> "#") Then : NLS = NLS + 1 : End If : Next I
        ' Console.WriteLine()
        ' Console.WriteLine("6.3A debug print subs array")
        ' Console.WriteLine("NLS= " & NLS.ToString)
        ' For I = 1 To NLS : For J = 1 To 2 : Console.Write(Subs(I, J) & " ") : Next J : Console.WriteLine() : Next I    'debug
        ' Console.WriteLine()



        '6.3B construct the SubExpression array (line numbers which have subexpressions bound by '+' and '-') (too complex to combine VarExpr and SubExpr?)
        M = 1 'subexpr Y
        For I = 1 To NLS
            K = 5 'subexpr X
            S = 1 'sub line number

            ' was Function FnLinelookup(A As String) As String
            A = (Subs(I, 1))
            FnLineLookup = "?"
            For L = 1 To NVE
                If VarExpr(L, 1) = A Then FnLineLookup = L
            Next L
            ' was End Sub
            J = FnLineLookup

            For L = 5 To CInt(VarExpr(J, 2)) + 2
                If (VarExpr(J, L) = "+" Or VarExpr(J, L) = "-") Then : K = 5 : M = M + 1 : L = L + 1 : S = S + 1 : End If 'revisit this
                SubExpr(M, 1) = Subs(I, 1) & "." & S  'decimal line number
                SubExpr(M, 2) = L '#tokens
                SubExpr(M, 3) = VarExpr(J, 3) 'var
                SubExpr(M, 4) = "="  'VarExpr(J, 3) 'var
                SubExpr(M, K) = VarExpr(J, L)
                'Console.WriteLine("debug I=" & I & " L=" & L & " J=" & J & " M=" & M & " K=" & K)
                'For O = 1 To 7 : Console.Write(SubExpr(M, O) & " ") : Next O : Console.WriteLine("")
                K = K + 1
            Next L
            M = M + 1
            S = S + 1
            NLSE = M
        Next I
        If (NLSE > 0) Then NLSE = NLSE - 1
        For I = 1 To NLSE 'count the reduced number of tokens
            TC = 0
            For J = 3 To 22
                If SubExpr(I, J) <> "#" Then : TC = TC + 1 : End If
                ' Console.WriteLine("I=" & I & " J=" & J & " TC=" & TC) 'debug token count
            Next J
            SubExpr(I, 2) = TC.ToString
        Next I
        ' Console.WriteLine()
        '  Console.WriteLine("6.3B debug print the SubExpr array") 'debug
        '  Console.WriteLine("NLSE=" & NLSE)
        ' For I = 1 To NLSE : For J = 1 To SubExpr(I, 2) + 2 : Console.Write(SubExpr(I, J) & " ") : Next J : Console.WriteLine() : Next I  'debug


        '6.3C construct the SubExponents array (line numbers which have subexpressions bound by '+' and '-') (too complex to combine VarExpr and SubExpr?)
        M = 1 'subexponents Y
        For I = 1 To NLS
            K = 5 'subexponents X
            S = 1 'sub line number
            ' was Function FnLinelookup(A As String) As String
            A = (Subs(I, 1))
            FnLineLookup = "?"
            For L = 1 To NVE
                If VarExpr(L, 1) = A Then FnLineLookup = L
            Next L
            ' was End Sub
            ' J = FnLinelookup(Subs(I, 1))  
            J = FnLineLookup

            For L = 5 To CInt(VarMath(J, 2)) + 2
                If (VarMath(J, L) = "+" Or VarMath(J, L) = "-") Then : K = 5 : M = M + 1 : L = L + 1 : S = S + 1 : End If 'revisit this
                SubExponents(M, 1) = Subs(I, 1) & "." & S  'decimal line number
                SubExponents(M, 2) = VarMath(J, 2) '#tokens
                SubExponents(M, 3) = VarMath(J, 3) 'var
                SubExponents(M, 4) = "1" ' overwrite the =; first var has index of 1
                ' If VarMath(J, L) = "#" Then : Exit For : End If
                If Left(VarMath(J, L), 1) Like "#" Then : SubExponents(M, K) = "1"
                ElseIf VarMath(J, L) = "*" Then : SubExponents(M, K) = "1"
                    ' Console.WriteLine("debug NEW 6.4 I=" & I & " K=" & K & "J=" & J & " " & VarExponents(K, 1) & " " & VarExponents(K, 2) & " " & VarExponents(K, 3) & " " & VarExponents(K, 4) & " " & VarExponents(K, 5)) 'debug
                ElseIf VarMath(J, L) = "/" Then : SubExponents(M, K) = "-1"
                ElseIf VarMath(J, L) = "**" Then
                    SubExponents(M, K) = "1"
                    SubExponents(M, K - 2) = (CInt(VarMath(J, L + 1)) * CInt(SubExponents(M, K - 2))).ToString
                Else : SubExponents(M, K) = VarMath(J, L)
                End If
                'Console.WriteLine("debug I=" & I & " L=" & L & " J=" & J & " M=" & M & " K=" & K)
                'For O = 1 To 7 : Console.Write(SubExponents(M, O) & " ") : Next O : Console.WriteLine("")
                K = K + 1
            Next L
            M = M + 1
            S = S + 1
            NLSE = M
        Next I
        If (NLSE > 0) Then NLSE = NLSE - 1
        For I = 1 To NLSE 'count the reduced number of tokens
            TC = 0
            For J = 3 To 22
                If SubExponents(I, J) <> "#" Then : TC = TC + 1 : End If
                ' Console.WriteLine("I=" & I & " J=" & J & " TC=" & TC) 'debug token count
            Next J
            SubExponents(I, 2) = TC.ToString
        Next I
        ' Console.WriteLine()
        ' Console.WriteLine("6.3C debug print the SubExponents array") 'debug
        ' Console.WriteLine("NLSE=" & NLSE)
        ' For I = 1 To NLSE : For J = 1 To SubExponents(I, 2) + 2 : Console.Write(SubExponents(I, J) & " ") : Next J : Console.WriteLine() : Next I  'debug

        ' Console.WriteLine()
        ' Console.WriteLine("Debug 6.3D Check the SubExponents array for repeated variables and consolidate")
        ' new in v0.99: 0 for var name is used to indicate that varexp has been consolidated and should not be looked up for varname
        For I = 1 To NLSE
            For J = 5 To CInt(SubExponents(I, 2)) + 2 Step 2
                For K = J + 2 To CInt(SubExponents(I, 2)) + 2
                    'these <> expressions never work on strings!
                    ' If (SubExponents(I, J)) <> "1" Then
                    If ((SubExponents(I, J) <> "1") And (SubExponents(I, J) <> "0")) Then 'new in v0.99
                        If (SubExponents(I, J)) = SubExponents(I, K) Then
                            SubExponents(I, J - 1) = (CInt(SubExponents(I, J - 1)) + CInt(SubExponents(I, K - 1))).ToString
                            SubExponents(I, K - 1) = "0" : SubExponents(I, K) = "0"
                        End If
                    End If
                Next K
            Next J
            For J = 5 To CInt(SubExponents(I, 2)) + 2 Step 2
                If (SubExponents(I, J - 1)) = "0" Then
                    'SubExponents(I, J - 1) = "1"
                    SubExponents(I, J) = "0"
                End If
            Next J
        Next I
        ' Console.WriteLine()
        ' Console.WriteLine("6.3D debug print the consolidated SubExponents array") 'debug
        ' Console.WriteLine("NLSE=" & NLSE)
        ' For I = 1 To NLSE : For J = 1 To SubExponents(I, 2) + 2 : Console.Write(SubExponents(I, J) & " ") : Next J : Console.WriteLine() : Next I  'debug



        ' 6.4 construct the VarExponents array 
        K = 0
        For I = 1 To NVE 'check for subexpressions, build array of counters
            Subflag = 0
            For J = 1 To NLS
                If VarMath(I, 1) = Subs(J, 1) Then
                    Subflag = 1
                End If
            Next J
            If Subflag = 0 Then
                K = K + 1
                VarExponents(K, 1) = VarMath(I, 1) 'line#
                VarExponents(K, 2) = VarMath(I, 2) '#tokens
                VarExponents(K, 3) = VarMath(I, 3) 'var
                VarExponents(K, 4) = "1" ' overwrite the =; first var has index of 1
                For J = 5 To CInt(VarMath(I, 2)) + 2
                    If VarMath(I, J) = "#" Then : Exit For : End If
                    If Left(VarMath(I, J), 1) Like "#" Then : VarExponents(K, J) = "1" 'numerical factor
                    ElseIf VarMath(I, J) = "*" Then : VarExponents(K, J) = "1"
                        ' Console.WriteLine("debug NEW 6.4 I=" & I & " K=" & K & "J=" & J & " " & VarExponents(K, 1) & " " & VarExponents(K, 2) & " " & VarExponents(K, 3) & " " & VarExponents(K, 4) & " " & VarExponents(K, 5)) 'debug
                    ElseIf VarMath(I, J) = "/" Then : VarExponents(K, J) = "-1"
                    ElseIf VarMath(I, J) = "**" Then
                        VarExponents(K, J) = "1"
                        VarExponents(K, J - 2) = (CInt(VarMath(I, J + 1)) * CInt(VarExponents(K, J - 2))).ToString
                    Else : VarExponents(K, J) = VarMath(I, J)
                    End If
                Next J
            End If
        Next I
        NE = K
        For I = 1 To NVE 'count the reduced number of tokens
            TC = 0
            For J = 3 To 22
                If VarExponents(I, J) <> "#" Then : TC = TC + 1 : End If
                ' Console.WriteLine("I=" & I & " J=" & J & " TC=" & TC) 'debug token count
            Next J
            VarExponents(I, 2) = TC.ToString
        Next I
        ' Console.WriteLine()
        ' Console.WriteLine("6.4A debug print the raw VarExponents array") 
        ' Console.WriteLine("NE=" & NE)
        ' For I = 1 To NE : For J = 1 To CInt(VarExponents(I, 2)) + 2 : Console.Write(VarExponents(I, J) & " ") : Next J : Console.WriteLine() : Next I  

        ' Console.WriteLine("Debug 6.4B Check the expression for repeated variables and consolidate")
        ' new in v0.99: 0 for var name is used to indicate that varexp has been consolidated and should not be looked up for varname
        For I = 1 To NE
            For J = 5 To CInt(VarExponents(I, 2)) + 2 Step 2
                For K = J + 2 To CInt(VarExponents(I, 2)) + 2
                    'these <> expressions never work on strings!
                    If ((VarExponents(I, J) <> "1") And (VarExponents(I, J) <> "0")) Then
                        If (VarExponents(I, J)) = VarExponents(I, K) Then
                            VarExponents(I, J - 1) = (CInt(VarExponents(I, J - 1)) + CInt(VarExponents(I, K - 1))).ToString
                            VarExponents(I, K - 1) = "0" : VarExponents(I, K) = "0"
                        End If
                    End If
                Next K
            Next J
            ' If some exponent is zero, then remove that variable
            For J = 5 To CInt(VarExponents(I, 2)) + 2
                If (VarExponents(I, J - 1)) = "0" Then
                    '     VarExponents(I, J - 1) = "1"
                    VarExponents(I, J) = "0"
                End If
            Next J
        Next I
        ' Console.WriteLine()
        ' Console.WriteLine("6.4B debug print the consolidated VarExponents array") 'debug
        ' Console.WriteLine("NE=" & NE)
        ' For I = 1 To NE : For J = 1 To CInt(VarExponents(I, 2)) + 2 : Console.Write(VarExponents(I, J) & " ") : Next J : Console.WriteLine() : Next I  'debug


        ' 6.5 construct the KOQ Exponents array 
        For I = 1 To NKOQE
            KOQExponents(I, 1) = KOQMath(I, 1) 'line#
            KOQExponents(I, 2) = KOQMath(I, 2) '#tokens
            KOQExponents(I, 3) = KOQMath(I, 3) 'var
            KOQExponents(I, 4) = "1" 'skip the =; first var has index of 1
            KOQExponents(I, 5) = KOQMath(I, 5) 'var
            For J = 6 To CInt(KOQMath(I, 2)) + 2
                If KOQMath(I, J) = "#" Then : Exit For : End If
                If Left(KOQMath(I, J), 1) Like "#" Then : KOQExponents(I, J) = "1"
                ElseIf KOQMath(I, J) = "*" Then : KOQExponents(I, J) = "1"
                ElseIf KOQMath(I, J) = "/" Then : KOQExponents(I, J) = "-1"
                ElseIf KOQMath(I, J) = "**" Then
                    KOQExponents(I, J) = "1"
                    KOQExponents(I, J - 2) = (CInt(KOQMath(I, J + 1)) * CInt(KOQExponents(I, J - 2))).ToString
                Else : KOQExponents(I, J) = KOQMath(I, J)
                End If
            Next J
        Next I
        For I = 1 To NKOQE 'count the reduced number of tokens
            TC = 0
            For J = 3 To 22
                If KOQExponents(I, J) <> "#" Then : TC = TC + 1 : End If
                ' Console.WriteLine("I=" & I & " J=" & J & " TC=" & TC) 'debug token count
            Next J
            KOQExponents(I, 2) = TC.ToString
        Next I
        ' Console.WriteLine()
        ' Console.WriteLine("6.5 debug print the KOQExponents array") 'debug
        ' Console.WriteLine("NKOQE=" & NKOQE)
        ' For I = 1 To NKOQE : For J = 1 To CInt(KOQExponents(I, 2)) + 2 : Console.Write(KOQExponents(I, J) & " ") : Next J : Console.WriteLine() : Next I  'debug
        ' Console.WriteLine()

        ' Console.WriteLine("Debug 6.6 Check the KOQExponents array for repeated variables and consolidate")
        ' new in v0.99: 0 for var name is used to indicate that varexp has been consolidated and should not be looked up for varname
        For I = 1 To NKOQE
            For J = 5 To CInt(KOQExponents(I, 2)) + 2 Step 2
                For K = J + 2 To CInt(KOQExponents(I, 2)) + 2
                    If ((KOQExponents(I, J) <> "1") And (KOQExponents(I, J) <> "0")) Then 'new in v0.99
                        If (KOQExponents(I, J)) = KOQExponents(I, K) Then
                            KOQExponents(I, J - 1) = (CInt(KOQExponents(I, J - 1)) + CInt(KOQExponents(I, K - 1))).ToString
                            KOQExponents(I, K - 1) = "0" : KOQExponents(I, K) = "0"
                        End If
                    End If
                    ' Console.Write("I=" & I & " J=" & J & " K= " & K & " ")
                    ' For M = 1 To CInt(KOQExponents(I, 2)) + 2 : Console.Write(KOQExponents(I, M) & " ") : Next M : Console.WriteLine()
                Next K
            Next J

            ' new in v0.99: ditto, If some consolidated exponent is zero, then set KOQ name = 0
            For J = 5 To CInt(KOQExponents(I, 2)) + 2 Step 2
                If (KOQExponents(I, J - 1)) = "0" Then
                    '                KOQExponents(I, J - 1) = "1"
                    KOQExponents(I, J) = "0"
                End If
            Next J
        Next I
        ' Console.WriteLine()
        ' Console.WriteLine("6.6 debug print the consolidated KOQExponents array") 'debug
        ' Console.WriteLine("NKOQE=" & NKOQE)
        ' For I = 1 To NKOQE : For J = 1 To CInt(KOQExponents(I, 2)) + 2 : Console.Write(KOQExponents(I, J) & " ") : Next J : Console.WriteLine() : Next I  'debug


        '7.1 Comparison #1: VarExponents array with KOQExponents array, find KOQ mismatch and line number
        '  Console.WriteLine()
        ' Console.WriteLine("7.1 Compare VarExponents array with KOQExponents array")
        FL = 0
        For I = 1 To NVE
            For J = 1 To NKOQE
                If VarExponents(I, 3) = KOQExponents(J, 3) Then
                    For L = 4 To CInt(VarExponents(I, 2)) + 1 Step 2
                        For K = 4 To CInt(KOQExponents(J, 2)) + 1 Step 2
                            If Not (VarExponents(I, L + 1) = "0" Or KOQExponents(J, K + 1) = "0") Then 'don't lookup this consolidated var
                                If (VarExponents(I, L) = KOQExponents(J, K)) And (VarExponents(I, L + 1) = KOQExponents(J, K + 1)) Then
                                    FL = FL + 1
                                    'Console.WriteLine("debug: FL= " & FL & " " & VarExponents(I, L) & " " & VarExponents(I, L + 1) & " " & KOQExponents(J, K) & " " & KOQExponents(J, K + 1) & " " & ErrorFlag(FL, 4))
                                    'Console.WriteLine("debug ")
                                    ErrorFlag(FL, 1) = VarExponents(I, 1)
                                    ErrorFlag(FL, 2) = VarExponents(I, L + 1) 'KOQ
                                    'was   Function FnVarLookup(A As Integer, B As Integer) As String
                                    A = I : B = L + 1
                                    FnVarlookup = "?Var"
                                    For C = 1 To NVE
                                        If VarExpr(C, 1) = VarExponents(A, 1) Then FnVarlookup = VarExpr(C, B)
                                    Next C
                                    ' was End Function
                                    ErrorFlag(FL, 3) = FnVarlookup 'Var itself?
                                    ErrorFlag(FL, 4) = "match"
                                    'Console.WriteLine("I=" & I & " J=" & J & " L=" & L & " K=" & K)
                                    'Console.WriteLine(VarExponents(I, L) & " " & VarExponents(I, L + 1) & " " & KOQExponents(J, K) & " " & KOQExponents(J, K + 1) & " " & ErrorFlag(FL, 4))
                                Else
                                    FL = FL + 1
                                    ErrorFlag(FL, 1) = VarExponents(I, 1)
                                    ErrorFlag(FL, 2) = VarExponents(I, L + 1) 'KOQ
                                    'was   Function FnVarLookup(A As Integer, B As Integer) As String
                                    A = I : B = L + 1
                                    FnVarlookup = "?Var"
                                    For C = 1 To NVE
                                        If VarExpr(C, 1) = VarExponents(A, 1) Then FnVarlookup = VarExpr(C, B)
                                    Next C
                                    ' was End Sub
                                    ErrorFlag(FL, 3) = FnVarlookup 'Var itself?
                                    ErrorFlag(FL, 4) = "error"
                                    ' Console.WriteLine("I=" & I & " J=" & J & " L=" & L & " K=" & K)
                                    ' Console.WriteLine(VarExponents(I, L) & " " & VarExponents(I, L + 1) & " " & KOQExponents(J, K) & " " & KOQExponents(J, K + 1) & " " & ErrorFlag(FL, 4))
                                End If
                            End If
                        Next K
                    Next L
                End If
            Next J
        Next I
        'Console.WriteLine()
        ' Console.WriteLine("7.1 Debug print Errorflag")
        '  Console.WriteLine("FL=" & FL)
        '  For I = 1 To FL
        ' Console.WriteLine("Line " & ErrorFlag(I, 1) & " KOQ " & ErrorFlag(I, 2) & " Variable " & ErrorFlag(I, 3) & " has " & ErrorFlag(I, 4))
        '   Next I


        '7.2 Comparison #2: SubExponents array with KOQExponents array, find KOQ mismatch and line number
        ' Console.WriteLine()
        ' Console.WriteLine("7.2 Compare SubExponents array with KOQExponents array-requires additional internal comparison")
        SubFL = 0
        For I = 1 To NLSE
            ' This is the internal comparison for the case of exactly one variable on the RHS
            If (SubExponents(I, 2) = "3") Then
                If ((SubExponents(I, 3) = SubExponents(I, 5)) And (SubExponents(I, 4) = "1")) Then
                    ' Console.WriteLine("DEBUG -subexponents internal match... I=" & I & " SubExponents(I,5)=" & SubExponents(I, 5))
                    SubFL = SubFL + 1
                    SubErrorFlag(SubFL, 1) = SubExponents(I, 1) 'decimal line number
                    SubErrorFlag(SubFL, 2) = SubExponents(I, 3) 'KOQ
                    'was  Function FnSubLookup(A As Integer, B As Integer) As String
                    FnSublookup = "?Sub"
                    A = I : B = 5
                    For C = 1 To NLSE
                        If SubExpr(C, 1) = SubExponents(A, 1) Then FnSublookup = SubExpr(C, B)
                    Next C
                    'was End Sub
                    SubErrorFlag(SubFL, 3) = FnSublookup
                    SubErrorFlag(SubFL, 4) = "match"
                    ' Console.WriteLine("I=" & I & " " & (3 + 2 * CInt(Right(SubExponents(I, 1), 1)).ToString & " variable=" & SubErrorFlag(SubFL, 3)))
                Else
                    '  Console.WriteLine("DEBUG -subexponents internal error... I=" & I & " SubExponents(I,5)=" & SubExponents(I, 5))
                    SubFL = SubFL + 1
                    SubErrorFlag(SubFL, 1) = SubExponents(I, 1) 'decimal line number
                    SubErrorFlag(SubFL, 2) = SubExponents(I, 3) 'KOQ
                    'was  Function FnSubLookup(A As Integer, B As Integer) As String
                    FnSublookup = "?Sub"
                    A = I : B = 5
                    For C = 1 To NLSE
                        If SubExpr(C, 1) = SubExponents(A, 1) Then FnSublookup = SubExpr(C, B)
                    Next C
                    'was End Sub
                    SubErrorFlag(SubFL, 3) = FnSublookup
                    SubErrorFlag(SubFL, 4) = "error"
                    '  Console.WriteLine("I=" & I & " " & (3 + 2 * CInt(Right(SubExponents(I, 1), 1)).ToString & " variable=" & SubErrorFlag(SubFL, 3)))
                End If
                GoTo 100
            End If
            For J = 1 To NKOQE
                If SubExponents(I, 3) = KOQExponents(J, 3) Then
                    For L = 4 To CInt(SubExponents(I, 2)) + 1 Step 2
                        For K = 4 To CInt(KOQExponents(J, 2)) + 1 Step 2
                            If Not (SubExponents(I, L + 1) = "0" Or KOQExponents(J, K + 1) = "0") Then 'don't lookup this consolidated var
                                If ((SubExponents(I, L) = KOQExponents(J, K)) And (SubExponents(I, L + 1) = KOQExponents(J, K + 1))) Then
                                    SubFL = SubFL + 1
                                    SubErrorFlag(SubFL, 1) = SubExponents(I, 1) 'decimal line number
                                    SubErrorFlag(SubFL, 2) = SubExponents(I, L + 1) 'KOQ
                                    'was  Function FnSubLookup(A As Integer, B As Integer) As String
                                    FnSublookup = "?Sub"
                                    A = I : B = L + 1
                                    '+++++++++++++++ For C = 1 To NVE 'SHOULDNT THIS BE NLSE NOT NVE???+++++++++++++++++++++++++++++++
                                    For C = 1 To NLSE
                                        If SubExpr(C, 1) = SubExponents(A, 1) Then FnSublookup = SubExpr(C, B)
                                    Next C
                                    'was End Sub
                                    SubErrorFlag(SubFL, 3) = FnSublookup
                                    SubErrorFlag(SubFL, 4) = "match"
                                    'Console.WriteLine("I=" & I & " J=" & J & " L=" & L & " K=" & K)
                                    'Console.WriteLine(SubExponents(I, L) & " " & SubExponents(I, L + 1) & " " & KOQExponents(J, K) & " " & KOQExponents(J, K + 1) & " " & SubErrorFlag(FL, 4))
                                Else
                                    SubFL = SubFL + 1
                                    SubErrorFlag(SubFL, 1) = SubExponents(I, 1) 'decimal line number
                                    SubErrorFlag(SubFL, 2) = SubExponents(I, L + 1) 'KOQ
                                    'was Function FnSubLookup(A As Integer, B As Integer) As String
                                    FnSublookup = "?Sub"
                                    A = I : B = L + 1
                                    '+++++++++++++++ For C = 1 To NVE 'SHOULDNT THIS BE NLSE NOT NVE???+++++++++++++++++++++++++++++++
                                    For C = 1 To NLSE
                                        If SubExpr(C, 1) = SubExponents(A, 1) Then FnSublookup = SubExpr(C, B)
                                    Next C
                                    'was End Sub
                                    SubErrorFlag(SubFL, 3) = FnSublookup
                                    SubErrorFlag(SubFL, 4) = "error"
                                    'Console.WriteLine("I=" & I & " J=" & J & " L=" & L & " K=" & K)
                                    'Console.WriteLine(SubExponents(I, L) & " " & SubExponents(I, L + 1) & " " & KOQExponents(J, K) & " " & KOQExponents(J, K + 1) & " " & SubErrorFlag(FL, 4))
                                End If
                            End If
                        Next K
                    Next L
                End If
            Next J
100:    Next I
        ' Console.WriteLine()
        ' Console.WriteLine("7.2 Debug print SubErrorflag")
        ' Console.WriteLine("SubFL= " & SubFL)
        ' For I = 1 To SubFL
        ' Console.WriteLine("Line " & SubErrorFlag(I, 1) & " KOQ " & SubErrorFlag(I, 2) & " Variable " & SubErrorFlag(I, 3) & " has " & SubErrorFlag(I, 4))
        ' Next I
        '  Console.WriteLine()


        '8.1 Interpret the Errorflag and SubErrorflag arrays and generate the final output
        ' Console.WriteLine()
        ' Console.WriteLine("8.0 Interpret the Errorflag array and generate the final output")
        LineFlag = ErrorFlag(1, 1)
        KOQFlag = ErrorFlag(1, 2)
        VarFlag = ErrorFlag(1, 3)
        StateFlag = ErrorFlag(1, 4)
        For I = 2 To FL
            If ErrorFlag(I, 1) = LineFlag Then 'if the same line, step through the line data and update var and state
                If ErrorFlag(I, 3) = VarFlag Then 'if the same var, update the state
                    If ErrorFlag(I, 4) = "match" Then
                        StateFlag = ErrorFlag(I, 4)
                    End If
                Else 'if not the same var, output then update var and state
                    If (KOQFlag <> "1" And StateFlag = "error") Then FinalFlag = "error" : Console.WriteLine("Line " & LineFlag & " variable " & VarFlag & " has KOQ error")
                    ' Console.WriteLine("NewVar: " & " " & LineFlag & " " & KOQFlag & " " & VarFlag & " " & Stateflag)
                    KOQFlag = ErrorFlag(I, 2)
                    VarFlag = ErrorFlag(I, 3)
                    StateFlag = ErrorFlag(I, 4)
                End If
            Else 'must be a new line number
                'If Stateflag = "error" Then
                ' Console.WriteLine("NewVar: " & " " & LineFlag & " " & KOQFlag & " " & VarFlag & " " & Stateflag)
                If (KOQFlag <> "1" And StateFlag = "error") Then FinalFlag = "error" : Console.WriteLine("Line " & LineFlag & " variable " & VarFlag & " has KOQ error")
                LineFlag = ErrorFlag(I, 1)
                KOQFlag = ErrorFlag(I, 2)
                VarFlag = ErrorFlag(I, 3)
                StateFlag = ErrorFlag(I, 4)
            End If
        Next I
        ' process the last variable
        If (KOQFlag <> "1" And StateFlag = "error") Then FinalFlag = "error" : Console.WriteLine("Line " & LineFlag & " variable " & VarFlag & " has KOQ error")


        ' 8.2 Interpret the  SubErrorflag array and generate the final output
        LineFlag = SubErrorFlag(1, 1)
        KOQFlag = SubErrorFlag(1, 2)
        VarFlag = SubErrorFlag(1, 3)
        StateFlag = SubErrorFlag(1, 4)
        For I = 2 To SubFL 'changed in v0.99 from FL 
            If SubErrorFlag(I, 1) = LineFlag Then 'if the same line, step thruogh the line data and update var and state
                If SubErrorFlag(I, 3) = VarFlag Then 'if the same var, update the state
                    If SubErrorFlag(I, 4) = "match" Then
                        StateFlag = SubErrorFlag(I, 4)
                    End If
                Else
                    ' Console.WriteLine("'if not the same var, output then update var and state")
                    ' Console.WriteLine("NewVar: " & " " & LineFlag & " " & KOQFlag & " " & VarFlag & " " & StateFlag)
                    If (KOQFlag <> "1" And StateFlag = "error") Then FinalFlag = "error" : Console.WriteLine("Line " & Left(LineFlag, InStr(LineFlag, ".") - 1) & " variable " & VarFlag & " has KOQ error")
                    KOQFlag = SubErrorFlag(I, 2)
                    VarFlag = SubErrorFlag(I, 3)
                    StateFlag = SubErrorFlag(I, 4)
                End If
            Else 'must be a new line number
                'If Stateflag = "error" Then
                ' Console.WriteLine("NewVar: " & " " & LineFlag & " " & KOQFlag & " " & VarFlag & " " & Stateflag)
                If (KOQFlag <> "1" And StateFlag = "error") Then FinalFlag = "error" : Console.WriteLine("Line " & Left(LineFlag, InStr(LineFlag, ".") - 1) & " variable " & VarFlag & " has KOQ error")
                LineFlag = SubErrorFlag(I, 1)
                KOQFlag = SubErrorFlag(I, 2)
                VarFlag = SubErrorFlag(I, 3)
                StateFlag = SubErrorFlag(I, 4)
            End If
        Next I
        ' process the last variable
        If (KOQFlag <> "1" And StateFlag = "error") Then FinalFlag = "error" : Console.WriteLine("Line " & Left(LineFlag, InStr(LineFlag, ".") - 1) & " variable " & VarFlag & " has KOQ error")


        If FinalFlag = "" Then Console.WriteLine("NO KOQ errors detected")
    End Sub ' Main

End Module

