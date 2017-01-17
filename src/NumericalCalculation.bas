Attribute VB_Name = "NumericalCalculation"
Function NumericCalc(Scheme As String, VarNames As Range, VarValues As Range, Constants As Range) As Double
    
    Dim ValueMap As New Collection
    'Var Set-----------------------------start
    For varCnt = 1 To VarNames.Columns.Count

        ValueMap.Add CDbl(VarValues.Cells(1, varCnt)), VarNames.Cells(1, varCnt)
    Next
    'Var Set-----------------------------end
    
    'Constant Set-----------------------------start
    For constCnt = 1 To Constants.Columns.Count

        ValueMap.Add CDbl(Constants.Cells(2, constCnt)), Constants.Cells(1, constCnt)
    Next
    'Constant Set-----------------------------end
    
    NumericCalc = XCalc(Scheme, ValueMap)
    
End Function

Function XCalc(Scheme As String, ValueMap As Collection) As Double
    
    Scheme = Replace(Scheme, " ", "")
    
    'Check First and Last Brackets, Remove them if they are couple
    Do While CheckFirstLast(Scheme)
        Scheme = Mid(Scheme, 2, Len(Scheme) - 2)
    Loop
    
    
    Dim foundFlg As Boolean
    
    foundFlg = False
    
    If IsNumeric(Scheme) Then
    
        XCalc = CDbl(Scheme)
        foundFlg = True
    
    End If
    
    If Not foundFlg Then
    
        Dim n As Integer
        n = 0
           
        Dim operator As String
        Dim opeIndex As Integer
        
        Dim j As Integer
        
        'Search + or - not in ()
        For j = 1 To Len(Scheme)
            
            Dim s As String
            s = Mid(Scheme, Len(Scheme) + 1 - j, 1)
                
            If s = ")" Then
                n = n + 1
            ElseIf s = "(" Then
                n = n - 1
            Else
                If n = 0 And (s = "+" Or s = "-") Then
                    operator = s
                    opeIndex = Len(Scheme) + 1 - j
                    foundFlg = True
                    Exit For
                End If
            End If
        Next
        
        'Search * or / not in () if not found + or -
        If Not foundFlg Then
            
            n = 0
            For j = 1 To Len(Scheme)
                
                Dim ss As String
                ss = Mid(Scheme, Len(Scheme) + 1 - j, 1)
                   
                If ss = ")" Then
                    n = n + 1
                ElseIf ss = "(" Then
                    n = n - 1
                Else
                    If n = 0 And (ss = "*" Or ss = "/") Then
                        operator = ss
                        opeIndex = Len(Scheme) + 1 - j
                        foundFlg = True
                        Exit For
                    End If
                End If
            Next

        End If
                
        'Search ^ not in () if not found +, -, * or /
        If Not foundFlg Then
            
            n = 0
            For j = 1 To Len(Scheme)
                
                Dim sss As String
                sss = Mid(Scheme, Len(Scheme) + 1 - j, 1)
                   
                If sss = ")" Then
                        n = n + 1
                ElseIf sss = "(" Then
                        n = n - 1
                Else
                    If n = 0 And sss = "^" Then
                        operator = sss
                        opeIndex = Len(Scheme) + 1 - j
                        foundFlg = True
                        Exit For
                    End If
                End If
            Next

        End If
        
        
        If foundFlg Then
        
        
            Dim Scheme1 As String
            Dim Scheme2 As String
            
            'if + or - is first Letter
            If opeIndex = 1 And (operator = "+" Or operator = "-") Then
                Scheme1 = 0
            Else
                Scheme1 = Mid(Scheme, 1, opeIndex - 1)
            End If
            
            Scheme2 = Mid(Scheme, opeIndex + 1, Len(Scheme) - opeIndex)
            
            If operator = "+" Then
            
                XCalc = XCalc(Scheme1, ValueMap) + XCalc(Scheme2, ValueMap)
                
            ElseIf operator = "-" Then
            
                XCalc = XCalc(Scheme1, ValueMap) - XCalc(Scheme2, ValueMap)
            
            ElseIf operator = "*" Then
            
                XCalc = XCalc(Scheme1, ValueMap) * XCalc(Scheme2, ValueMap)
                
            ElseIf operator = "/" Then
            
                XCalc = XCalc(Scheme1, ValueMap) / XCalc(Scheme2, ValueMap)
            
            ElseIf operator = "^" Then
            
                XCalc = Application.WorksheetFunction.Power(XCalc(Scheme1, ValueMap), XCalc(Scheme2, ValueMap))
                
            End If
        
        
        Else
        
            'If not found +, -, *, / ----------------------------start
            If Left(Scheme, 4) = "sin(" Then
            
                XCalc = Math.Sin(XCalc(Mid(Scheme, 5, Len(Scheme) - 5), ValueMap))
                
            ElseIf Left(Scheme, 4) = "cos(" Then
            
                XCalc = Math.Cos(XCalc(Mid(Scheme, 5, Len(Scheme) - 5), ValueMap))
                    
            ElseIf Left(Scheme, 4) = "tan(" Then
            
                XCalc = Math.Tan(XCalc(Mid(Scheme, 5, Len(Scheme) - 5), ValueMap))
                
            ElseIf Left(Scheme, 6) = "log10(" Then
             
                XCalc = Application.WorksheetFunction.Log10(XCalc(Mid(Scheme, 7, Len(Scheme) - 7), ValueMap))
                
            ElseIf Left(Scheme, 5) = "loge(" Then
            
                XCalc = Math.Log(XCalc(Mid(Scheme, 6, Len(Scheme) - 6), ValueMap))
                
            ElseIf Left(Scheme, 3) = "ln(" Then
            
                XCalc = Math.Log(XCalc(Mid(Scheme, 4, Len(Scheme) - 4), ValueMap))
                            
            ElseIf Left(Scheme, 4) = "abs(" Then
            
                XCalc = Math.Abs(XCalc(Mid(Scheme, 5, Len(Scheme) - 5), ValueMap))
                            
            Else
            
                'Change Name to Value
                XCalc = ValueMap(Scheme)

            End If
            'If not found +, -, *, / ----------------------------end
               
        End If
    
    End If
    

    
End Function

Function CheckFirstLast(Scheme) As Boolean

    'True   Couple
    'False  Not couple

    Dim intScLen As Integer
    intScLen = Len(Scheme)
    
    
    Dim flg As Boolean
    flg = Not (Left(Scheme, 1) = "(" And Right(Scheme, 1) = ")")
    
    Dim n  As Integer
    Dim targetIndex As Integer
   
    n = 0
    targetIndex = -1
        
    Dim i As Integer
   
    For i = 1 To intScLen
    
    
        Dim s As String
        s = Mid(Scheme, i, 1)
       
        If s = "(" Then
       
            n = n + 1
           
        ElseIf s = ")" Then
       
            n = n - 1
       
        End If
       
        If n < 0 Then
       
            Err.Raise (1)
        End If
       
        If targetIndex = -1 And n = 0 Then
       
            targetIndex = i
        End If
    Next
    
    If Not n = 0 Then
    
        Err.Raise (1)
        
    Else
    
        If flg Then
        
            CheckFirstLast = False
        
        Else
        
            If targetIndex = intScLen Then
            
                CheckFirstLast = True
                
            Else
            
                CheckFirstLast = False
                
            End If
        
        End If
    
    End If

End Function
