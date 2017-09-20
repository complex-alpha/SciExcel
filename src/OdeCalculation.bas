Attribute VB_Name = "OdeCalculation"
Function Euler(IndependentVariable As String, ObjectiveVariable As String, VarSchemes As Range, StepValue As Double, VarNames As Range, VarValues As Range, Constants As Range) As Double

    Dim f As String
    
    For cnt = 1 To VarSchemes.Columns.Count

        If VarSchemes.Cells(1, cnt) = ObjectiveVariable Then
        
            f = VarSchemes.Cells(2, cnt)
            Exit For
            
        End If
    Next
    
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
    
    Euler = ValueMap(ObjectiveVariable) + XCalc(f, ValueMap) * StepValue

End Function

Function RK4Calc(IndependentVariable As String, ObjectiveVariable As String, VarSchemes As Range, StepValue As Double, VarNames As Range, VarValues As Range, Constants As Range)

    Dim fMap As New Collection
    For cnt = 1 To VarSchemes.Columns.Count

        fMap.Add VarSchemes.Cells(2, cnt), VarSchemes.Cells(1, cnt)
            
    Next
  
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
    
    Dim k1 As Double
    Dim ValueMap2 As New Collection
    
    For varCnt = 1 To VarNames.Columns.Count
    
        varName = VarNames.Cells(1, varCnt).Value
        
        
        If IndependentVariable = varName Then
        
            v = ValueMap(varName) + StepValue / 2
        
        Else
        
            v = XCalc(fMap(varName), ValueMap)
            
            If ObjectiveVariable = varName Then
        
                k1 = v
            End If
            
        End If
        
        ValueMap2.Add ValueMap(varName) + v * StepValue / 2, varName
        
    Next
    'Constant Set-----------------------------start
    For constCnt = 1 To Constants.Columns.Count

        ValueMap2.Add CDbl(Constants.Cells(2, constCnt)), Constants.Cells(1, constCnt)
    Next
    'Constant Set-----------------------------end
    
    Dim k2 As Double
    Dim ValueMap3 As New Collection
    
    For varCnt = 1 To VarNames.Columns.Count

        varName = VarNames.Cells(1, varCnt).Value


        If IndependentVariable = varName Then

            v = ValueMap(varName) + StepValue / 2

        Else

            v = XCalc(fMap(varName), ValueMap2)

            If ObjectiveVariable = varName Then

                k2 = v
            End If

        End If

        ValueMap3.Add ValueMap2(varName) + v * StepValue / 2, varName

    Next
    'Constant Set-----------------------------start
    For constCnt = 1 To Constants.Columns.Count

        ValueMap3.Add CDbl(Constants.Cells(2, constCnt)), Constants.Cells(1, constCnt)
    Next
    'Constant Set-----------------------------end
    
    Dim k3 As Double
    Dim ValueMap4 As New Collection
    
    For varCnt = 1 To VarNames.Columns.Count

        varName = VarNames.Cells(1, varCnt).Value


        If IndependentVariable = varName Then

            v = ValueMap(varName) + StepValue

        Else

            v = XCalc(fMap(varName), ValueMap3)

            If ObjectiveVariable = varName Then

                k3 = v
            End If

        End If

        ValueMap4.Add ValueMap3(varName) + v * StepValue, varName

    Next
    'Constant Set-----------------------------start
    For constCnt = 1 To Constants.Columns.Count

        ValueMap4.Add CDbl(Constants.Cells(2, constCnt)), Constants.Cells(1, constCnt)
    Next
    'Constant Set-----------------------------end
    
    Dim k4 As Double
    k4 = XCalc(fMap(ObjectiveVariable), ValueMap3)
    
    RK4Calc = ValueMap(ObjectiveVariable) + (k1 + 2 * k2 + 2 * k3 + k4) / 6 * StepValue
  

End Function
