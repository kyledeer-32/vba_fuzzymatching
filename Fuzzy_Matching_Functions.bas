Attribute VB_Name = "Fuzzy_Matching_Functions"
Public Function LevD(ByRef String1 As String, ByRef String2 As String) As Long

    String1 = Trim(String1)
    String2 = Trim(String2)
    
    String1 = LCase(String1)
    String2 = LCase(String2)
    
    Dim count_S1 As Long
    Dim count_S2 As Long
    
    count_S1 = Len(String1)
    count_S2 = Len(String2)
    
    Dim LevMatrix() As Long
    ReDim LevMatrix(0 To count_S1, 0 To count_S2)
    
    Dim i As Long
    Dim j As Long
    
        For i = 0 To count_S1:
            LevMatrix(i, 0) = i
            Next i
        
        For j = 0 To count_S2:
            LevMatrix(0, j) = j
            Next j
    
    Dim CharCol1 As Object
    Set CharCol1 = New Collection
    
    Dim CharCol2 As Object
    Set CharCol2 = New Collection
    
    Dim x As Long
    Dim y As Long
    
        For x = 1 To count_S1:
            CharCol1.Add Mid(String1, x, 1)
            Next x
        
        For y = 1 To count_S2:
            CharCol2.Add Mid(String2, y, 1)
            Next y
            
    Dim Cost As Long
    
    Dim Calc1 As Long
    Dim Calc2 As Long
    Dim Calc3 As Long
    
        For i = 1 To count_S1:
            For j = 1 To count_S2:
                If CharCol1(i) <> CharCol2(j) Then
                    Cost = 1
                Else
                    Cost = 0
                End If
                              
                Calc1 = LevMatrix(i - 1, j) + 1
                Calc2 = LevMatrix(i, j - 1) + 1
                Calc3 = LevMatrix(i - 1, j - 1) + Cost

                LevMatrix(i, j) = WorksheetFunction.Min(Calc1, Calc2, Calc3)
                               
            Next j
        Next i
      
    LevD = LevMatrix(count_S1, count_S2)
    
End Function

Public Function Fuzzy_Match(ByRef targetS As String, ByRef MatchRange As Range) As String

    Dim MatchArray() As Variant
    MatchArray = MatchRange.Value2
        
    'Must add "Microsoft Scripting Runtime" library for dictionaries
    Dim EditDistance_Dict As Object
    Set EditDistance_Dict = New Scripting.Dictionary
    
    Dim S2 As String
    Dim EditD As Long
    
    Dim x As Long
    Dim y As Long
   
    For x = LBound(MatchArray, 1) To UBound(MatchArray, 1):
        For y = LBound(MatchArray, 2) To UBound(MatchArray, 2):
            If EditDistance_Dict.Exists(MatchArray(x, y)) Then
                GoTo Skip_Add
            End If
            

'Begin Change





'End change



Skip_Add:

        Next y
    Next x
   
   
    Dim key As Variant
    Dim Closest_Match As String
    Dim Distance As Long
    
    Distance = 1000

    For Each key In EditDistance_Dict.Keys
        If EditDistance_Dict(key) < Distance Then
            Closest_Match = key
            Distance = EditDistance_Dict(key)
        End If
    Next key

    Fuzzy_Match = Closest_Match

End Function


Public Function String_Similarity(ByRef Str1 As String, ByRef Str2 As String)

'Calculate string similarity % to determine strength of mathcing words
'Function Source: http://adamfortuno.com/index.php/2021/07/05/levenshtein-distance-and-distance-similarity-functions/?msclkid=aa598709a8b611ec8de2ac84df81a9da

    Dim Similarity As Long
    
    Dim LevenDist As Long
    LevenDist = LevD(Str1, Str2)
        
    String_Similarity = (100 - (LevenDist / WorksheetFunction.Max(Len(Str1), Len(Str2)) * 100)) / 100
        
End Function
