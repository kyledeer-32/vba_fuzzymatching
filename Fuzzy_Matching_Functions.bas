Attribute VB_Name = "Fuzzy_Matching_Functions"
Public Function LevD(ByRef string1 As String, ByRef string2 As String) As Long

    string1 = Trim(string1)
    string2 = Trim(string2)
    
    string1 = LCase(string1)
    string2 = LCase(string2)
    
    Dim count_S1 As Long
    Dim count_S2 As Long
    
    count_S1 = Len(string1)
    count_S2 = Len(string2)
    
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
            CharCol1.Add Mid(string1, x, 1)
            Next x
        
        For y = 1 To count_S2:
            CharCol2.Add Mid(string2, y, 1)
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

Public Function Fuzzy_Match(ByRef targetS As String, ByRef MatchRange As Range, Threshold As Double) As String

    'Check to ensure user input a valid threshold
    If Threshold < 0 Or Threshold > 1 Then
        MsgBox "Input a number for Threshold between 0 to 1" & (Chr(13) & Chr(10)) & "For example: if you want the function to return matches that have a minimum string similarity of 75% and above then you would input .75", 0, "Threshold Error"
        Exit Function
    End If

    Dim MatchArray() As Variant
    MatchArray = MatchRange.Value2
        
    'Must add "Microsoft Scripting Runtime" library for dictionaries
    Dim EditDistance_Dict As Object
    Set EditDistance_Dict = New Scripting.Dictionary
    
    Dim S2 As String
    Dim EditD As Integer
    Dim tcheck As Double
           
    Dim x As Long
    Dim y As Long
   
    For x = LBound(MatchArray, 1) To UBound(MatchArray, 1)
        For y = LBound(MatchArray, 2) To UBound(MatchArray, 2)
            If EditDistance_Dict.Exists(MatchArray(x, y)) Then
                GoTo Skip_Add
            End If

            S2 = MatchArray(x, y)
            
            'If string contains spaces, i.e., more than one word, then call function "bestword()" to partition string to find closest matching substring
            If InStr(S2, " ") > 0 Then
                
                Dim bw As String
                bw = bestword(targetS, S2)
                
                'If threshold isn't met, then string isn't added to dictionary
                tcheck = String_Similarity(targetS, bw)
                
                If tcheck >= Threshold Then
                    EditD = LevD(targetS, bw)
                    EditDistance_Dict.Add MatchArray(x, y), EditD
                End If
                        
            Else
                
                tcheck = String_Similarity(targetS, bw)
                
                'If threshold isn't met, then string isn't added to dictionary
                If tcheck >= Threshold Then
                    EditD = LevD(targetS, S2)
                    EditDistance_Dict.Add MatchArray(x, y), EditD
                End If
                
            End If
        
Skip_Add:

        Next y
    Next x
   
    'If count of dictionary is 0, then no strings in the array met the threshold
    If EditDistance_Dict.count = 0 Then
        Fuzzy_Match = "No Match Found"
        MsgBox "You can retry by lowering the Threshold", vbOKOnly, "NO MATCH FOUND"
        Exit Function
    End If
   
    Dim key As Variant
    Dim closest_match As String
    Dim bestED As Long
    bestED = Application.Min(EditDistance_Dict.Items)
    
        For Each key In EditDistance_Dict.Keys
            If EditDistance_Dict(key) = bestED Then
                Fuzzy_Match = key
            End If
        Next key
       
End Function
 
'Breaks sentences into substrings, then evaluates and returns the substring with closest match to target string

Public Function bestword(ByRef string1 As String, ByRef string2 As String) As String

    'Partition a string into substrings by using a single space as the delimeter
    If InStr(string2, " ") > 0 Then
         Dim SplitStr() As String
         SplitStr = Split(string2)
         
         Dim str As String
         Dim n As Long
         
         Dim ed As Object
         Set ed = New Scripting.Dictionary
         
         Dim ss As Object
         Set ss = New Scripting.Dictionary
         
             For n = 0 To UBound(SplitStr)
                 str = SplitStr(n)
    
                 Dim editdist As Long
                 editdist = LevD(str, string1)
                 
                    If ed.Exists(str) Then
                        GoTo Skip_Duplicate
                    Else: ed.Add str, editdist
                    End If
                 
                 Dim ssim As Double
                 ssim = String_Similarity(str, string1)
                 ss.Add str, ssim
                 
Skip_Duplicate:

             Next n
             
             
         'Evaluates which substring is the closest match
         Dim bestdist As Long
         bestdist = Application.Min(ed.Items)
         
         Dim bestssim As Double
         bestssim = Application.Max(ss.Items)
         
         Dim x As Variant
         
         For Each x In ed.Keys
             If ed.Item(x) = bestdist And ss.Item(x) = bestssim Then
                 bestword = x
                 GoTo Done
             End If
         Next x
         
         For Each x In ed.Keys
             If ed.Item(x) = bestdist Then
                 bestword = x
             End If
         Next x
         
     End If
            
Done:

End Function
 
Public Function String_Similarity(ByRef Str1 As String, ByRef Str2 As String)

'Calculate string similarity % to determine strength of matching words
'Function Source: http://adamfortuno.com/index.php/2021/07/05/levenshtein-distance-and-distance-similarity-functions/?msclkid=aa598709a8b611ec8de2ac84df81a9da

    Dim Similarity As Long
    
    Dim LevenDist As Long
    LevenDist = LevD(Str1, Str2)
        
    String_Similarity = (100 - (LevenDist / WorksheetFunction.Max(Len(Str1), Len(Str2)) * 100)) / 100
        
End Function
