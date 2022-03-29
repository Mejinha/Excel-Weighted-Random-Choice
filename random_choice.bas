Attribute VB_Name = "Módulo1"
Function CHOOSE(labels, Optional weights)

    Dim weights_array
    Dim sum_weights As Double
    Dim normalized_weights() As Variant
    Dim i As Integer
    Randomize

    If IsMissing(weights) = True Then
    
        ReDim weights_array(WorksheetFunction.CountA(labels) - 1)
        
        
        For w = 0 To WorksheetFunction.CountA(labels) - 1
            weights_array(w) = 1 / WorksheetFunction.CountA(labels)
        Next w
        
        sum_weights = 1

    Else
        weights_array = weights.Value
        sum_weights = WorksheetFunction.Sum(weights)
    
    End If


    i = 0
    
    For Each w In weights_array
        ReDim Preserve normalized_weights(i + 1)
        normalized_weights(i) = w / sum_weights
        i = i + 1
    
    Next w
    
    Dim choice As Double
    Dim sum_past As Double
    Dim rand_value As Double

    choice = 0
    sum_past = 0
    
    rand_value = Rnd
    
    For Each w In normalized_weights
    
        sum_past = sum_past + w
        
        If (rand_value <= sum_past) Then
            Debug.Print rand_value
            Exit For
            
        Else
            choice = choice + 1
            
        End If
         
    Next w
   
   CHOOSE = WorksheetFunction.Index(labels, choice + 1)

End Function
