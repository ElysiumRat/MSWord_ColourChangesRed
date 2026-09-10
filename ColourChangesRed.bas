Attribute VB_Name = "ColourChangesRed"
Sub ColourChangesRed()

    Dim rev As Revision
    Dim rng As Range
    
    ' Turn off screen updating for speed
    Application.ScreenUpdating = False
    
    For Each rev In ActiveDocument.Revisions
        
        Select Case rev.Type
        
            ' Added text
            Case wdRevisionInsert
                Set rng = rev.Range
                rng.Font.Color = wdColorRed
            
            ' Moved text (new location only)
            Case wdRevisionMovedTo
                Set rng = rev.Range
                rng.Font.Color = wdColorRed
            
            ' Ignore moved-from text and all other revision types
            Case Else
                ' Do nothing
                
        End Select
        
    Next rev
    
    Application.ScreenUpdating = True
    
    MsgBox "Added and moved-to text has been coloured red.", vbInformation

End Sub
