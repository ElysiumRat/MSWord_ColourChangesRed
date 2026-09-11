# MSWord_ColourChangesRed

A VBA macro for Word that changes the text colour of all active tracked changes to red. When submitting articles for publication to some journals (I've especially encountered this in Taylor & Francis journals), the R&R stage requires one shows exactly what was changed and it is sometimes specified that the authors should not use track changes to show what was changed. As such, this macro inspects all revisions in the active document and selectively applies a red font color (`wdColorRed`) specifically to insertions and text relocated to a new position.

This way, multiple authors can collaborate on the changes they wish to make, and once the final revisions have been agreed upon, the macro can be run and all relevant changes (additions and movements) will remain highlighted when changes are accepted. This overcomes the need to otherwise do so manually, or remembering to change one's font colour when adding text, saving time and effort and allowing for better focus on what matters: the content of revisions.

## Code Breakdown

### Initialization and Performance Optimization

The macro begins by defining the necessary object variables (`rev` as a `Revision` and `rng` as a `Range`). To maximize performance—especially in large documents containing thousands of changes—screen updating is temporarily disabled:

```vba
Dim rev As Revision
Dim rng As Range

' Turn off screen updating for speed
Application.ScreenUpdating = False
```

### Iterating and Filtering Revisions

The macro loops through every revision object in `ActiveDocument.Revisions` and uses a `Select Case` statement evaluated against `rev.Type` to determine the specific nature of the change:

```vba
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
```

 - `wdRevisionInsert`: Captures newly typed or inserted text, capturing the revision's range and updating its font color.
 - `wdRevisionMovedTo`: Captures text that has been moved into the document section from elsewhere, highlighting the destination text.
 - `Case Else`: Ignores deletions (`wdRevisionDelete`), formatting changes, and `wdRevisionMovedFrom` (the original origin point of moved text) to prevent redundant or confusing markup.

### Cleanup and Completion

Once the loop completes, screen updating is restored, and a confirmation message box notifies the user that the operation has finished:

```vba
Application.ScreenUpdating = True

MsgBox "Added and moved-to text has been coloured red.", vbInformation
```
