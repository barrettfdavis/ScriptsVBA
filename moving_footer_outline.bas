Attribute VB_Name = "Module1"
Sub moving_footer_outline()

Dim osld As Slide
Dim oshp As Shape

Dim str_pre As String
Dim str_post As String
Dim section_dummy As String

If ActivePresentation.SectionProperties.Count > 0 Then

    For Each osld In ActivePresentation.Slides
        
        With ActivePresentation.SectionProperties
            
            str_pre = ""
            str_post = ""
            section_dummy = "Default Section"

            For x = 1 To .Count
            
                If StrComp(section_dummy, .Name(x)) Then
                    section_dummy = .Name(x)
                    
                    If x < osld.sectionIndex Then
                        str_pre = str_pre & .Name(x) & " - "
                    ElseIf x > osld.sectionIndex Then
                        str_post = str_post & " - " & .Name(x)
                    End If
                    
                End If
            Next x
            
        End With
    

        For Each oshp In osld.Shapes
        
            If oshp.Type = msoPlaceholder Then
                If oshp.PlaceholderFormat.Type = ppPlaceholderFooter Then _
                
                section = ActivePresentation.SectionProperties.Name(osld.sectionIndex)
                oshp.TextFrame.TextRange = str_pre & section & str_post
                
                oshp.TextFrame.TextRange.Font.Bold = msoFalse
                oshp.TextFrame.TextRange.Font.Color = RGB(150, 150, 150)
                
                li = Len(str_pre)
                ri = Len(section) + 1
                oshp.TextFrame.TextRange.Characters(li, ri).Font.Bold = msoTrue
                oshp.TextFrame.TextRange.Characters(li, ri).Font.Color = RGB(0, 0, 0)
                
                oshp.TextFrame.WordWrap = msoFalse
                oshp.TextFrame.AutoSize = ppAutoSizeShapeToFitText
                
                End If
            End If
            
        Next oshp
        
        If (osld.SlideIndex > 1) And (osld.SlideIndex < ActivePresentation.Slides.Count) Then
            osld.HeadersFooters.footer.Visible = True
        Else
            osld.HeadersFooters.footer.Visible = False
        End If
        
    Next osld
    
End If

End Sub
