Attribute VB_Name = "SetShapeSizePositionHorizontal"
Sub SetShapeSizePositionHorizontal()

    Dim mSlide As Slide
    Dim mShape As Shape
    Dim mShapeCount As Single
    Const mSlideShowShapeCount = 3 '界面完整显示Shape的个数
    
    
    On Error Resume Next
    Err.Clear
    
    Debug.Print "ActiveWindow Width：" & ActiveWindow.Width
    Debug.Print "ActiveWindow Height：" & ActiveWindow.Height
    Debug.Print "ActivePresentation Width：" & ActivePresentation.PageSetup.SlideWidth
    Debug.Print "ActivePresentation Height：" & ActivePresentation.PageSetup.SlideHeight
    Debug.Print "------------------------------------------------------"
    
    Set mSlide = ActiveWindow.View.Slide
    Debug.Print "Active SlideID：" & mSlide.SlideID
    Debug.Print "Active SlideIndex：" & mSlide.SlideIndex
    Debug.Print "Active SlideNumber：" & mSlide.SlideNumber
    
    
    Debug.Print "Active Selection SlideRange.Count：" & ActiveWindow.Selection.SlideRange.Count
    Debug.Print "Active Selection ShapeRange.Count：" & ActiveWindow.Selection.ShapeRange.Count
    Debug.Print "------------------------------------------------------"
    
    If Err <> 0 Then
        MsgBox "Looks like no shape is selected!", vbCritical
        Exit Sub
    End If
    
    
     
    '给每一个Shape设置宽高、位置
    mShapeCount = ActiveWindow.Selection.ShapeRange.Count
    For i = 1 To mShapeCount Step 1
        Set mShape = ActiveWindow.Selection.ShapeRange(i)
        Debug.Print "Shape " & i & " Id：" & mShape.Id & " Name：" & mShape.Name & " Type：" & mShape.Type & " Visible：" & mShape.Visible & " LockAspectRatio：" & mShape.LockAspectRatio
        
        mShape.LockAspectRatio = msoFalse
        
        '设置宽高
        Debug.Print "Shape " & i & " Before Set Width：" & mShape.Width
        Debug.Print "Shape " & i & " Before Set Height：" & mShape.Height
        
        mShape.Width = ActivePresentation.PageSetup.SlideWidth / mSlideShowShapeCount
        mShape.Height = mShape.Width * 3 / 4 '16:9
        'mShape.Height = mShape.Width * 9 / 16 '16:9
        'mShape.Height = mShape.Width * 10 / 16 '16:4
                        
        Debug.Print "Shape " & i & " After Set Width：" & mShape.Width
        Debug.Print "Shape " & i & " After Set Height：" & mShape.Height
        
        
        '设置位置
        Debug.Print "Shape " & i & " Before Set Left：" & mShape.Left
        Debug.Print "Shape " & i & " Before Set Top：" & mShape.Top
        
        mShape.Left = ActivePresentation.PageSetup.SlideWidth - mShape.Width
        mShape.Top = ActivePresentation.PageSetup.SlideHeight - mShape.Height
        
        Debug.Print "Shape " & i & " After Set Left：" & mShape.Left
        Debug.Print "Shape " & i & " After Set Top：" & mShape.Top
        
        Debug.Print "------------------------------------------------------"
        
    Next
     
End Sub







