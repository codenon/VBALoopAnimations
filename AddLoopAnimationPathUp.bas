Attribute VB_Name = "AddLoopAnimationPathUp"
'添加垂直向上循环动画
Sub AddLoopAnimationPathUp()

    Dim mSlide As Slide
    Dim mShape As Shape
    Dim mShapeCount As Single
    Dim mSequence As Sequence
    Dim mEffect As Effect
    
    Const mMotionEffectDuration As Single = 6 'alter this to suit
    Const mSlideShowShapeCount = 1 '界面完整显示Shape的个数，Shape总数需 >= mSlideShowShapeCount + 2
    Const mDelayFactor = mSlideShowShapeCount + 1 '计算启动延时和动画等待延时用
    
    
    On Error Resume Next
    Err.Clear
    
    Set mSlide = ActiveWindow.View.Slide
    Set mSequence = mSlide.TimeLine.MainSequence
    
    Debug.Print "Active Selection ShapeRange.Count：" & ActiveWindow.Selection.ShapeRange.Count
    If Err <> 0 Then
        MsgBox "Looks like no shape is selected!", vbCritical
        Exit Sub
    End If
    
     
    '给每一个Shape添加路径动画和延时等待循环结束动画
    mShapeCount = ActiveWindow.Selection.ShapeRange.Count
    For i = 1 To mShapeCount Step 1
        Set mShape = ActiveWindow.Selection.ShapeRange(i)
        Debug.Print "Shape " & i & " Id：" & mShape.Id & " Name：" & mShape.Name & " Type：" & mShape.Type & " Visible：" & mShape.Visible
        
        '删除选中Shape原有动画
        Dim effectFirst As Effect
        Set effectFirst = mSequence.FindFirstAnimationFor(mShape)
        Do While Not effectFirst Is Nothing
            Debug.Print "Delete MainSequence Effect " & " Index：" & effectFirst.Index & "；Shape Id：" & mShape.Id & "；Shape Name：" & mShape.Name
            effectFirst.Delete
            Set effectFirst = mSequence.FindFirstAnimationFor(mShape)
        Loop
        
        ' 添加一个msoAnimEffectCustom动画
        Set mEffect = mSequence.AddEffect(Shape:=mShape, effectId:=msoAnimEffectCustom, Trigger:=msoAnimTriggerWithPrevious, Index:=-1)
       
        With mEffect
            '设置Effect 属性
            .Exit = msoFalse
            .Timing.SmoothStart = msoFalse
            .Timing.SmoothEnd = msoFalse
            .Timing.RewindAtEnd = msoTrue
            .Timing.RepeatCount = 1000 '-2147483648 没有找到设置‘直到下一次单击’或‘直到幻灯片末尾’的设置方法
            .Timing.TriggerType = msoAnimTriggerWithPrevious
            .Timing.TriggerDelayTime = (i - 1) * (mMotionEffectDuration / mDelayFactor)
            
            
            '添加msoAnimTypeMotion垂直向上循环动画
            .Behaviors.Add(msoAnimTypeMotion).MotionEffect.Path = "M 0 1 L 0 -1 E"
            .Behaviors(.Behaviors.Count).Timing.TriggerDelayTime = 0
            .Behaviors(.Behaviors.Count).Timing.Duration = mMotionEffectDuration
            
            '添加延时等待循环结束动画
            .Behaviors.Add(msoAnimTypeSet).SetEffect.Property = msoAnimVisibility
            .Behaviors(.Behaviors.Count).Timing.TriggerDelayTime = .Timing.Duration
            .Behaviors(.Behaviors.Count).Timing.Duration = (mShapeCount - mDelayFactor) * (mMotionEffectDuration / mDelayFactor)
            .Behaviors(.Behaviors.Count).SetEffect.To = 1 ' aShape not hidden use 0 for hidden while delayed
               
        End With
    Next
End Sub
