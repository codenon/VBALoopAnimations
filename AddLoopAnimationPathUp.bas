Attribute VB_Name = "AddLooAnimationpPathUp"
Sub AddLooAnimationpPathUp()

    'Debug.Print Application.Version
    'Debug.Print Application.ActivePresentation.Name
    'Debug.Print ActivePresentation.Name
    


    Dim mSlide As Slide
    Dim mShape As Shape
    Dim mShapeCount As Single
    Dim mSequence As Sequence
    Dim mSequenceEffectCount As Integer
    Dim mEffect As Effect
    Dim mBehaviors As AnimationBehaviors
    Dim mBehaviorCount As Integer
    Const mMotionEffectDuration As Single = 6 'alter this to suit
    Const mSlideShowShapeCount = 1 '界面完整显示Shape的个数，Shape总数需 >= mSlideShowShapeCount + 2
    Const mDelayCalFactor = mSlideShowShapeCount + 1 '计算启动延时和动画等待延时用
    
    
    
    On Error Resume Next
    Err.Clear
    
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
    
    
    '删除选中Shape原有动画
    Set mSequence = mSlide.TimeLine.MainSequence
    mSequenceEffectCount = mSequence.Count
    Debug.Print "Slide TimeLine MainSequence Effect Count：" & mSequenceEffectCount
    For i = mSequenceEffectCount To 1 Step -1
        
        'Debug.Print "Delete MainSequence Effect " & i & " Index：" & mSequence(i).Index & "； Id：" & mSequence(i).Shape.Id & "； Name：" & mSequence(i).Shape.Name
        'mSequence(i).Delete
        
    Next
    Debug.Print "------------------------------------------------------"
    
    '给每一个Shape添加路径动画和延时等待循环结束动画
    mShapeCount = ActiveWindow.Selection.ShapeRange.Count
    For i = 1 To mShapeCount Step 1
        Set mShape = ActiveWindow.Selection.ShapeRange(i)
        Debug.Print "Shape " & i & " Id：" & mShape.Id & " Name：" & mShape.Name & " Type：" & mShape.Type & " Visible：" & mShape.Visible
        Debug.Print "------------------------------------------------------"
        
        
        '删除选中Shape原有动画
        Dim effectFirst As Effect
        Set effectFirst = mSequence.FindFirstAnimationFor(mShape)
        Do While Not effectFirst Is Nothing
            Debug.Print "Delete MainSequence Effect " & " Index：" & effectFirst.Index & "；Shape Id：" & mShape.Id & "；Shape Name：" & mShape.Name
            effectFirst.Delete
            Set effectFirst = mSequence.FindFirstAnimationFor(mShape)
        Loop
        
        
        'mShape.Visible = msoTrue
        ' 添加一个msoAnimEffectAppear动画，不直接添加msoAnimEffectPathUp动画是因为动画开始前图片会显示
        Set mEffect = mSequence.AddEffect(Shape:=mShape, effectId:=msoAnimEffectAppear, Trigger:=msoAnimTriggerWithPrevious, Index:=-1)
       
        With mEffect
            '设置Effect 属性
            .Exit = msoFalse
            .Timing.SmoothStart = msoFalse
            .Timing.SmoothEnd = msoFalse
            .Timing.Duration = 0.001
            .Timing.RepeatCount = 1000 '-2147483648
            .Timing.TriggerType = msoAnimTriggerWithPrevious
            .Timing.TriggerDelayTime = (i - 1) * (mMotionEffectDuration / mDelayCalFactor)
            
            mBehaviorCount = .Behaviors.Count
            Debug.Print "Shape " & i & " AddEffect Behaviors Count Before：" & mBehaviorCount
            Debug.Print "Shape " & i & " AddEffect Timing Duration Before：" & .Timing.Duration
            Debug.Print "Shape " & i & " AddEffect Behavior " & (mBehaviorCount) & " Timing.TriggerDelayTime：" & .Behaviors(mBehaviorCount).Timing.TriggerDelayTime
            Debug.Print "Shape " & i & " AddEffect Behavior " & (mBehaviorCount) & " Timing.Duration：" & .Behaviors(mBehaviorCount).Timing.Duration
            
            
            '添加msoAnimEffectPathUp动画
            .Behaviors.Add(msoAnimTypeMotion).MotionEffect.Path = "M 0 1 L 0 -1 E"
            .Behaviors(mBehaviorCount + 1).Timing.TriggerDelayTime = 0
            .Behaviors(mBehaviorCount + 1).Timing.Duration = mMotionEffectDuration
            Debug.Print "Shape " & i & " AddEffect Behavior " & (mBehaviorCount + 1) & " Timing.TriggerDelayTime：" & .Behaviors(mBehaviorCount + 1).Timing.TriggerDelayTime
            Debug.Print "Shape " & i & " AddEffect Behavior " & (mBehaviorCount + 1) & " Timing.Duration：" & .Behaviors(mBehaviorCount + 1).Timing.Duration
            
            
            '添加延时等待循环结束动画
            .Behaviors.Add(msoAnimTypeSet).SetEffect.Property = msoAnimVisibility
            .Behaviors(mBehaviorCount + 2).Timing.TriggerDelayTime = .Timing.Duration
            '.Behaviors(mBehaviorCount + 2).Timing.Duration = mShapeCount * mMotionEffectDuration / 2 - mMotionEffectDuration
            .Behaviors(mBehaviorCount + 2).Timing.Duration = (mShapeCount - mDelayCalFactor) * (mMotionEffectDuration / mDelayCalFactor)
            .Behaviors(mBehaviorCount + 2).SetEffect.To = 1 ' aShape not hidden use 0 for hidden while delayed
            Debug.Print "Shape " & i & " AddEffect Behavior " & (mBehaviorCount + 2) & " Timing.TriggerDelayTime：" & .Behaviors(mBehaviorCount + 2).Timing.TriggerDelayTime
            Debug.Print "Shape " & i & " AddEffect Behavior " & (mBehaviorCount + 2) & " Timing.Duration：" & .Behaviors(mBehaviorCount + 2).Timing.Duration
            
            
            Debug.Print "Shape " & i & " AddEffect Behaviors Count after：" & .Behaviors.Count
            Debug.Print "Shape " & i & " AddEffect Timing Duration after：" & .Timing.Duration
            Debug.Print "------------------------------------------------------"
               
        End With
    Next
    
    Exit Sub
    
    Debug.Print "==================================================================="
     
    Set mSequence = mSlide.TimeLine.MainSequence
    mSequenceEffectCount = mSequence.Count
    Debug.Print "Slide TimeLine MainSequence Effect Count：" & mSequenceEffectCount
    Debug.Print "------------------------------------------------------"
    
    For i = 1 To mSequenceEffectCount Step 1
        Set mEffect = mSequence(i)
        With mEffect
            Debug.Print "Slide TimeLine MainSequence Effect " & i & " Index：" & .Index
            Debug.Print "Slide TimeLine MainSequence Effect " & i & " Shape.Name：" & .Shape.Name
            Debug.Print "Slide TimeLine MainSequence Effect " & i & " DisplayName：" & .DisplayName
            Debug.Print "Slide TimeLine MainSequence Effect " & i & " EffectType：" & .EffectType
            Debug.Print "Slide TimeLine MainSequence Effect " & i & " Timing.RepeatCount：" & .Timing.RepeatCount
            Debug.Print "Slide TimeLine MainSequence Effect " & i & " Timing.Duration：" & .Timing.Duration
            Debug.Print "Slide TimeLine MainSequence Effect " & i & " Behaviors.Count：" & .Behaviors.Count
            
           
            Debug.Print "------------------------------------------------------"
            
            For j = 1 To .Behaviors.Count Step 1
                Debug.Print "Slide TimeLine MainSequence Effect " & i & " Behavior " & j & " Type：" & .Behaviors(j).Type
                Debug.Print "Slide TimeLine MainSequence Effect " & i & " Behavior " & j & " Timing.TriggerDelayTime：" & .Behaviors(j).Timing.TriggerDelayTime
                Debug.Print "Slide TimeLine MainSequence Effect " & i & " Behavior " & j & " Timing.Duration：" & .Behaviors(j).Timing.Duration
            Next
            
            Debug.Print "------------------------------------------------------"
            
        End With
    Next

End Sub



