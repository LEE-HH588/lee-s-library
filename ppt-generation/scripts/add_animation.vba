' PowerPoint 动画批量添加宏
' 用法：在 PowerPoint 中按 Alt+F11 打开 VBA 编辑器，插入模块后运行
' 或按 Alt+F8 打开宏对话框运行

Sub AddFadeInAnimations()
    ' 为所有幻灯片的所有形状添加淡入动画
    Dim slide As slide
    Dim shape As shape
    
    For Each slide In ActivePresentation.Slides
        For Each shape In slide.Shapes
            shape.AnimationSettings.EntryEffect = ppEffectFade
            shape.AnimationSettings.Animate = msoTrue
            shape.AnimationSettings.TextUnitEffect = ppAnimateByParagraph
        Next shape
    Next slide
    
    MsgBox "淡入动画添加完成！共处理 " & ActivePresentation.Slides.Count & " 页幻灯片"
End Sub

Sub AddFlyInFromLeft()
    ' 从左侧飞入动画
    Dim slide As slide
    Dim shape As shape
    
    For Each slide In ActivePresentation.Slides
        For Each shape In slide.Shapes
            shape.AnimationSettings.EntryEffect = ppEffectFlyFromLeft
            shape.AnimationSettings.Animate = msoTrue
        Next shape
    Next slide
    
    MsgBox "飞入动画添加完成！"
End Sub

Sub AddZoomAnimation()
    ' 缩放动画
    Dim slide As slide
    Dim shape As shape
    
    For Each slide In ActivePresentation.Slides
        For Each shape In slide.Shapes
            shape.AnimationSettings.EntryEffect = ppEffectZoomIn
            shape.AnimationSettings.Animate = msoTrue
        Next shape
    Next slide
    
    MsgBox "缩放动画添加完成！"
End Sub

Sub AddStaggeredAnimations()
    ' 逐项动画（标题先出现，内容后出现）
    Dim slide As slide
    Dim shape As shape
    Dim idx As Integer
    
    For Each slide In ActivePresentation.Slides
        idx = 0
        For Each shape In slide.Shapes
            idx = idx + 1
            shape.AnimationSettings.Animate = msoTrue
            ' 标题用淡入，其他用飞入
            If idx = 1 Then
                shape.AnimationSettings.EntryEffect = ppEffectFade
            Else
                shape.AnimationSettings.EntryEffect = ppEffectFlyFromLeft
            End If
        Next shape
    Next slide
    
    MsgBox "逐项动画添加完成！"
End Sub

Sub AddTransitionEffects()
    ' 为所有幻灯片添加页面切换效果
    Dim slide As slide
    Dim idx As Integer
    
    idx = 1
    For Each slide In ActivePresentation.Slides
        ' 交替使用不同的切换效果
        Select Case (idx Mod 5)
            Case 0: slide.SlideShowTransition.EntryEffect = ppEffectFade
            Case 1: slide.SlideShowTransition.EntryEffect = ppEffectPushLeft
            Case 2: slide.SlideShowTransition.EntryEffect = ppEffectWipeLeft
            Case 3: slide.SlideShowTransition.EntryEffect = ppEffectCoverLeft
            Case 4: slide.SlideShowTransition.EntryEffect = ppEffectUncoverLeft
        End Select
        slide.SlideShowTransition.Duration = 1.0
        idx = idx + 1
    Next slide
    
    MsgBox "切换效果添加完成！"
End Sub

Sub RemoveAllAnimations()
    ' 移除所有动画
    Dim slide As slide
    Dim shape As shape
    
    For Each slide In ActivePresentation.Slides
        For Each shape In slide.Shapes
            shape.AnimationSettings.Animate = msoFalse
        Next shape
    Next slide
    
    MsgBox "所有动画已移除"
End Sub
