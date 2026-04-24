; ***************************************************************************************
; 示例3: 添加形状
; 说明: 演示如何添加各种自动形状（矩形、圆形、三角形等）并设置填充和线条
; ***************************************************************************************
XIncludeFile "..\PbPPT.pb"

Define outputDir.s = "output\"
If FileSize(outputDir) = -1
  CreateDirectory(outputDir)
EndIf

PbPPT_Init()

If PbPPT_Create()
  Define slideIdx.i = PbPPT_AddSlide(0)

  ; 添加矩形 - 蓝色填充
  Define rectId.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstRectangle$, PbPPT_InchesToEmu(0.5), PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2.5), PbPPT_InchesToEmu(1.5))
  PbPPT_SetFillSolid(rectId, "4472C4")
  PbPPT_SetLineFormat(rectId, PbPPT_PtToEmu(1), "2F528F")
  PbPPT_SetShapeText(rectId, "矩形")
  PbPPT_SetFont(rectId, 1, 1, "微软雅黑", 16, #True, #False, #False, "FFFFFF")

  ; 添加圆形 - 橙色填充
  Define ellipseId.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstEllipse$, PbPPT_InchesToEmu(3.5), PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2.5), PbPPT_InchesToEmu(1.5))
  PbPPT_SetFillSolid(ellipseId, "ED7D31")
  PbPPT_SetLineFormat(ellipseId, PbPPT_PtToEmu(1), "C55A11")
  PbPPT_SetShapeText(ellipseId, "圆形")
  PbPPT_SetFont(ellipseId, 1, 1, "微软雅黑", 16, #True, #False, #False, "FFFFFF")

  ; 添加三角形 - 绿色填充
  Define triId.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstTriangle$, PbPPT_InchesToEmu(6.5), PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2.5), PbPPT_InchesToEmu(1.5))
  PbPPT_SetFillSolid(triId, "70AD47")
  PbPPT_SetLineFormat(triId, PbPPT_PtToEmu(1), "548235")
  PbPPT_SetShapeText(triId, "三角形")
  PbPPT_SetFont(triId, 1, 1, "微软雅黑", 16, #True, #False, #False, "FFFFFF")

  ; 添加圆角矩形 - 紫色填充
  Define roundRectId.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstRoundRectangle$, PbPPT_InchesToEmu(0.5), PbPPT_InchesToEmu(3.5), PbPPT_InchesToEmu(2.5), PbPPT_InchesToEmu(1.5))
  PbPPT_SetFillSolid(roundRectId, "7030A0")
  PbPPT_SetLineFormat(roundRectId, PbPPT_PtToEmu(1), "5B2C8C")
  PbPPT_SetShapeText(roundRectId, "圆角矩形")
  PbPPT_SetFont(roundRectId, 1, 1, "微软雅黑", 14, #True, #False, #False, "FFFFFF")

  ; 添加菱形 - 红色填充
  Define diamondId.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstDiamond$, PbPPT_InchesToEmu(3.5), PbPPT_InchesToEmu(3.5), PbPPT_InchesToEmu(2.5), PbPPT_InchesToEmu(1.5))
  PbPPT_SetFillSolid(diamondId, "FF0000")
  PbPPT_SetLineFormat(diamondId, PbPPT_PtToEmu(1), "C00000")
  PbPPT_SetShapeText(diamondId, "菱形")
  PbPPT_SetFont(diamondId, 1, 1, "微软雅黑", 14, #True, #False, #False, "FFFFFF")

  ; 添加五角星 - 金色填充
  Define starId.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstStar5$, PbPPT_InchesToEmu(6.5), PbPPT_InchesToEmu(3.5), PbPPT_InchesToEmu(2.5), PbPPT_InchesToEmu(1.5))
  PbPPT_SetFillSolid(starId, "FFC000")
  PbPPT_SetLineFormat(starId, PbPPT_PtToEmu(1), "BF8F00")
  PbPPT_SetShapeText(starId, "五角星")
  PbPPT_SetFont(starId, 1, 1, "微软雅黑", 14, #True, #False, #False, "FFFFFF")

  ; 添加带旋转的箭头
  Define arrowId.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstArrow$, PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(6), PbPPT_InchesToEmu(5), PbPPT_InchesToEmu(1))
  PbPPT_SetFillSolid(arrowId, "00B0F0")
  PbPPT_SetShapeRotation(arrowId, 15)
  PbPPT_SetShapeText(arrowId, "旋转箭头")
  PbPPT_SetFont(arrowId, 1, 1, "微软雅黑", 16, #True, #False, #False, "FFFFFF")

  Define filename.s = outputDir + "example3_shapes.pptx"
  If PbPPT_Save(filename)
    MessageRequester("示例3 - 成功", "形状演示文稿已创建：" + #LF$ + filename)
  Else
    MessageRequester("示例3 - 错误", "保存文件失败！")
  EndIf

  PbPPT_Close()
EndIf
