﻿; ***************************************************************************************
; 示例9: 样式设置
; 说明: 演示如何设置形状的填充色、线条、字体样式、旋转等
; ***************************************************************************************
XIncludeFile "..\PbPPT.pb"

Define outputDir.s = "output\"
If FileSize(outputDir) = -1
  CreateDirectory(outputDir)
EndIf

PbPPT_Init()

If PbPPT_Create()
  Define slideIdx.i = PbPPT_AddSlide(0)

  ; === 填充样式演示 ===
  ; 纯色填充 - 红色矩形
  Define s1.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstRectangle$, PbPPT_InchesToEmu(0.3), PbPPT_InchesToEmu(0.5), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(1.2))
  PbPPT_SetFillSolid(s1, "FF0000")
  PbPPT_SetShapeText(s1, "纯色填充")
  PbPPT_SetFont(s1, 1, 1, "微软雅黑", 11, #True, #False, #False, "FFFFFF")

  ; 纯色填充 - 绿色圆形
  Define s2.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstEllipse$, PbPPT_InchesToEmu(2.6), PbPPT_InchesToEmu(0.5), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(1.2))
  PbPPT_SetFillSolid(s2, "00B050")
  PbPPT_SetShapeText(s2, "纯色填充")
  PbPPT_SetFont(s2, 1, 1, "微软雅黑", 11, #True, #False, #False, "FFFFFF")

  ; 无填充 - 仅线条
  Define s3.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstRectangle$, PbPPT_InchesToEmu(4.9), PbPPT_InchesToEmu(0.5), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(1.2))
  PbPPT_SetFillNone(s3)
  PbPPT_SetLineFormat(s3, PbPPT_PtToEmu(2), "7030A0")
  PbPPT_SetShapeText(s3, "无填充+线条")
  PbPPT_SetFont(s3, 1, 1, "微软雅黑", 11, #True, #False, #False, "7030A0")

  ; === 线条样式演示 ===
  ; 粗线条
  Define s4.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstRectangle$, PbPPT_InchesToEmu(7.2), PbPPT_InchesToEmu(0.5), PbPPT_InchesToEmu(2.3), PbPPT_InchesToEmu(1.2))
  PbPPT_SetFillSolid(s4, "FFC000")
  PbPPT_SetLineFormat(s4, PbPPT_PtToEmu(3), "BF8F00")
  PbPPT_SetShapeText(s4, "粗线条")
  PbPPT_SetFont(s4, 1, 1, "微软雅黑", 11, #True, #False, #False, "000000")

  ; === 旋转演示 ===
  Define s5.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstRectangle$, PbPPT_InchesToEmu(0.3), PbPPT_InchesToEmu(2.2), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(1.2))
  PbPPT_SetFillSolid(s5, "2E75B6")
  PbPPT_SetShapeRotation(s5, 15)
  PbPPT_SetShapeText(s5, "旋转15°")
  PbPPT_SetFont(s5, 1, 1, "微软雅黑", 11, #True, #False, #False, "FFFFFF")

  Define s6.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstRectangle$, PbPPT_InchesToEmu(2.6), PbPPT_InchesToEmu(2.2), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(1.2))
  PbPPT_SetFillSolid(s6, "2E75B6")
  PbPPT_SetShapeRotation(s6, -15)
  PbPPT_SetShapeText(s6, "旋转-15°")
  PbPPT_SetFont(s6, 1, 1, "微软雅黑", 11, #True, #False, #False, "FFFFFF")

  Define s7.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstRectangle$, PbPPT_InchesToEmu(4.9), PbPPT_InchesToEmu(2.2), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(1.2))
  PbPPT_SetFillSolid(s7, "2E75B6")
  PbPPT_SetShapeRotation(s7, 45)
  PbPPT_SetShapeText(s7, "旋转45°")
  PbPPT_SetFont(s7, 1, 1, "微软雅黑", 11, #True, #False, #False, "FFFFFF")

  ; === 连接符演示 ===
  Define connId.i = PbPPT_AddConnector(slideIdx, PbPPT_InchesToEmu(7.5), PbPPT_InchesToEmu(2.5), PbPPT_InchesToEmu(9), PbPPT_InchesToEmu(3.5))
  PbPPT_SetLineFormat(connId, PbPPT_PtToEmu(2), "FF0000")

  ; === 字体样式演示 ===
  ; 粗体文本
  Define tb1.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(0.3), PbPPT_InchesToEmu(4), PbPPT_InchesToEmu(4.5), PbPPT_InchesToEmu(0.6))
  PbPPT_SetShapeText(tb1, "粗体文本示例")
  PbPPT_SetFont(tb1, 1, 1, "微软雅黑", 18, #True, #False, #False, "000000")

  ; 斜体文本
  Define tb2.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(0.3), PbPPT_InchesToEmu(4.7), PbPPT_InchesToEmu(4.5), PbPPT_InchesToEmu(0.6))
  PbPPT_SetShapeText(tb2, "斜体文本示例")
  PbPPT_SetFont(tb2, 1, 1, "微软雅黑", 18, #False, #True, #False, "000000")

  ; 彩色文本
  Define tb3.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(0.3), PbPPT_InchesToEmu(5.4), PbPPT_InchesToEmu(4.5), PbPPT_InchesToEmu(0.6))
  PbPPT_SetShapeText(tb3, "红色粗斜体文本")
  PbPPT_SetFont(tb3, 1, 1, "微软雅黑", 18, #True, #True, #False, "FF0000")

  ; 大号文本
  Define tb4.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(5), PbPPT_InchesToEmu(4), PbPPT_InchesToEmu(4.5), PbPPT_InchesToEmu(2))
  PbPPT_SetShapeText(tb4, "大号文本")
  PbPPT_SetFont(tb4, 1, 1, "微软雅黑", 36, #True, #False, #False, "2E75B6")

  Define filename.s = outputDir + "example9_styles.pptx"
  If PbPPT_Save(filename)
    MessageRequester("示例9 - 成功", "样式演示文稿已创建：" + #LF$ + filename)
  Else
    MessageRequester("示例9 - 错误", "保存文件失败！")
  EndIf

  PbPPT_Close()
EndIf
