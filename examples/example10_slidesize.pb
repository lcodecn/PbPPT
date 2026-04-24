; ***************************************************************************************
; 示例10: 幻灯片尺寸
; 说明: 演示如何设置不同的幻灯片尺寸（标准4:3和宽屏16:9）
; ***************************************************************************************
XIncludeFile "..\PbPPT.pb"

Define outputDir.s = "output\"
If FileSize(outputDir) = -1
  CreateDirectory(outputDir)
EndIf

PbPPT_Init()

; === 创建标准4:3尺寸的演示文稿 ===
PbPPT_Create()
; 默认尺寸为标准4:3（10x7.5英寸）
Define slide1.i = PbPPT_AddSlide(0)
Define tb1.i = PbPPT_AddTextbox(slide1, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(2))
PbPPT_SetShapeText(tb1, "标准4:3尺寸（10×7.5英寸）")
PbPPT_SetFont(tb1, 1, 1, "微软雅黑", 28, #True, #False, #False, "2E75B6")
PbPPT_SetParagraphAlignment(tb1, 1, #PbPPT_AlignCenter)

Define info1.i = PbPPT_AddTextbox(slide1, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(4.5), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(2))
PbPPT_SetShapeText(info1, "宽度：" + StrF(PbPPT_EmuToInches(PbPPT_GetSlideWidth()), 1) + " 英寸" + #LF$ + "高度：" + StrF(PbPPT_EmuToInches(PbPPT_GetSlideHeight()), 1) + " 英寸" + #LF$ + "宽EMU：" + Str(PbPPT_GetSlideWidth()) + #LF$ + "高EMU：" + Str(PbPPT_GetSlideHeight()))
PbPPT_SetParagraphAlignment(info1, 1, #PbPPT_AlignCenter)

PbPPT_Save(outputDir + "example10_standard_4x3.pptx")
PbPPT_Close()

; === 创建宽屏16:9尺寸的演示文稿 ===
PbPPT_Create()
; 设置为宽屏16:9（13.333x7.5英寸）
PbPPT_SetSlideWidth(#PbPPT_WIDESCREEN_SLIDE_WIDTH)
PbPPT_SetSlideHeight(#PbPPT_WIDESCREEN_SLIDE_HEIGHT)

Define slide2.i = PbPPT_AddSlide(0)
Define tb2.i = PbPPT_AddTextbox(slide2, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(11), PbPPT_InchesToEmu(2))
PbPPT_SetShapeText(tb2, "宽屏16:9尺寸（13.33×7.5英寸）")
PbPPT_SetFont(tb2, 1, 1, "微软雅黑", 28, #True, #False, #False, "2E75B6")
PbPPT_SetParagraphAlignment(tb2, 1, #PbPPT_AlignCenter)

Define info2.i = PbPPT_AddTextbox(slide2, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(4.5), PbPPT_InchesToEmu(11), PbPPT_InchesToEmu(2))
PbPPT_SetShapeText(info2, "宽度：" + StrF(PbPPT_EmuToInches(PbPPT_GetSlideWidth()), 2) + " 英寸" + #LF$ + "高度：" + StrF(PbPPT_EmuToInches(PbPPT_GetSlideHeight()), 1) + " 英寸" + #LF$ + "宽EMU：" + Str(PbPPT_GetSlideWidth()) + #LF$ + "高EMU：" + Str(PbPPT_GetSlideHeight()))
PbPPT_SetParagraphAlignment(info2, 1, #PbPPT_AlignCenter)

PbPPT_Save(outputDir + "example10_widescreen_16x9.pptx")
PbPPT_Close()

; === 创建自定义尺寸的演示文稿 ===
PbPPT_Create()
; A4纸张尺寸（横向）
PbPPT_SetSlideWidth(PbPPT_CmToEmu(29.7))
PbPPT_SetSlideHeight(PbPPT_CmToEmu(21.0))

Define slide3.i = PbPPT_AddSlide(0)
Define tb3.i = PbPPT_AddTextbox(slide3, PbPPT_CmToEmu(2), PbPPT_CmToEmu(5), PbPPT_CmToEmu(25), PbPPT_CmToEmu(3))
PbPPT_SetShapeText(tb3, "A4横向自定义尺寸（29.7×21厘米）")
PbPPT_SetFont(tb3, 1, 1, "微软雅黑", 24, #True, #False, #False, "2E75B6")
PbPPT_SetParagraphAlignment(tb3, 1, #PbPPT_AlignCenter)

PbPPT_Save(outputDir + "example10_custom_a4.pptx")
PbPPT_Close()

MessageRequester("示例10 - 成功", "三种尺寸的演示文稿已创建：" + #LF$ + "1. 标准4:3" + #LF$ + "2. 宽屏16:9" + #LF$ + "3. A4自定义")
