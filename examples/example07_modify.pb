; ***************************************************************************************
; 示例7: 修改并保存
; 说明: 演示如何打开已有文件、修改内容后另存为新文件
; ***************************************************************************************
XIncludeFile "..\PbPPT.pb"

Define outputDir.s = "output\"
If FileSize(outputDir) = -1
  CreateDirectory(outputDir)
EndIf

PbPPT_Init()

; 先创建一个基础文件
PbPPT_Create()
Define slide1.i = PbPPT_AddSlide(0)
Define tb1.i = PbPPT_AddTextbox(slide1, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(2))
PbPPT_SetShapeText(tb1, "这是原始文本内容")
PbPPT_SetTitle("原始标题")
PbPPT_Save(outputDir + "example7_original.pptx")
PbPPT_Close()

; 打开文件并修改
If PbPPT_Open(outputDir + "example7_original.pptx")
  ; 修改文档属性
  PbPPT_SetTitle("修改后的标题")
  PbPPT_SetAuthor("PbPPT修改测试")
  PbPPT_SetSubject("修改演示")

  ; 添加新的幻灯片
  Define newSlide.i = PbPPT_AddSlide(0)
  Define newTb.i = PbPPT_AddTextbox(newSlide, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(2))
  PbPPT_SetShapeText(newTb, "这是修改后新增的幻灯片")
  PbPPT_SetFont(newTb, 1, 1, "微软雅黑", 24, #True, #False, #False, "FF0000")

  ; 在新幻灯片上添加形状
  Define shapeId.i = PbPPT_AddShape(newSlide, #PbPPT_PrstRectangle$, PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(4.5), PbPPT_InchesToEmu(6), PbPPT_InchesToEmu(1.5))
  PbPPT_SetFillSolid(shapeId, "4472C4")
  PbPPT_SetShapeText(shapeId, "新增的形状")
  PbPPT_SetFont(shapeId, 1, 1, "微软雅黑", 18, #True, #False, #False, "FFFFFF")

  ; 另存为新文件
  Define filename.s = outputDir + "example7_modified.pptx"
  If PbPPT_Save(filename)
    MessageRequester("示例7 - 成功", "文件已修改并保存：" + #LF$ + filename + #LF$ + "原幻灯片数：1，修改后：" + Str(PbPPT_GetSlideCount()))
  Else
    MessageRequester("示例7 - 错误", "保存文件失败！")
  EndIf

  PbPPT_Close()
Else
  MessageRequester("示例7 - 错误", "打开文件失败！")
EndIf
