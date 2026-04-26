﻿; ***************************************************************************************
; 示例4: 添加图片
; 说明: 演示如何向幻灯片添加图片（支持PNG/JPG/BMP等格式）
; ***************************************************************************************
XIncludeFile "..\PbPPT.pb"

Define outputDir.s = "output\"
If FileSize(outputDir) = -1
  CreateDirectory(outputDir)
EndIf

PbPPT_Init()

If PbPPT_Create()
  Define slideIdx.i = PbPPT_AddSlide(0)

  ; 添加标题
  Define titleId.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(0.3), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(1))
  PbPPT_SetShapeText(titleId, "图片示例")
  PbPPT_SetFont(titleId, 1, 1, "微软雅黑", 28, #True, #False, #False, "2E75B6")
  PbPPT_SetParagraphAlignment(titleId, 1, #PbPPT_AlignCenter)

  ; 添加图片（使用python-pptx测试目录中的图片）
  Define imagePath.s = "test.jpg"
  If FileSize(imagePath) > 0
    ; 指定位置和大小添加图片
    Define picId.i = PbPPT_AddPicture(slideIdx, imagePath, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(3.5), PbPPT_InchesToEmu(3.5))

    ; 添加图片说明
    Define descId.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(5.8), PbPPT_InchesToEmu(3.5), PbPPT_InchesToEmu(0.5))
    PbPPT_SetShapeText(descId, "Python Powered（指定大小）")
    PbPPT_SetParagraphAlignment(descId, 1, #PbPPT_AlignCenter)
  Else
    ; 图片不存在时添加提示文本
    Define noteId.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(2))
    PbPPT_SetShapeText(noteId, "提示：请确保python-pptx目录中存在测试图片文件")
  EndIf

  ; 添加第二张图片（自动尺寸）
  Define imagePath2.s = "test.jpg"
  If FileSize(imagePath2) > 0
    Define picId2.i = PbPPT_AddPicture(slideIdx, imagePath2, PbPPT_InchesToEmu(5.5), PbPPT_InchesToEmu(2), 0, 0)

    Define descId2.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(5), PbPPT_InchesToEmu(5.8), PbPPT_InchesToEmu(4), PbPPT_InchesToEmu(0.5))
    PbPPT_SetShapeText(descId2, "Monty Truth（自动大小）")
    PbPPT_SetParagraphAlignment(descId2, 1, #PbPPT_AlignCenter)
  EndIf

  Define filename.s = outputDir + "example4_picture.pptx"
  If PbPPT_Save(filename)
    MessageRequester("示例4 - 成功", "图片演示文稿已创建：" + #LF$ + filename)
  Else
    MessageRequester("示例4 - 错误", "保存文件失败！")
  EndIf

  PbPPT_Close()
EndIf

; IDE Options = PureBasic 6.40 (Windows - x86)
; CursorPosition = 48
; FirstLine = 30
; EnableThread
; EnableXP
; DPIAware
; CompileSourceDirectory