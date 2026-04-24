; ***************************************************************************************
; 示例2: 添加文本
; 说明: 演示如何添加文本框、设置多段落文本和富文本格式
; ***************************************************************************************
XIncludeFile "..\PbPPT.pb"

Define outputDir.s = "output\"
If FileSize(outputDir) = -1
  CreateDirectory(outputDir)
EndIf

PbPPT_Init()

If PbPPT_Create()
  ; 添加一张幻灯片
  Define slideIdx.i = PbPPT_AddSlide(0)

  ; 添加标题文本框
  Define titleId.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(0.5), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(1.5))
  PbPPT_SetShapeText(titleId, "PbPPT 文本示例")
  ; 设置标题字体：微软雅黑，28号，粗体，蓝色
  PbPPT_SetFont(titleId, 1, 1, "微软雅黑", 28, #True, #False, #False, "2E75B6")

  ; 添加内容文本框（多段落）
  Define contentId.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2.5), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(4))
  ; 设置多行文本（换行符分隔段落）
  PbPPT_SetShapeText(contentId, "这是第一段文本，演示基本文本设置。" + #LF$ + "这是第二段文本，演示多段落支持。" + #LF$ + "这是第三段文本，PbPPT可以轻松创建演示文稿。")

  ; 添加带富文本格式的文本框
  Define richId.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(6.5), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(1))
  PbPPT_SetShapeText(richId, "普通文本")
  ; 在第1段添加带格式的Run
  PbPPT_AddRun(richId, 1, "（粗体红色）", "宋体", 14, #True, #False, #False, "FF0000")
  PbPPT_AddRun(richId, 1, "（斜体绿色）", "宋体", 14, #False, #True, #False, "00B050")

  ; 设置段落对齐
  PbPPT_SetParagraphAlignment(titleId, 1, #PbPPT_AlignCenter)

  ; 保存
  Define filename.s = outputDir + "example2_text.pptx"
  If PbPPT_Save(filename)
    MessageRequester("示例2 - 成功", "文本演示文稿已创建：" + #LF$ + filename)
  Else
    MessageRequester("示例2 - 错误", "保存文件失败！")
  EndIf

  PbPPT_Close()
EndIf
