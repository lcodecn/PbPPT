; ***************************************************************************************
; 示例6: 打开已有文件
; 说明: 演示如何打开已有的pptx文件并读取其信息
; ***************************************************************************************
XIncludeFile "..\PbPPT.pb"

; 先确保有测试文件可打开（使用示例1创建的文件）
Define outputDir.s = "output\"
If FileSize(outputDir) = -1
  CreateDirectory(outputDir)
EndIf

PbPPT_Init()

; 先创建一个测试文件
PbPPT_Create()
PbPPT_AddSlide(0)
PbPPT_AddSlide(0)
PbPPT_SetTitle("测试演示文稿")
PbPPT_SetAuthor("PbPPT")
PbPPT_Save(outputDir + "example6_source.pptx")
PbPPT_Close()

; 打开刚创建的文件
If PbPPT_Open(outputDir + "example6_source.pptx")
  ; 读取演示文稿信息
  Define slideCount.i = PbPPT_GetSlideCount()
  Define width.i = PbPPT_GetSlideWidth()
  Define height.i = PbPPT_GetSlideHeight()
  Define title.s = PbPPT_GetTitle()
  Define author.s = PbPPT_GetAuthor()

  ; 构建信息字符串
  Define info.s = ""
  info + "幻灯片数量：" + Str(slideCount) + #LF$
  info + "幻灯片宽度：" + Str(width) + " EMU" + #LF$
  info + "幻灯片高度：" + Str(height) + " EMU" + #LF$
  info + "文档标题：" + title + #LF$
  info + "文档作者：" + author + #LF$
  info + "宽度（英寸）：" + StrF(PbPPT_EmuToInches(width), 2) + #LF$
  info + "高度（英寸）：" + StrF(PbPPT_EmuToInches(height), 2)

  MessageRequester("示例6 - 成功", "文件信息读取成功：" + #LF$ + #LF$ + info)

  PbPPT_Close()
Else
  MessageRequester("示例6 - 错误", "打开文件失败！")
EndIf
