﻿; ***************************************************************************************
; 示例1: 创建空白演示文稿
; 说明: 演示如何使用PbPPT创建一个包含多张空白幻灯片的演示文稿
; ***************************************************************************************
XIncludeFile "..\PbPPT.pb"

Define outputDir.s = "output\"
If FileSize(outputDir) = -1
  CreateDirectory(outputDir)
EndIf

; 初始化PbPPT库
PbPPT_Init()

; 创建新的演示文稿
If PbPPT_Create()
  ; 添加3张空白幻灯片
  PbPPT_AddSlide(0)
  PbPPT_AddSlide(0)
  PbPPT_AddSlide(0)

  ; 保存到文件
  Define filename.s = outputDir + "example1_blank.pptx"
  If PbPPT_Save(filename)
    MessageRequester("示例1 - 成功", "空白演示文稿已创建：" + #LF$ + filename + #LF$ + "幻灯片数量：" + Str(PbPPT_GetSlideCount()))
  Else
    MessageRequester("示例1 - 错误", "保存文件失败！")
  EndIf

  ; 关闭演示文稿
  PbPPT_Close()
Else
  MessageRequester("示例1 - 错误", "创建演示文稿失败！")
EndIf
