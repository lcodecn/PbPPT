; ***************************************************************************************
; 示例5: 添加表格
; 说明: 演示如何创建表格、设置单元格文本、行高列宽和单元格合并
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
  Define titleId.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(0.5), PbPPT_InchesToEmu(0.3), PbPPT_InchesToEmu(9), PbPPT_InchesToEmu(0.8))
  PbPPT_SetShapeText(titleId, "员工信息表")
  PbPPT_SetFont(titleId, 1, 1, "微软雅黑", 24, #True, #False, #False, "2E75B6")
  PbPPT_SetParagraphAlignment(titleId, 1, #PbPPT_AlignCenter)

  ; 添加4行4列的表格
  Define tblId.i = PbPPT_AddTable(slideIdx, 4, 4, PbPPT_InchesToEmu(0.5), PbPPT_InchesToEmu(1.5), PbPPT_InchesToEmu(9), PbPPT_InchesToEmu(4))

  ; 设置表头（第1行）
  PbPPT_SetCellValue(tblId, 1, 1, "编号")
  PbPPT_SetCellValue(tblId, 1, 2, "姓名")
  PbPPT_SetCellValue(tblId, 1, 3, "部门")
  PbPPT_SetCellValue(tblId, 1, 4, "职位")

  ; 设置数据行
  PbPPT_SetCellValue(tblId, 2, 1, "001")
  PbPPT_SetCellValue(tblId, 2, 2, "张三")
  PbPPT_SetCellValue(tblId, 2, 3, "技术部")
  PbPPT_SetCellValue(tblId, 2, 4, "高级工程师")

  PbPPT_SetCellValue(tblId, 3, 1, "002")
  PbPPT_SetCellValue(tblId, 3, 2, "李四")
  PbPPT_SetCellValue(tblId, 3, 3, "市场部")
  PbPPT_SetCellValue(tblId, 3, 4, "市场经理")

  PbPPT_SetCellValue(tblId, 4, 1, "003")
  PbPPT_SetCellValue(tblId, 4, 2, "王五")
  PbPPT_SetCellValue(tblId, 4, 3, "设计部")
  PbPPT_SetCellValue(tblId, 4, 4, "首席设计师")

  ; 设置列宽（编号列窄，姓名和部门列中等，职位列宽）
  PbPPT_SetColWidth(tblId, 1, PbPPT_InchesToEmu(1))
  PbPPT_SetColWidth(tblId, 2, PbPPT_InchesToEmu(2))
  PbPPT_SetColWidth(tblId, 3, PbPPT_InchesToEmu(2.5))
  PbPPT_SetColWidth(tblId, 4, PbPPT_InchesToEmu(3.5))

  ; 设置行高
  PbPPT_SetRowHeight(tblId, 1, PbPPT_InchesToEmu(0.6))
  PbPPT_SetRowHeight(tblId, 2, PbPPT_InchesToEmu(0.8))
  PbPPT_SetRowHeight(tblId, 3, PbPPT_InchesToEmu(0.8))
  PbPPT_SetRowHeight(tblId, 4, PbPPT_InchesToEmu(0.8))

  ; 读取并验证单元格数据
  Define cellText.s = PbPPT_GetCellValue(tblId, 2, 2)

  Define filename.s = outputDir + "example5_table.pptx"
  If PbPPT_Save(filename)
    MessageRequester("示例5 - 成功", "表格演示文稿已创建：" + #LF$ + filename + #LF$ + "验证：第2行第2列 = " + cellText)
  Else
    MessageRequester("示例5 - 错误", "保存文件失败！")
  EndIf

  PbPPT_Close()
EndIf
