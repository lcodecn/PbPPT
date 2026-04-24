; ***************************************************************************************
; 示例8: 添加图表
; 说明: 演示如何添加柱状图并设置图表标题和数据系列
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
  Define titleId.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(0.3), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(0.8))
  PbPPT_SetShapeText(titleId, "季度销售数据图表")
  PbPPT_SetFont(titleId, 1, 1, "微软雅黑", 24, #True, #False, #False, "2E75B6")
  PbPPT_SetParagraphAlignment(titleId, 1, #PbPPT_AlignCenter)

  ; 添加柱状图
  Define chartId.i = PbPPT_AddChart(slideIdx, #PbPPT_ChartTypeColumnClustered, PbPPT_InchesToEmu(0.5), PbPPT_InchesToEmu(1.5), PbPPT_InchesToEmu(9), PbPPT_InchesToEmu(5.5))

  ; 设置图表标题
  PbPPT_SetChartTitle(chartId, "2025年季度销售额")

  ; 添加分类标签
  PbPPT_AddChartCategory(chartId, "第一季度")
  PbPPT_AddChartCategory(chartId, "第二季度")
  PbPPT_AddChartCategory(chartId, "第三季度")
  PbPPT_AddChartCategory(chartId, "第四季度")

  ; 添加数据系列
  PbPPT_AddChartSeries(chartId, "产品A", "120,200,150,300")
  PbPPT_AddChartSeries(chartId, "产品B", "80,150,220,250")
  PbPPT_AddChartSeries(chartId, "产品C", "60,100,180,210")

  ; 设置图表样式
  PbPPT_SetChartHasLegend(chartId, #True)
  PbPPT_SetChartStyle(chartId, 2)

  Define filename.s = outputDir + "example8_chart.pptx"
  If PbPPT_Save(filename)
    MessageRequester("示例8 - 成功", "图表演示文稿已创建：" + #LF$ + filename)
  Else
    MessageRequester("示例8 - 错误", "保存文件失败！")
  EndIf

  PbPPT_Close()
EndIf
