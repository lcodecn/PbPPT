# PbPPT Library v1.8

PbPPT \- PureBasic PowerPoint pptx 操作库

- 作者  : lcode.cn
- 版本  : 1.8
- 许可证  : Apache 2.0
- 编译器  : PureBasic 6.40 (Windows - x86)

***

## 简介

PbPPT 是一个纯 PureBasic 实现的 PowerPoint 文件操作库，无需安装 Microsoft Office 或任何第三方依赖，即可创建和读取 PowerPoint pptx 文件。

该库基于 Python 的 python-pptx 项目（v1.0.2）移植而来，使用 PureBasic 内置的 XML 和 Packer（ZIP压缩）库实现。

## 主要功能

- 创建PPT文件  : 从零创建符合 Office Open XML 标准的 pptx 文件
- 读取PPT文件  : 解析现有 PowerPoint 文件内容（部分实现）
- 幻灯片操作  : 添加、删除幻灯片，设置幻灯片尺寸
- 文本框  : 添加文本框，设置多段落文本和富文本格式
- 自动形状  : 支持矩形、圆形、三角形、菱形、五角星、箭头等预设形状
- 图片嵌入  : 支持嵌入 PNG、JPG、BMP、GIF 等格式的图片
- 表格  : 创建表格，设置单元格文本、行高列宽
- 图表  : 支持柱状图、条形图、折线图、饼图等图表类型
- 连接符  : 添加直线连接符
- 形状样式  : 支持纯色填充、无填充、线条格式、旋转等样式设置
- 字体设置  : 支持字体名称、大小、颜色、粗体、斜体、下划线
- 段落对齐  : 支持左对齐、居中、右对齐等段落对齐方式
- 文档属性  : 设置标题、作者、主题、关键词等文档属性
- 幻灯片尺寸  : 支持标准4:3、宽屏16:9和自定义尺寸
- 图片去重  : 自动检测重复图片，减少文件大小

## 系统要求

- 本项目在PureBasic 6.40 （Windows x86）中编译通过，其他环境请自行测试。

## 快速开始

具体可参考开发文档：doc\PbPPT_Help.html

### 创建空白演示文稿

```purebasic
XIncludeFile "PbPPT.pb"

; 初始化PbPPT库
PbPPT_Init()

; 创建新的演示文稿
If PbPPT_Create()
  ; 添加3张空白幻灯片
  PbPPT_AddSlide(0)
  PbPPT_AddSlide(0)
  PbPPT_AddSlide(0)

  ; 保存文件
  If PbPPT_Save("output.pptx")
    MessageRequester("成功", "文件已保存！")
  EndIf

  ; 关闭演示文稿
  PbPPT_Close()
EndIf
```

### 添加文本和形状

```purebasic
XIncludeFile "PbPPT.pb"

PbPPT_Init()

If PbPPT_Create()
  Define slideIdx.i = PbPPT_AddSlide(0)

  ; 添加标题文本框
  Define titleId.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(0.5), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(1.5))
  PbPPT_SetShapeText(titleId, "PbPPT 示例")
  PbPPT_SetFont(titleId, 1, 1, "微软雅黑", 28, #True, #False, #False, "2E75B6")

  ; 添加矩形形状
  Define rectId.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstRectangle$, PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(3), PbPPT_InchesToEmu(6), PbPPT_InchesToEmu(2))
  PbPPT_SetFillSolid(rectId, "4472C4")
  PbPPT_SetShapeText(rectId, "蓝色矩形")
  PbPPT_SetFont(rectId, 1, 1, "微软雅黑", 18, #True, #False, #False, "FFFFFF")

  PbPPT_Save("output.pptx")
  PbPPT_Close()
EndIf
```

## API 文档

### 演示文稿操作

| 函数 | 说明 |
| --- | --- |
| `PbPPT_Init()` | 初始化PbPPT库 |
| `PbPPT_Create()` | 创建新的演示文稿 |
| `PbPPT_Open(filename.s)` | 打开已有的pptx文件 |
| `PbPPT_Save(filename.s)` | 保存演示文稿到文件 |
| `PbPPT_Close()` | 关闭演示文稿，释放资源 |
| `PbPPT_GetSlideCount()` | 获取幻灯片数量 |

### 幻灯片操作

| 函数 | 说明 |
| --- | --- |
| `PbPPT_AddSlide(layoutIndex.i)` | 添加新幻灯片，返回幻灯片索引 |
| `PbPPT_SetSlideWidth(width.i)` | 设置幻灯片宽度（EMU） |
| `PbPPT_SetSlideHeight(height.i)` | 设置幻灯片高度（EMU） |
| `PbPPT_GetSlideWidth()` | 获取幻灯片宽度（EMU） |
| `PbPPT_GetSlideHeight()` | 获取幻灯片高度（EMU） |

### 形状操作

| 函数 | 说明 |
| --- | --- |
| `PbPPT_AddTextbox(slideIdx, left, top, width, height)` | 添加文本框 |
| `PbPPT_AddShape(slideIdx, prstGeom, left, top, width, height)` | 添加自动形状 |
| `PbPPT_AddPicture(slideIdx, imagePath, left, top, width, height)` | 添加图片 |
| `PbPPT_AddTable(slideIdx, rows, cols, left, top, width, height)` | 添加表格 |
| `PbPPT_AddChart(slideIdx, chartType, left, top, width, height)` | 添加图表 |
| `PbPPT_AddConnector(slideIdx, beginX, beginY, endX, endY)` | 添加连接符 |

### 文本操作

| 函数 | 说明 |
| --- | --- |
| `PbPPT_SetShapeText(shapeId, text)` | 设置形状文本 |
| `PbPPT_SetFont(shapeId, paraIdx, runIdx, name, size, bold, italic, underline, color)` | 设置字体 |
| `PbPPT_AddRun(shapeId, paraIdx, text, name, size, bold, italic, underline, color)` | 添加富文本Run |
| `PbPPT_SetParagraphAlignment(shapeId, paraIdx, align)` | 设置段落对齐 |

### 样式操作

| 函数 | 说明 |
| --- | --- |
| `PbPPT_SetFillSolid(shapeId, color)` | 设置纯色填充 |
| `PbPPT_SetFillNone(shapeId)` | 设置无填充 |
| `PbPPT_SetLineFormat(shapeId, width, color)` | 设置线条格式 |
| `PbPPT_SetShapeRotation(shapeId, angle)` | 设置形状旋转角度 |

### 表格操作

| 函数 | 说明 |
| --- | --- |
| `PbPPT_SetCellValue(shapeId, row, col, text)` | 设置单元格值 |
| `PbPPT_GetCellValue(shapeId, row, col)` | 获取单元格值 |
| `PbPPT_SetColWidth(shapeId, col, width)` | 设置列宽 |
| `PbPPT_SetRowHeight(shapeId, row, height)` | 设置行高 |

### 图表操作

| 函数 | 说明 |
| --- | --- |
| `PbPPT_SetChartTitle(shapeId, title)` | 设置图表标题 |
| `PbPPT_AddChartCategory(shapeId, label)` | 添加图表分类标签 |
| `PbPPT_AddChartSeries(shapeId, name, values)` | 添加图表数据系列 |
| `PbPPT_SetChartHasLegend(shapeId, hasLegend)` | 设置是否显示图例 |
| `PbPPT_SetChartStyle(shapeId, style)` | 设置图表样式 |

### 文档属性

| 函数 | 说明 |
| --- | --- |
| `PbPPT_SetTitle(title)` | 设置文档标题 |
| `PbPPT_SetAuthor(author)` | 设置文档作者 |
| `PbPPT_SetSubject(subject)` | 设置文档主题 |
| `PbPPT_GetTitle()` | 获取文档标题 |
| `PbPPT_GetAuthor()` | 获取文档作者 |

### 单位转换

| 函数 | 说明 |
| --- | --- |
| `PbPPT_InchesToEmu(inches.f)` | 英寸转EMU |
| `PbPPT_CmToEmu(cm.f)` | 厘米转EMU |
| `PbPPT_PtToEmu(pt.f)` | 磅转EMU |
| `PbPPT_EmuToInches(emu.i)` | EMU转英寸 |

## 文件结构

PbPPT.pb 文件按照功能模块分为以下分区：

| 分区 | 内容 |
| --- | --- |
| 分区1 | 常量定义（OOXML规范、文件路径、XML命名空间、MIME类型等） |
| 分区2 | 枚举定义（形状类型、图表类型、对齐方式、填充类型等） |
| 分区3 | 结构体定义和全局数据存储 |
| 分区4 | 工具函数（单位转换、字符串处理、日期时间、XML/ZIP辅助） |
| 分区5 | XML写入器（生成PPTX文件各部分XML） |
| 分区6 | XML读取器（解析现有PPTX文件） |
| 分区7 | 形状创建函数（文本框、形状、图片、表格、图表、连接符） |
| 分区8 | 形状属性设置函数（文本、字体、填充、线条等） |
| 分区9 | 公共API |
| 分区10 | 初始化和清理 |

## 版本历史

### v1.8 (2026-04-24)

- \[修复] 图片无法显示的问题：修正slide rels中图片Target路径和r:embed引用
- \[修复] 图表无法显示的问题：新增完整的chart XML生成器，支持柱形图、条形图、折线图、饼图
- \[修复] 图表XML中c:chart元素嵌套错误
- \[修复] Content_Types.xml中缺少图片扩展名Default映射
- \[修复] WriteSlideRelsXML和WriteSlideXML调用顺序导致r:id为空的问题
- \[新增] 结构体PbPPT_Shapes添加chartStyle和chartHasLegend字段
- \[新增] 图表类型常量（LineStacked、LineStacked100、DoughnutExploded）
- \[优化] SetChartHasLegend和SetChartStyle改为使用独立字段存储
- \[优化] 统一图表数据格式解析（使用分号分隔系列）

### v1.7 (2026-04-24)

- \[修复] 生成的PPTX文件无法被Office工具打开的问题
- \[修复] presentation.xml中p:sldId的r:id属性为空（WritePresentationRelsXML和WritePresentationXML调用顺序问题）
- \[修复] .rels文件引用不存在的thumbnail.jpeg
- \[修复] 缺少p:notesSz元素
- \[修复] UTF-8编码缓冲区溢出问题
- \[新增] 添加p:notesSz元素（备注页尺寸）
- \[优化] 移除根关系中对不存在缩略图的引用

### v1.6 (2026-04-23)

- \[新增] 10个示例文件（空白演示文稿、文本、形状、图片、表格、打开文件、修改保存、图表、样式、幻灯片尺寸）
- \[新增] 示例文件扩展名改为.pb
- \[新增] 图片自动尺寸检测（使用PureBasic图像解码器）
- \[新增] 图片SHA1指纹去重功能
- \[修复] 示例文件保存路径问题
- \[优化] 所有示例文件转换为UTF-8 BOM格式

### v1.5 (2026-04-22)

- \[新增] 图表功能接口（AddChart、SetChartTitle、AddChartCategory、AddChartSeries）
- \[新增] 图表类型枚举（柱形图、条形图、折线图、饼图、圆环图、雷达图等）
- \[新增] 连接符功能（AddConnector）
- \[新增] 段落对齐功能（SetParagraphAlignment）
- \[新增] 富文本Run功能（AddRun）
- \[新增] 形状旋转功能（SetShapeRotation）

### v1.4 (2026-04-21)

- \[新增] 表格功能（AddTable、SetCellValue、GetCellValue、SetColWidth、SetRowHeight）
- \[新增] 图片嵌入功能（AddPicture），支持PNG/JPG/BMP/GIF等格式
- \[新增] 二进制部件管理（AddBinaryPart）
- \[新增] 文件指纹计算（CalcFileFingerprint）
- \[新增] 图片内容类型自动识别（GetImageContentType）

### v1.3 (2026-04-20)

- \[新增] 形状样式设置（SetFillSolid、SetFillNone、SetLineFormat）
- \[新增] 字体设置功能（SetFont），支持字体名称、大小、颜色、粗体、斜体、下划线
- \[新增] 自动形状支持（矩形、圆形、三角形、菱形、五角星、箭头等）
- \[新增] 预设形状常量定义（PrstRectangle、PrstEllipse等）
- \[新增] 形状填充类型枚举

### v1.2 (2026-04-19)

- \[新增] 文本框功能（AddTextbox、SetShapeText）
- \[新增] 多段落文本支持（使用换行符分隔段落）
- \[新增] 形状创建基础框架（CreateShape）
- \[新增] 形状类型枚举定义
- \[新增] 幻灯片尺寸设置（SetSlideWidth、SetSlideHeight）
- \[新增] 标准和宽屏幻灯片尺寸常量

### v1.1 (2026-04-18)

- \[新增] 幻灯片添加功能（AddSlide）
- \[新增] 幻灯片版式管理
- \[新增] 幻灯片母版XML生成（slideMaster1.xml）
- \[新增] 幻灯片版式XML生成（10种版式）
- \[新增] 主题XML生成（theme1.xml）
- \[新增] 演示文稿属性XML生成（presProps.xml、viewProps.xml、tableStyles.xml）
- \[新增] OPC包关系管理

### v1.0 (2026-04-17)

- \[新增] 初始项目创建，参考python-pptx（v1.0.2）编写
- \[新增] 演示文稿创建功能（PbPPT_Create、PbPPT_Save）
- \[新增] 演示文稿打开功能（PbPPT_Open，部分实现）
- \[新增] OOXML规范常量定义（命名空间、内容类型、关系类型等）
- \[新增] EMU单位系统（InchesToEmu、CmToEmu、PtToEmu等转换函数）
- \[新增] ZIP打包支持（UseZipPacker、CreatePack等）
- \[新增] XML辅助模块（节点创建、属性设置、文本设置、保存为字符串）
- \[新增] OPC包管理模块（部件、关系、内容类型）
- \[新增] PackURI路径处理
- \[新增] 文档属性设置（标题、作者、主题、关键词等）
- \[新增] 初始化和清理函数（PbPPT_Init、PbPPT_Close）

## 许可证

本库采用 Apache 2.0 许可证。

```
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```

本库参考的python-pptx 项目采用MIT许可证。

```
The MIT License (MIT)
Copyright (c) 2013 Steve Canny, https://github.com/scanny

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
```

## 致谢

- 感谢 python-pptx 项目提供了优秀的参考实现
- 感谢 PureBasic QQ群的支持
