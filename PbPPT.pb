; ***************************************************************************************
; PbPPT Library - PowerPoint pptx 操作库
; 版本: 1.8
; 作者: lcode.cn
; 许可证: Apache 2.0
;
; 说明: 无依赖操作PowerPoint pptx文件的PureBasic库
;       使用PureBasic内置XML和Packer库
;       无需安装Microsoft Office或任何第三方依赖
;       移植自python-pptx项目(v1.0.2)
;
; 支持本项目:
;   PayPal:  https://www.paypal.me/lcodecn
;   微信支付: #付款:lcodecn(经营_lcodecn)/openlib/003
;
; 主要功能:
;   - 创建和读取PowerPoint pptx文件
;   - 幻灯片管理（添加、删除、访问）
;   - 形状操作（自动形状、文本框、图片、表格、图表等）
;   - 文本操作（文本框、段落、字体、颜色）
;   - 表格操作（单元格读写、合并、样式）
;   - 图表操作（柱状图、折线图、饼图等）
;   - DrawingML样式（填充、颜色、线条、阴影）
;   - 幻灯片版式和母版
;   - 核心文档属性
;   - 完整的ZIP压缩支持
;
; 文件结构:
;   分区1:  常量定义（PPT规范、文件路径、XML命名空间、MIME类型、关系类型、枚举、单位）
;   分区2:  结构体定义（演示文稿、幻灯片、形状、文本、表格、图表等）
;   分区3:  全局变量和数据存储
;   分区4:  工具函数（单位转换、字符串处理、XML转义、SHA1哈希）
;   分区5:  ZIP辅助模块（创建、打开、添加、解压、关闭）
;   分区6:  XML辅助模块（创建、解析、节点操作、序列化）
;   分区7:  PackURI模块（包URI解析、路径计算）
;   分区8:  OPC包模块（读取、写入、内容类型、关系管理）
;   分区9:  XML写入器（生成pptx文件各部分XML）
;   分区10: XML读取器（解析pptx文件各部分XML）
;   分区11: 形状模块（形状创建、自动形状、图片、连接符、占位符）
;   分区12: 文本模块（文本框、段落、运行、字体）
;   分区13: 图表模块（图表创建、数据、坐标轴、系列）
;   分区14: DrawingML模块（填充、颜色、线条、阴影）
;   分区15: 表格模块（表格创建、单元格、行列、合并）
;   分区16: 公共API（用户可调用的所有函数接口）
;   分区17: 初始化和清理
;   分区18: 测试代码
; ***************************************************************************************

; 启用ZIP打包库（PureBasic的XML库是内置的，无需初始化）
UseZipPacker()

; ***************************************************************************************
; 前向声明（PureBasic要求函数在使用前声明或定义）
; ***************************************************************************************
Declare PbPPT_ClearAllData()
Declare PbPPT_CleanupTempDir()
Declare.b PbPPT_ExtractToTempDir(filename.s)
Declare PbPPT_ReadPresentationInfo()
Declare PbPPT_ReadCoreProps()
Declare PbPPT_ReadSlides()
Declare PbPPT_ReadAppProps()
Declare.i PbPPT_GetNextShapeId()
Declare PbPPT_AddBinaryPart(partName.s, contentType.s, *data, dataSize.i)
Declare PbPPT_SetShapeFill(shapeId.i, fillType.i, foreColor.s, backColor.s)
Declare PbPPT_SetShapeLine(shapeId.i, width.i, color.s, dashStyle.s)
Declare.b PbPPT_GetShapeById(shapeId.i)
Declare PbPPT_SetTextFrameText(shapeId.i, text.s)
Declare PbPPT_SetChartTitle(shapeId.i, title.s)
Declare.i PbPPT_WriteContentTypesXML()
Declare.i PbPPT_WriteRootRelsXML()
Declare.i PbPPT_WritePresentationXML()
Declare.i PbPPT_WritePresentationRelsXML()
Declare.i PbPPT_WriteCorePropsXML()
Declare.i PbPPT_WriteAppPropsXML()
Declare.i PbPPT_WritePresPropsXML()
Declare.i PbPPT_WriteViewPropsXML()
Declare.i PbPPT_WriteTableStylesXML()
Declare.i PbPPT_WriteThemeXML()
Declare.i PbPPT_WriteSlideMasterXML()
Declare.i PbPPT_WriteSlideMasterRelsXML()
Declare.i PbPPT_WriteSlideLayoutXML(layoutName.s, layoutIdx.i)
Declare.i PbPPT_WriteSlideXML(slideIndex.i)
Declare.i PbPPT_WriteSlideRelsXML(slideIndex.i, layoutRId.s)
Declare.i PbPPT_WriteChartXML(chartShapeId.i)
Declare PbPPT_AddXMLToZIP(packId.i, xmlId.i, archivePath.s)
Declare PbPPT_AddStringToZIP(packId.i, content.s, archivePath.s)
Declare PbPPT_AddBinaryToZIP(packId.i, *data, dataSize.i, archivePath.s)
Declare.i PbPPT_XMLCreateDocument()
Declare PbPPT_XMLFree(xmlId.i)
Declare.i PbPPT_XMLGetRoot(xmlId.i)
Declare.i PbPPT_XMLAddNode(xmlId.i, parentNode.i, nodeName.s)
Declare PbPPT_XMLSetAttribute(node.i, attrName.s, attrValue.s)
Declare PbPPT_XMLSetText(node.i, text.s)
Declare.i PbPPT_XMLParseFile(filename.s)
Declare.i PbPPT_XMLGetChild(node.i)
Declare.i PbPPT_XMLGetNext(node.i)
Declare.s PbPPT_XMLGetNodeName(node.i)
Declare.s PbPPT_XMLGetAttributeValue(node.i, attrName.s)
Declare.s PbPPT_XMLGetText(node.i)
Declare.i PbPPT_ZIPCreate(filename.s)
Declare PbPPT_ZIPClose(packId.i)
Declare.s PbPPT_CalcFileFingerprint(path.s)
Declare.i PbPPT_CreateShape(slideIndex.i, shapeType.i, left.i, top.i, width.i, height.i)
Declare.i PbPPT_CreateAutoShape(slideIndex.i, preset.s, left.i, top.i, width.i, height.i)
Declare.i PbPPT_CreateTextbox(slideIndex.i, left.i, top.i, width.i, height.i)
Declare.i PbPPT_CreatePicture(slideIndex.i, imagePath.s, left.i, top.i, width.i, height.i)
Declare.i PbPPT_CreateTable(slideIndex.i, rows.i, cols.i, left.i, top.i, width.i, height.i)
Declare.i PbPPT_CreateChart(slideIndex.i, chartType.i, left.i, top.i, width.i, height.i)
Declare.i PbPPT_CreateConnector(slideIndex.i, beginX.i, beginY.i, endX.i, endY.i)
Declare.s PbPPT_GetTextFrameText(shapeId.i)
Declare.s PbPPT_GetFileExtension(path.s)
Declare.s PbPPT_GetImageContentType(ext.s)
Declare.s PbPPT_GetCurrentDateTime()
Declare.i PbPPT_PtToEmu(pt.f)

; ***************************************************************************************
; 分区1: 常量定义
; ***************************************************************************************

; 1.1 PPT规范常量
#PbPPT_MaxSlides = 256
#PbPPT_MaxShapes = 1024
#PbPPT_Version$ = "1.0"

; 1.2 文件路径常量（定义pptx文件内部各XML文件的路径）
#PbPPT_ARCContentTypes$ = "[Content_Types].xml"
#PbPPT_ARCRootRels$ = "_rels/.rels"
#PbPPT_ARCPresentation$ = "ppt/presentation.xml"
#PbPPT_ARCPresentationRels$ = "ppt/_rels/presentation.xml.rels"
#PbPPT_ARCCore$ = "docProps/core.xml"
#PbPPT_ARCApp$ = "docProps/app.xml"
#PbPPT_ARCPresProps$ = "ppt/presProps.xml"
#PbPPT_ARCViewProps$ = "ppt/viewProps.xml"
#PbPPT_ARCTableStyles$ = "ppt/tableStyles.xml"
#PbPPT_ARCTheme$ = "ppt/theme/theme1.xml"
#PbPPT_ARCSlideMaster$ = "ppt/slideMasters/slideMaster1.xml"
#PbPPT_ARCSlideMasterRels$ = "ppt/slideMasters/_rels/slideMaster1.xml.rels"

; 1.3 XML命名空间常量（定义PPT XML文件中使用的各种XML命名空间URI）
#PbPPT_NS_P$ = "http://schemas.openxmlformats.org/presentationml/2006/main"
#PbPPT_NS_A$ = "http://schemas.openxmlformats.org/drawingml/2006/main"
#PbPPT_NS_R$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
#PbPPT_NS_C$ = "http://schemas.openxmlformats.org/drawingml/2006/chart"
#PbPPT_NS_REL$ = "http://schemas.openxmlformats.org/package/2006/relationships"
#PbPPT_NS_CT$ = "http://schemas.openxmlformats.org/package/2006/content-types"
#PbPPT_NS_CP$ = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
#PbPPT_NS_DC$ = "http://purl.org/dc/elements/1.1/"
#PbPPT_NS_DCTERMS$ = "http://purl.org/dc/terms/"
#PbPPT_NS_DCMITYPE$ = "http://purl.org/dc/dcmitype/"
#PbPPT_NS_XSI$ = "http://www.w3.org/2001/XMLSchema-instance"
#PbPPT_NS_EP$ = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
#PbPPT_NS_VT$ = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
#PbPPT_NS_P14$ = "http://schemas.microsoft.com/office/powerpoint/2010/main"
#PbPPT_NS_A14$ = "http://schemas.microsoft.com/office/drawing/2010/main"

; Clark标记法命名空间（用于PureBasic XML库的节点名称）
#PbPPT_CNS_P$ = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
#PbPPT_CNS_A$ = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
#PbPPT_CNS_R$ = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
#PbPPT_CNS_C$ = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
#PbPPT_CNS_REL$ = "{http://schemas.openxmlformats.org/package/2006/relationships}"
#PbPPT_CNS_CT$ = "{http://schemas.openxmlformats.org/package/2006/content-types}"
#PbPPT_CNS_CP$ = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
#PbPPT_CNS_DC$ = "{http://purl.org/dc/elements/1.1/}"
#PbPPT_CNS_DCTERMS$ = "{http://purl.org/dc/terms/}"
#PbPPT_CNS_XSI$ = "{http://www.w3.org/2001/XMLSchema-instance}"
#PbPPT_CNS_EP$ = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"
#PbPPT_CNS_VT$ = "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}"

; 1.4 MIME类型/内容类型常量
#PbPPT_CT_PRESENTATION_MAIN$ = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
#PbPPT_CT_PRES_MACRO_MAIN$ = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
#PbPPT_CT_TEMPLATE_MAIN$ = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
#PbPPT_CT_SLIDESHOW_MAIN$ = "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"
#PbPPT_CT_SLIDE$ = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
#PbPPT_CT_SLIDE_LAYOUT$ = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
#PbPPT_CT_SLIDE_MASTER$ = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
#PbPPT_CT_NOTES_MASTER$ = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
#PbPPT_CT_NOTES_SLIDE$ = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
#PbPPT_CT_PRES_PROPS$ = "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"
#PbPPT_CT_VIEW_PROPS$ = "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"
#PbPPT_CT_TABLE_STYLES$ = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"
#PbPPT_CT_THEME$ = "application/vnd.openxmlformats-officedocument.theme+xml"
#PbPPT_CT_CORE_PROPS$ = "application/vnd.openxmlformats-package.core-properties+xml"
#PbPPT_CT_EXT_PROPS$ = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
#PbPPT_CT_OPC_RELS$ = "application/vnd.openxmlformats-package.relationships+xml"
#PbPPT_CT_CHART$ = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
#PbPPT_CT_IMAGE_BMP$ = "image/bmp"
#PbPPT_CT_IMAGE_GIF$ = "image/gif"
#PbPPT_CT_IMAGE_JPEG$ = "image/jpeg"
#PbPPT_CT_IMAGE_PNG$ = "image/png"
#PbPPT_CT_IMAGE_TIFF$ = "image/tiff"
#PbPPT_CT_IMAGE_EMF$ = "image/x-emf"
#PbPPT_CT_IMAGE_WMF$ = "image/x-wmf"
#PbPPT_CT_XML$ = "application/xml"

; 1.5 关系类型常量
#PbPPT_RT_OFFICE_DOCUMENT$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
#PbPPT_RT_CORE_PROPERTIES$ = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
#PbPPT_RT_EXT_PROPERTIES$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
#PbPPT_RT_THUMBNAIL$ = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
#PbPPT_RT_SLIDE$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
#PbPPT_RT_SLIDE_LAYOUT$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
#PbPPT_RT_SLIDE_MASTER$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
#PbPPT_RT_THEME$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
#PbPPT_RT_NOTES_MASTER$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
#PbPPT_RT_NOTES_SLIDE$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
#PbPPT_RT_IMAGE$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
#PbPPT_RT_CHART$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
#PbPPT_RT_HYPERLINK$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
#PbPPT_RT_PRES_PROPS$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
#PbPPT_RT_VIEW_PROPS$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
#PbPPT_RT_TABLE_STYLES$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"

; 1.6 枚举常量 - 形状类型
#PbPPT_ShapeTypeAutoShape = 0
#PbPPT_ShapeTypePicture = 1
#PbPPT_ShapeTypeTextbox = 2
#PbPPT_ShapeTypeTable = 3
#PbPPT_ShapeTypeChart = 4
#PbPPT_ShapeTypeConnector = 5
#PbPPT_ShapeTypeGroup = 6
#PbPPT_ShapeTypePlaceholder = 7
#PbPPT_ShapeTypeFreeform = 8
#PbPPT_ShapeTypeGraphicFrame = 9

; 1.7 枚举常量 - 占位符类型
#PbPPT_PHTypeTitle = 0
#PbPPT_PHTypeBody = 1
#PbPPT_PHTypeCenterTitle = 2
#PbPPT_PHTypeSubtitle = 3
#PbPPT_PHTypeDate = 4
#PbPPT_PHTypeSlideNumber = 5
#PbPPT_PHTypeFooter = 6
#PbPPT_PHTypeHeader = 7
#PbPPT_PHTypeObject = 8
#PbPPT_PHTypeChart = 9
#PbPPT_PHTypeTable = 10
#PbPPT_PHTypePicture = 11
#PbPPT_PHTypeClipArt = 12
#PbPPT_PHTypeOrgChart = 13
#PbPPT_PHTypeMedia = 14

; 1.8 枚举常量 - 图表类型
#PbPPT_ChartTypeBarClustered = 0
#PbPPT_ChartTypeBarStacked = 1
#PbPPT_ChartTypeBarStacked100 = 2
#PbPPT_ChartTypeColumnClustered = 3
#PbPPT_ChartTypeColumnStacked = 4
#PbPPT_ChartTypeColumnStacked100 = 5
#PbPPT_ChartTypeLine = 6
#PbPPT_ChartTypeLineMarkers = 7
#PbPPT_ChartTypePie = 8
#PbPPT_ChartTypePieExploded = 9
#PbPPT_ChartTypeArea = 10
#PbPPT_ChartTypeAreaStacked = 11
#PbPPT_ChartTypeXYScatter = 12
#PbPPT_ChartTypeBubble = 13
#PbPPT_ChartTypeDoughnut = 14
#PbPPT_ChartTypeRadar = 15
#PbPPT_ChartTypeLineStacked = 16
#PbPPT_ChartTypeLineStacked100 = 17
#PbPPT_ChartTypeDoughnutExploded = 18

; 1.9 枚举常量 - 填充类型
#PbPPT_FillTypeNone = 0
#PbPPT_FillTypeSolid = 1
#PbPPT_FillTypeGradient = 2
#PbPPT_FillTypePattern = 3
#PbPPT_FillTypePicture = 4
#PbPPT_FillTypeBackground = 5

; 1.10 枚举常量 - 文本对齐方式
#PbPPT_AlignLeft = 0
#PbPPT_AlignCenter = 1
#PbPPT_AlignRight = 2
#PbPPT_AlignJustify = 3

; 1.11 枚举常量 - 垂直对齐方式
#PbPPT_ValignTop = 0
#PbPPT_ValignMiddle = 1
#PbPPT_ValignBottom = 2

; 1.12 单位转换常量（EMU - English Metric Units）
#PbPPT_EMU_PER_INCH = 914400
#PbPPT_EMU_PER_CM = 360000
#PbPPT_EMU_PER_MM = 36000
#PbPPT_EMU_PER_PT = 12700
#PbPPT_EMU_PER_CENTIPOINT = 127

; 1.13 默认幻灯片尺寸常量（EMU）
; 标准尺寸: 10英寸 x 7.5英寸 (9144000 x 6858000 EMU)
#PbPPT_DEFAULT_SLIDE_WIDTH = 9144000
#PbPPT_DEFAULT_SLIDE_HEIGHT = 6858000
; 宽屏尺寸: 13.333英寸 x 7.5英寸 (12192000 x 6858000 EMU)
#PbPPT_WIDESCREEN_SLIDE_WIDTH = 12192000
#PbPPT_WIDESCREEN_SLIDE_HEIGHT = 6858000

; 1.14 自动形状预设常量
#PbPPT_PrstRectangle$ = "rect"
#PbPPT_PrstRoundRectangle$ = "roundRect"
#PbPPT_PrstEllipse$ = "ellipse"
#PbPPT_PrstTriangle$ = "triangle"
#PbPPT_PrstRightTriangle$ = "rtTriangle"
#PbPPT_PrstDiamond$ = "diamond"
#PbPPT_PrstPentagon$ = "pentagon"
#PbPPT_PrstHexagon$ = "hexagon"
#PbPPT_PrstStar5$ = "star5"
#PbPPT_PrstArrow$ = "arrow"
#PbPPT_PrstChevron$ = "chevron"
#PbPPT_PrstHeart$ = "heart"
#PbPPT_PrstCloud$ = "cloud"

; ***************************************************************************************
; 分区2: 结构体定义
; ***************************************************************************************

; 关系结构体 - 描述包内部件之间的关系
Structure PbPPT_Relationship
  rId.s                  ; 关系ID，如"rId1"
  relType.s              ; 关系类型URI
  target.s               ; 目标路径
  isExternal.b           ; 是否为外部关系
EndStructure

; 部件结构体 - 描述包内的一个部件
Structure PbPPT_Part
  partName.s             ; 部件路径，如"/ppt/slides/slide1.xml"
  contentType.s          ; 内容类型（MIME类型）
  xmlContent.s           ; XML内容字符串
  blobSize.i             ; 二进制数据大小（用于图片等非XML部件）
  *blobData              ; 二进制数据指针
EndStructure

; 字体结构体
Structure PbPPT_Font
  name.s                 ; 字体名称
  size.f                 ; 字体大小（磅）
  bold.b                 ; 是否粗体
  italic.b               ; 是否斜体
  underline.b            ; 是否下划线
  strike.b               ; 是否删除线
  color.s                ; 字体颜色（十六进制RGB，如"FF0000"）
EndStructure

; 填充格式结构体
Structure PbPPT_FillFormat
  fillType.i             ; 填充类型（#PbPPT_FillType*）
  foreColor.s            ; 前景色（十六进制RGB）
  backColor.s            ; 背景色（十六进制RGB）
EndStructure

; 线条格式结构体
Structure PbPPT_LineFormat
  width.i                ; 线条宽度（EMU）
  color.s                ; 线条颜色（十六进制RGB）
  dashStyle.s            ; 虚线样式（"solid"/"dash"/"dashDot"等）
EndStructure

; 文本运行结构体 - 文本中最小的格式化单元
Structure PbPPT_Run
  text.s                 ; 运行文本内容
  font.PbPPT_Font        ; 字体属性
EndStructure

; 段落结构体
Structure PbPPT_Paragraph
  alignment.i            ; 对齐方式（#PbPPT_Align*）
  runs.s                 ; 运行列表（格式：文本#ATTR#字体名,字号,粗体,斜体,下划线,颜色#RUN#文本...）
  level.i                ; 缩进级别（0-8）
  spaceBefore.f          ; 段前间距（磅）
  spaceAfter.f           ; 段后间距（磅）
  lineSpacing.f          ; 行距（磅，0表示默认）
EndStructure

; 文本框结构体
Structure PbPPT_TextFrame
  wordWrap.b             ; 是否自动换行
  autoFit.i              ; 自动调整大小方式（0=无，1=文本溢出，2=形状调整）
  paragraphs.s           ; 段落列表（格式：段落内容#PARA#段落内容...，段落内含Run）
  verticalAnchor.i       ; 垂直对齐（#PbPPT_Valign*）
  marginLeft.i           ; 左边距（EMU，0表示默认）
  marginRight.i          ; 右边距（EMU，0表示默认）
  marginTop.i            ; 上边距（EMU，0表示默认）
  marginBottom.i         ; 下边距（EMU，0表示默认）
EndStructure

; 形状结构体 - 所有形状的通用属性
Structure PbPPT_Shape
  id.i                   ; 形状唯一ID
  slideIndex.i           ; 所属幻灯片索引（0-based，用于关联形状与幻灯片）
  shapeType.i            ; 形状类型（#PbPPT_ShapeType*）
  name.s                 ; 形状名称
  left.i                 ; 左边距（EMU）
  top.i                  ; 顶部距离（EMU）
  width.i                ; 宽度（EMU）
  height.i               ; 高度（EMU）
  rotation.f             ; 旋转角度（度）
  preset.s               ; 自动形状预设名称
  textFrame.PbPPT_TextFrame ; 文本框
  fill.PbPPT_FillFormat  ; 填充格式
  line.PbPPT_LineFormat  ; 线条格式
  ; 图片相关
  imagePath.s            ; 图片文件路径
  imageRId.s             ; 图片关系ID
  imageDesc.s            ; 图片描述
  ; 表格相关
  tableRows.i            ; 表格行数
  tableCols.i            ; 表格列数
  tableCells.s           ; 表格单元格数据（格式：行_列=文本，用|分隔）
  tableColWidths.s       ; 列宽列表（逗号分隔，单位EMU）
  tableRowHeights.s      ; 行高列表（逗号分隔，单位EMU）
  ; 图表相关
  chartType.i            ; 图表类型
  chartRId.s             ; 图表关系ID
  chartTitle.s           ; 图表标题
  chartData.s            ; 图表数据（分类|系列名:值1,值2,...格式）
  chartStyle.i           ; 图表样式
  chartHasLegend.b       ; 是否显示图例
  ; 占位符相关
  phType.i               ; 占位符类型
  phIdx.i                ; 占位符索引
  phOrient.s             ; 占位符方向
  phSz.s                 ; 占位符大小
  ; 超链接
  hyperlink.s            ; 超链接地址
EndStructure

; 幻灯片版式结构体
Structure PbPPT_SlideLayout
  id.i                   ; 版式ID
  name.s                 ; 版式名称
  partName.s             ; 部件路径
  rId.s                  ; 关系ID
  shapes.s               ; 形状列表（JSON格式存储）
EndStructure

; 幻灯片母版结构体
Structure PbPPT_SlideMaster
  id.i                   ; 母版ID
  name.s                 ; 母版名称
  partName.s             ; 部件路径
  rId.s                  ; 关系ID
  layouts.s              ; 版式列表（JSON格式存储）
EndStructure

; 幻灯片结构体
Structure PbPPT_Slide
  id.i                   ; 幻灯片唯一ID
  slideId.i              ; 幻灯片ID（用于presentation.xml中的sldId）
  name.s                 ; 幻灯片名称
  partName.s             ; 部件路径（如"/ppt/slides/slide1.xml"）
  rId.s                  ; 与演示文稿的关系ID
  layoutRId.s            ; 与版式的关系ID
  layoutIndex.i          ; 使用的版式索引
  shapes.s               ; 形状列表（JSON格式存储）
EndStructure

; 图表系列结构体
Structure PbPPT_ChartSeries
  name.s                 ; 系列名称
  values.s               ; 数值列表（逗号分隔）
  categories.s           ; 类别列表（逗号分隔）
EndStructure

; 演示文稿结构体
Structure PbPPT_Presentation
  filePath.s             ; 文件路径
  slideWidth.i           ; 幻灯片宽度（EMU）
  slideHeight.i          ; 幻灯片高度（EMU）
  isLoaded.b             ; 是否已加载
  isModified.b           ; 是否已修改
  ; 核心属性
  title.s                ; 文档标题
  subject.s              ; 主题
  creator.s              ; 创建者
  lastModifiedBy.s       ; 最后修改者
  description.s          ; 描述
  keywords.s             ; 关键词
  category.s             ; 分类
  created.s              ; 创建时间
  modified.s             ; 修改时间
  ; 幻灯片相关
  slideCount.i           ; 幻灯片数量
  nextShapeId.i          ; 下一个形状ID
  ; 图片去重
  imageCount.i           ; 图片数量
EndStructure

; ***************************************************************************************
; 分区3: 全局变量和数据存储
; ***************************************************************************************

; 全局演示文稿对象
Global PbPPT_Prs.PbPPT_Presentation

; 全局幻灯片列表
Global NewList PbPPT_Slides.PbPPT_Slide()

; 全局形状列表
Global NewList PbPPT_Shapes.PbPPT_Shape()

; 全局关系列表
Global NewList PbPPT_Rels.PbPPT_Relationship()

; 全局部件列表
Global NewList PbPPT_Parts.PbPPT_Part()

; 全局幻灯片版式列表
Global NewList PbPPT_Layouts.PbPPT_SlideLayout()

; 全局幻灯片母版列表
Global NewList PbPPT_Masters.PbPPT_SlideMaster()

; 全局图片SHA1映射（用于图片去重）
Global NewMap PbPPT_ImageSHA1Map.s()

; 临时目录路径
Global PbPPT_TempDir.s

; ***************************************************************************************
; 分区4: 工具函数
; ***************************************************************************************

; InchesToEmu - 英寸转EMU
Procedure.i PbPPT_InchesToEmu(inches.f)
  ProcedureReturn Int(inches * #PbPPT_EMU_PER_INCH)
EndProcedure

; CmToEmu - 厘米转EMU
Procedure.i PbPPT_CmToEmu(cm.f)
  ProcedureReturn Int(cm * #PbPPT_EMU_PER_CM)
EndProcedure

; MmToEmu - 毫米转EMU
Procedure.i PbPPT_MmToEmu(mm.f)
  ProcedureReturn Int(mm * #PbPPT_EMU_PER_MM)
EndProcedure

; PtToEmu - 磅转EMU
Procedure.i PbPPT_PtToEmu(pt.f)
  ProcedureReturn Int(pt * #PbPPT_EMU_PER_PT)
EndProcedure

; EmuToInches - EMU转英寸
Procedure.f PbPPT_EmuToInches(emu.i)
  ProcedureReturn emu / #PbPPT_EMU_PER_INCH
EndProcedure

; EmuToCm - EMU转厘米
Procedure.f PbPPT_EmuToCm(emu.i)
  ProcedureReturn emu / #PbPPT_EMU_PER_CM
EndProcedure

; EmuToPt - EMU转磅
Procedure.f PbPPT_EmuToPt(emu.i)
  ProcedureReturn emu / #PbPPT_EMU_PER_PT
EndProcedure

; EmuToMm - EMU转毫米
Procedure.f PbPPT_EmuToMm(emu.i)
  ProcedureReturn emu / #PbPPT_EMU_PER_MM
EndProcedure

; EscapeXML - 转义XML特殊字符
Procedure.s PbPPT_EscapeXML(text.s)
  text = ReplaceString(text, "&", "&amp;")
  text = ReplaceString(text, "<", "&lt;")
  text = ReplaceString(text, ">", "&gt;")
  text = ReplaceString(text, ~"\"", "&quot;")
  text = ReplaceString(text, "'", "&apos;")
  ProcedureReturn text
EndProcedure

; UnescapeXML - 反转义XML特殊字符
Procedure.s PbPPT_UnescapeXML(text.s)
  text = ReplaceString(text, "&amp;", "&")
  text = ReplaceString(text, "&lt;", "<")
  text = ReplaceString(text, "&gt;", ">")
  text = ReplaceString(text, "&quot;", ~"\"")
  text = ReplaceString(text, "&apos;", "'")
  ProcedureReturn text
EndProcedure

; GetCurrentDateTime - 获取当前日期时间（ISO 8601格式）
Procedure.s PbPPT_GetCurrentDateTime()
  Define now.i = Date()
  ProcedureReturn FormatDate("%yyyy-%mm-%ddT%hh:%nn:%ssZ", now)
EndProcedure

; GetFileExtension - 获取文件扩展名（小写）
Procedure.s PbPPT_GetFileExtension(path.s)
  Define ext.s = GetExtensionPart(path)
  ProcedureReturn LCase(ext)
EndProcedure

; GetImageContentType - 根据文件扩展名获取图片MIME类型
Procedure.s PbPPT_GetImageContentType(ext.s)
  ext = LCase(ext)
  Select ext
    Case "bmp"
      ProcedureReturn #PbPPT_CT_IMAGE_BMP$
    Case "gif"
      ProcedureReturn #PbPPT_CT_IMAGE_GIF$
    Case "jpg", "jpeg"
      ProcedureReturn #PbPPT_CT_IMAGE_JPEG$
    Case "png"
      ProcedureReturn #PbPPT_CT_IMAGE_PNG$
    Case "tiff", "tif"
      ProcedureReturn #PbPPT_CT_IMAGE_TIFF$
    Case "emf"
      ProcedureReturn #PbPPT_CT_IMAGE_EMF$
    Case "wmf"
      ProcedureReturn #PbPPT_CT_IMAGE_WMF$
    Default
      ProcedureReturn #PbPPT_CT_IMAGE_PNG$
  EndSelect
EndProcedure

; GetImageExtensionFromContentType - 根据MIME类型获取图片扩展名
Procedure.s PbPPT_GetImageExtFromCT(contentType.s)
  Select contentType
    Case #PbPPT_CT_IMAGE_BMP$
      ProcedureReturn "bmp"
    Case #PbPPT_CT_IMAGE_GIF$
      ProcedureReturn "gif"
    Case #PbPPT_CT_IMAGE_JPEG$
      ProcedureReturn "jpeg"
    Case #PbPPT_CT_IMAGE_PNG$
      ProcedureReturn "png"
    Case #PbPPT_CT_IMAGE_TIFF$
      ProcedureReturn "tiff"
    Case #PbPPT_CT_IMAGE_EMF$
      ProcedureReturn "emf"
    Case #PbPPT_CT_IMAGE_WMF$
      ProcedureReturn "wmf"
    Default
      ProcedureReturn "png"
  EndSelect
EndProcedure

; GetNextShapeId - 获取下一个形状ID
Procedure.i PbPPT_GetNextShapeId()
  PbPPT_Prs\nextShapeId + 1
  ProcedureReturn PbPPT_Prs\nextShapeId
EndProcedure

; SHA1哈希相关（简化版，用于图片去重）
; SHA1Init - 初始化SHA1上下文
Procedure PbPPT_SHA1Init(*ctx)
  Define i.i
  For i = 0 To 4
    PokeL(*ctx + i * 4, 0)
  Next
EndProcedure

; CalcSHA1File - 计算文件的SHA1哈希（简化版，使用文件大小+修改时间+MD5作为指纹）
Procedure.s PbPPT_CalcFileFingerprint(path.s)
  Define size.i = FileSize(path)
  Define dateVal.i = GetFileDate(path, #PB_Date_Modified)
  Define md5Val.s = ""
  ; 使用MD5计算文件前4KB的指纹
  If size > 0
    Define readSize.i = size
    If readSize > 4096
      readSize = 4096
    EndIf
    Define *buf = AllocateMemory(readSize)
    If *buf
      Define fh.i = ReadFile(#PB_Any, path)
      If fh
        ReadData(fh, *buf, readSize)
        CloseFile(fh)
        UseMD5Fingerprint()
        md5Val = Fingerprint(*buf, readSize, #PB_Cipher_MD5)
      EndIf
      FreeMemory(*buf)
    EndIf
  EndIf
  ProcedureReturn Hex(size) + "_" + Hex(dateVal) + "_" + md5Val
EndProcedure

; ***************************************************************************************
; 分区5: ZIP辅助模块
; ***************************************************************************************

; ZIPCreate - 创建新的ZIP包文件
Procedure.i PbPPT_ZIPCreate(filename.s)
  ProcedureReturn CreatePack(#PB_Any, filename)
EndProcedure

; ZIPOpen - 打开已有的ZIP包文件
Procedure.i PbPPT_ZIPOpen(filename.s)
  ProcedureReturn OpenPack(#PB_Any, filename)
EndProcedure

; ZIPAddFile - 将磁盘文件添加到ZIP包中
Procedure.b PbPPT_ZIPAddFile(packId.i, sourceFile.s, archivePath.s)
  ProcedureReturn AddPackFile(packId, sourceFile, archivePath)
EndProcedure

; ZIPAddMemory - 将内存数据添加到ZIP包中
Procedure.b PbPPT_ZIPAddMemory(packId.i, *data, dataSize.i, archivePath.s)
  ProcedureReturn AddPackMemory(packId, *data, dataSize, archivePath)
EndProcedure

; ZIPClose - 关闭ZIP包
Procedure PbPPT_ZIPClose(packId.i)
  ClosePack(packId)
EndProcedure

; ZIPExamine - 开始枚举ZIP包内的条目
Procedure.b PbPPT_ZIPExamine(packId.i)
  ProcedureReturn ExaminePack(packId)
EndProcedure

; ZIPNextEntry - 获取下一个ZIP条目
Procedure.b PbPPT_ZIPNextEntry(packId.i)
  ProcedureReturn NextPackEntry(packId)
EndProcedure

; ZIPEntryName - 获取当前ZIP条目的名称（路径）
Procedure.s PbPPT_ZIPEntryName(packId.i)
  ProcedureReturn PackEntryName(packId)
EndProcedure

; ZIPEntrySize - 获取当前ZIP条目的解压大小
Procedure.i PbPPT_ZIPEntrySize(packId.i)
  ProcedureReturn PackEntrySize(packId)
EndProcedure

; ZIPExtractToMemory - 将当前ZIP条目解压到内存缓冲区
Procedure.i PbPPT_ZIPExtractToMemory(packId.i, *buffer, bufferSize.i)
  ProcedureReturn UncompressPackMemory(packId, *buffer, bufferSize)
EndProcedure

; ***************************************************************************************
; 分区6: XML辅助模块
; ***************************************************************************************

; XMLCreateDocument - 创建新的XML文档
Procedure.i PbPPT_XMLCreateDocument()
  ProcedureReturn CreateXML(#PB_Any)
EndProcedure

; XMLAddNode - 在父节点下添加子节点
Procedure.i PbPPT_XMLAddNode(xmlId.i, parentNode.i, nodeName.s)
  ProcedureReturn CreateXMLNode(parentNode, nodeName)
EndProcedure

; XMLSetAttribute - 设置XML节点的属性
Procedure PbPPT_XMLSetAttribute(node.i, attrName.s, attrValue.s)
  SetXMLAttribute(node, attrName, attrValue)
EndProcedure

; XMLSetText - 设置XML节点的文本内容
Procedure PbPPT_XMLSetText(node.i, text.s)
  SetXMLNodeText(node, text)
EndProcedure

; XMLGetAttribute - 获取XML节点的属性值
Procedure.s PbPPT_XMLGetAttribute(node.i, attrName.s)
  ProcedureReturn GetXMLAttribute(node, attrName)
EndProcedure

; XMLGetText - 获取XML节点的文本内容
Procedure.s PbPPT_XMLGetText(node.i)
  ProcedureReturn GetXMLNodeText(node)
EndProcedure

; XMLSaveToString - 将XML文档保存为字符串（通过临时文件中转）
Procedure.s PbPPT_XMLSaveToString(xmlId.i)
  Define tempFile.s = GetTemporaryDirectory() + "PbPPT_temp_" + Str(Random(999999)) + ".xml"
  If SaveXML(xmlId, tempFile)
    Define file.i = ReadFile(#PB_Any, tempFile)
    If file
      Define content.s = ReadString(file, #PB_UTF8)
      CloseFile(file)
      DeleteFile(tempFile)
      ProcedureReturn content
    EndIf
    DeleteFile(tempFile)
  EndIf
  ProcedureReturn ""
EndProcedure

; XMLSaveToFile - 将XML文档保存到文件
Procedure.b PbPPT_XMLSaveToFile(xmlId.i, filename.s)
  ProcedureReturn SaveXML(xmlId, filename)
EndProcedure

; XMLParseString - 从字符串解析XML文档
Procedure.i PbPPT_XMLParseString(xmlString.s)
  ProcedureReturn ParseXML(#PB_Any, xmlString)
EndProcedure

; XMLParseFile - 从文件加载XML文档
Procedure.i PbPPT_XMLParseFile(filename.s)
  ProcedureReturn LoadXML(#PB_Any, filename)
EndProcedure

; XMLFree - 释放XML文档
Procedure PbPPT_XMLFree(xmlId.i)
  FreeXML(xmlId)
EndProcedure

; XMLGetRoot - 获取XML文档的根节点
Procedure.i PbPPT_XMLGetRoot(xmlId.i)
  ProcedureReturn RootXMLNode(xmlId)
EndProcedure

; XMLGetChild - 获取节点的第一个子节点
Procedure.i PbPPT_XMLGetChild(node.i)
  ProcedureReturn ChildXMLNode(node)
EndProcedure

; XMLGetNext - 获取节点的下一个兄弟节点
Procedure.i PbPPT_XMLGetNext(node.i)
  ProcedureReturn NextXMLNode(node)
EndProcedure

; XMLGetNodeName - 获取节点名称
Procedure.s PbPPT_XMLGetNodeName(node.i)
  ProcedureReturn GetXMLNodeName(node)
EndProcedure

; XMLGetAttributeValue - 获取节点属性值（安全版本，属性不存在返回空字符串）
Procedure.s PbPPT_XMLGetAttributeValue(node.i, attrName.s)
  Define result.s = GetXMLAttribute(node, attrName)
  ProcedureReturn result
EndProcedure

; AddXMLToZIP - 将XML文档写入ZIP包（添加UTF-8 BOM头）
Procedure PbPPT_AddXMLToZIP(packId.i, xmlId.i, archivePath.s)
  Define xmlString.s = PbPPT_XMLSaveToString(xmlId)
  If xmlString = ""
    ProcedureReturn
  EndIf
  Define strLen.i = StringByteLength(xmlString, #PB_UTF8)
  ; 缓冲区大小：3字节BOM + UTF8字节数 + 1字节null终止符
  Define bufSize.i = strLen + 3 + 1
  Define *buf = AllocateMemory(bufSize)
  If *buf
    ; 写入UTF-8 BOM (EF BB BF)
    PokeA(*buf, $EF)
    PokeA(*buf + 1, $BB)
    PokeA(*buf + 2, $BF)
    ; 写入XML字符串（PokeS会自动添加null终止符）
    PokeS(*buf + 3, xmlString, -1, #PB_UTF8)
    ; 添加到ZIP时使用strLen+3（不含null终止符）
    PbPPT_ZIPAddMemory(packId, *buf, strLen + 3, archivePath)
    FreeMemory(*buf)
  EndIf
EndProcedure

; AddStringToZIP - 将字符串写入ZIP包（添加UTF-8 BOM头）
Procedure PbPPT_AddStringToZIP(packId.i, text.s, archivePath.s)
  Define strLen.i = StringByteLength(text, #PB_UTF8)
  ; 缓冲区大小：3字节BOM + UTF8字节数 + 1字节null终止符
  Define bufSize.i = strLen + 3 + 1
  Define *buf = AllocateMemory(bufSize)
  If *buf
    ; 写入UTF-8 BOM (EF BB BF)
    PokeA(*buf, $EF)
    PokeA(*buf + 1, $BB)
    PokeA(*buf + 2, $BF)
    ; 写入字符串（PokeS会自动添加null终止符）
    PokeS(*buf + 3, text, -1, #PB_UTF8)
    ; 添加到ZIP时使用strLen+3（不含null终止符）
    PbPPT_ZIPAddMemory(packId, *buf, strLen + 3, archivePath)
    FreeMemory(*buf)
  EndIf
EndProcedure

; AddBinaryToZIP - 将二进制数据写入ZIP包
Procedure PbPPT_AddBinaryToZIP(packId.i, *data, dataSize.i, archivePath.s)
  PbPPT_ZIPAddMemory(packId, *data, dataSize, archivePath)
EndProcedure

; ***************************************************************************************
; 分区7: PackURI模块
; ***************************************************************************************

; PackURIGetBaseURI - 获取包URI的基础路径（目录部分）
; 例如: "/ppt/slides/slide1.xml" -> "/ppt/slides"
Procedure.s PbPPT_PackURIGetBaseURI(partName.s)
  Define lastSlash.i = FindString(partName, "/", 2)
  If lastSlash = 0
    ProcedureReturn "/"
  EndIf
  ; 从后向前查找最后一个斜杠
  Define pos.i = 0
  Define i.i
  For i = 2 To Len(partName)
    If Mid(partName, i, 1) = "/"
      pos = i
    EndIf
  Next
  If pos = 0
    ProcedureReturn "/"
  EndIf
  ProcedureReturn Left(partName, pos - 1)
EndProcedure

; PackURIGetFilename - 获取包URI的文件名部分
; 例如: "/ppt/slides/slide1.xml" -> "slide1.xml"
Procedure.s PbPPT_PackURIGetFilename(partName.s)
  Define pos.i = 0
  Define i.i
  For i = 2 To Len(partName)
    If Mid(partName, i, 1) = "/"
      pos = i
    EndIf
  Next
  If pos = 0
    ProcedureReturn Mid(partName, 2)
  EndIf
  ProcedureReturn Mid(partName, pos + 1)
EndProcedure

; PackURIGetExt - 获取包URI的扩展名部分
; 例如: "/ppt/slides/slide1.xml" -> "xml"
Procedure.s PbPPT_PackURIGetExt(partName.s)
  Define filename.s = PbPPT_PackURIGetFilename(partName)
  ProcedureReturn LCase(GetExtensionPart(filename))
EndProcedure

; PackURIGetMemberName - 获取ZIP包内的成员名（去掉前导斜杠）
; 例如: "/ppt/slides/slide1.xml" -> "ppt/slides/slide1.xml"
Procedure.s PbPPT_PackURIGetMemberName(partName.s)
  If Left(partName, 1) = "/"
    ProcedureReturn Mid(partName, 2)
  EndIf
  ProcedureReturn partName
EndProcedure

; PackURIGetRelsURI - 获取关系文件的URI
; 例如: "/ppt/slides/slide1.xml" -> "/ppt/slides/_rels/slide1.xml.rels"
Procedure.s PbPPT_PackURIGetRelsURI(partName.s)
  Define baseURI.s = PbPPT_PackURIGetBaseURI(partName)
  Define filename.s = PbPPT_PackURIGetFilename(partName)
  ProcedureReturn baseURI + "/_rels/" + filename + ".rels"
EndProcedure

; PackURIFromRelRef - 根据基础URI和相对引用计算绝对URI
; 例如: baseURI="/ppt/slides", relativeRef="../slideLayouts/slideLayout1.xml"
;       -> "/ppt/slideLayouts/slideLayout1.xml"
Procedure.s PbPPT_PackURIFromRelRef(baseURI.s, relativeRef.s)
  ; 处理"../"相对路径
  Define result.s = baseURI + "/" + relativeRef
  ; 简化路径：处理 "../" 和 "./"
  Define changed.b = #True
  While changed
    changed = #False
    Define pos.i = FindString(result, "/../")
    If pos > 0
      ; 找到"/../"前一个路径段
      Define prevSlash.i = 0
      Define j.i
      For j = pos - 1 To 1 Step -1
        If Mid(result, j, 1) = "/"
          prevSlash = j
          Break
        EndIf
      Next
      If prevSlash > 0
        result = Left(result, prevSlash) + Mid(result, pos + 4)
        changed = #True
      EndIf
    EndIf
    ; 处理 "./"
    result = ReplaceString(result, "/./", "/")
  Wend
  ; 确保以斜杠开头
  If Left(result, 1) <> "/"
    result = "/" + result
  EndIf
  ProcedureReturn result
EndProcedure

; PackURIRelativeRef - 计算从baseURI到targetURI的相对引用
; 例如: baseURI="/ppt/slides", targetURI="/ppt/slideLayouts/slideLayout1.xml"
;       -> "../slideLayouts/slideLayout1.xml"
Procedure.s PbPPT_PackURIRelativeRef(baseURI.s, targetURI.s)
  If baseURI = "/"
    ProcedureReturn Mid(targetURI, 2)
  EndIf
  ; 简化实现：计算相对路径
  Define baseParts.s = baseURI + "/"
  Define targetParts.s = targetURI
  ; 找到公共前缀
  Define commonLen.i = 0
  Define minLen.i = Len(baseParts)
  If Len(targetParts) < minLen
    minLen = Len(targetParts)
  EndIf
  Define i.i
  For i = 1 To minLen
    If Mid(baseParts, i, 1) <> Mid(targetParts, i, 1)
      Break
    EndIf
    commonLen = i
  Next
  ; 从公共前缀后找最后一个斜杠
  Define lastCommonSlash.i = 0
  For i = 1 To commonLen
    If Mid(baseParts, i, 1) = "/"
      lastCommonSlash = i
    EndIf
  Next
  ; 计算需要返回的层级数
  Define upCount.i = 0
  For i = lastCommonSlash + 1 To Len(baseURI)
    If Mid(baseURI, i, 1) = "/"
      upCount + 1
    EndIf
  Next
  ; 构建相对路径
  Define relPath.s = ""
  For i = 1 To upCount
    relPath + "../"
  Next
  relPath + Mid(targetURI, lastCommonSlash + 1)
  ProcedureReturn relPath
EndProcedure

; ***************************************************************************************
; 分区8: OPC包模块（读取、写入、内容类型、关系管理）
; ***************************************************************************************

; ClearAllData - 清空所有全局数据
Procedure PbPPT_ClearAllData()
  ClearList(PbPPT_Slides())
  ClearList(PbPPT_Shapes())
  ClearList(PbPPT_Rels())
  ClearList(PbPPT_Parts())
  ClearList(PbPPT_Layouts())
  ClearList(PbPPT_Masters())
  ClearMap(PbPPT_ImageSHA1Map())
  PbPPT_Prs\slideCount = 0
  PbPPT_Prs\nextShapeId = 2
  PbPPT_Prs\isLoaded = #False
  PbPPT_Prs\isModified = #False
  PbPPT_Prs\imageCount = 0
EndProcedure

; AddRelationship - 添加关系
Procedure.s PbPPT_AddRelationship(relType.s, target.s, isExternal.b)
  Define rId.s = "rId" + Str(ListSize(PbPPT_Rels()) + 1)
  AddElement(PbPPT_Rels())
  PbPPT_Rels()\rId = rId
  PbPPT_Rels()\relType = relType
  PbPPT_Rels()\target = target
  PbPPT_Rels()\isExternal = isExternal
  ProcedureReturn rId
EndProcedure

; FindRelationshipByType - 根据关系类型查找关系
Procedure.s PbPPT_FindRelByType(relType.s)
  ForEach PbPPT_Rels()
    If PbPPT_Rels()\relType = relType
      ProcedureReturn PbPPT_Rels()\rId
    EndIf
  Next
  ProcedureReturn ""
EndProcedure

; FindRelationshipTarget - 根据rId查找关系目标
Procedure.s PbPPT_FindRelTarget(rId.s)
  ForEach PbPPT_Rels()
    If PbPPT_Rels()\rId = rId
      ProcedureReturn PbPPT_Rels()\target
    EndIf
  Next
  ProcedureReturn ""
EndProcedure

; AddPart - 添加部件
Procedure PbPPT_AddPart(partName.s, contentType.s, xmlContent.s)
  AddElement(PbPPT_Parts())
  PbPPT_Parts()\partName = partName
  PbPPT_Parts()\contentType = contentType
  PbPPT_Parts()\xmlContent = xmlContent
  PbPPT_Parts()\blobSize = 0
  PbPPT_Parts()\blobData = 0
EndProcedure

; AddBinaryPart - 添加二进制部件（图片等）
Procedure PbPPT_AddBinaryPart(partName.s, contentType.s, *data, dataSize.i)
  AddElement(PbPPT_Parts())
  PbPPT_Parts()\partName = partName
  PbPPT_Parts()\contentType = contentType
  PbPPT_Parts()\xmlContent = ""
  PbPPT_Parts()\blobData = #Null
  PbPPT_Parts()\blobSize = 0
  If dataSize > 0 And *data <> #Null
    PbPPT_Parts()\blobData = AllocateMemory(dataSize)
    If PbPPT_Parts()\blobData
      CopyMemory(*data, PbPPT_Parts()\blobData, dataSize)
      PbPPT_Parts()\blobSize = dataSize
    EndIf
  EndIf
EndProcedure

; FindPart - 根据部件路径查找部件
Procedure.i PbPPT_FindPart(partName.s)
  ForEach PbPPT_Parts()
    If LCase(PbPPT_Parts()\partName) = LCase(partName)
      ProcedureReturn @PbPPT_Parts()
    EndIf
  Next
  ProcedureReturn 0
EndProcedure

; ***************************************************************************************
; 分区9: XML写入器（生成pptx文件各部分XML）
; ***************************************************************************************

; WriteContentTypesXML - 生成[Content_Types].xml
Procedure.i PbPPT_WriteContentTypesXML()
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define typesNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "Types")
  PbPPT_XMLSetAttribute(typesNode, "xmlns", #PbPPT_NS_CT$)
  ; 默认扩展名映射
  Define defaultNode.i
  defaultNode = PbPPT_XMLAddNode(xmlId, typesNode, "Default")
  PbPPT_XMLSetAttribute(defaultNode, "Extension", "xml")
  PbPPT_XMLSetAttribute(defaultNode, "ContentType", #PbPPT_CT_XML$)
  defaultNode = PbPPT_XMLAddNode(xmlId, typesNode, "Default")
  PbPPT_XMLSetAttribute(defaultNode, "Extension", "rels")
  PbPPT_XMLSetAttribute(defaultNode, "ContentType", #PbPPT_CT_OPC_RELS$)
  ; 图片扩展名映射（根据实际使用的图片类型动态添加）
  Define hasJpg.b = #False
  Define hasPng.b = #False
  Define hasBmp.b = #False
  Define hasGif.b = #False
  Define hasTiff.b = #False
  Define hasEmf.b = #False
  Define hasWmf.b = #False
  ForEach PbPPT_Parts()
    If PbPPT_Parts()\blobSize > 0
      Define partExt.s = PbPPT_GetFileExtension(PbPPT_Parts()\partName)
      Select partExt
        Case "jpg", "jpeg"
          hasJpg = #True
        Case "png"
          hasPng = #True
        Case "bmp"
          hasBmp = #True
        Case "gif"
          hasGif = #True
        Case "tiff", "tif"
          hasTiff = #True
        Case "emf"
          hasEmf = #True
        Case "wmf"
          hasWmf = #True
      EndSelect
    EndIf
  Next
  If hasJpg
    defaultNode = PbPPT_XMLAddNode(xmlId, typesNode, "Default")
    PbPPT_XMLSetAttribute(defaultNode, "Extension", "jpeg")
    PbPPT_XMLSetAttribute(defaultNode, "ContentType", #PbPPT_CT_IMAGE_JPEG$)
  EndIf
  If hasPng
    defaultNode = PbPPT_XMLAddNode(xmlId, typesNode, "Default")
    PbPPT_XMLSetAttribute(defaultNode, "Extension", "png")
    PbPPT_XMLSetAttribute(defaultNode, "ContentType", #PbPPT_CT_IMAGE_PNG$)
  EndIf
  If hasBmp
    defaultNode = PbPPT_XMLAddNode(xmlId, typesNode, "Default")
    PbPPT_XMLSetAttribute(defaultNode, "Extension", "bmp")
    PbPPT_XMLSetAttribute(defaultNode, "ContentType", #PbPPT_CT_IMAGE_BMP$)
  EndIf
  If hasGif
    defaultNode = PbPPT_XMLAddNode(xmlId, typesNode, "Default")
    PbPPT_XMLSetAttribute(defaultNode, "Extension", "gif")
    PbPPT_XMLSetAttribute(defaultNode, "ContentType", #PbPPT_CT_IMAGE_GIF$)
  EndIf
  If hasTiff
    defaultNode = PbPPT_XMLAddNode(xmlId, typesNode, "Default")
    PbPPT_XMLSetAttribute(defaultNode, "Extension", "tiff")
    PbPPT_XMLSetAttribute(defaultNode, "ContentType", #PbPPT_CT_IMAGE_TIFF$)
  EndIf
  If hasEmf
    defaultNode = PbPPT_XMLAddNode(xmlId, typesNode, "Default")
    PbPPT_XMLSetAttribute(defaultNode, "Extension", "emf")
    PbPPT_XMLSetAttribute(defaultNode, "ContentType", #PbPPT_CT_IMAGE_EMF$)
  EndIf
  If hasWmf
    defaultNode = PbPPT_XMLAddNode(xmlId, typesNode, "Default")
    PbPPT_XMLSetAttribute(defaultNode, "Extension", "wmf")
    PbPPT_XMLSetAttribute(defaultNode, "ContentType", #PbPPT_CT_IMAGE_WMF$)
  EndIf
  ; 图表内容类型映射
  Define hasChart.b = #False
  ForEach PbPPT_Shapes()
    If PbPPT_Shapes()\shapeType = #PbPPT_ShapeTypeChart
      hasChart = #True
      Break
    EndIf
  Next
  ; 根据幻灯片数量添加Override
  Define overrideNode.i
  overrideNode = PbPPT_XMLAddNode(xmlId, typesNode, "Override")
  PbPPT_XMLSetAttribute(overrideNode, "PartName", "/ppt/presentation.xml")
  PbPPT_XMLSetAttribute(overrideNode, "ContentType", #PbPPT_CT_PRESENTATION_MAIN$)
  overrideNode = PbPPT_XMLAddNode(xmlId, typesNode, "Override")
  PbPPT_XMLSetAttribute(overrideNode, "PartName", "/ppt/slideMasters/slideMaster1.xml")
  PbPPT_XMLSetAttribute(overrideNode, "ContentType", #PbPPT_CT_SLIDE_MASTER$)
  ; 幻灯片版式
  Define layoutCount.i = ListSize(PbPPT_Layouts())
  Define li.i
  For li = 1 To layoutCount
    overrideNode = PbPPT_XMLAddNode(xmlId, typesNode, "Override")
    PbPPT_XMLSetAttribute(overrideNode, "PartName", "/ppt/slideLayouts/slideLayout" + Str(li) + ".xml")
    PbPPT_XMLSetAttribute(overrideNode, "ContentType", #PbPPT_CT_SLIDE_LAYOUT$)
  Next
  ; 幻灯片
  Define slideCount.i = ListSize(PbPPT_Slides())
  Define si.i
  For si = 1 To slideCount
    overrideNode = PbPPT_XMLAddNode(xmlId, typesNode, "Override")
    PbPPT_XMLSetAttribute(overrideNode, "PartName", "/ppt/slides/slide" + Str(si) + ".xml")
    PbPPT_XMLSetAttribute(overrideNode, "ContentType", #PbPPT_CT_SLIDE$)
  Next
  ; 其他固定部件
  overrideNode = PbPPT_XMLAddNode(xmlId, typesNode, "Override")
  PbPPT_XMLSetAttribute(overrideNode, "PartName", "/ppt/presProps.xml")
  PbPPT_XMLSetAttribute(overrideNode, "ContentType", #PbPPT_CT_PRES_PROPS$)
  overrideNode = PbPPT_XMLAddNode(xmlId, typesNode, "Override")
  PbPPT_XMLSetAttribute(overrideNode, "PartName", "/ppt/viewProps.xml")
  PbPPT_XMLSetAttribute(overrideNode, "ContentType", #PbPPT_CT_VIEW_PROPS$)
  overrideNode = PbPPT_XMLAddNode(xmlId, typesNode, "Override")
  PbPPT_XMLSetAttribute(overrideNode, "PartName", "/ppt/theme/theme1.xml")
  PbPPT_XMLSetAttribute(overrideNode, "ContentType", #PbPPT_CT_THEME$)
  overrideNode = PbPPT_XMLAddNode(xmlId, typesNode, "Override")
  PbPPT_XMLSetAttribute(overrideNode, "PartName", "/ppt/tableStyles.xml")
  PbPPT_XMLSetAttribute(overrideNode, "ContentType", #PbPPT_CT_TABLE_STYLES$)
  ; 图表部件
  If hasChart
    ForEach PbPPT_Shapes()
      If PbPPT_Shapes()\shapeType = #PbPPT_ShapeTypeChart
        overrideNode = PbPPT_XMLAddNode(xmlId, typesNode, "Override")
        PbPPT_XMLSetAttribute(overrideNode, "PartName", "/ppt/charts/chart" + Str(PbPPT_Shapes()\id) + ".xml")
        PbPPT_XMLSetAttribute(overrideNode, "ContentType", #PbPPT_CT_CHART$)
      EndIf
    Next
  EndIf
  ; 核心属性
  overrideNode = PbPPT_XMLAddNode(xmlId, typesNode, "Override")
  PbPPT_XMLSetAttribute(overrideNode, "PartName", "/docProps/core.xml")
  PbPPT_XMLSetAttribute(overrideNode, "ContentType", #PbPPT_CT_CORE_PROPS$)
  overrideNode = PbPPT_XMLAddNode(xmlId, typesNode, "Override")
  PbPPT_XMLSetAttribute(overrideNode, "PartName", "/docProps/app.xml")
  PbPPT_XMLSetAttribute(overrideNode, "ContentType", #PbPPT_CT_EXT_PROPS$)
  ProcedureReturn xmlId
EndProcedure

; WriteRootRelsXML - 生成_rels/.rels
Procedure.i PbPPT_WriteRootRelsXML()
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define relsNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "Relationships")
  PbPPT_XMLSetAttribute(relsNode, "xmlns", #PbPPT_NS_REL$)
  Define relNode.i
  relNode = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
  PbPPT_XMLSetAttribute(relNode, "Id", "rId1")
  PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_OFFICE_DOCUMENT$)
  PbPPT_XMLSetAttribute(relNode, "Target", "ppt/presentation.xml")
  relNode = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
  PbPPT_XMLSetAttribute(relNode, "Id", "rId2")
  PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_CORE_PROPERTIES$)
  PbPPT_XMLSetAttribute(relNode, "Target", "docProps/core.xml")
  relNode = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
  PbPPT_XMLSetAttribute(relNode, "Id", "rId3")
  PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_EXT_PROPERTIES$)
  PbPPT_XMLSetAttribute(relNode, "Target", "docProps/app.xml")
  ProcedureReturn xmlId
EndProcedure

; WritePresentationXML - 生成ppt/presentation.xml
Procedure.i PbPPT_WritePresentationXML()
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define presNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "p:presentation")
  PbPPT_XMLSetAttribute(presNode, "xmlns:a", #PbPPT_NS_A$)
  PbPPT_XMLSetAttribute(presNode, "xmlns:r", #PbPPT_NS_R$)
  PbPPT_XMLSetAttribute(presNode, "xmlns:p", #PbPPT_NS_P$)
  PbPPT_XMLSetAttribute(presNode, "saveSubsetFonts", "1")
  PbPPT_XMLSetAttribute(presNode, "autoCompressPictures", "0")
  ; 幻灯片母版ID列表
  Define sldMasterIdLstNode.i = PbPPT_XMLAddNode(xmlId, presNode, "p:sldMasterIdLst")
  Define sldMasterIdNode.i = PbPPT_XMLAddNode(xmlId, sldMasterIdLstNode, "p:sldMasterId")
  PbPPT_XMLSetAttribute(sldMasterIdNode, "id", "2147483648")
  PbPPT_XMLSetAttribute(sldMasterIdNode, "r:id", "rId1")
  ; 幻灯片ID列表
  Define sldIdLstNode.i = PbPPT_XMLAddNode(xmlId, presNode, "p:sldIdLst")
  Define slideIdx.i = 0
  ForEach PbPPT_Slides()
    slideIdx + 1
    Define sldIdNode.i = PbPPT_XMLAddNode(xmlId, sldIdLstNode, "p:sldId")
    PbPPT_XMLSetAttribute(sldIdNode, "id", Str(256 + slideIdx - 1))
    PbPPT_XMLSetAttribute(sldIdNode, "r:id", PbPPT_Slides()\rId)
  Next
  ; 幻灯片尺寸
  Define sldSzNode.i = PbPPT_XMLAddNode(xmlId, presNode, "p:sldSz")
  PbPPT_XMLSetAttribute(sldSzNode, "cx", Str(PbPPT_Prs\slideWidth))
  PbPPT_XMLSetAttribute(sldSzNode, "cy", Str(PbPPT_Prs\slideHeight))
  ; 备注页尺寸（与幻灯片尺寸互换宽高）
  Define notesSzNode.i = PbPPT_XMLAddNode(xmlId, presNode, "p:notesSz")
  PbPPT_XMLSetAttribute(notesSzNode, "cx", Str(PbPPT_Prs\slideHeight))
  PbPPT_XMLSetAttribute(notesSzNode, "cy", Str(PbPPT_Prs\slideWidth))
  ; 默认文本样式
  Define defaultTxtStyleNode.i = PbPPT_XMLAddNode(xmlId, presNode, "p:defaultTextStyle")
  Define defPNode.i
  ; 默认标题样式
  defPNode = PbPPT_XMLAddNode(xmlId, defaultTxtStyleNode, "a:defPPr")
  PbPPT_XMLSetAttribute(defPNode, "lvl", "0")
  ProcedureReturn xmlId
EndProcedure

; WritePresentationRelsXML - 生成ppt/_rels/presentation.xml.rels
Procedure.i PbPPT_WritePresentationRelsXML()
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define relsNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "Relationships")
  PbPPT_XMLSetAttribute(relsNode, "xmlns", #PbPPT_NS_REL$)
  Define relNode.i
  Define rIdIdx.i = 0
  ; 幻灯片母版关系
  rIdIdx + 1
  relNode = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
  PbPPT_XMLSetAttribute(relNode, "Id", "rId" + Str(rIdIdx))
  PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_SLIDE_MASTER$)
  PbPPT_XMLSetAttribute(relNode, "Target", "slideMasters/slideMaster1.xml")
  ; 幻灯片关系
  Define slideIdx.i = 0
  ForEach PbPPT_Slides()
    slideIdx + 1
    rIdIdx + 1
    relNode = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
    PbPPT_XMLSetAttribute(relNode, "Id", "rId" + Str(rIdIdx))
    PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_SLIDE$)
    PbPPT_XMLSetAttribute(relNode, "Target", "slides/slide" + Str(slideIdx) + ".xml")
    PbPPT_Slides()\rId = "rId" + Str(rIdIdx)
  Next
  ; 其他关系
  rIdIdx + 1
  relNode = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
  PbPPT_XMLSetAttribute(relNode, "Id", "rId" + Str(rIdIdx))
  PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_PRES_PROPS$)
  PbPPT_XMLSetAttribute(relNode, "Target", "presProps.xml")
  rIdIdx + 1
  relNode = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
  PbPPT_XMLSetAttribute(relNode, "Id", "rId" + Str(rIdIdx))
  PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_VIEW_PROPS$)
  PbPPT_XMLSetAttribute(relNode, "Target", "viewProps.xml")
  rIdIdx + 1
  relNode = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
  PbPPT_XMLSetAttribute(relNode, "Id", "rId" + Str(rIdIdx))
  PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_TABLE_STYLES$)
  PbPPT_XMLSetAttribute(relNode, "Target", "tableStyles.xml")
  rIdIdx + 1
  relNode = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
  PbPPT_XMLSetAttribute(relNode, "Id", "rId" + Str(rIdIdx))
  PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_THEME$)
  PbPPT_XMLSetAttribute(relNode, "Target", "theme/theme1.xml")
  ProcedureReturn xmlId
EndProcedure

; WriteSlideXML - 生成幻灯片XML
Procedure.i PbPPT_WriteSlideXML(slideIndex.i)
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define sldNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "p:sld")
  PbPPT_XMLSetAttribute(sldNode, "xmlns:a", #PbPPT_NS_A$)
  PbPPT_XMLSetAttribute(sldNode, "xmlns:r", #PbPPT_NS_R$)
  PbPPT_XMLSetAttribute(sldNode, "xmlns:p", #PbPPT_NS_P$)
  ; 通用幻灯片数据
  Define cSldNode.i = PbPPT_XMLAddNode(xmlId, sldNode, "p:cSld")
  ; 形状树
  Define spTreeNode.i = PbPPT_XMLAddNode(xmlId, cSldNode, "p:spTree")
  ; 形状树根节点
  Define nvGrpSpPrNode.i = PbPPT_XMLAddNode(xmlId, spTreeNode, "p:nvGrpSpPr")
  Define cNvPrNode.i = PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode, "p:cNvPr")
  PbPPT_XMLSetAttribute(cNvPrNode, "id", "1")
  PbPPT_XMLSetAttribute(cNvPrNode, "name", "")
  Define cNvGrpSpPrNode.i = PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode, "p:cNvGrpSpPr")
  Define nvPrNode.i = PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode, "p:nvPr")
  Define grpSpPrNode.i = PbPPT_XMLAddNode(xmlId, spTreeNode, "p:grpSpPr")
  Define xfrmNode.i = PbPPT_XMLAddNode(xmlId, grpSpPrNode, "a:xfrm")
  Define offNode.i = PbPPT_XMLAddNode(xmlId, xfrmNode, "a:off")
  PbPPT_XMLSetAttribute(offNode, "x", "0")
  PbPPT_XMLSetAttribute(offNode, "y", "0")
  Define extNode.i = PbPPT_XMLAddNode(xmlId, xfrmNode, "a:ext")
  PbPPT_XMLSetAttribute(extNode, "cx", Str(PbPPT_Prs\slideWidth))
  PbPPT_XMLSetAttribute(extNode, "cy", Str(PbPPT_Prs\slideHeight))
  Define chOffNode.i = PbPPT_XMLAddNode(xmlId, xfrmNode, "a:chOff")
  PbPPT_XMLSetAttribute(chOffNode, "x", "0")
  PbPPT_XMLSetAttribute(chOffNode, "y", "0")
  Define chExtNode.i = PbPPT_XMLAddNode(xmlId, xfrmNode, "a:chExt")
  PbPPT_XMLSetAttribute(chExtNode, "cx", Str(PbPPT_Prs\slideWidth))
  PbPPT_XMLSetAttribute(chExtNode, "cy", Str(PbPPT_Prs\slideHeight))
  ; 添加属于当前幻灯片的所有形状
  ForEach PbPPT_Shapes()
    If PbPPT_Shapes()\id > 0 And PbPPT_Shapes()\slideIndex = slideIndex
      Select PbPPT_Shapes()\shapeType
        Case #PbPPT_ShapeTypeTextbox, #PbPPT_ShapeTypeAutoShape
          ; 添加p:sp元素（文本框或自动形状）
          Define spNode.i = PbPPT_XMLAddNode(xmlId, spTreeNode, "p:sp")
          ; 非可视属性
          Define nvSpPrNode.i = PbPPT_XMLAddNode(xmlId, spNode, "p:nvSpPr")
          Define cNvPrNode2.i = PbPPT_XMLAddNode(xmlId, nvSpPrNode, "p:cNvPr")
          PbPPT_XMLSetAttribute(cNvPrNode2, "id", Str(PbPPT_Shapes()\id))
          PbPPT_XMLSetAttribute(cNvPrNode2, "name", PbPPT_Shapes()\name)
          Define cNvSpPrNode.i = PbPPT_XMLAddNode(xmlId, nvSpPrNode, "p:cNvSpPr")
          If PbPPT_Shapes()\shapeType = #PbPPT_ShapeTypeTextbox
            PbPPT_XMLSetAttribute(cNvSpPrNode, "txBox", "1")
          EndIf
          Define nvPrNode2.i = PbPPT_XMLAddNode(xmlId, nvSpPrNode, "p:nvPr")
          ; 形状属性
          Define spPrNode.i = PbPPT_XMLAddNode(xmlId, spNode, "p:spPr")
          If PbPPT_Shapes()\rotation <> 0
            Define xfrmNode2.i = PbPPT_XMLAddNode(xmlId, spPrNode, "a:xfrm")
            PbPPT_XMLSetAttribute(xfrmNode2, "rot", Str(Int(PbPPT_Shapes()\rotation * 60000)))
            Define offNode2.i = PbPPT_XMLAddNode(xmlId, xfrmNode2, "a:off")
            PbPPT_XMLSetAttribute(offNode2, "x", Str(PbPPT_Shapes()\left))
            PbPPT_XMLSetAttribute(offNode2, "y", Str(PbPPT_Shapes()\top))
            Define extNode2.i = PbPPT_XMLAddNode(xmlId, xfrmNode2, "a:ext")
            PbPPT_XMLSetAttribute(extNode2, "cx", Str(PbPPT_Shapes()\width))
            PbPPT_XMLSetAttribute(extNode2, "cy", Str(PbPPT_Shapes()\height))
          Else
            Define xfrmNode2.i = PbPPT_XMLAddNode(xmlId, spPrNode, "a:xfrm")
            Define offNode2.i = PbPPT_XMLAddNode(xmlId, xfrmNode2, "a:off")
            PbPPT_XMLSetAttribute(offNode2, "x", Str(PbPPT_Shapes()\left))
            PbPPT_XMLSetAttribute(offNode2, "y", Str(PbPPT_Shapes()\top))
            Define extNode2.i = PbPPT_XMLAddNode(xmlId, xfrmNode2, "a:ext")
            PbPPT_XMLSetAttribute(extNode2, "cx", Str(PbPPT_Shapes()\width))
            PbPPT_XMLSetAttribute(extNode2, "cy", Str(PbPPT_Shapes()\height))
          EndIf
          ; 预设几何形状
          Define prstGeomNode.i = PbPPT_XMLAddNode(xmlId, spPrNode, "a:prstGeom")
          PbPPT_XMLSetAttribute(prstGeomNode, "prst", PbPPT_Shapes()\preset)
          Define avLstNode.i = PbPPT_XMLAddNode(xmlId, prstGeomNode, "a:avLst")
          ; 填充
          If PbPPT_Shapes()\fill\fillType = #PbPPT_FillTypeSolid
            Define solidFillNode.i = PbPPT_XMLAddNode(xmlId, spPrNode, "a:solidFill")
            Define srgbClrNode.i = PbPPT_XMLAddNode(xmlId, solidFillNode, "a:srgbClr")
            PbPPT_XMLSetAttribute(srgbClrNode, "val", PbPPT_Shapes()\fill\foreColor)
          ElseIf PbPPT_Shapes()\fill\fillType = #PbPPT_FillTypeNone
            Define noFillNode.i = PbPPT_XMLAddNode(xmlId, spPrNode, "a:noFill")
          EndIf
          ; 线条
          If PbPPT_Shapes()\line\width > 0 Or PbPPT_Shapes()\line\color <> ""
            Define lnNode.i = PbPPT_XMLAddNode(xmlId, spPrNode, "a:ln")
            If PbPPT_Shapes()\line\width > 0
              PbPPT_XMLSetAttribute(lnNode, "w", Str(PbPPT_Shapes()\line\width))
            EndIf
            If PbPPT_Shapes()\line\color <> ""
              Define lnFillNode.i = PbPPT_XMLAddNode(xmlId, lnNode, "a:solidFill")
              Define lnClrNode.i = PbPPT_XMLAddNode(xmlId, lnFillNode, "a:srgbClr")
              PbPPT_XMLSetAttribute(lnClrNode, "val", PbPPT_Shapes()\line\color)
            EndIf
          EndIf
          ; 文本框内容
          Define txBodyNode.i = PbPPT_XMLAddNode(xmlId, spNode, "p:txBody")
          Define bodyPrNode.i = PbPPT_XMLAddNode(xmlId, txBodyNode, "a:bodyPr")
          If PbPPT_Shapes()\shapeType = #PbPPT_ShapeTypeTextbox
            PbPPT_XMLSetAttribute(bodyPrNode, "wrap", "square")
          EndIf
          ; 垂直对齐
          Select PbPPT_Shapes()\textFrame\verticalAnchor
            Case #PbPPT_ValignMiddle
              PbPPT_XMLSetAttribute(bodyPrNode, "anchor", "ctr")
            Case #PbPPT_ValignBottom
              PbPPT_XMLSetAttribute(bodyPrNode, "anchor", "b")
          EndSelect
          ; 文本框边距
          If PbPPT_Shapes()\textFrame\marginLeft > 0
            PbPPT_XMLSetAttribute(bodyPrNode, "lIns", Str(PbPPT_Shapes()\textFrame\marginLeft))
          EndIf
          If PbPPT_Shapes()\textFrame\marginRight > 0
            PbPPT_XMLSetAttribute(bodyPrNode, "rIns", Str(PbPPT_Shapes()\textFrame\marginRight))
          EndIf
          If PbPPT_Shapes()\textFrame\marginTop > 0
            PbPPT_XMLSetAttribute(bodyPrNode, "tIns", Str(PbPPT_Shapes()\textFrame\marginTop))
          EndIf
          If PbPPT_Shapes()\textFrame\marginBottom > 0
            PbPPT_XMLSetAttribute(bodyPrNode, "bIns", Str(PbPPT_Shapes()\textFrame\marginBottom))
          EndIf
          Define lstStyleNode.i = PbPPT_XMLAddNode(xmlId, txBodyNode, "a:lstStyle")
          ; 解析并生成段落
          Define paraText.s = PbPPT_Shapes()\textFrame\paragraphs
          If paraText = ""
            ; 空段落
            Define pNode.i = PbPPT_XMLAddNode(xmlId, txBodyNode, "a:p")
            Define endParaRPrNode.i = PbPPT_XMLAddNode(xmlId, pNode, "a:endParaRPr")
            PbPPT_XMLSetAttribute(endParaRPrNode, "lang", "zh-CN")
          Else
            ; 按#PARA#分割段落
            Define paraCount.i = CountString(paraText, "#PARA#") + 1
            Define pi.i
            For pi = 1 To paraCount
              Define paraLine.s = StringField(paraText, pi, "#PARA#")
              Define pNode2.i = PbPPT_XMLAddNode(xmlId, txBodyNode, "a:p")
              ; 段落属性
              Define pPrNode.i = PbPPT_XMLAddNode(xmlId, pNode2, "a:pPr")
              ; 对齐方式
              Select PbPPT_Shapes()\textFrame\wordWrap
                Case #PbPPT_AlignCenter
                  PbPPT_XMLSetAttribute(pPrNode, "algn", "ctr")
                Case #PbPPT_AlignRight
                  PbPPT_XMLSetAttribute(pPrNode, "algn", "r")
                Case #PbPPT_AlignJustify
                  PbPPT_XMLSetAttribute(pPrNode, "algn", "just")
              EndSelect
              ; 解析Run列表（格式：文本#ATTR#字体名,字号,粗体,斜体,下划线,颜色#RUN#文本...）
              If paraLine <> ""
                Define runCount.i = CountString(paraLine, "#RUN#") + 1
                Define ri2.i
                For ri2 = 1 To runCount
                  Define runPart.s = StringField(paraLine, ri2, "#RUN#")
                  ; 分离文本和属性
                  Define runText.s = runPart
                  Define runAttr.s = ""
                  Define attrPos.i = FindString(runPart, "#ATTR#")
                  If attrPos > 0
                    runText = Left(runPart, attrPos - 1)
                    runAttr = Mid(runPart, attrPos + 6)
                  EndIf
                  Define rNode.i = PbPPT_XMLAddNode(xmlId, pNode2, "a:r")
                  Define rPrNode.i = PbPPT_XMLAddNode(xmlId, rNode, "a:rPr")
                  PbPPT_XMLSetAttribute(rPrNode, "lang", "zh-CN")
                  PbPPT_XMLSetAttribute(rPrNode, "dirty", "0")
                  ; 解析字体属性（格式：字体名,字号,粗体,斜体,下划线,颜色）
                  If runAttr <> ""
                    Define fnName.s = StringField(runAttr, 1, ",")
                    Define fnSize.s = StringField(runAttr, 2, ",")
                    Define fnBold.s = StringField(runAttr, 3, ",")
                    Define fnItalic.s = StringField(runAttr, 4, ",")
                    Define fnUnderline.s = StringField(runAttr, 5, ",")
                    Define fnColor.s = StringField(runAttr, 6, ",")
                    If fnName <> ""
                      Define latinNode.i = PbPPT_XMLAddNode(xmlId, rPrNode, "a:latin")
                      PbPPT_XMLSetAttribute(latinNode, "typeface", fnName)
                      Define eaNode.i = PbPPT_XMLAddNode(xmlId, rPrNode, "a:ea")
                      PbPPT_XMLSetAttribute(eaNode, "typeface", fnName)
                    EndIf
                    If fnSize <> "" And ValF(fnSize) > 0
                      ; 字号以百分之一磅为单位存储
                      PbPPT_XMLSetAttribute(rPrNode, "sz", Str(Int(ValF(fnSize) * 100)))
                    EndIf
                    If fnBold = "1"
                      PbPPT_XMLSetAttribute(rPrNode, "b", "1")
                    EndIf
                    If fnItalic = "1"
                      PbPPT_XMLSetAttribute(rPrNode, "i", "1")
                    EndIf
                    If fnUnderline = "1"
                      PbPPT_XMLSetAttribute(rPrNode, "u", "sng")
                    EndIf
                    If fnColor <> ""
                      Define solidFillNode2.i = PbPPT_XMLAddNode(xmlId, rPrNode, "a:solidFill")
                      Define srgbClrNode2.i = PbPPT_XMLAddNode(xmlId, solidFillNode2, "a:srgbClr")
                      PbPPT_XMLSetAttribute(srgbClrNode2, "val", fnColor)
                    EndIf
                  EndIf
                  Define tNode.i = PbPPT_XMLAddNode(xmlId, rNode, "a:t")
                  PbPPT_XMLSetText(tNode, runText)
                Next
              Else
                ; 空段落也需要一个endParaRPr
                Define endParaRPrNode2.i = PbPPT_XMLAddNode(xmlId, pNode2, "a:endParaRPr")
                PbPPT_XMLSetAttribute(endParaRPrNode2, "lang", "zh-CN")
              EndIf
            Next
          EndIf

        Case #PbPPT_ShapeTypePicture
          ; 添加p:pic元素（图片）
          Define picNode.i = PbPPT_XMLAddNode(xmlId, spTreeNode, "p:pic")
          Define nvPicPrNode.i = PbPPT_XMLAddNode(xmlId, picNode, "p:nvPicPr")
          Define cNvPrNode3.i = PbPPT_XMLAddNode(xmlId, nvPicPrNode, "p:cNvPr")
          PbPPT_XMLSetAttribute(cNvPrNode3, "id", Str(PbPPT_Shapes()\id))
          PbPPT_XMLSetAttribute(cNvPrNode3, "name", PbPPT_Shapes()\name)
          PbPPT_XMLSetAttribute(cNvPrNode3, "descr", PbPPT_Shapes()\imageDesc)
          Define cNvPicPrNode.i = PbPPT_XMLAddNode(xmlId, nvPicPrNode, "p:cNvPicPr")
          PbPPT_XMLSetAttribute(cNvPicPrNode, "preferRelativeResize", "1")
          Define nvPrNode3.i = PbPPT_XMLAddNode(xmlId, nvPicPrNode, "p:nvPr")
          ; 形状属性
          Define spPrNode2.i = PbPPT_XMLAddNode(xmlId, picNode, "p:spPr")
          Define xfrmNode3.i = PbPPT_XMLAddNode(xmlId, spPrNode2, "a:xfrm")
          Define offNode3.i = PbPPT_XMLAddNode(xmlId, xfrmNode3, "a:off")
          PbPPT_XMLSetAttribute(offNode3, "x", Str(PbPPT_Shapes()\left))
          PbPPT_XMLSetAttribute(offNode3, "y", Str(PbPPT_Shapes()\top))
          Define extNode3.i = PbPPT_XMLAddNode(xmlId, xfrmNode3, "a:ext")
          PbPPT_XMLSetAttribute(extNode3, "cx", Str(PbPPT_Shapes()\width))
          PbPPT_XMLSetAttribute(extNode3, "cy", Str(PbPPT_Shapes()\height))
          Define prstGeomNode2.i = PbPPT_XMLAddNode(xmlId, spPrNode2, "a:prstGeom")
          PbPPT_XMLSetAttribute(prstGeomNode2, "prst", "rect")
          Define avLstNode2.i = PbPPT_XMLAddNode(xmlId, prstGeomNode2, "a:avLst")
          ; 图片填充
          Define blipFillNode.i = PbPPT_XMLAddNode(xmlId, picNode, "p:blipFill")
          Define blipNode.i = PbPPT_XMLAddNode(xmlId, blipFillNode, "a:blip")
          PbPPT_XMLSetAttribute(blipNode, "r:embed", PbPPT_Shapes()\imageRId)
          Define stretchNode.i = PbPPT_XMLAddNode(xmlId, blipFillNode, "a:stretch")
          Define fillRectNode.i = PbPPT_XMLAddNode(xmlId, stretchNode, "a:fillRect")

        Case #PbPPT_ShapeTypeTable
          ; 添加p:graphicFrame元素（表格）
          Define gfNode.i = PbPPT_XMLAddNode(xmlId, spTreeNode, "p:graphicFrame")
          Define nvGrpSpPrNode2.i = PbPPT_XMLAddNode(xmlId, gfNode, "p:nvGraphicFramePr")
          Define cNvPrNode4.i = PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode2, "p:cNvPr")
          PbPPT_XMLSetAttribute(cNvPrNode4, "id", Str(PbPPT_Shapes()\id))
          PbPPT_XMLSetAttribute(cNvPrNode4, "name", PbPPT_Shapes()\name)
          Define cNvGrpSpPrNode2.i = PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode2, "p:cNvGraphicFramePr")
          Define nvPrNode4.i = PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode2, "p:nvPr")
          ; 位置
          Define xfrmNode4.i = PbPPT_XMLAddNode(xmlId, gfNode, "p:xfrm")
          Define offNode4.i = PbPPT_XMLAddNode(xmlId, xfrmNode4, "a:off")
          PbPPT_XMLSetAttribute(offNode4, "x", Str(PbPPT_Shapes()\left))
          PbPPT_XMLSetAttribute(offNode4, "y", Str(PbPPT_Shapes()\top))
          Define extNode4.i = PbPPT_XMLAddNode(xmlId, xfrmNode4, "a:ext")
          PbPPT_XMLSetAttribute(extNode4, "cx", Str(PbPPT_Shapes()\width))
          PbPPT_XMLSetAttribute(extNode4, "cy", Str(PbPPT_Shapes()\height))
          ; 图形数据
          Define graphicNode.i = PbPPT_XMLAddNode(xmlId, gfNode, "a:graphic")
          Define graphicDataNode.i = PbPPT_XMLAddNode(xmlId, graphicNode, "a:graphicData")
          PbPPT_XMLSetAttribute(graphicDataNode, "uri", #PbPPT_NS_A$ + "/table")
          ; 表格
          Define tblNode.i = PbPPT_XMLAddNode(xmlId, graphicDataNode, "a:tbl")
          Define rows.i = PbPPT_Shapes()\tableRows
          Define cols.i = PbPPT_Shapes()\tableCols
          ; 表格网格（列宽）
          Define tblGridNode.i = PbPPT_XMLAddNode(xmlId, tblNode, "a:tblGrid")
          Define ci.i
          For ci = 1 To cols
            Define gridColNode.i = PbPPT_XMLAddNode(xmlId, tblGridNode, "a:gridCol")
            Define colWidth.i = PbPPT_Shapes()\width / cols
            ; 尝试从tableColWidths获取自定义列宽
            If PbPPT_Shapes()\tableColWidths <> ""
              Define customWidth.s = StringField(PbPPT_Shapes()\tableColWidths, ci, ",")
              If customWidth <> "" And Val(customWidth) > 0
                colWidth = Val(customWidth)
              EndIf
            EndIf
            PbPPT_XMLSetAttribute(gridColNode, "w", Str(colWidth))
          Next
          ; 表格行
          Define ri.i
          For ri = 1 To rows
            Define trNode.i = PbPPT_XMLAddNode(xmlId, tblNode, "a:tr")
            Define rowHeight.i = PbPPT_Shapes()\height / rows
            ; 尝试从tableRowHeights获取自定义行高
            If PbPPT_Shapes()\tableRowHeights <> ""
              Define customHeight.s = StringField(PbPPT_Shapes()\tableRowHeights, ri, ",")
              If customHeight <> "" And Val(customHeight) > 0
                rowHeight = Val(customHeight)
              EndIf
            EndIf
            PbPPT_XMLSetAttribute(trNode, "h", Str(rowHeight))
            ; 表格单元格
            For ci = 1 To cols
              Define tcNode.i = PbPPT_XMLAddNode(xmlId, trNode, "a:tc")
              ; 从tableCells提取单元格文本（格式：行_列=文本|行_列=文本|...）
              Define cellKey.s = Str(ri) + "_" + Str(ci)
              Define cellText.s = ""
              Define cellData.s = PbPPT_Shapes()\tableCells
              If cellData <> ""
                Define cellPairCount.i = CountString(cellData, "|") + 1
                Define cpi.i
                For cpi = 1 To cellPairCount
                  Define cellPair.s = StringField(cellData, cpi, "|")
                  Define eqPos.i = FindString(cellPair, "=")
                  If eqPos > 0
                    Define pairKey.s = Left(cellPair, eqPos - 1)
                    If pairKey = cellKey
                      cellText = Mid(cellPair, eqPos + 1)
                      Break
                    EndIf
                  EndIf
                Next
              EndIf
              ; 单元格文本体
              Define txBodyNode2.i = PbPPT_XMLAddNode(xmlId, tcNode, "a:txBody")
              Define bodyPrNode2.i = PbPPT_XMLAddNode(xmlId, txBodyNode2, "a:bodyPr")
              Define lstStyleNode2.i = PbPPT_XMLAddNode(xmlId, txBodyNode2, "a:lstStyle")
              Define pNode3.i = PbPPT_XMLAddNode(xmlId, txBodyNode2, "a:p")
              If cellText <> ""
                Define rNode2.i = PbPPT_XMLAddNode(xmlId, pNode3, "a:r")
                Define rPrNode2.i = PbPPT_XMLAddNode(xmlId, rNode2, "a:rPr")
                PbPPT_XMLSetAttribute(rPrNode2, "lang", "zh-CN")
                PbPPT_XMLSetAttribute(rPrNode2, "dirty", "0")
                Define tNode2.i = PbPPT_XMLAddNode(xmlId, rNode2, "a:t")
                PbPPT_XMLSetText(tNode2, cellText)
              Else
                Define endParaRPrNode3.i = PbPPT_XMLAddNode(xmlId, pNode3, "a:endParaRPr")
                PbPPT_XMLSetAttribute(endParaRPrNode3, "lang", "zh-CN")
              EndIf
              Define tcPrNode.i = PbPPT_XMLAddNode(xmlId, tcNode, "a:tcPr")
            Next
          Next

        Case #PbPPT_ShapeTypeChart
          ; 添加p:graphicFrame元素（图表）
          Define gfNode2.i = PbPPT_XMLAddNode(xmlId, spTreeNode, "p:graphicFrame")
          Define nvGrpSpPrNode3.i = PbPPT_XMLAddNode(xmlId, gfNode2, "p:nvGraphicFramePr")
          Define cNvPrNode5.i = PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode3, "p:cNvPr")
          PbPPT_XMLSetAttribute(cNvPrNode5, "id", Str(PbPPT_Shapes()\id))
          PbPPT_XMLSetAttribute(cNvPrNode5, "name", PbPPT_Shapes()\name)
          Define cNvGrpSpPrNode3.i = PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode3, "p:cNvGraphicFramePr")
          Define nvPrNode5.i = PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode3, "p:nvPr")
          ; 位置
          Define xfrmNode5.i = PbPPT_XMLAddNode(xmlId, gfNode2, "p:xfrm")
          Define offNode5.i = PbPPT_XMLAddNode(xmlId, xfrmNode5, "a:off")
          PbPPT_XMLSetAttribute(offNode5, "x", Str(PbPPT_Shapes()\left))
          PbPPT_XMLSetAttribute(offNode5, "y", Str(PbPPT_Shapes()\top))
          Define extNode5.i = PbPPT_XMLAddNode(xmlId, xfrmNode5, "a:ext")
          PbPPT_XMLSetAttribute(extNode5, "cx", Str(PbPPT_Shapes()\width))
          PbPPT_XMLSetAttribute(extNode5, "cy", Str(PbPPT_Shapes()\height))
          ; 图形数据
          Define graphicNode2.i = PbPPT_XMLAddNode(xmlId, gfNode2, "a:graphic")
          Define graphicDataNode2.i = PbPPT_XMLAddNode(xmlId, graphicNode2, "a:graphicData")
          PbPPT_XMLSetAttribute(graphicDataNode2, "uri", #PbPPT_NS_C$)
          Define chartNode.i = PbPPT_XMLAddNode(xmlId, graphicDataNode2, "c:chart")
          PbPPT_XMLSetAttribute(chartNode, "xmlns:c", #PbPPT_NS_C$)
          PbPPT_XMLSetAttribute(chartNode, "xmlns:r", #PbPPT_NS_R$)
          PbPPT_XMLSetAttribute(chartNode, "r:id", PbPPT_Shapes()\chartRId)

        Case #PbPPT_ShapeTypeConnector
          ; 添加p:cxnSp元素（连接符）
          Define cxnSpNode.i = PbPPT_XMLAddNode(xmlId, spTreeNode, "p:cxnSp")
          Define nvCxnSpPrNode.i = PbPPT_XMLAddNode(xmlId, cxnSpNode, "p:nvCxnSpPr")
          Define cNvPrNode6.i = PbPPT_XMLAddNode(xmlId, nvCxnSpPrNode, "p:cNvPr")
          PbPPT_XMLSetAttribute(cNvPrNode6, "id", Str(PbPPT_Shapes()\id))
          PbPPT_XMLSetAttribute(cNvPrNode6, "name", PbPPT_Shapes()\name)
          Define cNvCxnSpPrNode.i = PbPPT_XMLAddNode(xmlId, nvCxnSpPrNode, "p:cNvCxnSpPr")
          Define nvPrNode6.i = PbPPT_XMLAddNode(xmlId, nvCxnSpPrNode, "p:nvPr")
          ; 连接符属性
          Define spPrNode3.i = PbPPT_XMLAddNode(xmlId, cxnSpNode, "p:spPr")
          Define xfrmNode6.i = PbPPT_XMLAddNode(xmlId, spPrNode3, "a:xfrm")
          Define offNode6.i = PbPPT_XMLAddNode(xmlId, xfrmNode6, "a:off")
          PbPPT_XMLSetAttribute(offNode6, "x", Str(PbPPT_Shapes()\left))
          PbPPT_XMLSetAttribute(offNode6, "y", Str(PbPPT_Shapes()\top))
          Define extNode6.i = PbPPT_XMLAddNode(xmlId, xfrmNode6, "a:ext")
          PbPPT_XMLSetAttribute(extNode6, "cx", Str(PbPPT_Shapes()\width))
          PbPPT_XMLSetAttribute(extNode6, "cy", Str(PbPPT_Shapes()\height))
          Define prstGeomNode3.i = PbPPT_XMLAddNode(xmlId, spPrNode3, "a:prstGeom")
          PbPPT_XMLSetAttribute(prstGeomNode3, "prst", "line")
          Define avLstNode3.i = PbPPT_XMLAddNode(xmlId, prstGeomNode3, "a:avLst")
          ; 线条样式
          If PbPPT_Shapes()\line\color <> ""
            Define lnNode2.i = PbPPT_XMLAddNode(xmlId, spPrNode3, "a:ln")
            If PbPPT_Shapes()\line\width > 0
              PbPPT_XMLSetAttribute(lnNode2, "w", Str(PbPPT_Shapes()\line\width))
            EndIf
            Define lnFillNode2.i = PbPPT_XMLAddNode(xmlId, lnNode2, "a:solidFill")
            Define lnClrNode2.i = PbPPT_XMLAddNode(xmlId, lnFillNode2, "a:srgbClr")
            PbPPT_XMLSetAttribute(lnClrNode2, "val", PbPPT_Shapes()\line\color)
          EndIf

      EndSelect
    EndIf
  Next
  ; 颜色映射
  Define clrMapNode.i = PbPPT_XMLAddNode(xmlId, cSldNode, "p:clrMapOvr")
  Define aClrMapNode.i = PbPPT_XMLAddNode(xmlId, clrMapNode, "a:masterClrMapping")
  ProcedureReturn xmlId
EndProcedure

; WriteSlideRelsXML - 生成幻灯片关系文件XML
Procedure.i PbPPT_WriteSlideRelsXML(slideIndex.i, layoutRId.s)
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define relsNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "Relationships")
  PbPPT_XMLSetAttribute(relsNode, "xmlns", #PbPPT_NS_REL$)
  ; 版式关系
  Define relNode.i = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
  PbPPT_XMLSetAttribute(relNode, "Id", "rId1")
  PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_SLIDE_LAYOUT$)
  PbPPT_XMLSetAttribute(relNode, "Target", "../slideLayouts/slideLayout1.xml")
  ; 图片和图表关系（遍历该幻灯片上的形状）
  Define rIdIdx.i = 1
  ForEach PbPPT_Shapes()
    If PbPPT_Shapes()\slideIndex = slideIndex
      If PbPPT_Shapes()\shapeType = #PbPPT_ShapeTypePicture And PbPPT_Shapes()\imageRId <> ""
        rIdIdx + 1
        Define actualRId.s = "rId" + Str(rIdIdx)
        relNode = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
        PbPPT_XMLSetAttribute(relNode, "Id", actualRId)
        PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_IMAGE$)
        ; 从imageRId中提取图片索引，构建正确的media路径
        ; imageRId格式为"rId_img_N"，对应partName为"/ppt/media/imageN.ext"
        Define imgKey.s = PbPPT_Shapes()\imageRId
        Define imgIdxStr.s = Mid(imgKey, 9)
        Define imgIdx.i = Val(imgIdxStr)
        ; 查找对应的Parts获取扩展名
        Define imgExt.s = "png"
        ForEach PbPPT_Parts()
          If FindString(PbPPT_Parts()\partName, "/ppt/media/image" + imgIdxStr + ".") > 0
            imgExt = PbPPT_GetFileExtension(PbPPT_Parts()\partName)
            Break
          EndIf
        Next
        PbPPT_XMLSetAttribute(relNode, "Target", "../media/image" + imgIdxStr + "." + imgExt)
        ; 回写实际的rId到形状数据，以便WriteSlideXML使用
        PbPPT_Shapes()\imageRId = actualRId
      ElseIf PbPPT_Shapes()\shapeType = #PbPPT_ShapeTypeChart And PbPPT_Shapes()\chartRId <> ""
        rIdIdx + 1
        Define chartRId.s = "rId" + Str(rIdIdx)
        relNode = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
        PbPPT_XMLSetAttribute(relNode, "Id", chartRId)
        PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_CHART$)
        ; 图表关系Target指向charts目录下的chart XML文件
        Define chartIdx.i = PbPPT_Shapes()\id
        PbPPT_XMLSetAttribute(relNode, "Target", "../charts/chart" + Str(chartIdx) + ".xml")
        ; 回写实际的rId到形状数据
        PbPPT_Shapes()\chartRId = chartRId
      EndIf
    EndIf
  Next
  ProcedureReturn xmlId
EndProcedure

; WriteSlideMasterXML - 生成幻灯片母版XML
Procedure.i PbPPT_WriteSlideMasterXML()
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define sldMasterNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "p:sldMaster")
  PbPPT_XMLSetAttribute(sldMasterNode, "xmlns:a", #PbPPT_NS_A$)
  PbPPT_XMLSetAttribute(sldMasterNode, "xmlns:r", #PbPPT_NS_R$)
  PbPPT_XMLSetAttribute(sldMasterNode, "xmlns:p", #PbPPT_NS_P$)
  ; 通用幻灯片数据
  Define cSldNode.i = PbPPT_XMLAddNode(xmlId, sldMasterNode, "p:cSld")
  Define bgNode.i = PbPPT_XMLAddNode(xmlId, cSldNode, "p:bg")
  Define bgRefNode.i = PbPPT_XMLAddNode(xmlId, bgNode, "p:bgRef")
  PbPPT_XMLSetAttribute(bgRefNode, "idx", "1001")
  Define schemeClrNode.i = PbPPT_XMLAddNode(xmlId, bgRefNode, "a:schemeClr")
  PbPPT_XMLSetAttribute(schemeClrNode, "val", "bg1")
  ; 形状树
  Define spTreeNode.i = PbPPT_XMLAddNode(xmlId, cSldNode, "p:spTree")
  Define nvGrpSpPrNode.i = PbPPT_XMLAddNode(xmlId, spTreeNode, "p:nvGrpSpPr")
  Define cNvPrNode.i = PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode, "p:cNvPr")
  PbPPT_XMLSetAttribute(cNvPrNode, "id", "1")
  PbPPT_XMLSetAttribute(cNvPrNode, "name", "")
  PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode, "p:cNvGrpSpPr")
  PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode, "p:nvPr")
  Define grpSpPrNode.i = PbPPT_XMLAddNode(xmlId, spTreeNode, "p:grpSpPr")
  ; 版式ID列表
  Define sldLayoutIdLstNode.i = PbPPT_XMLAddNode(xmlId, sldMasterNode, "p:sldLayoutIdLst")
  Define layoutCount.i = ListSize(PbPPT_Layouts())
  Define li.i
  For li = 1 To layoutCount
    Define sldLayoutIdNode.i = PbPPT_XMLAddNode(xmlId, sldLayoutIdLstNode, "p:sldLayoutId")
    PbPPT_XMLSetAttribute(sldLayoutIdNode, "id", Str(2147483649 + li - 1))
    PbPPT_XMLSetAttribute(sldLayoutIdNode, "r:id", "rId" + Str(li + 1))
  Next
  ProcedureReturn xmlId
EndProcedure

; WriteChartXML - 生成嵌入式图表XML文件
; chartShapeId: 图表形状的ID
Procedure.i PbPPT_WriteChartXML(chartShapeId.i)
  ; 查找图表形状数据
  Define foundChart.b = #False
  Define chartType.i
  Define chartTitle.s
  Define chartData.s
  Define chartStyle.i
  Define chartHasLegend.b
  ForEach PbPPT_Shapes()
    If PbPPT_Shapes()\id = chartShapeId And PbPPT_Shapes()\shapeType = #PbPPT_ShapeTypeChart
      chartType = PbPPT_Shapes()\chartType
      chartTitle = PbPPT_Shapes()\chartTitle
      chartData = PbPPT_Shapes()\chartData
      chartStyle = PbPPT_Shapes()\chartStyle
      chartHasLegend = PbPPT_Shapes()\chartHasLegend
      foundChart = #True
      Break
    EndIf
  Next
  If Not foundChart
    ProcedureReturn 0
  EndIf
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  ; chartSpace根元素
  Define chartSpaceNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "c:chartSpace")
  PbPPT_XMLSetAttribute(chartSpaceNode, "xmlns:c", #PbPPT_NS_C$)
  PbPPT_XMLSetAttribute(chartSpaceNode, "xmlns:a", #PbPPT_NS_A$)
  PbPPT_XMLSetAttribute(chartSpaceNode, "xmlns:r", #PbPPT_NS_R$)
  ; 图表样式
  If chartStyle > 0
    Define styleNode.i = PbPPT_XMLAddNode(xmlId, chartSpaceNode, "c:style")
    PbPPT_XMLSetAttribute(styleNode, "val", Str(chartStyle))
  EndIf
  ; chart元素
  Define chartNode.i = PbPPT_XMLAddNode(xmlId, chartSpaceNode, "c:chart")
  ; 自动标题删除标记
  Define autoTitleDelNode.i = PbPPT_XMLAddNode(xmlId, chartNode, "c:autoTitleDeleted")
  PbPPT_XMLSetAttribute(autoTitleDelNode, "val", "0")
  ; 如果有标题，添加标题元素
  If chartTitle <> ""
    Define titleNode.i = PbPPT_XMLAddNode(xmlId, chartNode, "c:title")
    Define titleTxNode.i = PbPPT_XMLAddNode(xmlId, titleNode, "c:tx")
    Define titleRichNode.i = PbPPT_XMLAddNode(xmlId, titleTxNode, "c:rich")
    Define titleBodyPrNode.i = PbPPT_XMLAddNode(xmlId, titleRichNode, "a:bodyPr")
    Define titleLstStyleNode.i = PbPPT_XMLAddNode(xmlId, titleRichNode, "a:lstStyle")
    Define titlePNode.i = PbPPT_XMLAddNode(xmlId, titleRichNode, "a:p")
    Define titlePPrNode.i = PbPPT_XMLAddNode(xmlId, titlePNode, "a:pPr")
    PbPPT_XMLSetAttribute(titlePPrNode, "algn", "ctr")
    Define titleRNode.i = PbPPT_XMLAddNode(xmlId, titlePNode, "a:r")
    Define titleRPrNode.i = PbPPT_XMLAddNode(xmlId, titleRNode, "a:rPr")
    PbPPT_XMLSetAttribute(titleRPrNode, "lang", "zh-CN")
    PbPPT_XMLSetAttribute(titleRPrNode, "dirty", "0")
    PbPPT_XMLSetAttribute(titleRPrNode, "sz", "1400")
    PbPPT_XMLSetAttribute(titleRPrNode, "b", "1")
    Define titleLatinNode.i = PbPPT_XMLAddNode(xmlId, titleRPrNode, "a:latin")
    PbPPT_XMLSetAttribute(titleLatinNode, "typeface", "+mn-lt")
    Define titleEaNode.i = PbPPT_XMLAddNode(xmlId, titleRPrNode, "a:ea")
    PbPPT_XMLSetAttribute(titleEaNode, "typeface", "+mn-ea")
    Define titleTNode.i = PbPPT_XMLAddNode(xmlId, titleRNode, "a:t")
    PbPPT_XMLSetText(titleTNode, chartTitle)
  EndIf
  ; 绘图区
  Define plotAreaNode.i = PbPPT_XMLAddNode(xmlId, chartNode, "c:plotArea")
  ; 解析图表数据（格式：分类1,分类2,...|系列名1:值1,值2,...;系列名2:值1,值2,...）
  Define pipePos.i = FindString(chartData, "|")
  Define catPart.s
  Define seriesPart.s
  If pipePos > 0
    catPart = Left(chartData, pipePos - 1)
    seriesPart = Mid(chartData, pipePos + 1)
  Else
    catPart = chartData
    seriesPart = ""
  EndIf
  Define catCount.i = 0
  If catPart <> ""
    catCount = CountString(catPart, ",") + 1
  EndIf
  Define serCount.i = 0
  If seriesPart <> ""
    serCount = CountString(seriesPart, ";") + 1
  EndIf
  ; 生成坐标轴ID
  Define catAxId.s = Str(2000000000 + chartShapeId * 2)
  Define valAxId.s = Str(2000000000 + chartShapeId * 2 + 1)
  ; 根据图表类型生成xChart元素
  Select chartType
    Case #PbPPT_ChartTypeColumnClustered, #PbPPT_ChartTypeColumnStacked, #PbPPT_ChartTypeColumnStacked100
      Define barChartNode.i = PbPPT_XMLAddNode(xmlId, plotAreaNode, "c:barChart")
      PbPPT_XMLSetAttribute(barChartNode, "xmlns:c", #PbPPT_NS_C$)
      Define barDirNode.i = PbPPT_XMLAddNode(xmlId, barChartNode, "c:barDir")
      PbPPT_XMLSetAttribute(barDirNode, "val", "col")
      Define groupingNode.i = PbPPT_XMLAddNode(xmlId, barChartNode, "c:grouping")
      If chartType = #PbPPT_ChartTypeColumnClustered
        PbPPT_XMLSetAttribute(groupingNode, "val", "clustered")
      ElseIf chartType = #PbPPT_ChartTypeColumnStacked
        PbPPT_XMLSetAttribute(groupingNode, "val", "stacked")
      Else
        PbPPT_XMLSetAttribute(groupingNode, "val", "percentStacked")
      EndIf
      ; 生成系列
      Define si.i
      For si = 1 To serCount
        Define serPart.s = StringField(seriesPart, si, ";")
        Define serName.s = StringField(serPart, 1, ":")
        Define serValues.s = StringField(serPart, 2, ":")
        Define serNode.i = PbPPT_XMLAddNode(xmlId, barChartNode, "c:ser")
        Define idxNode.i = PbPPT_XMLAddNode(xmlId, serNode, "c:idx")
        PbPPT_XMLSetAttribute(idxNode, "val", Str(si - 1))
        Define orderNode.i = PbPPT_XMLAddNode(xmlId, serNode, "c:order")
        PbPPT_XMLSetAttribute(orderNode, "val", Str(si - 1))
        ; 系列名称
        Define txNode.i = PbPPT_XMLAddNode(xmlId, serNode, "c:tx")
        Define strRefNode.i = PbPPT_XMLAddNode(xmlId, txNode, "c:strRef")
        Define fNode.i = PbPPT_XMLAddNode(xmlId, strRefNode, "c:f")
        PbPPT_XMLSetText(fNode, "Sheet1!$" + Chr(65 + si) + "$1")
        Define strCacheNode.i = PbPPT_XMLAddNode(xmlId, strRefNode, "c:strCache")
        Define ptCountNode.i = PbPPT_XMLAddNode(xmlId, strCacheNode, "c:ptCount")
        PbPPT_XMLSetAttribute(ptCountNode, "val", "1")
        Define ptNode.i = PbPPT_XMLAddNode(xmlId, strCacheNode, "c:pt")
        PbPPT_XMLSetAttribute(ptNode, "idx", "0")
        Define vNode.i = PbPPT_XMLAddNode(xmlId, ptNode, "c:v")
        PbPPT_XMLSetText(vNode, serName)
        ; 分类
        Define catNode2.i = PbPPT_XMLAddNode(xmlId, serNode, "c:cat")
        Define catStrRefNode.i = PbPPT_XMLAddNode(xmlId, catNode2, "c:strRef")
        Define catFNode.i = PbPPT_XMLAddNode(xmlId, catStrRefNode, "c:f")
        PbPPT_XMLSetText(catFNode, "Sheet1!$A$2:$A$" + Str(catCount + 1))
        Define catStrCacheNode.i = PbPPT_XMLAddNode(xmlId, catStrRefNode, "c:strCache")
        Define catPtCountNode.i = PbPPT_XMLAddNode(xmlId, catStrCacheNode, "c:ptCount")
        PbPPT_XMLSetAttribute(catPtCountNode, "val", Str(catCount))
        Define ci2.i
        For ci2 = 1 To catCount
          Define catPtNode.i = PbPPT_XMLAddNode(xmlId, catStrCacheNode, "c:pt")
          PbPPT_XMLSetAttribute(catPtNode, "idx", Str(ci2 - 1))
          Define catVNode.i = PbPPT_XMLAddNode(xmlId, catPtNode, "c:v")
          PbPPT_XMLSetText(catVNode, StringField(catPart, ci2, ","))
        Next
        ; 数值
        Define valNode2.i = PbPPT_XMLAddNode(xmlId, serNode, "c:val")
        Define numRefNode.i = PbPPT_XMLAddNode(xmlId, valNode2, "c:numRef")
        Define numFNode.i = PbPPT_XMLAddNode(xmlId, numRefNode, "c:f")
        PbPPT_XMLSetText(numFNode, "Sheet1!$" + Chr(65 + si) + "$2:$" + Chr(65 + si) + "$" + Str(catCount + 1))
        Define numCacheNode.i = PbPPT_XMLAddNode(xmlId, numRefNode, "c:numCache")
        Define formatCodeNode.i = PbPPT_XMLAddNode(xmlId, numCacheNode, "c:formatCode")
        PbPPT_XMLSetText(formatCodeNode, "General")
        Define numPtCountNode.i = PbPPT_XMLAddNode(xmlId, numCacheNode, "c:ptCount")
        PbPPT_XMLSetAttribute(numPtCountNode, "val", Str(catCount))
        Define valCount.i = CountString(serValues, ",") + 1
        Define vi.i
        For vi = 1 To valCount
          Define numPtNode.i = PbPPT_XMLAddNode(xmlId, numCacheNode, "c:pt")
          PbPPT_XMLSetAttribute(numPtNode, "idx", Str(vi - 1))
          Define numVNode.i = PbPPT_XMLAddNode(xmlId, numPtNode, "c:v")
          PbPPT_XMLSetText(numVNode, StringField(serValues, vi, ","))
        Next
      Next
      ; 坐标轴ID
      Define axIdNode1.i = PbPPT_XMLAddNode(xmlId, barChartNode, "c:axId")
      PbPPT_XMLSetAttribute(axIdNode1, "val", catAxId)
      Define axIdNode2.i = PbPPT_XMLAddNode(xmlId, barChartNode, "c:axId")
      PbPPT_XMLSetAttribute(axIdNode2, "val", valAxId)

    Case #PbPPT_ChartTypeBarClustered, #PbPPT_ChartTypeBarStacked, #PbPPT_ChartTypeBarStacked100
      Define barChartNode2.i = PbPPT_XMLAddNode(xmlId, plotAreaNode, "c:barChart")
      PbPPT_XMLSetAttribute(barChartNode2, "xmlns:c", #PbPPT_NS_C$)
      Define barDirNode2.i = PbPPT_XMLAddNode(xmlId, barChartNode2, "c:barDir")
      PbPPT_XMLSetAttribute(barDirNode2, "val", "bar")
      Define groupingNode2.i = PbPPT_XMLAddNode(xmlId, barChartNode2, "c:grouping")
      If chartType = #PbPPT_ChartTypeBarClustered
        PbPPT_XMLSetAttribute(groupingNode2, "val", "clustered")
      ElseIf chartType = #PbPPT_ChartTypeBarStacked
        PbPPT_XMLSetAttribute(groupingNode2, "val", "stacked")
      Else
        PbPPT_XMLSetAttribute(groupingNode2, "val", "percentStacked")
      EndIf
      ; 系列生成（与柱形图相同逻辑）
      Define si3.i
      For si3 = 1 To serCount
        Define serPart3.s = StringField(seriesPart, si3, ";")
        Define serName3.s = StringField(serPart3, 1, ":")
        Define serValues3.s = StringField(serPart3, 2, ":")
        Define serNode3.i = PbPPT_XMLAddNode(xmlId, barChartNode2, "c:ser")
        Define idxNode3.i = PbPPT_XMLAddNode(xmlId, serNode3, "c:idx")
        PbPPT_XMLSetAttribute(idxNode3, "val", Str(si3 - 1))
        Define orderNode3.i = PbPPT_XMLAddNode(xmlId, serNode3, "c:order")
        PbPPT_XMLSetAttribute(orderNode3, "val", Str(si3 - 1))
        Define txNode3.i = PbPPT_XMLAddNode(xmlId, serNode3, "c:tx")
        Define strRefNode3.i = PbPPT_XMLAddNode(xmlId, txNode3, "c:strRef")
        Define fNode3.i = PbPPT_XMLAddNode(xmlId, strRefNode3, "c:f")
        PbPPT_XMLSetText(fNode3, "Sheet1!$" + Chr(65 + si3) + "$1")
        Define strCacheNode3.i = PbPPT_XMLAddNode(xmlId, strRefNode3, "c:strCache")
        Define ptCountNode3.i = PbPPT_XMLAddNode(xmlId, strCacheNode3, "c:ptCount")
        PbPPT_XMLSetAttribute(ptCountNode3, "val", "1")
        Define ptNode3.i = PbPPT_XMLAddNode(xmlId, strCacheNode3, "c:pt")
        PbPPT_XMLSetAttribute(ptNode3, "idx", "0")
        Define vNode3.i = PbPPT_XMLAddNode(xmlId, ptNode3, "c:v")
        PbPPT_XMLSetText(vNode3, serName3)
        Define catNode3.i = PbPPT_XMLAddNode(xmlId, serNode3, "c:cat")
        Define catStrRefNode3.i = PbPPT_XMLAddNode(xmlId, catNode3, "c:strRef")
        Define catFNode3.i = PbPPT_XMLAddNode(xmlId, catStrRefNode3, "c:f")
        PbPPT_XMLSetText(catFNode3, "Sheet1!$A$2:$A$" + Str(catCount + 1))
        Define catStrCacheNode3.i = PbPPT_XMLAddNode(xmlId, catStrRefNode3, "c:strCache")
        Define catPtCountNode3.i = PbPPT_XMLAddNode(xmlId, catStrCacheNode3, "c:ptCount")
        PbPPT_XMLSetAttribute(catPtCountNode3, "val", Str(catCount))
        Define ci3.i
        For ci3 = 1 To catCount
          Define catPtNode3.i = PbPPT_XMLAddNode(xmlId, catStrCacheNode3, "c:pt")
          PbPPT_XMLSetAttribute(catPtNode3, "idx", Str(ci3 - 1))
          Define catVNode3.i = PbPPT_XMLAddNode(xmlId, catPtNode3, "c:v")
          PbPPT_XMLSetText(catVNode3, StringField(catPart, ci3, ","))
        Next
        Define valNode3.i = PbPPT_XMLAddNode(xmlId, serNode3, "c:val")
        Define numRefNode3.i = PbPPT_XMLAddNode(xmlId, valNode3, "c:numRef")
        Define numFNode3.i = PbPPT_XMLAddNode(xmlId, numRefNode3, "c:f")
        PbPPT_XMLSetText(numFNode3, "Sheet1!$" + Chr(65 + si3) + "$2:$" + Chr(65 + si3) + "$" + Str(catCount + 1))
        Define numCacheNode3.i = PbPPT_XMLAddNode(xmlId, numRefNode3, "c:numCache")
        Define formatCodeNode3.i = PbPPT_XMLAddNode(xmlId, numCacheNode3, "c:formatCode")
        PbPPT_XMLSetText(formatCodeNode3, "General")
        Define numPtCountNode3.i = PbPPT_XMLAddNode(xmlId, numCacheNode3, "c:ptCount")
        PbPPT_XMLSetAttribute(numPtCountNode3, "val", Str(catCount))
        Define valCount3.i = CountString(serValues3, ",") + 1
        Define vi3.i
        For vi3 = 1 To valCount3
          Define numPtNode3.i = PbPPT_XMLAddNode(xmlId, numCacheNode3, "c:pt")
          PbPPT_XMLSetAttribute(numPtNode3, "idx", Str(vi3 - 1))
          Define numVNode3.i = PbPPT_XMLAddNode(xmlId, numPtNode3, "c:v")
          PbPPT_XMLSetText(numVNode3, StringField(serValues3, vi3, ","))
        Next
      Next
      Define axIdNode5.i = PbPPT_XMLAddNode(xmlId, barChartNode2, "c:axId")
      PbPPT_XMLSetAttribute(axIdNode5, "val", catAxId)
      Define axIdNode6.i = PbPPT_XMLAddNode(xmlId, barChartNode2, "c:axId")
      PbPPT_XMLSetAttribute(axIdNode6, "val", valAxId)

    Case #PbPPT_ChartTypeLine, #PbPPT_ChartTypeLineMarkers, #PbPPT_ChartTypeLineStacked, #PbPPT_ChartTypeLineStacked100
      Define lineChartNode.i = PbPPT_XMLAddNode(xmlId, plotAreaNode, "c:lineChart")
      PbPPT_XMLSetAttribute(lineChartNode, "xmlns:c", #PbPPT_NS_C$)
      Define lineGroupingNode.i = PbPPT_XMLAddNode(xmlId, lineChartNode, "c:grouping")
      If chartType = #PbPPT_ChartTypeLine Or chartType = #PbPPT_ChartTypeLineMarkers
        PbPPT_XMLSetAttribute(lineGroupingNode, "val", "standard")
      ElseIf chartType = #PbPPT_ChartTypeLineStacked
        PbPPT_XMLSetAttribute(lineGroupingNode, "val", "stacked")
      Else
        PbPPT_XMLSetAttribute(lineGroupingNode, "val", "percentStacked")
      EndIf
      Define varyColorsNode.i = PbPPT_XMLAddNode(xmlId, lineChartNode, "c:varyColors")
      PbPPT_XMLSetAttribute(varyColorsNode, "val", "0")
      ; 系列生成
      Define si4.i
      For si4 = 1 To serCount
        Define serPart4.s = StringField(seriesPart, si4, ";")
        Define serName4.s = StringField(serPart4, 1, ":")
        Define serValues4.s = StringField(serPart4, 2, ":")
        Define serNode4.i = PbPPT_XMLAddNode(xmlId, lineChartNode, "c:ser")
        Define idxNode4.i = PbPPT_XMLAddNode(xmlId, serNode4, "c:idx")
        PbPPT_XMLSetAttribute(idxNode4, "val", Str(si4 - 1))
        Define orderNode4.i = PbPPT_XMLAddNode(xmlId, serNode4, "c:order")
        PbPPT_XMLSetAttribute(orderNode4, "val", Str(si4 - 1))
        Define txNode4.i = PbPPT_XMLAddNode(xmlId, serNode4, "c:tx")
        Define strRefNode4.i = PbPPT_XMLAddNode(xmlId, txNode4, "c:strRef")
        Define fNode4.i = PbPPT_XMLAddNode(xmlId, strRefNode4, "c:f")
        PbPPT_XMLSetText(fNode4, "Sheet1!$" + Chr(65 + si4) + "$1")
        Define strCacheNode4.i = PbPPT_XMLAddNode(xmlId, strRefNode4, "c:strCache")
        Define ptCountNode4.i = PbPPT_XMLAddNode(xmlId, strCacheNode4, "c:ptCount")
        PbPPT_XMLSetAttribute(ptCountNode4, "val", "1")
        Define ptNode4.i = PbPPT_XMLAddNode(xmlId, strCacheNode4, "c:pt")
        PbPPT_XMLSetAttribute(ptNode4, "idx", "0")
        Define vNode4.i = PbPPT_XMLAddNode(xmlId, ptNode4, "c:v")
        PbPPT_XMLSetText(vNode4, serName4)
        If chartType = #PbPPT_ChartTypeLine
          Define markerNode.i = PbPPT_XMLAddNode(xmlId, serNode4, "c:marker")
          Define symbolNode.i = PbPPT_XMLAddNode(xmlId, markerNode, "c:symbol")
          PbPPT_XMLSetAttribute(symbolNode, "val", "none")
        EndIf
        Define catNode4.i = PbPPT_XMLAddNode(xmlId, serNode4, "c:cat")
        Define catStrRefNode4.i = PbPPT_XMLAddNode(xmlId, catNode4, "c:strRef")
        Define catFNode4.i = PbPPT_XMLAddNode(xmlId, catStrRefNode4, "c:f")
        PbPPT_XMLSetText(catFNode4, "Sheet1!$A$2:$A$" + Str(catCount + 1))
        Define catStrCacheNode4.i = PbPPT_XMLAddNode(xmlId, catStrRefNode4, "c:strCache")
        Define catPtCountNode4.i = PbPPT_XMLAddNode(xmlId, catStrCacheNode4, "c:ptCount")
        PbPPT_XMLSetAttribute(catPtCountNode4, "val", Str(catCount))
        Define ci4.i
        For ci4 = 1 To catCount
          Define catPtNode4.i = PbPPT_XMLAddNode(xmlId, catStrCacheNode4, "c:pt")
          PbPPT_XMLSetAttribute(catPtNode4, "idx", Str(ci4 - 1))
          Define catVNode4.i = PbPPT_XMLAddNode(xmlId, catPtNode4, "c:v")
          PbPPT_XMLSetText(catVNode4, StringField(catPart, ci4, ","))
        Next
        Define valNode4.i = PbPPT_XMLAddNode(xmlId, serNode4, "c:val")
        Define numRefNode4.i = PbPPT_XMLAddNode(xmlId, valNode4, "c:numRef")
        Define numFNode4.i = PbPPT_XMLAddNode(xmlId, numRefNode4, "c:f")
        PbPPT_XMLSetText(numFNode4, "Sheet1!$" + Chr(65 + si4) + "$2:$" + Chr(65 + si4) + "$" + Str(catCount + 1))
        Define numCacheNode4.i = PbPPT_XMLAddNode(xmlId, numRefNode4, "c:numCache")
        Define formatCodeNode4.i = PbPPT_XMLAddNode(xmlId, numCacheNode4, "c:formatCode")
        PbPPT_XMLSetText(formatCodeNode4, "General")
        Define numPtCountNode4.i = PbPPT_XMLAddNode(xmlId, numCacheNode4, "c:ptCount")
        PbPPT_XMLSetAttribute(numPtCountNode4, "val", Str(catCount))
        Define valCount4.i = CountString(serValues4, ",") + 1
        Define vi4.i
        For vi4 = 1 To valCount4
          Define numPtNode4.i = PbPPT_XMLAddNode(xmlId, numCacheNode4, "c:pt")
          PbPPT_XMLSetAttribute(numPtNode4, "idx", Str(vi4 - 1))
          Define numVNode4.i = PbPPT_XMLAddNode(xmlId, numPtNode4, "c:v")
          PbPPT_XMLSetText(numVNode4, StringField(serValues4, vi4, ","))
        Next
        Define smoothNode.i = PbPPT_XMLAddNode(xmlId, serNode4, "c:smooth")
        PbPPT_XMLSetAttribute(smoothNode, "val", "0")
      Next
      Define axIdNode7.i = PbPPT_XMLAddNode(xmlId, lineChartNode, "c:axId")
      PbPPT_XMLSetAttribute(axIdNode7, "val", catAxId)
      Define axIdNode8.i = PbPPT_XMLAddNode(xmlId, lineChartNode, "c:axId")
      PbPPT_XMLSetAttribute(axIdNode8, "val", valAxId)

    Case #PbPPT_ChartTypePie, #PbPPT_ChartTypePieExploded
      Define pieChartNode.i = PbPPT_XMLAddNode(xmlId, plotAreaNode, "c:pieChart")
      PbPPT_XMLSetAttribute(pieChartNode, "xmlns:c", #PbPPT_NS_C$)
      Define varyColorsNode2.i = PbPPT_XMLAddNode(xmlId, pieChartNode, "c:varyColors")
      PbPPT_XMLSetAttribute(varyColorsNode2, "val", "1")
      ; 饼图系列（通常只有一个系列）
      Define si5.i
      For si5 = 1 To serCount
        Define serPart5.s = StringField(seriesPart, si5, ";")
        Define serName5.s = StringField(serPart5, 1, ":")
        Define serValues5.s = StringField(serPart5, 2, ":")
        Define serNode5.i = PbPPT_XMLAddNode(xmlId, pieChartNode, "c:ser")
        Define idxNode5.i = PbPPT_XMLAddNode(xmlId, serNode5, "c:idx")
        PbPPT_XMLSetAttribute(idxNode5, "val", Str(si5 - 1))
        Define orderNode5.i = PbPPT_XMLAddNode(xmlId, serNode5, "c:order")
        PbPPT_XMLSetAttribute(orderNode5, "val", Str(si5 - 1))
        Define txNode5.i = PbPPT_XMLAddNode(xmlId, serNode5, "c:tx")
        Define strRefNode5.i = PbPPT_XMLAddNode(xmlId, txNode5, "c:strRef")
        Define fNode5.i = PbPPT_XMLAddNode(xmlId, strRefNode5, "c:f")
        PbPPT_XMLSetText(fNode5, "Sheet1!$" + Chr(65 + si5) + "$1")
        Define strCacheNode5.i = PbPPT_XMLAddNode(xmlId, strRefNode5, "c:strCache")
        Define ptCountNode5.i = PbPPT_XMLAddNode(xmlId, strCacheNode5, "c:ptCount")
        PbPPT_XMLSetAttribute(ptCountNode5, "val", "1")
        Define ptNode5.i = PbPPT_XMLAddNode(xmlId, strCacheNode5, "c:pt")
        PbPPT_XMLSetAttribute(ptNode5, "idx", "0")
        Define vNode5.i = PbPPT_XMLAddNode(xmlId, ptNode5, "c:v")
        PbPPT_XMLSetText(vNode5, serName5)
        If chartType = #PbPPT_ChartTypePieExploded
          Define explosionNode.i = PbPPT_XMLAddNode(xmlId, serNode5, "c:explosion")
          PbPPT_XMLSetAttribute(explosionNode, "val", "25")
        EndIf
        Define catNode5.i = PbPPT_XMLAddNode(xmlId, serNode5, "c:cat")
        Define catStrRefNode5.i = PbPPT_XMLAddNode(xmlId, catNode5, "c:strRef")
        Define catFNode5.i = PbPPT_XMLAddNode(xmlId, catStrRefNode5, "c:f")
        PbPPT_XMLSetText(catFNode5, "Sheet1!$A$2:$A$" + Str(catCount + 1))
        Define catStrCacheNode5.i = PbPPT_XMLAddNode(xmlId, catStrRefNode5, "c:strCache")
        Define catPtCountNode5.i = PbPPT_XMLAddNode(xmlId, catStrCacheNode5, "c:ptCount")
        PbPPT_XMLSetAttribute(catPtCountNode5, "val", Str(catCount))
        Define ci5.i
        For ci5 = 1 To catCount
          Define catPtNode5.i = PbPPT_XMLAddNode(xmlId, catStrCacheNode5, "c:pt")
          PbPPT_XMLSetAttribute(catPtNode5, "idx", Str(ci5 - 1))
          Define catVNode5.i = PbPPT_XMLAddNode(xmlId, catPtNode5, "c:v")
          PbPPT_XMLSetText(catVNode5, StringField(catPart, ci5, ","))
        Next
        Define valNode5.i = PbPPT_XMLAddNode(xmlId, serNode5, "c:val")
        Define numRefNode5.i = PbPPT_XMLAddNode(xmlId, valNode5, "c:numRef")
        Define numFNode5.i = PbPPT_XMLAddNode(xmlId, numRefNode5, "c:f")
        PbPPT_XMLSetText(numFNode5, "Sheet1!$" + Chr(65 + si5) + "$2:$" + Chr(65 + si5) + "$" + Str(catCount + 1))
        Define numCacheNode5.i = PbPPT_XMLAddNode(xmlId, numRefNode5, "c:numCache")
        Define formatCodeNode5.i = PbPPT_XMLAddNode(xmlId, numCacheNode5, "c:formatCode")
        PbPPT_XMLSetText(formatCodeNode5, "General")
        Define numPtCountNode5.i = PbPPT_XMLAddNode(xmlId, numCacheNode5, "c:ptCount")
        PbPPT_XMLSetAttribute(numPtCountNode5, "val", Str(catCount))
        Define valCount5.i = CountString(serValues5, ",") + 1
        Define vi5.i
        For vi5 = 1 To valCount5
          Define numPtNode5.i = PbPPT_XMLAddNode(xmlId, numCacheNode5, "c:pt")
          PbPPT_XMLSetAttribute(numPtNode5, "idx", Str(vi5 - 1))
          Define numVNode5.i = PbPPT_XMLAddNode(xmlId, numPtNode5, "c:v")
          PbPPT_XMLSetText(numVNode5, StringField(serValues5, vi5, ","))
        Next
      Next

    Default
      ; 默认使用柱形图
      Define defBarChartNode.i = PbPPT_XMLAddNode(xmlId, plotAreaNode, "c:barChart")
      PbPPT_XMLSetAttribute(defBarChartNode, "xmlns:c", #PbPPT_NS_C$)
      Define defBarDirNode.i = PbPPT_XMLAddNode(xmlId, defBarChartNode, "c:barDir")
      PbPPT_XMLSetAttribute(defBarDirNode, "val", "col")
      Define defGroupingNode.i = PbPPT_XMLAddNode(xmlId, defBarChartNode, "c:grouping")
      PbPPT_XMLSetAttribute(defGroupingNode, "val", "clustered")
      Define defAxId1.i = PbPPT_XMLAddNode(xmlId, defBarChartNode, "c:axId")
      PbPPT_XMLSetAttribute(defAxId1, "val", catAxId)
      Define defAxId2.i = PbPPT_XMLAddNode(xmlId, defBarChartNode, "c:axId")
      PbPPT_XMLSetAttribute(defAxId2, "val", valAxId)
  EndSelect
  ; 坐标轴（饼图和圆环图不需要坐标轴）
  If chartType <> #PbPPT_ChartTypePie And chartType <> #PbPPT_ChartTypePieExploded And chartType <> #PbPPT_ChartTypeDoughnut And chartType <> #PbPPT_ChartTypeDoughnutExploded
    ; 分类轴
    Define catAxNode.i = PbPPT_XMLAddNode(xmlId, plotAreaNode, "c:catAx")
    Define catAxIdNode.i = PbPPT_XMLAddNode(xmlId, catAxNode, "c:axId")
    PbPPT_XMLSetAttribute(catAxIdNode, "val", catAxId)
    Define catScalingNode.i = PbPPT_XMLAddNode(xmlId, catAxNode, "c:scaling")
    Define catOrientationNode.i = PbPPT_XMLAddNode(xmlId, catScalingNode, "c:orientation")
    PbPPT_XMLSetAttribute(catOrientationNode, "val", "minMax")
    Define catDeleteNode.i = PbPPT_XMLAddNode(xmlId, catAxNode, "c:delete")
    PbPPT_XMLSetAttribute(catDeleteNode, "val", "0")
    Define catAxPosNode.i = PbPPT_XMLAddNode(xmlId, catAxNode, "c:axPos")
    PbPPT_XMLSetAttribute(catAxPosNode, "val", "b")
    Define catMajorTickNode.i = PbPPT_XMLAddNode(xmlId, catAxNode, "c:majorTickMark")
    PbPPT_XMLSetAttribute(catMajorTickNode, "val", "out")
    Define catMinorTickNode.i = PbPPT_XMLAddNode(xmlId, catAxNode, "c:minorTickMark")
    PbPPT_XMLSetAttribute(catMinorTickNode, "val", "none")
    Define catTickLblPosNode.i = PbPPT_XMLAddNode(xmlId, catAxNode, "c:tickLblPos")
    PbPPT_XMLSetAttribute(catTickLblPosNode, "val", "nextTo")
    Define catCrossAxNode.i = PbPPT_XMLAddNode(xmlId, catAxNode, "c:crossAx")
    PbPPT_XMLSetAttribute(catCrossAxNode, "val", valAxId)
    Define catCrossesNode.i = PbPPT_XMLAddNode(xmlId, catAxNode, "c:crosses")
    PbPPT_XMLSetAttribute(catCrossesNode, "val", "autoZero")
    Define catAutoNode.i = PbPPT_XMLAddNode(xmlId, catAxNode, "c:auto")
    PbPPT_XMLSetAttribute(catAutoNode, "val", "1")
    Define catLblAlgnNode.i = PbPPT_XMLAddNode(xmlId, catAxNode, "c:lblAlgn")
    PbPPT_XMLSetAttribute(catLblAlgnNode, "val", "ctr")
    Define catLblOffsetNode.i = PbPPT_XMLAddNode(xmlId, catAxNode, "c:lblOffset")
    PbPPT_XMLSetAttribute(catLblOffsetNode, "val", "100")
    ; 数值轴
    Define valAxNode.i = PbPPT_XMLAddNode(xmlId, plotAreaNode, "c:valAx")
    Define valAxIdNode.i = PbPPT_XMLAddNode(xmlId, valAxNode, "c:axId")
    PbPPT_XMLSetAttribute(valAxIdNode, "val", valAxId)
    Define valScalingNode.i = PbPPT_XMLAddNode(xmlId, valAxNode, "c:scaling")
    Define valOrientationNode.i = PbPPT_XMLAddNode(xmlId, valScalingNode, "c:orientation")
    PbPPT_XMLSetAttribute(valOrientationNode, "val", "minMax")
    Define valDeleteNode.i = PbPPT_XMLAddNode(xmlId, valAxNode, "c:delete")
    PbPPT_XMLSetAttribute(valDeleteNode, "val", "0")
    Define valAxPosNode.i = PbPPT_XMLAddNode(xmlId, valAxNode, "c:axPos")
    PbPPT_XMLSetAttribute(valAxPosNode, "val", "l")
    Define valMajorGridlinesNode.i = PbPPT_XMLAddNode(xmlId, valAxNode, "c:majorGridlines")
    Define valNumFmtNode.i = PbPPT_XMLAddNode(xmlId, valAxNode, "c:numFmt")
    PbPPT_XMLSetAttribute(valNumFmtNode, "formatCode", "General")
    PbPPT_XMLSetAttribute(valNumFmtNode, "sourceLinked", "1")
    Define valMajorTickNode.i = PbPPT_XMLAddNode(xmlId, valAxNode, "c:majorTickMark")
    PbPPT_XMLSetAttribute(valMajorTickNode, "val", "out")
    Define valMinorTickNode.i = PbPPT_XMLAddNode(xmlId, valAxNode, "c:minorTickMark")
    PbPPT_XMLSetAttribute(valMinorTickNode, "val", "none")
    Define valTickLblPosNode.i = PbPPT_XMLAddNode(xmlId, valAxNode, "c:tickLblPos")
    PbPPT_XMLSetAttribute(valTickLblPosNode, "val", "nextTo")
    Define valCrossAxNode.i = PbPPT_XMLAddNode(xmlId, valAxNode, "c:crossAx")
    PbPPT_XMLSetAttribute(valCrossAxNode, "val", catAxId)
    Define valCrossesNode.i = PbPPT_XMLAddNode(xmlId, valAxNode, "c:crosses")
    PbPPT_XMLSetAttribute(valCrossesNode, "val", "autoZero")
    Define valCrossBetweenNode.i = PbPPT_XMLAddNode(xmlId, valAxNode, "c:crossBetween")
    PbPPT_XMLSetAttribute(valCrossBetweenNode, "val", "midCat")
  EndIf
  ; 图例
  If chartHasLegend
    Define legendNode.i = PbPPT_XMLAddNode(xmlId, chartNode, "c:legend")
    Define legendPosNode.i = PbPPT_XMLAddNode(xmlId, legendNode, "c:legendPos")
    PbPPT_XMLSetAttribute(legendPosNode, "val", "r")
    Define legendLayoutNode.i = PbPPT_XMLAddNode(xmlId, legendNode, "c:layout")
    Define legendOverlayNode.i = PbPPT_XMLAddNode(xmlId, legendNode, "c:overlay")
    PbPPT_XMLSetAttribute(legendOverlayNode, "val", "0")
  EndIf
  ; 绘图区属性
  Define plotVisOnlyNode.i = PbPPT_XMLAddNode(xmlId, chartNode, "c:plotVisOnly")
  PbPPT_XMLSetAttribute(plotVisOnlyNode, "val", "1")
  Define dispBlanksAsNode.i = PbPPT_XMLAddNode(xmlId, chartNode, "c:dispBlanksAs")
  PbPPT_XMLSetAttribute(dispBlanksAsNode, "val", "gap")
  ; 默认文本属性
  Define txPrNode.i = PbPPT_XMLAddNode(xmlId, chartSpaceNode, "c:txPr")
  Define txPrBodyPrNode.i = PbPPT_XMLAddNode(xmlId, txPrNode, "a:bodyPr")
  Define txPrLstStyleNode.i = PbPPT_XMLAddNode(xmlId, txPrNode, "a:lstStyle")
  Define txPrPNode.i = PbPPT_XMLAddNode(xmlId, txPrNode, "a:p")
  Define txPrPPrNode.i = PbPPT_XMLAddNode(xmlId, txPrPNode, "a:pPr")
  Define txPrDefRPrNode.i = PbPPT_XMLAddNode(xmlId, txPrPPrNode, "a:defRPr")
  PbPPT_XMLSetAttribute(txPrDefRPrNode, "sz", "1800")
  ProcedureReturn xmlId
EndProcedure

; WriteSlideMasterRelsXML - 生成幻灯片母版关系文件XML
Procedure.i PbPPT_WriteSlideMasterRelsXML()
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define relsNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "Relationships")
  PbPPT_XMLSetAttribute(relsNode, "xmlns", #PbPPT_NS_REL$)
  ; 主题关系
  Define relNode.i = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
  PbPPT_XMLSetAttribute(relNode, "Id", "rId1")
  PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_THEME$)
  PbPPT_XMLSetAttribute(relNode, "Target", "../theme/theme1.xml")
  ; 版式关系
  Define layoutCount.i = ListSize(PbPPT_Layouts())
  Define li.i
  For li = 1 To layoutCount
    relNode = PbPPT_XMLAddNode(xmlId, relsNode, "Relationship")
    PbPPT_XMLSetAttribute(relNode, "Id", "rId" + Str(li + 1))
    PbPPT_XMLSetAttribute(relNode, "Type", #PbPPT_RT_SLIDE_LAYOUT$)
    PbPPT_XMLSetAttribute(relNode, "Target", "../slideLayouts/slideLayout" + Str(li) + ".xml")
  Next
  ProcedureReturn xmlId
EndProcedure

; WriteSlideLayoutXML - 生成幻灯片版式XML
Procedure.i PbPPT_WriteSlideLayoutXML(layoutName.s, layoutIndex.i)
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define sldLayoutNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "p:sldLayout")
  PbPPT_XMLSetAttribute(sldLayoutNode, "xmlns:a", #PbPPT_NS_A$)
  PbPPT_XMLSetAttribute(sldLayoutNode, "xmlns:r", #PbPPT_NS_R$)
  PbPPT_XMLSetAttribute(sldLayoutNode, "xmlns:p", #PbPPT_NS_P$)
  PbPPT_XMLSetAttribute(sldLayoutNode, "type", layoutName)
  PbPPT_XMLSetAttribute(sldLayoutNode, "preserve", "1")
  ; 通用幻灯片数据
  Define cSldNode.i = PbPPT_XMLAddNode(xmlId, sldLayoutNode, "p:cSld")
  PbPPT_XMLSetAttribute(cSldNode, "name", layoutName)
  ; 形状树
  Define spTreeNode.i = PbPPT_XMLAddNode(xmlId, cSldNode, "p:spTree")
  Define nvGrpSpPrNode.i = PbPPT_XMLAddNode(xmlId, spTreeNode, "p:nvGrpSpPr")
  Define cNvPrNode.i = PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode, "p:cNvPr")
  PbPPT_XMLSetAttribute(cNvPrNode, "id", "1")
  PbPPT_XMLSetAttribute(cNvPrNode, "name", "")
  PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode, "p:cNvGrpSpPr")
  PbPPT_XMLAddNode(xmlId, nvGrpSpPrNode, "p:nvPr")
  PbPPT_XMLAddNode(xmlId, spTreeNode, "p:grpSpPr")
  ; 根据版式类型添加占位符
  If layoutName = "title" Or layoutName = "ctrTitle"
    ; 标题占位符
    Define spNode.i = PbPPT_XMLAddNode(xmlId, spTreeNode, "p:sp")
    Define nvSpPrNode.i = PbPPT_XMLAddNode(xmlId, spNode, "p:nvSpPr")
    Define cNvPrNode2.i = PbPPT_XMLAddNode(xmlId, nvSpPrNode, "p:cNvPr")
    PbPPT_XMLSetAttribute(cNvPrNode2, "id", "2")
    PbPPT_XMLSetAttribute(cNvPrNode2, "name", "Title 1")
    PbPPT_XMLAddNode(xmlId, nvSpPrNode, "p:cNvSpPr")
    Define nvPrNode.i = PbPPT_XMLAddNode(xmlId, nvSpPrNode, "p:nvPr")
    Define phNode.i = PbPPT_XMLAddNode(xmlId, nvPrNode, "p:ph")
    PbPPT_XMLSetAttribute(phNode, "type", "ctrTitle")
  ElseIf layoutName = "obj"
    ; 标题和内容占位符
    Define spNode2.i = PbPPT_XMLAddNode(xmlId, spTreeNode, "p:sp")
    Define nvSpPrNode2.i = PbPPT_XMLAddNode(xmlId, spNode2, "p:nvSpPr")
    Define cNvPrNode3.i = PbPPT_XMLAddNode(xmlId, nvSpPrNode2, "p:cNvPr")
    PbPPT_XMLSetAttribute(cNvPrNode3, "id", "2")
    PbPPT_XMLSetAttribute(cNvPrNode3, "name", "Title 1")
    PbPPT_XMLAddNode(xmlId, nvSpPrNode2, "p:cNvSpPr")
    Define nvPrNode2.i = PbPPT_XMLAddNode(xmlId, nvSpPrNode2, "p:nvPr")
    Define phNode2.i = PbPPT_XMLAddNode(xmlId, nvPrNode2, "p:ph")
    PbPPT_XMLSetAttribute(phNode2, "type", "title")
  EndIf
  ; 颜色映射
  Define clrMapNode.i = PbPPT_XMLAddNode(xmlId, sldLayoutNode, "p:clrMap")
  PbPPT_XMLSetAttribute(clrMapNode, "bg1", "lt1")
  PbPPT_XMLSetAttribute(clrMapNode, "tx1", "dk1")
  PbPPT_XMLSetAttribute(clrMapNode, "bg2", "lt2")
  PbPPT_XMLSetAttribute(clrMapNode, "tx2", "dk2")
  PbPPT_XMLSetAttribute(clrMapNode, "accent1", "accent1")
  PbPPT_XMLSetAttribute(clrMapNode, "accent2", "accent2")
  PbPPT_XMLSetAttribute(clrMapNode, "accent3", "accent3")
  PbPPT_XMLSetAttribute(clrMapNode, "accent4", "accent4")
  PbPPT_XMLSetAttribute(clrMapNode, "accent5", "accent5")
  PbPPT_XMLSetAttribute(clrMapNode, "accent6", "accent6")
  PbPPT_XMLSetAttribute(clrMapNode, "hlink", "hlink")
  PbPPT_XMLSetAttribute(clrMapNode, "folHlink", "folHlink")
  ProcedureReturn xmlId
EndProcedure

; WriteThemeXML - 生成主题XML
Procedure.i PbPPT_WriteThemeXML()
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define themeNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "a:theme")
  PbPPT_XMLSetAttribute(themeNode, "xmlns:a", #PbPPT_NS_A$)
  PbPPT_XMLSetAttribute(themeNode, "name", "Office 主题")
  ; 主题元素
  Define themeElementsNode.i = PbPPT_XMLAddNode(xmlId, themeNode, "a:themeElements")
  ; 颜色方案
  Define clrSchemeNode.i = PbPPT_XMLAddNode(xmlId, themeElementsNode, "a:clrScheme")
  PbPPT_XMLSetAttribute(clrSchemeNode, "name", "Office")
  Define colors.s = "dk1,lt1,dk2,lt2,accent1,accent2,accent3,accent4,accent5,accent6,hlink,folHlink"
  Define sysColors.s = "windowText,window,windowText,window,4F81BD,C0504D,9BBB59,8064A2,4BACC6,F79646,0000FF,800080"
  Define ci.i
  For ci = 1 To 12
    Define colorName.s = StringField(colors, ci, ",")
    Define sysColorVal.s = StringField(sysColors, ci, ",")
    Define colorNode.i
    If ci <= 2
      colorNode = PbPPT_XMLAddNode(xmlId, clrSchemeNode, "a:" + colorName)
      Define sysClrNode.i = PbPPT_XMLAddNode(xmlId, colorNode, "a:sysClr")
      PbPPT_XMLSetAttribute(sysClrNode, "val", sysColorVal)
      If ci = 1
        PbPPT_XMLSetAttribute(sysClrNode, "lastClr", "000000")
      Else
        PbPPT_XMLSetAttribute(sysClrNode, "lastClr", "FFFFFF")
      EndIf
    ElseIf ci <= 4
      colorNode = PbPPT_XMLAddNode(xmlId, clrSchemeNode, "a:" + colorName)
      Define srgbClrNode.i = PbPPT_XMLAddNode(xmlId, colorNode, "a:srgbClr")
      PbPPT_XMLSetAttribute(srgbClrNode, "val", sysColorVal)
    Else
      colorNode = PbPPT_XMLAddNode(xmlId, clrSchemeNode, "a:" + colorName)
      Define srgbClrNode2.i = PbPPT_XMLAddNode(xmlId, colorNode, "a:srgbClr")
      PbPPT_XMLSetAttribute(srgbClrNode2, "val", sysColorVal)
    EndIf
  Next
  ; 字体方案
  Define fontSchemeNode.i = PbPPT_XMLAddNode(xmlId, themeElementsNode, "a:fontScheme")
  PbPPT_XMLSetAttribute(fontSchemeNode, "name", "Office")
  Define majorFontNode.i = PbPPT_XMLAddNode(xmlId, fontSchemeNode, "a:majorFont")
  Define latinNode.i = PbPPT_XMLAddNode(xmlId, majorFontNode, "a:latin")
  PbPPT_XMLSetAttribute(latinNode, "typeface", "Calibri")
  Define eaNode.i = PbPPT_XMLAddNode(xmlId, majorFontNode, "a:ea")
  PbPPT_XMLSetAttribute(eaNode, "typeface", "")
  Define csNode.i = PbPPT_XMLAddNode(xmlId, majorFontNode, "a:cs")
  PbPPT_XMLSetAttribute(csNode, "typeface", "")
  Define minorFontNode.i = PbPPT_XMLAddNode(xmlId, fontSchemeNode, "a:minorFont")
  Define latinNode2.i = PbPPT_XMLAddNode(xmlId, minorFontNode, "a:latin")
  PbPPT_XMLSetAttribute(latinNode2, "typeface", "Calibri")
  Define eaNode2.i = PbPPT_XMLAddNode(xmlId, minorFontNode, "a:ea")
  PbPPT_XMLSetAttribute(eaNode2, "typeface", "")
  Define csNode2.i = PbPPT_XMLAddNode(xmlId, minorFontNode, "a:cs")
  PbPPT_XMLSetAttribute(csNode2, "typeface", "")
  ; 格式方案
  Define fmtSchemeNode.i = PbPPT_XMLAddNode(xmlId, themeElementsNode, "a:fmtScheme")
  PbPPT_XMLSetAttribute(fmtSchemeNode, "name", "Office")
  ; 填充样式
  Define fillStyleLstNode.i = PbPPT_XMLAddNode(xmlId, fmtSchemeNode, "a:fillStyleLst")
  Define solidFillNode.i = PbPPT_XMLAddNode(xmlId, fillStyleLstNode, "a:solidFill")
  Define srgbClrNode3.i = PbPPT_XMLAddNode(xmlId, solidFillNode, "a:schemeClr")
  PbPPT_XMLSetAttribute(srgbClrNode3, "val", "phClr")
  ; 线条样式
  Define lnStyleLstNode.i = PbPPT_XMLAddNode(xmlId, fmtSchemeNode, "a:lnStyleLst")
  ; 效果样式
  Define effectStyleLstNode.i = PbPPT_XMLAddNode(xmlId, fmtSchemeNode, "a:effectStyleLst")
  ; 背景填充
  Define bgFillStyleLstNode.i = PbPPT_XMLAddNode(xmlId, fmtSchemeNode, "a:bgFillStyleLst")
  ProcedureReturn xmlId
EndProcedure

; WriteCorePropsXML - 生成核心属性XML
Procedure.i PbPPT_WriteCorePropsXML()
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define cpNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "cp:coreProperties")
  PbPPT_XMLSetAttribute(cpNode, "xmlns:cp", #PbPPT_NS_CP$)
  PbPPT_XMLSetAttribute(cpNode, "xmlns:dc", #PbPPT_NS_DC$)
  PbPPT_XMLSetAttribute(cpNode, "xmlns:dcterms", #PbPPT_NS_DCTERMS$)
  PbPPT_XMLSetAttribute(cpNode, "xmlns:dcmitype", #PbPPT_NS_DCMITYPE$)
  PbPPT_XMLSetAttribute(cpNode, "xmlns:xsi", #PbPPT_NS_XSI$)
  ; 标题
  If PbPPT_Prs\title <> ""
    Define titleNode.i = PbPPT_XMLAddNode(xmlId, cpNode, "dc:title")
    PbPPT_XMLSetText(titleNode, PbPPT_Prs\title)
  EndIf
  ; 主题
  If PbPPT_Prs\subject <> ""
    Define subjectNode.i = PbPPT_XMLAddNode(xmlId, cpNode, "dc:subject")
    PbPPT_XMLSetText(subjectNode, PbPPT_Prs\subject)
  EndIf
  ; 创建者
  If PbPPT_Prs\creator <> ""
    Define creatorNode.i = PbPPT_XMLAddNode(xmlId, cpNode, "dc:creator")
    PbPPT_XMLSetText(creatorNode, PbPPT_Prs\creator)
  EndIf
  ; 描述
  If PbPPT_Prs\description <> ""
    Define descNode.i = PbPPT_XMLAddNode(xmlId, cpNode, "dc:description")
    PbPPT_XMLSetText(descNode, PbPPT_Prs\description)
  EndIf
  ; 关键词
  If PbPPT_Prs\keywords <> ""
    Define keywordsNode.i = PbPPT_XMLAddNode(xmlId, cpNode, "cp:keywords")
    PbPPT_XMLSetText(keywordsNode, PbPPT_Prs\keywords)
  EndIf
  ; 最后修改者
  If PbPPT_Prs\lastModifiedBy <> ""
    Define lastModByNode.i = PbPPT_XMLAddNode(xmlId, cpNode, "cp:lastModifiedBy")
    PbPPT_XMLSetText(lastModByNode, PbPPT_Prs\lastModifiedBy)
  EndIf
  ; 创建时间
  Define createdNode.i = PbPPT_XMLAddNode(xmlId, cpNode, "dcterms:created")
  PbPPT_XMLSetAttribute(createdNode, "xsi:type", "dcterms:W3CDTF")
  If PbPPT_Prs\created <> ""
    PbPPT_XMLSetText(createdNode, PbPPT_Prs\created)
  Else
    PbPPT_XMLSetText(createdNode, PbPPT_GetCurrentDateTime())
  EndIf
  ; 修改时间
  Define modifiedNode.i = PbPPT_XMLAddNode(xmlId, cpNode, "dcterms:modified")
  PbPPT_XMLSetAttribute(modifiedNode, "xsi:type", "dcterms:W3CDTF")
  PbPPT_XMLSetText(modifiedNode, PbPPT_GetCurrentDateTime())
  ; 分类
  If PbPPT_Prs\category <> ""
    Define categoryNode.i = PbPPT_XMLAddNode(xmlId, cpNode, "cp:category")
    PbPPT_XMLSetText(categoryNode, PbPPT_Prs\category)
  EndIf
  ProcedureReturn xmlId
EndProcedure

; WriteAppPropsXML - 生成应用程序属性XML
Procedure.i PbPPT_WriteAppPropsXML()
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define propsNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "Properties")
  PbPPT_XMLSetAttribute(propsNode, "xmlns", #PbPPT_NS_EP$)
  PbPPT_XMLSetAttribute(propsNode, "xmlns:vt", #PbPPT_NS_VT$)
  ; 应用程序
  Define appNode.i = PbPPT_XMLAddNode(xmlId, propsNode, "Application")
  PbPPT_XMLSetText(appNode, "PbPPT")
  ; 幻灯片数量
  Define slidesNode.i = PbPPT_XMLAddNode(xmlId, propsNode, "Slides")
  PbPPT_XMLSetText(slidesNode, Str(ListSize(PbPPT_Slides())))
  ; 隐藏幻灯片
  Define hiddenNode.i = PbPPT_XMLAddNode(xmlId, propsNode, "HiddenSlides")
  PbPPT_XMLSetText(hiddenNode, "0")
  ; 标题对
  Define hpNode.i = PbPPT_XMLAddNode(xmlId, propsNode, "HeadingPairs")
  Define vecNode.i = PbPPT_XMLAddNode(xmlId, hpNode, "vt:vector")
  PbPPT_XMLSetAttribute(vecNode, "size", "4")
  PbPPT_XMLSetAttribute(vecNode, "baseType", "variant")
  Define varNode.i = PbPPT_XMLAddNode(xmlId, vecNode, "vt:variant")
  Define lpstrNode.i = PbPPT_XMLAddNode(xmlId, varNode, "vt:lpstr")
  PbPPT_XMLSetText(lpstrNode, "Theme")
  varNode = PbPPT_XMLAddNode(xmlId, vecNode, "vt:variant")
  Define i4Node.i = PbPPT_XMLAddNode(xmlId, varNode, "vt:i4")
  PbPPT_XMLSetText(i4Node, "1")
  varNode = PbPPT_XMLAddNode(xmlId, vecNode, "vt:variant")
  lpstrNode = PbPPT_XMLAddNode(xmlId, varNode, "vt:lpstr")
  PbPPT_XMLSetText(lpstrNode, "幻灯片标题")
  varNode = PbPPT_XMLAddNode(xmlId, vecNode, "vt:variant")
  i4Node = PbPPT_XMLAddNode(xmlId, varNode, "vt:i4")
  PbPPT_XMLSetText(i4Node, Str(ListSize(PbPPT_Slides())))
  ; 部件标题
  Define topNode.i = PbPPT_XMLAddNode(xmlId, propsNode, "TitlesOfParts")
  Define vecNode2.i = PbPPT_XMLAddNode(xmlId, topNode, "vt:vector")
  PbPPT_XMLSetAttribute(vecNode2, "size", Str(1 + ListSize(PbPPT_Slides())))
  PbPPT_XMLSetAttribute(vecNode2, "baseType", "lpstr")
  Define lpstrNode2.i = PbPPT_XMLAddNode(xmlId, vecNode2, "vt:lpstr")
  PbPPT_XMLSetText(lpstrNode2, "Office 主题")
  ForEach PbPPT_Slides()
    lpstrNode2 = PbPPT_XMLAddNode(xmlId, vecNode2, "vt:lpstr")
    PbPPT_XMLSetText(lpstrNode2, "幻灯片 " + Str(ListIndex(PbPPT_Slides()) + 1))
  Next
  ProcedureReturn xmlId
EndProcedure

; WritePresPropsXML - 生成演示文稿属性XML
Procedure.i PbPPT_WritePresPropsXML()
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define presPrNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "p:presentationPr")
  PbPPT_XMLSetAttribute(presPrNode, "xmlns:a", #PbPPT_NS_A$)
  PbPPT_XMLSetAttribute(presPrNode, "xmlns:r", #PbPPT_NS_R$)
  PbPPT_XMLSetAttribute(presPrNode, "xmlns:p", #PbPPT_NS_P$)
  ProcedureReturn xmlId
EndProcedure

; WriteViewPropsXML - 生成视图属性XML
Procedure.i PbPPT_WriteViewPropsXML()
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define viewPrNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "p:viewPr")
  PbPPT_XMLSetAttribute(viewPrNode, "xmlns:a", #PbPPT_NS_A$)
  PbPPT_XMLSetAttribute(viewPrNode, "xmlns:r", #PbPPT_NS_R$)
  PbPPT_XMLSetAttribute(viewPrNode, "xmlns:p", #PbPPT_NS_P$)
  ; 普通视图
  Define normalViewPrNode.i = PbPPT_XMLAddNode(xmlId, viewPrNode, "p:normalViewPr")
  Define restoredLeftNode.i = PbPPT_XMLAddNode(xmlId, normalViewPrNode, "p:restoredLeft")
  PbPPT_XMLSetAttribute(restoredLeftNode, "sz", "15620")
  Define restoredTopNode.i = PbPPT_XMLAddNode(xmlId, normalViewPrNode, "p:restoredTop")
  PbPPT_XMLSetAttribute(restoredTopNode, "sz", "94660")
  ; 幻灯片视图
  Define slideViewPrNode.i = PbPPT_XMLAddNode(xmlId, viewPrNode, "p:slideViewPr")
  Define cSldViewPrNode.i = PbPPT_XMLAddNode(xmlId, slideViewPrNode, "p:cSldViewPr")
  ProcedureReturn xmlId
EndProcedure

; WriteTableStylesXML - 生成表格样式XML
Procedure.i PbPPT_WriteTableStylesXML()
  Define xmlId.i = PbPPT_XMLCreateDocument()
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define tblStyleLstNode.i = PbPPT_XMLAddNode(xmlId, rootNode, "a:tblStyleLst")
  PbPPT_XMLSetAttribute(tblStyleLstNode, "xmlns:a", #PbPPT_NS_A$)
  PbPPT_XMLSetAttribute(tblStyleLstNode, "def", "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
  ProcedureReturn xmlId
EndProcedure

; ***************************************************************************************
; 分区10: XML读取器（解析pptx文件各部分XML）
; ***************************************************************************************

; ExtractToTempDir - 将pptx文件解压到临时目录
Procedure.b PbPPT_ExtractToTempDir(filename.s)
  Define packId.i = PbPPT_ZIPOpen(filename)
  If packId = 0
    ProcedureReturn #False
  EndIf
  ; 创建临时目录
  PbPPT_TempDir = GetTemporaryDirectory() + "PbPPT_" + Str(Random(999999)) + "/"
  CreateDirectory(PbPPT_TempDir)
  ; 枚举并解压所有条目
  If PbPPT_ZIPExamine(packId)
    While PbPPT_ZIPNextEntry(packId)
      Define entryName.s = PbPPT_ZIPEntryName(packId)
      Define entrySize.i = PbPPT_ZIPEntrySize(packId)
      ; 创建目录结构
      Define dirPath.s = PbPPT_TempDir + GetPathPart(entryName)
      If FileSize(dirPath) = -1
        CreateDirectory(dirPath)
      EndIf
      ; 解压到临时文件
      Define outputPath.s = PbPPT_TempDir + entryName
      If entrySize > 0
        Define *buffer = AllocateMemory(entrySize)
        If *buffer
          If PbPPT_ZIPExtractToMemory(packId, *buffer, entrySize)
            Define outFile.i = CreateFile(#PB_Any, outputPath)
            If outFile
              WriteData(outFile, *buffer, entrySize)
              CloseFile(outFile)
            EndIf
          EndIf
          FreeMemory(*buffer)
        EndIf
      EndIf
    Wend
  EndIf
  PbPPT_ZIPClose(packId)
  ProcedureReturn #True
EndProcedure

; CleanupTempDir - 清理临时目录
Procedure PbPPT_CleanupTempDir()
  If PbPPT_TempDir <> "" And FileSize(PbPPT_TempDir) = -2
    DeleteDirectory(PbPPT_TempDir, "", #PB_FileSystem_Recursive)
  EndIf
EndProcedure

; ReadPresentationInfo - 从presentation.xml读取演示文稿信息
Procedure PbPPT_ReadPresentationInfo()
  Define xmlFile.s = PbPPT_TempDir + "ppt/presentation.xml"
  If FileSize(xmlFile) <= 0
    ProcedureReturn
  EndIf
  Define xmlId.i = PbPPT_XMLParseFile(xmlFile)
  If xmlId = 0
    ProcedureReturn
  EndIf
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  ; 查找sldSz节点获取幻灯片尺寸
  Define node.i = PbPPT_XMLGetChild(rootNode)
  While node
    Define nodeName.s = PbPPT_XMLGetNodeName(node)
    If FindString(nodeName, "sldSz")
      Define cx.s = PbPPT_XMLGetAttributeValue(node, "cx")
      Define cy.s = PbPPT_XMLGetAttributeValue(node, "cy")
      If cx <> ""
        PbPPT_Prs\slideWidth = Val(cx)
      EndIf
      If cy <> ""
        PbPPT_Prs\slideHeight = Val(cy)
      EndIf
    EndIf
    node = PbPPT_XMLGetNext(node)
  Wend
  PbPPT_XMLFree(xmlId)
EndProcedure

; ReadSlides - 从pptx文件读取所有幻灯片信息
Procedure PbPPT_ReadSlides()
  Define xmlFile.s = PbPPT_TempDir + "ppt/_rels/presentation.xml.rels"
  If FileSize(xmlFile) <= 0
    ProcedureReturn
  EndIf
  Define xmlId.i = PbPPT_XMLParseFile(xmlFile)
  If xmlId = 0
    ProcedureReturn
  EndIf
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define node.i = PbPPT_XMLGetChild(rootNode)
  Define slideIdx.i = 0
  While node
    Define nodeType.s = PbPPT_XMLGetAttributeValue(node, "Type")
    If FindString(nodeType, "/slide")
      Define target.s = PbPPT_XMLGetAttributeValue(node, "Target")
      Define rId.s = PbPPT_XMLGetAttributeValue(node, "Id")
      If target <> ""
        slideIdx + 1
        AddElement(PbPPT_Slides())
        PbPPT_Slides()\id = slideIdx
        PbPPT_Slides()\slideId = 255 + slideIdx
        PbPPT_Slides()\name = "幻灯片 " + Str(slideIdx)
        PbPPT_Slides()\partName = "/ppt/" + target
        PbPPT_Slides()\rId = rId
        PbPPT_Slides()\layoutIndex = 0
      EndIf
    EndIf
    node = PbPPT_XMLGetNext(node)
  Wend
  PbPPT_XMLFree(xmlId)
  ; 读取每张幻灯片的形状
  Define si.i
  For si = 1 To ListSize(PbPPT_Slides())
    SelectElement(PbPPT_Slides(), si - 1)
    Define slidePartName.s = Mid(PbPPT_Slides()\partName, 6)
    Define slideXmlFile.s = PbPPT_TempDir + "ppt/" + slidePartName
    If FileSize(slideXmlFile) > 0
      Define sXmlId.i = PbPPT_XMLParseFile(slideXmlFile)
      If sXmlId
        Define sRoot.i = PbPPT_XMLGetRoot(sXmlId)
        Define sNode.i = PbPPT_XMLGetChild(sRoot)
        While sNode
          Define sNodeName.s = PbPPT_XMLGetNodeName(sNode)
          If FindString(sNodeName, "cSld")
            Define cSldChild.i = PbPPT_XMLGetChild(sNode)
            While cSldChild
              Define cSldName.s = PbPPT_XMLGetNodeName(cSldChild)
              If FindString(cSldName, "spTree")
                Define spTreeNode.i = PbPPT_XMLGetChild(cSldChild)
                While spTreeNode
                  Define spName.s = PbPPT_XMLGetNodeName(spTreeNode)
                  If FindString(spName, ">sp") And Not FindString(spName, "nvSpPr")
                    ; 文本框或自动形状
                    AddElement(PbPPT_Shapes())
                    PbPPT_Shapes()\id = PbPPT_GetNextShapeId()
                    PbPPT_Shapes()\slideIndex = si - 1
                    PbPPT_Shapes()\shapeType = #PbPPT_ShapeTypeAutoShape
                    PbPPT_Shapes()\preset = "rect"
                    ; 解析形状属性
                    Define spChild.i = PbPPT_XMLGetChild(spTreeNode)
                    While spChild
                      Define spChildName.s = PbPPT_XMLGetNodeName(spChild)
                      If FindString(spChildName, "nvSpPr")
                        Define nvChild.i = PbPPT_XMLGetChild(spChild)
                        While nvChild
                          Define nvName.s = PbPPT_XMLGetNodeName(nvChild)
                          If FindString(nvName, "cNvPr")
                            PbPPT_Shapes()\id = Val(PbPPT_XMLGetAttributeValue(nvChild, "id"))
                            PbPPT_Shapes()\name = PbPPT_XMLGetAttributeValue(nvChild, "name")
                          EndIf
                          nvChild = PbPPT_XMLGetNext(nvChild)
                        Wend
                      ElseIf FindString(spChildName, "spPr")
                        Define prChild.i = PbPPT_XMLGetChild(spChild)
                        While prChild
                          Define prName.s = PbPPT_XMLGetNodeName(prChild)
                          If FindString(prName, "xfrm")
                            Define xfChild.i = PbPPT_XMLGetChild(prChild)
                            While xfChild
                              Define xfName.s = PbPPT_XMLGetNodeName(xfChild)
                              If FindString(xfName, "off")
                                PbPPT_Shapes()\left = Val(PbPPT_XMLGetAttributeValue(xfChild, "x"))
                                PbPPT_Shapes()\top = Val(PbPPT_XMLGetAttributeValue(xfChild, "y"))
                              ElseIf FindString(xfName, "ext")
                                PbPPT_Shapes()\width = Val(PbPPT_XMLGetAttributeValue(xfChild, "cx"))
                                PbPPT_Shapes()\height = Val(PbPPT_XMLGetAttributeValue(xfChild, "cy"))
                              EndIf
                              xfChild = PbPPT_XMLGetNext(xfChild)
                            Wend
                          ElseIf FindString(prName, "prstGeom")
                            PbPPT_Shapes()\preset = PbPPT_XMLGetAttributeValue(prChild, "prst")
                          EndIf
                          prChild = PbPPT_XMLGetNext(prChild)
                        Wend
                      ElseIf FindString(spChildName, "txBody")
                        PbPPT_Shapes()\shapeType = #PbPPT_ShapeTypeTextbox
                        ; 读取文本内容
                        Define txChild.i = PbPPT_XMLGetChild(spChild)
                        Define paraTexts.s = ""
                        While txChild
                          Define txName.s = PbPPT_XMLGetNodeName(txChild)
                          If FindString(txName, ">p") And Not FindString(txName, "pPr")
                            Define pChild.i = PbPPT_XMLGetChild(txChild)
                            Define runTexts.s = ""
                            While pChild
                              Define pName.s = PbPPT_XMLGetNodeName(pChild)
                              If FindString(pName, ">r")
                                Define rChild.i = PbPPT_XMLGetChild(pChild)
                                Define rText.s = ""
                                While rChild
                                  Define rName.s = PbPPT_XMLGetNodeName(rChild)
                                  If FindString(rName, ">t")
                                    rText = PbPPT_XMLGetText(rChild)
                                  EndIf
                                  rChild = PbPPT_XMLGetNext(rChild)
                                Wend
                                If runTexts <> ""
                                  runTexts + "#RUN#"
                                EndIf
                                runTexts + rText
                              EndIf
                              pChild = PbPPT_XMLGetNext(pChild)
                            Wend
                            If paraTexts <> ""
                              paraTexts + "#PARA#"
                            EndIf
                            paraTexts + runTexts
                          EndIf
                          txChild = PbPPT_XMLGetNext(txChild)
                        Wend
                        PbPPT_Shapes()\textFrame\paragraphs = paraTexts
                      EndIf
                      spChild = PbPPT_XMLGetNext(spChild)
                    Wend
                  ElseIf FindString(spName, ">pic")
                    ; 图片形状
                    AddElement(PbPPT_Shapes())
                    PbPPT_Shapes()\id = PbPPT_GetNextShapeId()
                    PbPPT_Shapes()\slideIndex = si - 1
                    PbPPT_Shapes()\shapeType = #PbPPT_ShapeTypePicture
                    PbPPT_Shapes()\preset = "rect"
                    Define picChild.i = PbPPT_XMLGetChild(spTreeNode)
                    While picChild
                      Define picChildName.s = PbPPT_XMLGetNodeName(picChild)
                      If FindString(picChildName, "nvPicPr")
                        Define nvPicChild.i = PbPPT_XMLGetChild(picChild)
                        While nvPicChild
                          Define nvPicName.s = PbPPT_XMLGetNodeName(nvPicChild)
                          If FindString(nvPicName, "cNvPr")
                            PbPPT_Shapes()\id = Val(PbPPT_XMLGetAttributeValue(nvPicChild, "id"))
                            PbPPT_Shapes()\name = PbPPT_XMLGetAttributeValue(nvPicChild, "name")
                          EndIf
                          nvPicChild = PbPPT_XMLGetNext(nvPicChild)
                        Wend
                      ElseIf FindString(picChildName, "spPr")
                        Define picPrChild.i = PbPPT_XMLGetChild(picChild)
                        While picPrChild
                          Define picPrName.s = PbPPT_XMLGetNodeName(picPrChild)
                          If FindString(picPrName, "xfrm")
                            Define picXfChild.i = PbPPT_XMLGetChild(picPrChild)
                            While picXfChild
                              Define picXfName.s = PbPPT_XMLGetNodeName(picXfChild)
                              If FindString(picXfName, "off")
                                PbPPT_Shapes()\left = Val(PbPPT_XMLGetAttributeValue(picXfChild, "x"))
                                PbPPT_Shapes()\top = Val(PbPPT_XMLGetAttributeValue(picXfChild, "y"))
                              ElseIf FindString(picXfName, "ext")
                                PbPPT_Shapes()\width = Val(PbPPT_XMLGetAttributeValue(picXfChild, "cx"))
                                PbPPT_Shapes()\height = Val(PbPPT_XMLGetAttributeValue(picXfChild, "cy"))
                              EndIf
                              picXfChild = PbPPT_XMLGetNext(picXfChild)
                            Wend
                          EndIf
                          picPrChild = PbPPT_XMLGetNext(picPrChild)
                        Wend
                      EndIf
                      picChild = PbPPT_XMLGetNext(picChild)
                    Wend
                  EndIf
                  spTreeNode = PbPPT_XMLGetNext(spTreeNode)
                Wend
              EndIf
              cSldChild = PbPPT_XMLGetNext(cSldChild)
            Wend
          EndIf
          sNode = PbPPT_XMLGetNext(sNode)
        Wend
        PbPPT_XMLFree(sXmlId)
      EndIf
    EndIf
  Next
EndProcedure

; ReadAppProps - 读取应用属性
Procedure PbPPT_ReadAppProps()
  Define xmlFile.s = PbPPT_TempDir + "docProps/app.xml"
  If FileSize(xmlFile) <= 0
    ProcedureReturn
  EndIf
  Define xmlId.i = PbPPT_XMLParseFile(xmlFile)
  If xmlId = 0
    ProcedureReturn
  EndIf
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define node.i = PbPPT_XMLGetChild(rootNode)
  While node
    Define nodeName.s = PbPPT_XMLGetNodeName(node)
    If FindString(nodeName, "Slides")
      ; 已通过幻灯片列表获取
    EndIf
    node = PbPPT_XMLGetNext(node)
  Wend
  PbPPT_XMLFree(xmlId)
EndProcedure

; ***************************************************************************************
; 分区11: 形状模块（形状创建、自动形状、图片、连接符、占位符）
; ***************************************************************************************

; CreateShape - 通用形状创建（内部函数）
Procedure.i PbPPT_CreateShape(slideIndex.i, shapeType.i, left.i, top.i, width.i, height.i)
  AddElement(PbPPT_Shapes())
  PbPPT_Shapes()\id = PbPPT_Prs\nextShapeId
  PbPPT_Prs\nextShapeId + 1
  PbPPT_Shapes()\slideIndex = slideIndex
  PbPPT_Shapes()\shapeType = shapeType
  PbPPT_Shapes()\left = left
  PbPPT_Shapes()\top = top
  PbPPT_Shapes()\width = width
  PbPPT_Shapes()\height = height
  PbPPT_Shapes()\rotation = 0
  PbPPT_Shapes()\preset = "rect"
  PbPPT_Shapes()\name = "形状 " + Str(PbPPT_Shapes()\id)
  PbPPT_Shapes()\fill\fillType = #PbPPT_FillTypeNone
  PbPPT_Shapes()\fill\foreColor = ""
  PbPPT_Shapes()\fill\backColor = ""
  PbPPT_Shapes()\line\width = 0
  PbPPT_Shapes()\line\color = ""
  PbPPT_Shapes()\line\dashStyle = "solid"
  PbPPT_Shapes()\textFrame\paragraphs = ""
  PbPPT_Shapes()\textFrame\wordWrap = #PbPPT_AlignLeft
  PbPPT_Shapes()\textFrame\verticalAnchor = #PbPPT_ValignTop
  PbPPT_Shapes()\textFrame\autoFit = 0
  PbPPT_Shapes()\textFrame\marginLeft = 0
  PbPPT_Shapes()\textFrame\marginRight = 0
  PbPPT_Shapes()\textFrame\marginTop = 0
  PbPPT_Shapes()\textFrame\marginBottom = 0
  PbPPT_Shapes()\imagePath = ""
  PbPPT_Shapes()\imageRId = ""
  PbPPT_Shapes()\imageDesc = ""
  PbPPT_Shapes()\tableRows = 0
  PbPPT_Shapes()\tableCols = 0
  PbPPT_Shapes()\tableCells = ""
  PbPPT_Shapes()\tableColWidths = ""
  PbPPT_Shapes()\tableRowHeights = ""
  PbPPT_Shapes()\chartType = 0
  PbPPT_Shapes()\chartRId = ""
  PbPPT_Shapes()\chartTitle = ""
  PbPPT_Shapes()\chartData = ""
  PbPPT_Shapes()\phType = 0
  PbPPT_Shapes()\phIdx = 0
  PbPPT_Shapes()\hyperlink = ""
  ProcedureReturn PbPPT_Shapes()\id
EndProcedure

; CreateAutoShape - 创建自动形状
Procedure.i PbPPT_CreateAutoShape(slideIndex.i, preset.s, left.i, top.i, width.i, height.i)
  Define shapeId.i = PbPPT_CreateShape(slideIndex, #PbPPT_ShapeTypeAutoShape, left, top, width, height)
  If shapeId > 0
    PbPPT_Shapes()\preset = preset
    PbPPT_Shapes()\name = "自动形状 " + Str(shapeId)
    PbPPT_Shapes()\fill\fillType = #PbPPT_FillTypeSolid
    PbPPT_Shapes()\fill\foreColor = "4472C4"
    PbPPT_Shapes()\line\width = PbPPT_PtToEmu(0.75)
    PbPPT_Shapes()\line\color = "2F528F"
  EndIf
  ProcedureReturn shapeId
EndProcedure

; CreateTextbox - 创建文本框
Procedure.i PbPPT_CreateTextbox(slideIndex.i, left.i, top.i, width.i, height.i)
  Define shapeId.i = PbPPT_CreateShape(slideIndex, #PbPPT_ShapeTypeTextbox, left, top, width, height)
  If shapeId > 0
    PbPPT_Shapes()\name = "文本框 " + Str(shapeId)
    PbPPT_Shapes()\fill\fillType = #PbPPT_FillTypeNone
    PbPPT_Shapes()\line\width = 0
  EndIf
  ProcedureReturn shapeId
EndProcedure

; CreatePicture - 创建图片形状
Procedure.i PbPPT_CreatePicture(slideIndex.i, imagePath.s, left.i, top.i, width.i, height.i)
  Define shapeId.i = PbPPT_CreateShape(slideIndex, #PbPPT_ShapeTypePicture, left, top, width, height)
  If shapeId > 0
    PbPPT_Shapes()\name = "图片 " + Str(shapeId)
    PbPPT_Shapes()\imagePath = imagePath
    PbPPT_Shapes()\imageDesc = GetFilePart(imagePath)
    ; 添加图片到包中并获取关系ID
    Define fp.s = PbPPT_CalcFileFingerprint(imagePath)
    Define imgIdx.i = PbPPT_Prs\imageCount + 1
    Define foundRId.s = ""
    ; 检查图片是否已存在（去重）
    If FindMapElement(PbPPT_ImageSHA1Map(), fp)
      foundRId = PbPPT_ImageSHA1Map(fp)
    EndIf
    If foundRId = ""
      ; 新图片，添加到包中
      PbPPT_Prs\imageCount + 1
      imgIdx = PbPPT_Prs\imageCount
      Define ext.s = PbPPT_GetFileExtension(imagePath)
      Define contentType.s = PbPPT_GetImageContentType(ext)
      Define partName.s = "/ppt/media/image" + Str(imgIdx) + "." + ext
      ; 读取图片文件到内存
      Define imgSize.i = FileSize(imagePath)
      If imgSize > 0
        Define *imgBuf = AllocateMemory(imgSize)
        If *imgBuf
          Define imgFh.i = ReadFile(#PB_Any, imagePath)
          If imgFh
            ReadData(imgFh, *imgBuf, imgSize)
            CloseFile(imgFh)
            PbPPT_AddBinaryPart(partName, contentType, *imgBuf, imgSize)
          EndIf
          FreeMemory(*imgBuf)
        EndIf
      EndIf
      ; 计算关系ID（在WriteSlideRelsXML中分配）
      PbPPT_Shapes()\imageRId = "rId_img_" + Str(imgIdx)
      PbPPT_ImageSHA1Map(fp) = PbPPT_Shapes()\imageRId
    Else
      PbPPT_Shapes()\imageRId = foundRId
    EndIf
    ; 如果未指定尺寸，尝试自动获取
    If width = 0 Or height = 0
      UsePNGImageDecoder()
      UseJPEGImageDecoder()
      Define imgId.i = LoadImage(#PB_Any, imagePath)
      If imgId
        If width = 0
          PbPPT_Shapes()\width = Int(ImageWidth(imgId) / 96.0 * #PbPPT_EMU_PER_INCH)
        EndIf
        If height = 0
          PbPPT_Shapes()\height = Int(ImageHeight(imgId) / 96.0 * #PbPPT_EMU_PER_INCH)
        EndIf
        FreeImage(imgId)
      EndIf
    EndIf
  EndIf
  ProcedureReturn shapeId
EndProcedure

; CreateTable - 创建表格形状
Procedure.i PbPPT_CreateTable(slideIndex.i, rows.i, cols.i, left.i, top.i, width.i, height.i)
  Define shapeId.i = PbPPT_CreateShape(slideIndex, #PbPPT_ShapeTypeTable, left, top, width, height)
  If shapeId > 0
    PbPPT_Shapes()\name = "表格 " + Str(shapeId)
    PbPPT_Shapes()\tableRows = rows
    PbPPT_Shapes()\tableCols = cols
    PbPPT_Shapes()\tableCells = ""
    PbPPT_Shapes()\fill\fillType = #PbPPT_FillTypeNone
    PbPPT_Shapes()\line\width = 0
  EndIf
  ProcedureReturn shapeId
EndProcedure

; CreateChart - 创建图表形状
Procedure.i PbPPT_CreateChart(slideIndex.i, chartType.i, left.i, top.i, width.i, height.i)
  Define shapeId.i = PbPPT_CreateShape(slideIndex, #PbPPT_ShapeTypeChart, left, top, width, height)
  If shapeId > 0
    PbPPT_Shapes()\name = "图表 " + Str(shapeId)
    PbPPT_Shapes()\chartType = chartType
    PbPPT_Shapes()\chartRId = "rId_chart_" + Str(shapeId)
    PbPPT_Shapes()\chartTitle = ""
    PbPPT_Shapes()\chartData = ""
    PbPPT_Shapes()\fill\fillType = #PbPPT_FillTypeNone
    PbPPT_Shapes()\line\width = 0
  EndIf
  ProcedureReturn shapeId
EndProcedure

; CreateConnector - 创建连接符
Procedure.i PbPPT_CreateConnector(slideIndex.i, beginX.i, beginY.i, endX.i, endY.i)
  Define left.i = beginX
  Define top.i = beginY
  Define width.i = endX - beginX
  Define height.i = endY - beginY
  If width < 0
    left = endX
    width = -width
  EndIf
  If height < 0
    top = endY
    height = -height
  EndIf
  Define shapeId.i = PbPPT_CreateShape(slideIndex, #PbPPT_ShapeTypeConnector, left, top, width, height)
  If shapeId > 0
    PbPPT_Shapes()\name = "连接符 " + Str(shapeId)
    PbPPT_Shapes()\line\width = PbPPT_PtToEmu(1)
    PbPPT_Shapes()\line\color = "000000"
  EndIf
  ProcedureReturn shapeId
EndProcedure

; SetShapePosition - 设置形状位置和大小
Procedure PbPPT_SetShapePosition(shapeId.i, left.i, top.i, width.i, height.i)
  ForEach PbPPT_Shapes()
    If PbPPT_Shapes()\id = shapeId
      If left >= 0
        PbPPT_Shapes()\left = left
      EndIf
      If top >= 0
        PbPPT_Shapes()\top = top
      EndIf
      If width > 0
        PbPPT_Shapes()\width = width
      EndIf
      If height > 0
        PbPPT_Shapes()\height = height
      EndIf
      Break
    EndIf
  Next
EndProcedure

; SetShapeRotation - 设置形状旋转角度
Procedure PbPPT_SetShapeRotation(shapeId.i, rotation.f)
  ForEach PbPPT_Shapes()
    If PbPPT_Shapes()\id = shapeId
      PbPPT_Shapes()\rotation = rotation
      Break
    EndIf
  Next
EndProcedure

; SetShapeName - 设置形状名称
Procedure PbPPT_SetShapeName(shapeId.i, name.s)
  ForEach PbPPT_Shapes()
    If PbPPT_Shapes()\id = shapeId
      PbPPT_Shapes()\name = name
      Break
    EndIf
  Next
EndProcedure

; SetShapeFill - 设置形状填充
Procedure PbPPT_SetShapeFill(shapeId.i, fillType.i, foreColor.s, backColor.s)
  ForEach PbPPT_Shapes()
    If PbPPT_Shapes()\id = shapeId
      PbPPT_Shapes()\fill\fillType = fillType
      PbPPT_Shapes()\fill\foreColor = foreColor
      PbPPT_Shapes()\fill\backColor = backColor
      Break
    EndIf
  Next
EndProcedure

; SetShapeLine - 设置形状线条
Procedure PbPPT_SetShapeLine(shapeId.i, width.i, color.s, dashStyle.s)
  ForEach PbPPT_Shapes()
    If PbPPT_Shapes()\id = shapeId
      PbPPT_Shapes()\line\width = width
      PbPPT_Shapes()\line\color = color
      PbPPT_Shapes()\line\dashStyle = dashStyle
      Break
    EndIf
  Next
EndProcedure

; GetShapeById - 按ID查找形状（将指针定位到该形状）
Procedure.b PbPPT_GetShapeById(shapeId.i)
  ForEach PbPPT_Shapes()
    If PbPPT_Shapes()\id = shapeId
      ProcedureReturn #True
    EndIf
  Next
  ProcedureReturn #False
EndProcedure

; GetShapeCount - 获取指定幻灯片上的形状数量
Procedure.i PbPPT_GetShapeCount(slideIndex.i)
  Define count.i = 0
  ForEach PbPPT_Shapes()
    If PbPPT_Shapes()\slideIndex = slideIndex
      count + 1
    EndIf
  Next
  ProcedureReturn count
EndProcedure

; ***************************************************************************************
; 分区12: 文本模块（文本框、段落、运行、字体）
; ***************************************************************************************

; SetTextFrameText - 设置文本框全部文本（简单模式，自动创建段落和Run）
Procedure PbPPT_SetTextFrameText(shapeId.i, text.s)
  If PbPPT_GetShapeById(shapeId)
    ; 将文本按换行符分割为段落，每段一个Run
    Define result.s = ""
    Define paraCount.i = CountString(text, #LF$) + 1
    Define pi.i
    For pi = 1 To paraCount
      Define paraLine.s = StringField(text, pi, #LF$)
      If pi > 1
        result + "#PARA#"
      EndIf
      ; 每个段落包含一个Run，无特殊字体属性
      result + paraLine
    Next
    PbPPT_Shapes()\textFrame\paragraphs = result
  EndIf
EndProcedure

; GetTextFrameText - 获取文本框全部文本
Procedure.s PbPPT_GetTextFrameText(shapeId.i)
  Define result.s = ""
  If PbPPT_GetShapeById(shapeId)
    Define paraText.s = PbPPT_Shapes()\textFrame\paragraphs
    If paraText <> ""
      Define paraCount.i = CountString(paraText, "#PARA#") + 1
      Define pi.i
      For pi = 1 To paraCount
        Define paraLine.s = StringField(paraText, pi, "#PARA#")
        If pi > 1
          result + #LF$
        EndIf
        ; 提取每个Run的纯文本（去掉#ATTR#属性部分）
        Define runCount.i = CountString(paraLine, "#RUN#") + 1
        Define ri.i
        For ri = 1 To runCount
          Define runPart.s = StringField(paraLine, ri, "#RUN#")
          Define attrPos.i = FindString(runPart, "#ATTR#")
          If attrPos > 0
            result + Left(runPart, attrPos - 1)
          Else
            result + runPart
          EndIf
        Next
      Next
    EndIf
  EndIf
  ProcedureReturn result
EndProcedure

; AddParagraph - 在文本框末尾添加段落
Procedure PbPPT_AddParagraph(shapeId.i, text.s)
  If PbPPT_GetShapeById(shapeId)
    Define current.s = PbPPT_Shapes()\textFrame\paragraphs
    If current <> ""
      current + "#PARA#"
    EndIf
    current + text
    PbPPT_Shapes()\textFrame\paragraphs = current
  EndIf
EndProcedure

; AddRun - 在指定段落末尾添加文本运行
Procedure PbPPT_AddRun(shapeId.i, paraIndex.i, text.s, fontName.s, fontSize.f, bold.b, italic.b, underline.b, color.s)
  If PbPPT_GetShapeById(shapeId)
    Define paraText.s = PbPPT_Shapes()\textFrame\paragraphs
    Define paraCount.i = CountString(paraText, "#PARA#") + 1
    If paraIndex < 1 Or paraIndex > paraCount
      ProcedureReturn
    EndIf
    ; 重建段落列表
    Define result.s = ""
    Define pi.i
    For pi = 1 To paraCount
      If pi > 1
        result + "#PARA#"
      EndIf
      Define paraLine.s = StringField(paraText, pi, "#PARA#")
      If pi = paraIndex
        ; 在该段落末尾添加Run
        Define runAttr.s = fontName + "," + StrF(fontSize, 1) + "," + Str(bold) + "," + Str(italic) + "," + Str(underline) + "," + color
        If paraLine <> ""
          paraLine + "#RUN#" + text + "#ATTR#" + runAttr
        Else
          paraLine = text + "#ATTR#" + runAttr
        EndIf
      EndIf
      result + paraLine
    Next
    PbPPT_Shapes()\textFrame\paragraphs = result
  EndIf
EndProcedure

; SetParagraphAlignment - 设置段落对齐方式
Procedure PbPPT_SetParagraphAlignment(shapeId.i, paraIndex.i, alignment.i)
  ; 对齐方式存储在textFrame.wordWrap中（简化处理，仅支持全局对齐）
  If PbPPT_GetShapeById(shapeId)
    PbPPT_Shapes()\textFrame\wordWrap = alignment
  EndIf
EndProcedure

; SetFont - 设置指定段落和运行的字体属性
Procedure PbPPT_SetFont(shapeId.i, paraIndex.i, runIndex.i, fontName.s, fontSize.f, bold.b, italic.b, underline.b, color.s)
  If PbPPT_GetShapeById(shapeId)
    Define paraText.s = PbPPT_Shapes()\textFrame\paragraphs
    Define paraCount.i = CountString(paraText, "#PARA#") + 1
    If paraIndex < 1 Or paraIndex > paraCount
      ProcedureReturn
    EndIf
    ; 重建段落列表
    Define result.s = ""
    Define pi.i
    For pi = 1 To paraCount
      If pi > 1
        result + "#PARA#"
      EndIf
      Define paraLine.s = StringField(paraText, pi, "#PARA#")
      If pi = paraIndex
        ; 重建Run列表
        Define runCount.i = CountString(paraLine, "#RUN#") + 1
        Define newParaLine.s = ""
        Define ri.i
        For ri = 1 To runCount
          If ri > 1
            newParaLine + "#RUN#"
          EndIf
          Define runPart.s = StringField(paraLine, ri, "#RUN#")
          Define runText.s = runPart
          Define runAttr.s = ""
          Define attrPos.i = FindString(runPart, "#ATTR#")
          If attrPos > 0
            runText = Left(runPart, attrPos - 1)
            runAttr = Mid(runPart, attrPos + 6)
          EndIf
          If ri = runIndex
            ; 更新字体属性
            runAttr = fontName + "," + StrF(fontSize, 1) + "," + Str(bold) + "," + Str(italic) + "," + Str(underline) + "," + color
          EndIf
          newParaLine + runText
          If runAttr <> ""
            newParaLine + "#ATTR#" + runAttr
          EndIf
        Next
        paraLine = newParaLine
      EndIf
      result + paraLine
    Next
    PbPPT_Shapes()\textFrame\paragraphs = result
  EndIf
EndProcedure

; SetTextFrameWordWrap - 设置文本框自动换行
Procedure PbPPT_SetTextFrameWordWrap(shapeId.i, wrap.b)
  If PbPPT_GetShapeById(shapeId)
    ; wordWrap字段复用为对齐方式，这里需要单独处理
    ; 通过设置bodyPr的wrap属性实现
    PbPPT_Shapes()\textFrame\wordWrap = wrap
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; SetTextFrameAutoFit - 设置文本框自动调整大小
Procedure PbPPT_SetTextFrameAutoFit(shapeId.i, autoFit.i)
  If PbPPT_GetShapeById(shapeId)
    PbPPT_Shapes()\textFrame\autoFit = autoFit
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; SetParagraphLevel - 设置段落缩进级别
Procedure PbPPT_SetParagraphLevel(shapeId.i, paraIndex.i, level.i)
  ; 级别信息存储在段落的#ATTR#中（扩展格式）
  ; 简化处理：暂不实现独立级别存储
  PbPPT_Prs\isModified = #True
EndProcedure

; SetParagraphSpacing - 设置段落间距
Procedure PbPPT_SetParagraphSpacing(shapeId.i, paraIndex.i, spaceBefore.f, spaceAfter.f)
  ; 间距信息存储在段落的#ATTR#中（扩展格式）
  ; 简化处理：暂不实现独立间距存储
  PbPPT_Prs\isModified = #True
EndProcedure

; GetParagraphText - 获取指定段落的纯文本
Procedure.s PbPPT_GetParagraphText(shapeId.i, paraIndex.i)
  Define result.s = ""
  If PbPPT_GetShapeById(shapeId)
    Define paraText.s = PbPPT_Shapes()\textFrame\paragraphs
    Define paraCount.i = CountString(paraText, "#PARA#") + 1
    If paraIndex >= 1 And paraIndex <= paraCount
      Define paraLine.s = StringField(paraText, paraIndex, "#PARA#")
      ; 提取每个Run的纯文本（去掉#ATTR#属性部分）
      Define runCount.i = CountString(paraLine, "#RUN#") + 1
      Define ri.i
      For ri = 1 To runCount
        Define runPart.s = StringField(paraLine, ri, "#RUN#")
        Define attrPos.i = FindString(runPart, "#ATTR#")
        If attrPos > 0
          result + Left(runPart, attrPos - 1)
        Else
          result + runPart
        EndIf
      Next
    EndIf
  EndIf
  ProcedureReturn result
EndProcedure

; SetParagraphText - 设置指定段落的文本（替换整个段落内容）
Procedure PbPPT_SetParagraphText(shapeId.i, paraIndex.i, text.s)
  If PbPPT_GetShapeById(shapeId)
    Define paraText.s = PbPPT_Shapes()\textFrame\paragraphs
    Define paraCount.i = CountString(paraText, "#PARA#") + 1
    If paraIndex >= 1 And paraIndex <= paraCount
      ; 重建段落列表
      Define result.s = ""
      Define pi.i
      For pi = 1 To paraCount
        If pi > 1
          result + "#PARA#"
        EndIf
        If pi = paraIndex
          result + text
        Else
          result + StringField(paraText, pi, "#PARA#")
        EndIf
      Next
      PbPPT_Shapes()\textFrame\paragraphs = result
      PbPPT_Prs\isModified = #True
    EndIf
  EndIf
EndProcedure

; SetRunText - 设置指定运行的文本
Procedure PbPPT_SetRunText(shapeId.i, paraIndex.i, runIndex.i, text.s)
  If PbPPT_GetShapeById(shapeId)
    Define paraText.s = PbPPT_Shapes()\textFrame\paragraphs
    Define paraCount.i = CountString(paraText, "#PARA#") + 1
    If paraIndex < 1 Or paraIndex > paraCount
      ProcedureReturn
    EndIf
    Define result.s = ""
    Define pi.i
    For pi = 1 To paraCount
      If pi > 1
        result + "#PARA#"
      EndIf
      Define paraLine.s = StringField(paraText, pi, "#PARA#")
      If pi = paraIndex
        Define runCount.i = CountString(paraLine, "#RUN#") + 1
        Define newParaLine.s = ""
        Define ri.i
        For ri = 1 To runCount
          If ri > 1
            newParaLine + "#RUN#"
          EndIf
          Define runPart.s = StringField(paraLine, ri, "#RUN#")
          Define runAttr.s = ""
          Define attrPos.i = FindString(runPart, "#ATTR#")
          If attrPos > 0
            runAttr = Mid(runPart, attrPos)
          EndIf
          If ri = runIndex
            newParaLine + text + runAttr
          Else
            newParaLine + runPart
          EndIf
        Next
        paraLine = newParaLine
      EndIf
      result + paraLine
    Next
    PbPPT_Shapes()\textFrame\paragraphs = result
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; GetRunText - 获取指定运行的文本
Procedure.s PbPPT_GetRunText(shapeId.i, paraIndex.i, runIndex.i)
  Define result.s = ""
  If PbPPT_GetShapeById(shapeId)
    Define paraText.s = PbPPT_Shapes()\textFrame\paragraphs
    Define paraCount.i = CountString(paraText, "#PARA#") + 1
    If paraIndex >= 1 And paraIndex <= paraCount
      Define paraLine.s = StringField(paraText, paraIndex, "#PARA#")
      Define runCount.i = CountString(paraLine, "#RUN#") + 1
      If runIndex >= 1 And runIndex <= runCount
        Define runPart.s = StringField(paraLine, runIndex, "#RUN#")
        Define attrPos.i = FindString(runPart, "#ATTR#")
        If attrPos > 0
          result = Left(runPart, attrPos - 1)
        Else
          result = runPart
        EndIf
      EndIf
    EndIf
  EndIf
  ProcedureReturn result
EndProcedure

; SetFontName - 设置字体名称
Procedure PbPPT_SetFontName(shapeId.i, paraIndex.i, runIndex.i, fontName.s)
  If PbPPT_GetShapeById(shapeId)
    Define paraText.s = PbPPT_Shapes()\textFrame\paragraphs
    Define paraCount.i = CountString(paraText, "#PARA#") + 1
    If paraIndex < 1 Or paraIndex > paraCount
      ProcedureReturn
    EndIf
    Define result.s = ""
    Define pi.i
    For pi = 1 To paraCount
      If pi > 1
        result + "#PARA#"
      EndIf
      Define paraLine.s = StringField(paraText, pi, "#PARA#")
      If pi = paraIndex
        Define runCount.i = CountString(paraLine, "#RUN#") + 1
        Define newParaLine.s = ""
        Define ri.i
        For ri = 1 To runCount
          If ri > 1
            newParaLine + "#RUN#"
          EndIf
          Define runPart.s = StringField(paraLine, ri, "#RUN#")
          Define runText.s = runPart
          Define runAttr.s = ""
          Define attrPos.i = FindString(runPart, "#ATTR#")
          If attrPos > 0
            runText = Left(runPart, attrPos - 1)
            runAttr = Mid(runPart, attrPos + 6)
          EndIf
          If ri = runIndex
            ; 更新字体名称（属性格式：字体名,字号,粗体,斜体,下划线,颜色）
            Define fnSize.s = StringField(runAttr, 2, ",")
            Define fnBold.s = StringField(runAttr, 3, ",")
            Define fnItalic.s = StringField(runAttr, 4, ",")
            Define fnUnderline.s = StringField(runAttr, 5, ",")
            Define fnColor.s = StringField(runAttr, 6, ",")
            runAttr = fontName + "," + fnSize + "," + fnBold + "," + fnItalic + "," + fnUnderline + "," + fnColor
          EndIf
          newParaLine + runText
          If runAttr <> ""
            newParaLine + "#ATTR#" + runAttr
          EndIf
        Next
        paraLine = newParaLine
      EndIf
      result + paraLine
    Next
    PbPPT_Shapes()\textFrame\paragraphs = result
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; SetFontSize - 设置字号
Procedure PbPPT_SetFontSize(shapeId.i, paraIndex.i, runIndex.i, fontSize.f)
  If PbPPT_GetShapeById(shapeId)
    Define paraText.s = PbPPT_Shapes()\textFrame\paragraphs
    Define paraCount.i = CountString(paraText, "#PARA#") + 1
    If paraIndex < 1 Or paraIndex > paraCount
      ProcedureReturn
    EndIf
    Define result.s = ""
    Define pi.i
    For pi = 1 To paraCount
      If pi > 1
        result + "#PARA#"
      EndIf
      Define paraLine.s = StringField(paraText, pi, "#PARA#")
      If pi = paraIndex
        Define runCount.i = CountString(paraLine, "#RUN#") + 1
        Define newParaLine.s = ""
        Define ri.i
        For ri = 1 To runCount
          If ri > 1
            newParaLine + "#RUN#"
          EndIf
          Define runPart.s = StringField(paraLine, ri, "#RUN#")
          Define runText.s = runPart
          Define runAttr.s = ""
          Define attrPos.i = FindString(runPart, "#ATTR#")
          If attrPos > 0
            runText = Left(runPart, attrPos - 1)
            runAttr = Mid(runPart, attrPos + 6)
          EndIf
          If ri = runIndex
            Define fnName.s = StringField(runAttr, 1, ",")
            Define fnBold.s = StringField(runAttr, 3, ",")
            Define fnItalic.s = StringField(runAttr, 4, ",")
            Define fnUnderline.s = StringField(runAttr, 5, ",")
            Define fnColor.s = StringField(runAttr, 6, ",")
            runAttr = fnName + "," + StrF(fontSize, 1) + "," + fnBold + "," + fnItalic + "," + fnUnderline + "," + fnColor
          EndIf
          newParaLine + runText
          If runAttr <> ""
            newParaLine + "#ATTR#" + runAttr
          EndIf
        Next
        paraLine = newParaLine
      EndIf
      result + paraLine
    Next
    PbPPT_Shapes()\textFrame\paragraphs = result
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; SetFontBold - 设置粗体
Procedure PbPPT_SetFontBold(shapeId.i, paraIndex.i, runIndex.i, bold.b)
  If PbPPT_GetShapeById(shapeId)
    Define paraText.s = PbPPT_Shapes()\textFrame\paragraphs
    Define paraCount.i = CountString(paraText, "#PARA#") + 1
    If paraIndex < 1 Or paraIndex > paraCount
      ProcedureReturn
    EndIf
    Define result.s = ""
    Define pi.i
    For pi = 1 To paraCount
      If pi > 1
        result + "#PARA#"
      EndIf
      Define paraLine.s = StringField(paraText, pi, "#PARA#")
      If pi = paraIndex
        Define runCount.i = CountString(paraLine, "#RUN#") + 1
        Define newParaLine.s = ""
        Define ri.i
        For ri = 1 To runCount
          If ri > 1
            newParaLine + "#RUN#"
          EndIf
          Define runPart.s = StringField(paraLine, ri, "#RUN#")
          Define runText.s = runPart
          Define runAttr.s = ""
          Define attrPos.i = FindString(runPart, "#ATTR#")
          If attrPos > 0
            runText = Left(runPart, attrPos - 1)
            runAttr = Mid(runPart, attrPos + 6)
          EndIf
          If ri = runIndex
            Define fnName.s = StringField(runAttr, 1, ",")
            Define fnSize.s = StringField(runAttr, 2, ",")
            Define fnItalic.s = StringField(runAttr, 4, ",")
            Define fnUnderline.s = StringField(runAttr, 5, ",")
            Define fnColor.s = StringField(runAttr, 6, ",")
            runAttr = fnName + "," + fnSize + "," + Str(bold) + "," + fnItalic + "," + fnUnderline + "," + fnColor
          EndIf
          newParaLine + runText
          If runAttr <> ""
            newParaLine + "#ATTR#" + runAttr
          EndIf
        Next
        paraLine = newParaLine
      EndIf
      result + paraLine
    Next
    PbPPT_Shapes()\textFrame\paragraphs = result
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; SetFontItalic - 设置斜体
Procedure PbPPT_SetFontItalic(shapeId.i, paraIndex.i, runIndex.i, italic.b)
  If PbPPT_GetShapeById(shapeId)
    Define paraText.s = PbPPT_Shapes()\textFrame\paragraphs
    Define paraCount.i = CountString(paraText, "#PARA#") + 1
    If paraIndex < 1 Or paraIndex > paraCount
      ProcedureReturn
    EndIf
    Define result.s = ""
    Define pi.i
    For pi = 1 To paraCount
      If pi > 1
        result + "#PARA#"
      EndIf
      Define paraLine.s = StringField(paraText, pi, "#PARA#")
      If pi = paraIndex
        Define runCount.i = CountString(paraLine, "#RUN#") + 1
        Define newParaLine.s = ""
        Define ri.i
        For ri = 1 To runCount
          If ri > 1
            newParaLine + "#RUN#"
          EndIf
          Define runPart.s = StringField(paraLine, ri, "#RUN#")
          Define runText.s = runPart
          Define runAttr.s = ""
          Define attrPos.i = FindString(runPart, "#ATTR#")
          If attrPos > 0
            runText = Left(runPart, attrPos - 1)
            runAttr = Mid(runPart, attrPos + 6)
          EndIf
          If ri = runIndex
            Define fnName.s = StringField(runAttr, 1, ",")
            Define fnSize.s = StringField(runAttr, 2, ",")
            Define fnBold.s = StringField(runAttr, 3, ",")
            Define fnUnderline.s = StringField(runAttr, 5, ",")
            Define fnColor.s = StringField(runAttr, 6, ",")
            runAttr = fnName + "," + fnSize + "," + fnBold + "," + Str(italic) + "," + fnUnderline + "," + fnColor
          EndIf
          newParaLine + runText
          If runAttr <> ""
            newParaLine + "#ATTR#" + runAttr
          EndIf
        Next
        paraLine = newParaLine
      EndIf
      result + paraLine
    Next
    PbPPT_Shapes()\textFrame\paragraphs = result
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; SetFontColor - 设置字体颜色
Procedure PbPPT_SetFontColor(shapeId.i, paraIndex.i, runIndex.i, color.s)
  If PbPPT_GetShapeById(shapeId)
    Define paraText.s = PbPPT_Shapes()\textFrame\paragraphs
    Define paraCount.i = CountString(paraText, "#PARA#") + 1
    If paraIndex < 1 Or paraIndex > paraCount
      ProcedureReturn
    EndIf
    Define result.s = ""
    Define pi.i
    For pi = 1 To paraCount
      If pi > 1
        result + "#PARA#"
      EndIf
      Define paraLine.s = StringField(paraText, pi, "#PARA#")
      If pi = paraIndex
        Define runCount.i = CountString(paraLine, "#RUN#") + 1
        Define newParaLine.s = ""
        Define ri.i
        For ri = 1 To runCount
          If ri > 1
            newParaLine + "#RUN#"
          EndIf
          Define runPart.s = StringField(paraLine, ri, "#RUN#")
          Define runText.s = runPart
          Define runAttr.s = ""
          Define attrPos.i = FindString(runPart, "#ATTR#")
          If attrPos > 0
            runText = Left(runPart, attrPos - 1)
            runAttr = Mid(runPart, attrPos + 6)
          EndIf
          If ri = runIndex
            Define fnName.s = StringField(runAttr, 1, ",")
            Define fnSize.s = StringField(runAttr, 2, ",")
            Define fnBold.s = StringField(runAttr, 3, ",")
            Define fnItalic.s = StringField(runAttr, 4, ",")
            Define fnUnderline.s = StringField(runAttr, 5, ",")
            runAttr = fnName + "," + fnSize + "," + fnBold + "," + fnItalic + "," + fnUnderline + "," + color
          EndIf
          newParaLine + runText
          If runAttr <> ""
            newParaLine + "#ATTR#" + runAttr
          EndIf
        Next
        paraLine = newParaLine
      EndIf
      result + paraLine
    Next
    PbPPT_Shapes()\textFrame\paragraphs = result
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; SetChartTitle - 设置图表标题
Procedure PbPPT_SetChartTitle(shapeId.i, title.s)
  If PbPPT_GetShapeById(shapeId)
    PbPPT_Shapes()\chartTitle = title
  EndIf
EndProcedure

; AddChartCategory - 添加图表分类标签
Procedure PbPPT_AddChartCategory(shapeId.i, label.s)
  If PbPPT_GetShapeById(shapeId)
    Define current.s = PbPPT_Shapes()\chartData
    ; 格式：分类1,分类2,...|系列名:值1,值2,...
    Define pipePos.i = FindString(current, "|")
    If pipePos > 0
      Define cats.s = Left(current, pipePos - 1)
      Define series.s = Mid(current, pipePos + 1)
      cats + "," + label
      PbPPT_Shapes()\chartData = cats + "|" + series
    Else
      If current <> ""
        PbPPT_Shapes()\chartData = current + "," + label
      Else
        PbPPT_Shapes()\chartData = label
      EndIf
    EndIf
  EndIf
EndProcedure

; AddChartSeries - 添加图表数据系列
Procedure PbPPT_AddChartSeries(shapeId.i, name.s, values.s)
  If PbPPT_GetShapeById(shapeId)
    Define current.s = PbPPT_Shapes()\chartData
    Define pipePos.i = FindString(current, "|")
    If pipePos > 0
      Define cats.s = Left(current, pipePos - 1)
      Define series.s = Mid(current, pipePos + 1)
      If series <> ""
        series + ";" + name + ":" + values
      Else
        series = name + ":" + values
      EndIf
      PbPPT_Shapes()\chartData = cats + "|" + series
    Else
      ; 还没有分类，先添加系列
      PbPPT_Shapes()\chartData = current + "|" + name + ":" + values
    EndIf
  EndIf
EndProcedure

; SetChartHasLegend - 设置图表是否显示图例
Procedure PbPPT_SetChartHasLegend(shapeId.i, hasLegend.b)
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType <> #PbPPT_ShapeTypeChart
      ProcedureReturn
    EndIf
    PbPPT_Shapes()\chartHasLegend = hasLegend
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; SetChartStyle - 设置图表样式
Procedure PbPPT_SetChartStyle(shapeId.i, style.i)
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType <> #PbPPT_ShapeTypeChart
      ProcedureReturn
    EndIf
    PbPPT_Shapes()\chartStyle = style
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; SetFillSolid - 设置纯色填充
Procedure PbPPT_SetFillSolid(shapeId.i, color.s)
  PbPPT_SetShapeFill(shapeId, #PbPPT_FillTypeSolid, color, "")
EndProcedure

; SetFillGradient - 设置渐变填充
Procedure PbPPT_SetFillGradient(shapeId.i, angle.f, foreColor.s, backColor.s)
  PbPPT_SetShapeFill(shapeId, #PbPPT_FillTypeGradient, foreColor, backColor.s)
EndProcedure

; SetFillPattern - 设置图案填充
Procedure PbPPT_SetFillPattern(shapeId.i, foreColor.s, backColor.s)
  PbPPT_SetShapeFill(shapeId, #PbPPT_FillTypePattern, foreColor, backColor)
EndProcedure

; SetFillNone - 设置无填充
Procedure PbPPT_SetFillNone(shapeId.i)
  PbPPT_SetShapeFill(shapeId, #PbPPT_FillTypeNone, "", "")
EndProcedure

; SetLineFormat - 设置线条格式
Procedure PbPPT_SetLineFormat(shapeId.i, width.i, color.s)
  PbPPT_SetShapeLine(shapeId, width, color, "solid")
EndProcedure

; SetLineDashStyle - 设置线条虚线样式
Procedure PbPPT_SetLineDashStyle(shapeId.i, dashStyle.s)
  If PbPPT_GetShapeById(shapeId)
    PbPPT_Shapes()\line\dashStyle = dashStyle
  EndIf
EndProcedure

; SetFillPicture - 设置图片填充
Procedure PbPPT_SetFillPicture(shapeId.i, imagePath.s)
  If PbPPT_GetShapeById(shapeId)
    PbPPT_Shapes()\fill\fillType = #PbPPT_FillTypePicture
    PbPPT_Shapes()\fill\foreColor = imagePath
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; SetShadow - 设置阴影效果
Procedure PbPPT_SetShadow(shapeId.i, shadowType.i, color.s, blur.i, offset.i)
  ; 阴影信息暂存储在fill的backColor字段中（简化处理）
  ; 完整实现需要在WriteSlideXML中生成a:effectLst/a:outerShdw节点
  If PbPPT_GetShapeById(shapeId)
    PbPPT_Shapes()\fill\backColor = "SHADOW:" + Str(shadowType) + "," + color + "," + Str(blur) + "," + Str(offset)
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; SetLineHeadEnd - 设置线头样式（箭头）
Procedure PbPPT_SetLineHeadEnd(shapeId.i, headStyle.s)
  ; 线头样式存储在line的dashStyle字段中（扩展格式）
  If PbPPT_GetShapeById(shapeId)
    PbPPT_Shapes()\line\dashStyle = PbPPT_Shapes()\line\dashStyle + "#HEAD#" + headStyle
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; SetLineTailEnd - 设置线尾样式（箭头）
Procedure PbPPT_SetLineTailEnd(shapeId.i, tailStyle.s)
  ; 线尾样式存储在line的dashStyle字段中（扩展格式）
  If PbPPT_GetShapeById(shapeId)
    PbPPT_Shapes()\line\dashStyle = PbPPT_Shapes()\line\dashStyle + "#TAIL#" + tailStyle
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; ***************************************************************************************
; 分区15: 表格模块（表格创建、单元格、行列、合并）
; ***************************************************************************************

; SetCellValue - 设置表格单元格文本
Procedure PbPPT_SetCellValue(shapeId.i, row.i, col.i, text.s)
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType <> #PbPPT_ShapeTypeTable
      ProcedureReturn
    EndIf
    Define cellKey.s = Str(row) + "_" + Str(col)
    Define cellData.s = PbPPT_Shapes()\tableCells
    ; 检查是否已存在该单元格
    Define found.b = #False
    Define result.s = ""
    If cellData <> ""
      Define pairCount.i = CountString(cellData, "|") + 1
      Define pi.i
      For pi = 1 To pairCount
        Define pair.s = StringField(cellData, pi, "|")
        Define eqPos.i = FindString(pair, "=")
        If eqPos > 0
          Define pairKey.s = Left(pair, eqPos - 1)
          If pairKey = cellKey
            ; 替换现有值
            If result <> ""
              result + "|"
            EndIf
            result + cellKey + "=" + text
            found = #True
          Else
            If result <> ""
              result + "|"
            EndIf
            result + pair
          EndIf
        EndIf
      Next
    EndIf
    If Not found
      If result <> ""
        result + "|"
      EndIf
      result + cellKey + "=" + text
    EndIf
    PbPPT_Shapes()\tableCells = result
  EndIf
EndProcedure

; GetCellValue - 获取表格单元格文本
Procedure.s PbPPT_GetCellValue(shapeId.i, row.i, col.i)
  Define result.s = ""
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType <> #PbPPT_ShapeTypeTable
      ProcedureReturn ""
    EndIf
    Define cellKey.s = Str(row) + "_" + Str(col)
    Define cellData.s = PbPPT_Shapes()\tableCells
    If cellData <> ""
      Define pairCount.i = CountString(cellData, "|") + 1
      Define pi.i
      For pi = 1 To pairCount
        Define pair.s = StringField(cellData, pi, "|")
        Define eqPos.i = FindString(pair, "=")
        If eqPos > 0
          Define pairKey.s = Left(pair, eqPos - 1)
          If pairKey = cellKey
            result = Mid(pair, eqPos + 1)
            Break
          EndIf
        EndIf
      Next
    EndIf
  EndIf
  ProcedureReturn result
EndProcedure

; SetCellFont - 设置表格单元格字体
Procedure PbPPT_SetCellFont(shapeId.i, row.i, col.i, fontName.s, fontSize.f, bold.b, italic.b, color.s)
  ; 简化处理：表格单元格字体暂不支持单独设置
  ; 后续版本可扩展tableCells格式来支持单元格级别格式
EndProcedure

; SetCellFill - 设置表格单元格填充色
Procedure PbPPT_SetCellFill(shapeId.i, row.i, col.i, color.s)
  ; 单元格填充色存储在tableCells扩展格式中
  ; 简化处理：通过扩展tableCells格式存储填充色
  ; 格式扩展：行_列=文本#FILL#颜色
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType <> #PbPPT_ShapeTypeTable
      ProcedureReturn
    EndIf
    Define cellKey.s = Str(row) + "_" + Str(col)
    Define cellData.s = PbPPT_Shapes()\tableCells
    Define cellText.s = ""
    ; 先获取现有文本
    If cellData <> ""
      Define pairCount.i = CountString(cellData, "|") + 1
      Define pi.i
      For pi = 1 To pairCount
        Define pair.s = StringField(cellData, pi, "|")
        Define eqPos.i = FindString(pair, "=")
        If eqPos > 0
          Define pairKey.s = Left(pair, eqPos - 1)
          If pairKey = cellKey
            Define pairVal.s = Mid(pair, eqPos + 1)
            Define fillPos.i = FindString(pairVal, "#FILL#")
            If fillPos > 0
              cellText = Left(pairVal, fillPos - 1)
            Else
              cellText = pairVal
            EndIf
            Break
          EndIf
        EndIf
      Next
    EndIf
    ; 更新带填充色的单元格数据
    PbPPT_SetCellValue(shapeId, row, col, cellText + "#FILL#" + color)
  EndIf
EndProcedure

; SetCellAlignment - 设置表格单元格水平对齐
Procedure PbPPT_SetCellAlignment(shapeId.i, row.i, col.i, alignment.i)
  ; 对齐方式存储在扩展格式中：行_列=文本#FILL#颜色#ALIGN#对齐值
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType <> #PbPPT_ShapeTypeTable
      ProcedureReturn
    EndIf
    Define cellKey.s = Str(row) + "_" + Str(col)
    Define cellData.s = PbPPT_Shapes()\tableCells
    Define cellText.s = ""
    Define cellFill.s = ""
    ; 获取现有数据
    If cellData <> ""
      Define pairCount.i = CountString(cellData, "|") + 1
      Define pi.i
      For pi = 1 To pairCount
        Define pair.s = StringField(cellData, pi, "|")
        Define eqPos.i = FindString(pair, "=")
        If eqPos > 0
          Define pairKey.s = Left(pair, eqPos - 1)
          If pairKey = cellKey
            Define pairVal.s = Mid(pair, eqPos + 1)
            Define fillPos.i = FindString(pairVal, "#FILL#")
            If fillPos > 0
              cellText = Left(pairVal, fillPos - 1)
              Define afterFill.s = Mid(pairVal, fillPos + 6)
              Define alignPos.i = FindString(afterFill, "#ALIGN#")
              If alignPos > 0
                cellFill = Left(afterFill, alignPos - 1)
              Else
                cellFill = afterFill
              EndIf
            Else
              Define alignPos2.i = FindString(pairVal, "#ALIGN#")
              If alignPos2 > 0
                cellText = Left(pairVal, alignPos2 - 1)
              Else
                cellText = pairVal
              EndIf
            EndIf
            Break
          EndIf
        EndIf
      Next
    EndIf
    ; 重建单元格数据
    Define newVal.s = cellText
    If cellFill <> ""
      newVal + "#FILL#" + cellFill
    EndIf
    newVal + "#ALIGN#" + Str(alignment)
    PbPPT_SetCellValue(shapeId, row, col, newVal)
  EndIf
EndProcedure

; SetCellVerticalAlignment - 设置表格单元格垂直对齐
Procedure PbPPT_SetCellVerticalAlignment(shapeId.i, row.i, col.i, valign.i)
  ; 垂直对齐存储在扩展格式中：#VALIGN#值
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType <> #PbPPT_ShapeTypeTable
      ProcedureReturn
    EndIf
    Define cellKey.s = Str(row) + "_" + Str(col)
    Define cellData.s = PbPPT_Shapes()\tableCells
    Define cellText.s = ""
    ; 获取现有文本（去掉旧VALIGN）
    If cellData <> ""
      Define pairCount.i = CountString(cellData, "|") + 1
      Define pi.i
      For pi = 1 To pairCount
        Define pair.s = StringField(cellData, pi, "|")
        Define eqPos.i = FindString(pair, "=")
        If eqPos > 0
          Define pairKey.s = Left(pair, eqPos - 1)
          If pairKey = cellKey
            Define pairVal.s = Mid(pair, eqPos + 1)
            Define valignPos.i = FindString(pairVal, "#VALIGN#")
            If valignPos > 0
              cellText = Left(pairVal, valignPos - 1)
            Else
              cellText = pairVal
            EndIf
            Break
          EndIf
        EndIf
      Next
    EndIf
    PbPPT_SetCellValue(shapeId, row, col, cellText + "#VALIGN#" + Str(valign))
  EndIf
EndProcedure

; SetCellMargin - 设置表格单元格边距
Procedure PbPPT_SetCellMargin(shapeId.i, row.i, col.i, left.i, top.i, right.i, bottom.i)
  ; 边距存储在扩展格式中：#MARGIN#左,上,右,下
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType <> #PbPPT_ShapeTypeTable
      ProcedureReturn
    EndIf
    Define cellKey.s = Str(row) + "_" + Str(col)
    Define cellData.s = PbPPT_Shapes()\tableCells
    Define cellText.s = ""
    If cellData <> ""
      Define pairCount.i = CountString(cellData, "|") + 1
      Define pi.i
      For pi = 1 To pairCount
        Define pair.s = StringField(cellData, pi, "|")
        Define eqPos.i = FindString(pair, "=")
        If eqPos > 0
          Define pairKey.s = Left(pair, eqPos - 1)
          If pairKey = cellKey
            Define pairVal.s = Mid(pair, eqPos + 1)
            Define marginPos.i = FindString(pairVal, "#MARGIN#")
            If marginPos > 0
              cellText = Left(pairVal, marginPos - 1)
            Else
              cellText = pairVal
            EndIf
            Break
          EndIf
        EndIf
      Next
    EndIf
    PbPPT_SetCellValue(shapeId, row, col, cellText + "#MARGIN#" + Str(left) + "," + Str(top) + "," + Str(right) + "," + Str(bottom))
  EndIf
EndProcedure

; MergeCells - 合并单元格
Procedure PbPPT_MergeCells(shapeId.i, startRow.i, startCol.i, endRow.i, endCol.i)
  ; 合并信息存储在tableCells扩展格式中：#MERGE#起始行,起始列,结束行,结束列
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType <> #PbPPT_ShapeTypeTable
      ProcedureReturn
    EndIf
    Define ri.i
    For ri = startRow To endRow
      Define ci.i
      For ci = startCol To endCol
        Define cellKey.s = Str(ri) + "_" + Str(ci)
        Define cellData.s = PbPPT_Shapes()\tableCells
        Define cellText.s = ""
        ; 获取现有文本
        If cellData <> ""
          Define pairCount.i = CountString(cellData, "|") + 1
          Define pi.i
          For pi = 1 To pairCount
            Define pair.s = StringField(cellData, pi, "|")
            Define eqPos.i = FindString(pair, "=")
            If eqPos > 0
              Define pairKey.s = Left(pair, eqPos - 1)
              If pairKey = cellKey
                Define pairVal.s = Mid(pair, eqPos + 1)
                Define mergePos.i = FindString(pairVal, "#MERGE#")
                If mergePos > 0
                  cellText = Left(pairVal, mergePos - 1)
                Else
                  cellText = pairVal
                EndIf
                Break
              EndIf
            EndIf
          Next
        EndIf
        ; 标记合并信息
        Define mergeInfo.s = Str(startRow) + "," + Str(startCol) + "," + Str(endRow) + "," + Str(endCol)
        PbPPT_SetCellValue(shapeId, ri, ci, cellText + "#MERGE#" + mergeInfo)
      Next
    Next
  EndIf
EndProcedure

; SplitCells - 拆分合并单元格
Procedure PbPPT_SplitCells(shapeId.i, row.i, col.i)
  ; 移除合并标记
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType <> #PbPPT_ShapeTypeTable
      ProcedureReturn
    EndIf
    Define cellKey.s = Str(row) + "_" + Str(col)
    Define cellData.s = PbPPT_Shapes()\tableCells
    Define cellText.s = ""
    If cellData <> ""
      Define pairCount.i = CountString(cellData, "|") + 1
      Define pi.i
      For pi = 1 To pairCount
        Define pair.s = StringField(cellData, pi, "|")
        Define eqPos.i = FindString(pair, "=")
        If eqPos > 0
          Define pairKey.s = Left(pair, eqPos - 1)
          If pairKey = cellKey
            Define pairVal.s = Mid(pair, eqPos + 1)
            Define mergePos.i = FindString(pairVal, "#MERGE#")
            If mergePos > 0
              cellText = Left(pairVal, mergePos - 1)
            Else
              cellText = pairVal
            EndIf
            Break
          EndIf
        EndIf
      Next
    EndIf
    PbPPT_SetCellValue(shapeId, row, col, cellText)
  EndIf
EndProcedure

; SetRowHeight - 设置表格行高
Procedure PbPPT_SetRowHeight(shapeId.i, rowIndex.i, height.i)
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType <> #PbPPT_ShapeTypeTable
      ProcedureReturn
    EndIf
    Define heights.s = PbPPT_Shapes()\tableRowHeights
    If heights = ""
      ; 初始化行高列表
      Define ri.i
      For ri = 1 To PbPPT_Shapes()\tableRows
        If ri > 1
          heights + ","
        EndIf
        heights + Str(PbPPT_Shapes()\height / PbPPT_Shapes()\tableRows)
      Next
    EndIf
    ; 更新指定行高
    Define result.s = ""
    Define count.i = CountString(heights, ",") + 1
    Define ci.i
    For ci = 1 To count
      If ci > 1
        result + ","
      EndIf
      If ci = rowIndex
        result + Str(height)
      Else
        result + StringField(heights, ci, ",")
      EndIf
    Next
    PbPPT_Shapes()\tableRowHeights = result
  EndIf
EndProcedure

; SetColWidth - 设置表格列宽
Procedure PbPPT_SetColWidth(shapeId.i, colIndex.i, width.i)
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType <> #PbPPT_ShapeTypeTable
      ProcedureReturn
    EndIf
    Define widths.s = PbPPT_Shapes()\tableColWidths
    If widths = ""
      ; 初始化列宽列表
      Define ci.i
      For ci = 1 To PbPPT_Shapes()\tableCols
        If ci > 1
          widths + ","
        EndIf
        widths + Str(PbPPT_Shapes()\width / PbPPT_Shapes()\tableCols)
      Next
    EndIf
    ; 更新指定列宽
    Define result.s = ""
    Define count.i = CountString(widths, ",") + 1
    Define ci.i
    For ci = 1 To count
      If ci > 1
        result + ","
      EndIf
      If ci = colIndex
        result + Str(width)
      Else
        result + StringField(widths, ci, ",")
      EndIf
    Next
    PbPPT_Shapes()\tableColWidths = result
  EndIf
EndProcedure

; GetRowCount - 获取表格行数
Procedure.i PbPPT_GetRowCount(shapeId.i)
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType = #PbPPT_ShapeTypeTable
      ProcedureReturn PbPPT_Shapes()\tableRows
    EndIf
  EndIf
  ProcedureReturn 0
EndProcedure

; GetColCount - 获取表格列数
Procedure.i PbPPT_GetColCount(shapeId.i)
  If PbPPT_GetShapeById(shapeId)
    If PbPPT_Shapes()\shapeType = #PbPPT_ShapeTypeTable
      ProcedureReturn PbPPT_Shapes()\tableCols
    EndIf
  EndIf
  ProcedureReturn 0
EndProcedure

; ***************************************************************************************
; 分区16: 公共API（用户可调用的所有函数接口）
; ***************************************************************************************

; PbPPT_Create - 创建新的演示文稿
Procedure.b PbPPT_Create()
  ; 清除所有数据
  PbPPT_ClearAllData()
  ; 设置默认值
  PbPPT_Prs\slideWidth = #PbPPT_DEFAULT_SLIDE_WIDTH
  PbPPT_Prs\slideHeight = #PbPPT_DEFAULT_SLIDE_HEIGHT
  PbPPT_Prs\isLoaded = #True
  PbPPT_Prs\isModified = #False
  PbPPT_Prs\nextShapeId = 2
  PbPPT_Prs\imageCount = 0
  PbPPT_Prs\creator = "PbPPT"
  PbPPT_Prs\lastModifiedBy = "PbPPT"
  PbPPT_Prs\created = PbPPT_GetCurrentDateTime()
  PbPPT_Prs\modified = PbPPT_GetCurrentDateTime()
  ; 添加默认版式
  Define layoutNames.s = "标题幻灯片,标题和内容,节标题,两栏内容,仅标题,空白,内容与标题,图片与标题,标题和竖排文字,竖排标题与文本"
  Define li.i
  For li = 1 To 10
    AddElement(PbPPT_Layouts())
    PbPPT_Layouts()\id = li
    PbPPT_Layouts()\name = StringField(layoutNames, li, ",")
    PbPPT_Layouts()\partName = "/ppt/slideLayouts/slideLayout" + Str(li) + ".xml"
    PbPPT_Layouts()\rId = "rId" + Str(li + 1)
  Next
  ; 添加默认母版
  AddElement(PbPPT_Masters())
  PbPPT_Masters()\id = 1
  PbPPT_Masters()\name = "幻灯片母版"
  PbPPT_Masters()\partName = "/ppt/slideMasters/slideMaster1.xml"
  PbPPT_Masters()\rId = "rId1"
  ProcedureReturn #True
EndProcedure

; PbPPT_Open - 打开已有的演示文稿
Procedure.b PbPPT_Open(filename.s)
  If FileSize(filename) <= 0
    ProcedureReturn #False
  EndIf
  ; 清除所有数据
  PbPPT_ClearAllData()
  ; 解压到临时目录
  If Not PbPPT_ExtractToTempDir(filename)
    ProcedureReturn #False
  EndIf
  ; 读取演示文稿信息
  PbPPT_ReadPresentationInfo()
  ; 读取核心属性
  PbPPT_ReadCoreProps()
  ; 读取幻灯片
  PbPPT_ReadSlides()
  ; 读取应用属性
  PbPPT_ReadAppProps()
  PbPPT_Prs\filePath = filename
  PbPPT_Prs\isLoaded = #True
  PbPPT_Prs\isModified = #False
  PbPPT_Prs\slideCount = ListSize(PbPPT_Slides())
  ; 添加默认版式
  Define li.i
  For li = 1 To 10
    AddElement(PbPPT_Layouts())
    PbPPT_Layouts()\id = li
    PbPPT_Layouts()\name = "版式 " + Str(li)
    PbPPT_Layouts()\partName = "/ppt/slideLayouts/slideLayout" + Str(li) + ".xml"
    PbPPT_Layouts()\rId = "rId" + Str(li + 1)
  Next
  AddElement(PbPPT_Masters())
  PbPPT_Masters()\id = 1
  PbPPT_Masters()\name = "幻灯片母版"
  PbPPT_Masters()\partName = "/ppt/slideMasters/slideMaster1.xml"
  PbPPT_Masters()\rId = "rId1"
  ProcedureReturn #True
EndProcedure

; PbPPT_Save - 保存演示文稿到文件
Procedure.b PbPPT_Save(filename.s)
  If Not PbPPT_Prs\isLoaded
    ProcedureReturn #False
  EndIf
  Define packId.i = PbPPT_ZIPCreate(filename)
  If packId = 0
    ProcedureReturn #False
  EndIf
  ; 写入[Content_Types].xml
  Define xmlId.i = PbPPT_WriteContentTypesXML()
  If xmlId = 0 : PbPPT_ZIPClose(packId) : ProcedureReturn #False : EndIf
  PbPPT_AddXMLToZIP(packId, xmlId, #PbPPT_ARCContentTypes$)
  PbPPT_XMLFree(xmlId)
  ; 写入_rels/.rels
  xmlId = PbPPT_WriteRootRelsXML()
  If xmlId = 0 : PbPPT_ZIPClose(packId) : ProcedureReturn #False : EndIf
  PbPPT_AddXMLToZIP(packId, xmlId, #PbPPT_ARCRootRels$)
  PbPPT_XMLFree(xmlId)
  ; 写入ppt/_rels/presentation.xml.rels（必须先于presentation.xml，因为此函数会设置幻灯片的rId）
  xmlId = PbPPT_WritePresentationRelsXML()
  If xmlId = 0 : PbPPT_ZIPClose(packId) : ProcedureReturn #False : EndIf
  PbPPT_AddXMLToZIP(packId, xmlId, #PbPPT_ARCPresentationRels$)
  PbPPT_XMLFree(xmlId)
  ; 写入ppt/presentation.xml（此时幻灯片的rId已被正确设置）
  xmlId = PbPPT_WritePresentationXML()
  If xmlId = 0 : PbPPT_ZIPClose(packId) : ProcedureReturn #False : EndIf
  PbPPT_AddXMLToZIP(packId, xmlId, #PbPPT_ARCPresentation$)
  PbPPT_XMLFree(xmlId)
  ; 写入docProps/core.xml
  xmlId = PbPPT_WriteCorePropsXML()
  If xmlId = 0 : PbPPT_ZIPClose(packId) : ProcedureReturn #False : EndIf
  PbPPT_AddXMLToZIP(packId, xmlId, #PbPPT_ARCCore$)
  PbPPT_XMLFree(xmlId)
  ; 写入docProps/app.xml
  xmlId = PbPPT_WriteAppPropsXML()
  If xmlId = 0 : PbPPT_ZIPClose(packId) : ProcedureReturn #False : EndIf
  PbPPT_AddXMLToZIP(packId, xmlId, #PbPPT_ARCApp$)
  PbPPT_XMLFree(xmlId)
  ; 写入ppt/presProps.xml
  xmlId = PbPPT_WritePresPropsXML()
  If xmlId = 0 : PbPPT_ZIPClose(packId) : ProcedureReturn #False : EndIf
  PbPPT_AddXMLToZIP(packId, xmlId, #PbPPT_ARCPresProps$)
  PbPPT_XMLFree(xmlId)
  ; 写入ppt/viewProps.xml
  xmlId = PbPPT_WriteViewPropsXML()
  If xmlId = 0 : PbPPT_ZIPClose(packId) : ProcedureReturn #False : EndIf
  PbPPT_AddXMLToZIP(packId, xmlId, #PbPPT_ARCViewProps$)
  PbPPT_XMLFree(xmlId)
  ; 写入ppt/tableStyles.xml
  xmlId = PbPPT_WriteTableStylesXML()
  If xmlId = 0 : PbPPT_ZIPClose(packId) : ProcedureReturn #False : EndIf
  PbPPT_AddXMLToZIP(packId, xmlId, #PbPPT_ARCTableStyles$)
  PbPPT_XMLFree(xmlId)
  ; 写入ppt/theme/theme1.xml
  xmlId = PbPPT_WriteThemeXML()
  If xmlId = 0 : PbPPT_ZIPClose(packId) : ProcedureReturn #False : EndIf
  PbPPT_AddXMLToZIP(packId, xmlId, #PbPPT_ARCTheme$)
  PbPPT_XMLFree(xmlId)
  ; 写入ppt/slideMasters/slideMaster1.xml
  xmlId = PbPPT_WriteSlideMasterXML()
  If xmlId = 0 : PbPPT_ZIPClose(packId) : ProcedureReturn #False : EndIf
  PbPPT_AddXMLToZIP(packId, xmlId, #PbPPT_ARCSlideMaster$)
  PbPPT_XMLFree(xmlId)
  ; 写入ppt/slideMasters/_rels/slideMaster1.xml.rels
  xmlId = PbPPT_WriteSlideMasterRelsXML()
  If xmlId = 0 : PbPPT_ZIPClose(packId) : ProcedureReturn #False : EndIf
  PbPPT_AddXMLToZIP(packId, xmlId, #PbPPT_ARCSlideMasterRels$)
  PbPPT_XMLFree(xmlId)
  ; 写入幻灯片版式
  Define layoutNames.s = "标题幻灯片,标题和内容,节标题,两栏内容,仅标题,空白,内容与标题,图片与标题,标题和竖排文字,竖排标题与文本"
  Define li.i
  For li = 1 To 10
    xmlId = PbPPT_WriteSlideLayoutXML(StringField(layoutNames, li, ","), li)
    PbPPT_AddXMLToZIP(packId, xmlId, "ppt/slideLayouts/slideLayout" + Str(li) + ".xml")
    PbPPT_XMLFree(xmlId)
    ; 版式关系文件
    Define layoutRelsXml.s = ~"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + #PbPPT_NS_REL$ + ~"\"><Relationship Id=\"rId1\" Type=\"" + #PbPPT_RT_SLIDE_MASTER$ + ~"\" Target=\"../slideMasters/slideMaster1.xml\"/></Relationships>"
    PbPPT_AddStringToZIP(packId, layoutRelsXml, "ppt/slideLayouts/_rels/slideLayout" + Str(li) + ".xml.rels")
  Next
  ; 写入每张幻灯片
  Define slideIdx.i = 0
  ForEach PbPPT_Slides()
    slideIdx + 1
    ; 先写入幻灯片关系文件（会设置图片和图表的实际rId）
    xmlId = PbPPT_WriteSlideRelsXML(slideIdx - 1, PbPPT_Slides()\layoutRId)
    PbPPT_AddXMLToZIP(packId, xmlId, "ppt/slides/_rels/slide" + Str(slideIdx) + ".xml.rels")
    PbPPT_XMLFree(xmlId)
    ; 再写入幻灯片XML（此时rId已被正确设置）
    xmlId = PbPPT_WriteSlideXML(slideIdx - 1)
    PbPPT_AddXMLToZIP(packId, xmlId, "ppt/slides/slide" + Str(slideIdx) + ".xml")
    PbPPT_XMLFree(xmlId)
  Next
  ; 写入图片二进制数据
  ForEach PbPPT_Parts()
    If PbPPT_Parts()\blobSize > 0 And PbPPT_Parts()\blobData
      Define archivePath.s = Mid(PbPPT_Parts()\partName, 2)
      PbPPT_AddBinaryToZIP(packId, PbPPT_Parts()\blobData, PbPPT_Parts()\blobSize, archivePath)
    EndIf
  Next
  ; 写入图表XML文件
  ForEach PbPPT_Shapes()
    If PbPPT_Shapes()\shapeType = #PbPPT_ShapeTypeChart
      Define chartXmlId.i = PbPPT_WriteChartXML(PbPPT_Shapes()\id)
      If chartXmlId
        PbPPT_AddXMLToZIP(packId, chartXmlId, "ppt/charts/chart" + Str(PbPPT_Shapes()\id) + ".xml")
        PbPPT_XMLFree(chartXmlId)
      EndIf
    EndIf
  Next
  PbPPT_ZIPClose(packId)
  PbPPT_Prs\filePath = filename
  PbPPT_Prs\isModified = #False
  ProcedureReturn #True
EndProcedure

; PbPPT_Close - 关闭演示文稿
Procedure PbPPT_Close()
  ; 释放所有二进制数据
  ForEach PbPPT_Parts()
    If PbPPT_Parts()\blobSize > 0 And PbPPT_Parts()\blobData <> #Null
      FreeMemory(PbPPT_Parts()\blobData)
      PbPPT_Parts()\blobData = #Null
      PbPPT_Parts()\blobSize = 0
    EndIf
  Next
  ; 清理临时目录
  PbPPT_CleanupTempDir()
  ; 清除所有数据
  PbPPT_ClearAllData()
EndProcedure

; PbPPT_SetSlideWidth - 设置幻灯片宽度
Procedure PbPPT_SetSlideWidth(width.i)
  PbPPT_Prs\slideWidth = width
  PbPPT_Prs\isModified = #True
EndProcedure

; PbPPT_SetSlideHeight - 设置幻灯片高度
Procedure PbPPT_SetSlideHeight(height.i)
  PbPPT_Prs\slideHeight = height
  PbPPT_Prs\isModified = #True
EndProcedure

; PbPPT_GetSlideWidth - 获取幻灯片宽度
Procedure.i PbPPT_GetSlideWidth()
  ProcedureReturn PbPPT_Prs\slideWidth
EndProcedure

; PbPPT_GetSlideHeight - 获取幻灯片高度
Procedure.i PbPPT_GetSlideHeight()
  ProcedureReturn PbPPT_Prs\slideHeight
EndProcedure

; PbPPT_GetSlideCount - 获取幻灯片数量
Procedure.i PbPPT_GetSlideCount()
  ProcedureReturn ListSize(PbPPT_Slides())
EndProcedure

; PbPPT_AddSlide - 添加幻灯片
Procedure.i PbPPT_AddSlide(layoutIndex.i)
  AddElement(PbPPT_Slides())
  Define slideIdx.i = ListSize(PbPPT_Slides())
  PbPPT_Slides()\id = slideIdx
  PbPPT_Slides()\slideId = 255 + slideIdx
  PbPPT_Slides()\name = "幻灯片 " + Str(slideIdx)
  PbPPT_Slides()\partName = "/ppt/slides/slide" + Str(slideIdx) + ".xml"
  PbPPT_Slides()\rId = ""
  PbPPT_Slides()\layoutRId = "rId1"
  PbPPT_Slides()\layoutIndex = layoutIndex
  PbPPT_Prs\slideCount = slideIdx
  PbPPT_Prs\isModified = #True
  ProcedureReturn slideIdx - 1
EndProcedure

; PbPPT_DeleteSlide - 删除幻灯片
Procedure PbPPT_DeleteSlide(slideIndex.i)
  Define idx.i = 0
  ForEach PbPPT_Slides()
    If idx = slideIndex
      ; 删除该幻灯片上的所有形状
      Define delSlideId.i = PbPPT_Slides()\id
      ForEach PbPPT_Shapes()
        If PbPPT_Shapes()\slideIndex = slideIndex
          DeleteElement(PbPPT_Shapes())
        EndIf
      Next
      ; 删除幻灯片
      DeleteElement(PbPPT_Slides())
      PbPPT_Prs\slideCount = ListSize(PbPPT_Slides())
      PbPPT_Prs\isModified = #True
      ; 更新后续幻灯片的索引和形状的slideIndex
      Define newIdx.i = slideIndex
      ForEach PbPPT_Slides()
        If ListIndex(PbPPT_Slides()) >= slideIndex
          ; 更新形状的slideIndex引用
          ForEach PbPPT_Shapes()
            If PbPPT_Shapes()\slideIndex > slideIndex
              PbPPT_Shapes()\slideIndex - 1
            EndIf
          Next
        EndIf
      Next
      Break
    EndIf
    idx + 1
  Next
EndProcedure

; PbPPT_GetSlide - 获取幻灯片ID（通过索引）
Procedure.i PbPPT_GetSlide(slideIndex.i)
  Define idx.i = 0
  ForEach PbPPT_Slides()
    If idx = slideIndex
      ProcedureReturn PbPPT_Slides()\id
    EndIf
    idx + 1
  Next
  ProcedureReturn -1
EndProcedure

; PbPPT_GetSlideName - 获取幻灯片名称
Procedure.s PbPPT_GetSlideName(slideIndex.i)
  Define idx.i = 0
  ForEach PbPPT_Slides()
    If idx = slideIndex
      ProcedureReturn PbPPT_Slides()\name
    EndIf
    idx + 1
  Next
  ProcedureReturn ""
EndProcedure

; PbPPT_AddShape - 添加自动形状
Procedure.i PbPPT_AddShape(slideIndex.i, preset.s, left.i, top.i, width.i, height.i)
  Define shapeId.i = PbPPT_CreateAutoShape(slideIndex, preset, left, top, width, height)
  PbPPT_Prs\isModified = #True
  ProcedureReturn shapeId
EndProcedure

; PbPPT_AddTextbox - 添加文本框
Procedure.i PbPPT_AddTextbox(slideIndex.i, left.i, top.i, width.i, height.i)
  Define shapeId.i = PbPPT_CreateTextbox(slideIndex, left, top, width, height)
  PbPPT_Prs\isModified = #True
  ProcedureReturn shapeId
EndProcedure

; PbPPT_AddPicture - 添加图片
Procedure.i PbPPT_AddPicture(slideIndex.i, imagePath.s, left.i, top.i, width.i, height.i)
  Define shapeId.i = PbPPT_CreatePicture(slideIndex, imagePath, left, top, width, height)
  PbPPT_Prs\isModified = #True
  ProcedureReturn shapeId
EndProcedure

; PbPPT_AddTable - 添加表格
Procedure.i PbPPT_AddTable(slideIndex.i, rows.i, cols.i, left.i, top.i, width.i, height.i)
  Define shapeId.i = PbPPT_CreateTable(slideIndex, rows, cols, left, top, width, height)
  PbPPT_Prs\isModified = #True
  ProcedureReturn shapeId
EndProcedure

; PbPPT_AddChart - 添加图表
Procedure.i PbPPT_AddChart(slideIndex.i, chartType.i, left.i, top.i, width.i, height.i)
  Define shapeId.i = PbPPT_CreateChart(slideIndex, chartType, left, top, width, height)
  PbPPT_Prs\isModified = #True
  ProcedureReturn shapeId
EndProcedure

; PbPPT_AddConnector - 添加连接符
Procedure.i PbPPT_AddConnector(slideIndex.i, beginX.i, beginY.i, endX.i, endY.i)
  Define shapeId.i = PbPPT_CreateConnector(slideIndex, beginX, beginY, endX, endY)
  PbPPT_Prs\isModified = #True
  ProcedureReturn shapeId
EndProcedure

; PbPPT_SetShapeText - 设置形状文本
Procedure PbPPT_SetShapeText(shapeId.i, text.s)
  PbPPT_SetTextFrameText(shapeId, text)
  PbPPT_Prs\isModified = #True
EndProcedure

; PbPPT_GetShapeText - 获取形状文本
Procedure.s PbPPT_GetShapeText(shapeId.i)
  ProcedureReturn PbPPT_GetTextFrameText(shapeId)
EndProcedure

; PbPPT_SetTitle - 设置文档标题
Procedure PbPPT_SetTitle(title.s)
  PbPPT_Prs\title = title
  PbPPT_Prs\isModified = #True
EndProcedure

; PbPPT_GetTitle - 获取文档标题
Procedure.s PbPPT_GetTitle()
  ProcedureReturn PbPPT_Prs\title
EndProcedure

; PbPPT_SetAuthor - 设置文档作者
Procedure PbPPT_SetAuthor(author.s)
  PbPPT_Prs\creator = author
  PbPPT_Prs\isModified = #True
EndProcedure

; PbPPT_GetAuthor - 获取文档作者
Procedure.s PbPPT_GetAuthor()
  ProcedureReturn PbPPT_Prs\creator
EndProcedure

; PbPPT_SetSubject - 设置文档主题
Procedure PbPPT_SetSubject(subject.s)
  PbPPT_Prs\subject = subject
  PbPPT_Prs\isModified = #True
EndProcedure

; PbPPT_SetKeywords - 设置文档关键词
Procedure PbPPT_SetKeywords(keywords.s)
  PbPPT_Prs\keywords = keywords
  PbPPT_Prs\isModified = #True
EndProcedure

; PbPPT_SetCategory - 设置文档分类
Procedure PbPPT_SetCategory(category.s)
  PbPPT_Prs\category = category
  PbPPT_Prs\isModified = #True
EndProcedure

; PbPPT_AddHyperlink - 添加超链接
Procedure PbPPT_AddHyperlink(shapeId.i, paraIdx.i, runIdx.i, url.s)
  If PbPPT_GetShapeById(shapeId)
    PbPPT_Shapes()\hyperlink = url
    PbPPT_Prs\isModified = #True
  EndIf
EndProcedure

; ***************************************************************************************
; 分区17: 初始化和清理
; ***************************************************************************************

; PbPPT_Init - 库初始化
Procedure PbPPT_Init()
  UseZipPacker()
  PbPPT_Prs\slideWidth = #PbPPT_DEFAULT_SLIDE_WIDTH
  PbPPT_Prs\slideHeight = #PbPPT_DEFAULT_SLIDE_HEIGHT
  PbPPT_Prs\isLoaded = #False
  PbPPT_Prs\nextShapeId = 2
  PbPPT_Prs\imageCount = 0
EndProcedure

; ***************************************************************************************
; 分区18: 测试代码
; ***************************************************************************************

; Test1 - 创建空白演示文稿
Procedure.b PbPPT_Test1()
  PbPPT_Create()
  PbPPT_AddSlide(0)
  PbPPT_AddSlide(0)
  PbPPT_AddSlide(0)
  Define result.b = PbPPT_Save("test1.pptx")
  PbPPT_Close()
  ProcedureReturn result
EndProcedure

; Test2 - 添加文本
Procedure.b PbPPT_Test2()
  PbPPT_Create()
  Define slideIdx.i = PbPPT_AddSlide(0)
  Define tbId.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(2))
  PbPPT_SetShapeText(tbId, "PbPPT测试标题")
  Define tbId2.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(3.5), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(3))
  PbPPT_SetShapeText(tbId2, "这是第一行文本" + #LF$ + "这是第二行文本" + #LF$ + "这是第三行文本")
  Define result.b = PbPPT_Save("test2.pptx")
  PbPPT_Close()
  ProcedureReturn result
EndProcedure

; Test3 - 添加形状
Procedure.b PbPPT_Test3()
  PbPPT_Create()
  Define slideIdx.i = PbPPT_AddSlide(0)
  ; 矩形
  Define rectId.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstRectangle$, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(1.5))
  PbPPT_SetShapeFill(rectId, #PbPPT_FillTypeSolid, "4472C4", "")
  ; 圆形
  Define ellipseId.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstEllipse$, PbPPT_InchesToEmu(4), PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(1.5))
  PbPPT_SetShapeFill(ellipseId, #PbPPT_FillTypeSolid, "ED7D31", "")
  ; 三角形
  Define triId.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstTriangle$, PbPPT_InchesToEmu(7), PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(1.5))
  PbPPT_SetShapeFill(triId, #PbPPT_FillTypeSolid, "70AD47", "")
  Define result.b = PbPPT_Save("test3.pptx")
  PbPPT_Close()
  ProcedureReturn result
EndProcedure

; Test4 - 添加图片
Procedure.b PbPPT_Test4(imagePath.s)
  If FileSize(imagePath) <= 0
    ProcedureReturn #False
  EndIf
  PbPPT_Create()
  Define slideIdx.i = PbPPT_AddSlide(0)
  PbPPT_AddPicture(slideIdx, imagePath, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(4), PbPPT_InchesToEmu(3))
  Define result.b = PbPPT_Save("test4.pptx")
  PbPPT_Close()
  ProcedureReturn result
EndProcedure

; Test5 - 添加表格
Procedure.b PbPPT_Test5()
  PbPPT_Create()
  Define slideIdx.i = PbPPT_AddSlide(0)
  Define tblId.i = PbPPT_AddTable(slideIdx, 3, 4, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(4))
  ; 设置单元格文本
  PbPPT_SetCellValue(tblId, 1, 1, "姓名")
  PbPPT_SetCellValue(tblId, 1, 2, "年龄")
  PbPPT_SetCellValue(tblId, 1, 3, "城市")
  PbPPT_SetCellValue(tblId, 1, 4, "职业")
  PbPPT_SetCellValue(tblId, 2, 1, "张三")
  PbPPT_SetCellValue(tblId, 2, 2, "28")
  PbPPT_SetCellValue(tblId, 2, 3, "北京")
  PbPPT_SetCellValue(tblId, 2, 4, "工程师")
  PbPPT_SetCellValue(tblId, 3, 1, "李四")
  PbPPT_SetCellValue(tblId, 3, 2, "35")
  PbPPT_SetCellValue(tblId, 3, 3, "上海")
  PbPPT_SetCellValue(tblId, 3, 4, "设计师")
  Define result.b = PbPPT_Save("test5.pptx")
  PbPPT_Close()
  ProcedureReturn result
EndProcedure

; Test6 - 打开已有文件
Procedure.b PbPPT_Test6(filename.s)
  If FileSize(filename) <= 0
    ProcedureReturn #False
  EndIf
  If Not PbPPT_Open(filename)
    ProcedureReturn #False
  EndIf
  Define slideCount.i = PbPPT_GetSlideCount()
  Define width.i = PbPPT_GetSlideWidth()
  Define height.i = PbPPT_GetSlideHeight()
  Define title.s = PbPPT_GetTitle()
  Define author.s = PbPPT_GetAuthor()
  PbPPT_Close()
  ProcedureReturn #True
EndProcedure

; Test7 - 修改并保存
Procedure.b PbPPT_Test7(inputFile.s, outputFile.s)
  If Not PbPPT_Open(inputFile)
    ProcedureReturn #False
  EndIf
  ; 修改属性
  PbPPT_SetTitle("修改后的标题")
  PbPPT_SetAuthor("PbPPT测试")
  ; 添加新幻灯片
  Define slideIdx.i = PbPPT_AddSlide(0)
  Define tbId.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(2))
  PbPPT_SetShapeText(tbId, "这是新添加的幻灯片")
  Define result.b = PbPPT_Save(outputFile)
  PbPPT_Close()
  ProcedureReturn result
EndProcedure

; Test8 - 添加图表
Procedure.b PbPPT_Test8()
  PbPPT_Create()
  Define slideIdx.i = PbPPT_AddSlide(0)
  Define chartId.i = PbPPT_AddChart(slideIdx, #PbPPT_ChartTypeColumnClustered, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(5))
  PbPPT_SetChartTitle(chartId, "销售数据")
  PbPPT_AddChartCategory(chartId, "第一季度")
  PbPPT_AddChartCategory(chartId, "第二季度")
  PbPPT_AddChartCategory(chartId, "第三季度")
  PbPPT_AddChartCategory(chartId, "第四季度")
  PbPPT_AddChartSeries(chartId, "产品A", "100,200,150,300")
  PbPPT_AddChartSeries(chartId, "产品B", "80,150,200,250")
  Define result.b = PbPPT_Save("test8.pptx")
  PbPPT_Close()
  ProcedureReturn result
EndProcedure

; Test9 - 样式设置
Procedure.b PbPPT_Test9()
  PbPPT_Create()
  Define slideIdx.i = PbPPT_AddSlide(0)
  ; 带填充和线条的矩形
  Define rectId.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstRectangle$, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(3), PbPPT_InchesToEmu(2))
  PbPPT_SetFillSolid(rectId, "FF0000")
  PbPPT_SetLineFormat(rectId, PbPPT_PtToEmu(2), "000000")
  ; 带旋转的形状
  Define rectId2.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstRectangle$, PbPPT_InchesToEmu(5), PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(3), PbPPT_InchesToEmu(2))
  PbPPT_SetFillSolid(rectId2, "00FF00")
  PbPPT_SetShapeRotation(rectId2, 30)
  ; 带字体的文本框
  Define tbId.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(4), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(2))
  PbPPT_SetShapeText(tbId, "粗体红色文本")
  Define result.b = PbPPT_Save("test9.pptx")
  PbPPT_Close()
  ProcedureReturn result
EndProcedure

; Test10 - 幻灯片尺寸（16:9宽屏）
Procedure.b PbPPT_Test10()
  PbPPT_Create()
  PbPPT_SetSlideWidth(#PbPPT_WIDESCREEN_SLIDE_WIDTH)
  PbPPT_SetSlideHeight(#PbPPT_WIDESCREEN_SLIDE_HEIGHT)
  PbPPT_AddSlide(0)
  Define tbId.i = PbPPT_AddTextbox(0, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(11), PbPPT_InchesToEmu(2))
  PbPPT_SetShapeText(tbId, "16:9宽屏演示文稿")
  Define result.b = PbPPT_Save("test10.pptx")
  PbPPT_Close()
  ProcedureReturn result
EndProcedure

; PbPPT_RunAllTests - 运行所有测试
Procedure PbPPT_RunAllTests()
  Define results.s = ""
  If PbPPT_Test1()
    results + "测试1（空白演示文稿）: 通过" + #LF$
  Else
    results + "测试1（空白演示文稿）: 失败" + #LF$
  EndIf
  If PbPPT_Test2()
    results + "测试2（添加文本）: 通过" + #LF$
  Else
    results + "测试2（添加文本）: 失败" + #LF$
  EndIf
  If PbPPT_Test3()
    results + "测试3（添加形状）: 通过" + #LF$
  Else
    results + "测试3（添加形状）: 失败" + #LF$
  EndIf
  If PbPPT_Test5()
    results + "测试5（添加表格）: 通过" + #LF$
  Else
    results + "测试5（添加表格）: 失败" + #LF$
  EndIf
  If PbPPT_Test6("test1.pptx")
    results + "测试6（打开文件）: 通过" + #LF$
  Else
    results + "测试6（打开文件）: 失败" + #LF$
  EndIf
  If PbPPT_Test9()
    results + "测试9（样式设置）: 通过" + #LF$
  Else
    results + "测试9（样式设置）: 失败" + #LF$
  EndIf
  If PbPPT_Test10()
    results + "测试10（幻灯片尺寸）: 通过" + #LF$
  Else
    results + "测试10（幻灯片尺寸）: 失败" + #LF$
  EndIf
  MessageRequester("PbPPT 测试结果", results)
EndProcedure

; ReadSlideCount - 从presentation.xml读取幻灯片数量
Procedure.i PbPPT_ReadSlideCount()
  Define count.i = 0
  Define xmlFile.s = PbPPT_TempDir + "ppt/presentation.xml"
  If FileSize(xmlFile) <= 0
    ProcedureReturn 0
  EndIf
  Define xmlId.i = PbPPT_XMLParseFile(xmlFile)
  If xmlId = 0
    ProcedureReturn 0
  EndIf
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  ; 查找sldIdLst节点
  Define node.i = PbPPT_XMLGetChild(rootNode)
  While node
    Define nodeName.s = PbPPT_XMLGetNodeName(node)
    If FindString(nodeName, "sldIdLst")
      Define childNode.i = PbPPT_XMLGetChild(node)
      While childNode
        count + 1
        childNode = PbPPT_XMLGetNext(childNode)
      Wend
      Break
    EndIf
    node = PbPPT_XMLGetNext(node)
  Wend
  PbPPT_XMLFree(xmlId)
  ProcedureReturn count
EndProcedure

; ReadCoreProps - 读取核心属性
Procedure PbPPT_ReadCoreProps()
  Define xmlFile.s = PbPPT_TempDir + "docProps/core.xml"
  If FileSize(xmlFile) <= 0
    ProcedureReturn
  EndIf
  Define xmlId.i = PbPPT_XMLParseFile(xmlFile)
  If xmlId = 0
    ProcedureReturn
  EndIf
  Define rootNode.i = PbPPT_XMLGetRoot(xmlId)
  Define node.i = PbPPT_XMLGetChild(rootNode)
  While node
    Define nodeName.s = PbPPT_XMLGetNodeName(node)
    If FindString(nodeName, "title")
      PbPPT_Prs\title = PbPPT_XMLGetText(node)
    ElseIf FindString(nodeName, "creator")
      PbPPT_Prs\creator = PbPPT_XMLGetText(node)
    ElseIf FindString(nodeName, "subject")
      PbPPT_Prs\subject = PbPPT_XMLGetText(node)
    ElseIf FindString(nodeName, "description")
      PbPPT_Prs\description = PbPPT_XMLGetText(node)
    ElseIf FindString(nodeName, "keywords")
      PbPPT_Prs\keywords = PbPPT_XMLGetText(node)
    ElseIf FindString(nodeName, "created")
      PbPPT_Prs\created = PbPPT_XMLGetText(node)
    ElseIf FindString(nodeName, "modified")
      PbPPT_Prs\modified = PbPPT_XMLGetText(node)
    ElseIf FindString(nodeName, "lastModifiedBy")
      PbPPT_Prs\lastModifiedBy = PbPPT_XMLGetText(node)
    ElseIf FindString(nodeName, "category")
      PbPPT_Prs\category = PbPPT_XMLGetText(node)
    EndIf
    node = PbPPT_XMLGetNext(node)
  Wend
  PbPPT_XMLFree(xmlId)
EndProcedure
