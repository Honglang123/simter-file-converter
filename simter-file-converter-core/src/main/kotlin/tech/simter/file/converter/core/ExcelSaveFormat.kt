package tech.simter.file.converter.core

/**
 * 转换 Excel 文档格式时使用的格式参数常数定义。
 */
enum class ExcelSaveFormat(val format: String, val value: Int) {
  /**
   * PDF 格式 (*.pdf)
   *
   * 调用 ExportAsFixedFormat 方法进行转换所需
   */
  PDF("pdf", 0),

  /**
   * XPS 文档 (*.xps)
   *
   * 调用 ExportAsFixedFormat 方法进行转换所需
   */
  XPS("xps", 1),

  /**
   * Excel 97-2003 工作簿 (*.xls)
   */
  XLS("xls", 56),

  /**
   * Excel 2007+ 工作簿 (*.xlsx)
   */
  XLSX("xlsx", 51),

  /**
   * Excel 2007+ 启用宏的工作簿 (*.xlsm)
   */
  XLSM("xlsm", 52),

  /**
   * Excel 2003 XML 电子表格 (*.xml)
   */
  XML_2003("xml", 46),

  /**
   * 纯文本 (*.txt)
   */
  TXT("txt", 42),

  /**
   * CSV 格式 (*.csv)
   */
  CSV("csv", 6),

  /**
   * HTML 格式 (*.html)
   */
  HTML("html", 44),

  /**
   * 单个文件网页 (*.mht)
   */
  MHT("mht", 45)
}