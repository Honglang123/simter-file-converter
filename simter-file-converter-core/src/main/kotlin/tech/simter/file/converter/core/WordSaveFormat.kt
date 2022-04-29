package tech.simter.file.converter.core

/**
 * 转换 Word 文档格式时使用的格式参数常数定义。
 */
enum class WordSaveFormat(val format: String, val value: Int) {
  /**
   * PDF 格式 (*.pdf)
   */
  PDF("pdf", 17),

  /**
   * Word 2007+ 文档 (*.docx)
   */
  DOCX("docx", 12),

  /**
   * Word 2007+ 启用宏的文档 (*.docm)
   */
  DOCM("docm", 13),

  /**
   * Word 2007+ XML 文档 (*.xml)
   */
  XML_2007PLUS("xml", 19),

  /**
   * Word 97-2003 文件 (*.doc)
   */
  DOC("doc", 0),

  /**
   * Word 2003 XML 文档 (*.xml)
   */
  XML_2003("xml", 11),

  /**
   * 纯文本 (*.txt)
   */
  TXT("txt", 2),

  /**
   * RTF 格式 (*.rtf)
   */
  RTF("rtf", 6),

  /**
   * HTML 格式 (*.html)
   */
  HTML("html", 8),

  /**
   * 单个文件网页 (*.mht)
   */
  MHT("mht", 9)
}