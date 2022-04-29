package tech.simter.file.converter.core

/**
 * 转换 PowerPoint 文档格式时使用的格式参数常数定义。
 */
enum class PowerPointSaveFormat(val format: String, val value: Int) {
  /**
   * PDF 格式 (*.pdf)
   */
  PDF("pdf", 32),

  /**
   * PowerPoint 默认格式(*.pptx)
   */
  PPTX("pptx", 1),

  /**
   * PowerPoint 旧格式(*.ppt)
   */
  PPT("ppt", 0),

  /**
   * PowerPoint 默认格式(*.pptx)
   */
  DEFAULT("pptx", 11),

  /**
   * PowerPoint 2007+ 启用宏的文档 (*.pptm)
   */
  PPTM("pptm", 5),

  /**
   * PNG 格式 (*.png)
   */
  BMP("bmp", 19),

  /**
   * PNG 格式 (*.png)
   */
  PNG("png", 18),

  /**
   * JPG 格式 (*.jpg)
   */
  JPG("png", 17),

  /**
   * GIF 格式 (*.gif)
   */
  GIF("gif", 7),

  /**
   * slideshow 格式 (*.ppsx)
   */
  SHOW("ppsx", 16),

  /**
   * GIF 格式 (*.tif)
   */
  TIF("tif", 21),

  /**
   * RTF 格式 (*.rtf)
   */
  RTF("rtf", 6),

  /**
   * HTML 格式 (*.html)
   */
  HTML("html", 12)
}