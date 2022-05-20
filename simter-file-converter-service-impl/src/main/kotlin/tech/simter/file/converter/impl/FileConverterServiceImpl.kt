package tech.simter.file.converter.impl

import com.jacob.activeX.ActiveXComponent
import com.jacob.com.ComThread
import com.jacob.com.Dispatch
import com.sun.org.slf4j.internal.LoggerFactory
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.beans.factory.annotation.Value
import org.springframework.stereotype.Service
import org.springframework.util.StringUtils
import reactor.core.publisher.Mono
import tech.simter.file.converter.core.*
import java.nio.file.Paths
import java.rmi.RemoteException
import java.util.*

/**
 * 文档转换 Service 实现。
 */
@Service
class FileConverterServiceImpl @Autowired constructor(
  @Value("\${simter-file-converter.from-root-dir:/gfdata/file}")
  private val fromRootDir: String, // 来源文件的根路径
  @Value("\${simter-file-converter.to-root-dir:/gfdata/file_converted}")
  private val toRootDir: String    // 转换后文件保存的根路径
) : FileConverterService {
  private val logger = LoggerFactory.getLogger(FileConverterServiceImpl::class.java)

  override fun convert(fromFile: String, toFile: String, password: Optional<String>): Mono<Conversion> {
    logger.warn("----Convert Starter----")

    // 获取原始文件的绝对路径
    val fromPath = Paths.get(fromRootDir, fromFile)
    if (!fromPath.toFile().exists()) {
      logger.warn("----Convert End：$fromPath 来源文件未找到！----")
      return Mono.just(Conversion(false, "$fromPath 来源文件未找到！"))
    }

    // 获取转换后文件的绝对路径
    val toPath = Paths.get(toRootDir, toFile)
    if (!toPath.toFile().parentFile.exists()) toPath.toFile().parentFile.mkdirs()
    logger.warn("Convert Info：fromFile=${fromPath}，toFile=${toPath}")

    val conversion = try {
      convertFormat(fromPath.toString(), toPath.toString(), password)
      logger.warn("----Convert End----")
      Conversion(true, "转换成功")
    } catch (e: RemoteException) {
      logger.warn("Convert Error：${e.message}")
      Conversion(false, e.message)
    }

    return Mono.just(conversion)
  }

  /**
   * 使用 jacob 调用 Office 服务执行 Office 文档格式转换
   *
   * @param fromPath 来源文件的绝对路径
   * @param toPath 存储转换后文件的绝对路径
   * @param password 打开来源文件所需的密码
   */
  private fun convertFormat(fromPath: String, toPath: String, password: Optional<String>) {
    // 获取文件格式
    val fromFormat = StringUtils.getFilenameExtension(fromPath)!!.toLowerCase()
    val toFormat = StringUtils.getFilenameExtension(toPath)!!.toLowerCase()

    var app: ActiveXComponent? = null
    try {
      when {
        isWordFormat(fromFormat) -> {
          val format = WordSaveFormat.values()
            .firstOrNull { it.format == toFormat } ?: throw RemoteException("不支持将 .$fromFormat 文件转换为 .$toFormat 文件")

          // 初始化 com 线程
          ComThread.InitSTA()
          // 打开 Word 应用程序
          app = ActiveXComponent("Word.Application")
          convertByWord(app, fromPath, toPath, format.value, password.orElse("bc"))
        }
        isExcelFormat(fromFormat) -> {
          val format = ExcelSaveFormat.values()
            .firstOrNull { it.format == toFormat } ?: throw RemoteException("不支持将 .$fromFormat 文件转换为 .$toFormat 文件")

          // 初始化 com 线程
          ComThread.InitSTA()
          // 打开 Excel 应用程序
          app = ActiveXComponent("Excel.Application")
          convertByExcel(app, fromPath, toPath, format.value, password.orElse("bc"))
        }
        isPowerPointFormat(fromFormat) -> {
          val format = PowerPointSaveFormat.values()
            .firstOrNull { it.format == toFormat } ?: throw RemoteException("不支持将 .$fromFormat 文件转换为 .$toFormat 文件")

          // 初始化 com 线程
          ComThread.InitSTA()
          // 打开 PowerPoint 应用程序
          app = ActiveXComponent("PowerPoint.Application")
          convertByPowerPoint(app, fromPath, toPath, format.value, password.orElse("bc"))
        }
        else -> {
          throw RemoteException("不支持 $fromFormat 格式的文档转换")
        }
      }
    } catch (e: Exception) {
      if (e.message!!.contains("密码不正确") || e.message!!.contains("重新输入打开权限密码")) {
        if (password.isPresent) {
          throw RemoteException("输入的密码错误！")
        } else {
          throw RemoteException("需输入文档密码！")
        }
      } else throw RemoteException(e.message)
    } finally {
      // 关闭应用程序
      if (app != null) {
        try {
          app.invoke("Quit")
          // 关闭 com 线程
          ComThread.Release()
        } catch (e: Exception) {
          throw RemoteException(e.message)
        }
      }
    }
  }

  /**
   * 使用 jacob 调用 Office 服务执行 Word 文档格式转换
   *
   * @param wordApp Word 应用程序
   * @param fromFile 来源文件的绝对路径
   * @param toFile   转换后的文件要保存到的绝对路径
   * @param toFormat 转换后的文件格式
   * @param password 打开来源文件所需的密码
   */
  private fun convertByWord(
    wordApp: ActiveXComponent,
    fromFile: String,
    toFile: String,
    toFormat: Int,
    password: String
  ) {
    // 设置 Word 不可见
    wordApp.setProperty("Visible", false)

    // 获 Documents 对象
    val documents = wordApp.getProperty("Documents").toDispatch()

    // 调用 Documents.Open 方法打开文档，并返回打开的文档对象 Document
    val document = Dispatch.call(documents, "Open", fromFile
      , false     // ConfirmConversions：false-不显示转换文件对话框
      , true      // ReadOnly：true-只读
      , false     // AddToRecentFiles：将该文档添加到最近使用的文件列表底部
      , password  // PasswordDocument: 打开文档所需密码
      , password  // PasswordTemplate: 打开模板所需密码
      , true      // Revert: 如果已打开同名文件，强制关闭重新打开
      , password  // WritePasswordDocument: 保存文档要更改的密码
      , password  // WritePasswordTemplate: 保存模板更改的密码
    ).toDispatch()

    // 调用 Document.SaveAs2 方法，将文档保存为指定格式
    Dispatch.call(document, "SaveAs2", toFile, toFormat)
    // 调用 Document.ExportAsFixedFormat 方法，将文档保存为 PDF 或 XPS 格式
    // Dispatch.call(document, "ExportAsFixedFormat", toFile, toFormat)

    // 调用 Document.Close 方法，关闭文档
    Dispatch.call(document, "Close", 0)
  }

  /**
   * 使用 jacob 调用 Office 服务执行 Excel 文档格式转换
   *
   * @param excelApp Excel 应用程序
   * @param fromFile 来源文件的绝对路径
   * @param toFile   转换后的文件要保存到的绝对路径
   * @param toFormat 转换后的文件格式
   * @param password 打开来源文件所需的密码
   */
  private fun convertByExcel(
    excelApp: ActiveXComponent,
    fromFile: String,
    toFile: String,
    toFormat: Int,
    password: String
  ) {
    // 设置 Excel 不可见
    excelApp.setProperty("Visible", false)

    // 获取 Workbooks 对象
    val workbooks = excelApp.getProperty("Workbooks").toDispatch()

    // 调用 Workbooks.Open 方法打开文档，并返回打开的文档对象 Workbook
    val workbook = Dispatch.call(workbooks, "Open", fromFile
      , false     // UpdateLinks
      , true      // ReadOnly：true-只读
      , 5         // Format: Nothing
      , password  // PasswordDocument: 打开文档所需密码
      , password  // PasswordTemplate: 打开模板所需密码
      , true      // Revert: 如果已打开同名文件，强制关闭重新打开
    ).toDispatch()

    if (toFormat == ExcelSaveFormat.PDF.value || toFormat == ExcelSaveFormat.XPS.value) {
      // 调用 Workbook.ExportAsFixedFormat 方法，将文档保存为 PDF 或 XPS 格式
      Dispatch.call(workbook, "ExportAsFixedFormat", toFormat, toFile)
    } else {
      // 调用 Workbook.SaveAs 方法，将文档保存为指定格式
      // 说明：SaveAs 方法的保存文件格式不支持 PDF，其保存格式枚举参考：https://docs.microsoft.com/zh-cn/office/vba/api/excel.xlfileformat
      Dispatch.call(workbook, "SaveAs", toFile, toFormat)
    }

    // 调用 Workbook.Close 方法，关闭文档
    Dispatch.call(workbook, "Close", false)
  }

  /**
   * 使用 jacob 调用 Office 服务执行 PowerPoint 文档格式转换
   *
   * @param pptApp PowerPoint 应用程序
   * @param fromFile 来源文件的绝对路径
   * @param toFile   转换后的文件要保存到的绝对路径
   * @param toFormat 转换后的文件格式
   * @param password 打开来源文件所需的密码
   */
  private fun convertByPowerPoint(
    pptApp: ActiveXComponent,
    fromFile: String,
    toFile: String,
    toFormat: Int,
    password: String
  ) {
    // 设置 PowerPoint 不可见：不支持
    // pptApp.setProperty("Visible", false);

    // 获取 Presentations 对象
    // val presentations = pptApp.getProperty("Presentations").toDispatch()
    // 调用 Presentations.Open 方法，打开文档，如果文档有密码，会弹出密码输入框，如果输入错误密码，才会报错
    /*val presentation = Dispatch.call(presentations, "Open", fromFile,
      true,  // ReadOnly
      true,  // Untitled指定文件是否有标题
      false  // WithWindow指定文件是否可见
    ).toDispatch()*/

    // 无论有无密码均可采用该方法打开文件并获取 Presentation 对象。
    val presentation = doSpecialPowerPointConvert(pptApp, fromFile, password)

    // 调用 Presentation.SaveAs 方法，将文档保存为指定格式
    Dispatch.call(presentation, "SaveAs", toFile, toFormat)

    // 调用 Presentation.Close 方法，关闭文档
    Dispatch.call(presentation, "Close")
  }

  private fun doSpecialPowerPointConvert(pptApp: ActiveXComponent, file: String, password: String): Dispatch {
    // 获取 Application 对象的 ProtectedViewWindows 属性
    val protectedViewWindows = pptApp.getProperty("ProtectedViewWindows").toDispatch()

    // 调用 ProtectedViewWindows 对象中 Open 方法打开密码文档，并返回 ProtectedViewWindow
    val protectedViewWindow = Dispatch.call(protectedViewWindows, "Open"  // 调用Documents对象的Open方法
      , file          // 文件全路径名
      , password      // ReadPassword: 密码
      , false         // OpenAndRepair
    ).toDispatch()

    // 启用编辑功能并返回 Presentation 对象
    return Dispatch.call(protectedViewWindow, "Edit").toDispatch()
  }

  // 判断是否是 Word 可以打开的文档格式
  private fun isWordFormat(fromFormat: String): Boolean {
    return "doc" == fromFormat || "docx" == fromFormat || "docm" == fromFormat
      || "txt" == fromFormat || "rtf" == fromFormat
  }

  // 判断是否是 Excel 可以打开的文档格式
  private fun isExcelFormat(fromFormat: String): Boolean {
    return "xls" == fromFormat || "xlsx" == fromFormat || "xlsm" == fromFormat || "csv" == fromFormat
  }

  // 判断是否是 PowerPoint 可以打开的文档格式
  private fun isPowerPointFormat(fromFormat: String): Boolean {
    return "ppt" == fromFormat || "pptx" == fromFormat || "pptm" == fromFormat
  }
}