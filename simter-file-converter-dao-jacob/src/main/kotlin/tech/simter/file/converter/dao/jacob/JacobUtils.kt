package tech.simter.file.converter.dao.jacob

import com.jacob.activeX.ActiveXComponent
import com.jacob.com.ComThread
import com.jacob.com.Dispatch
import tech.simter.file.converter.exception.PasswordException
import java.rmi.RemoteException
import java.util.*

object JacobUtils {
  fun convert2Pdf(source: String, type: String, password: Optional<String>): String {
    // 初始化 com 线程
    ComThread.InitSTA()
    var app: ActiveXComponent? = null
    try {
      val target = source.replace(type, "pdf")

      when {
        isWord(type) -> {
          // 打开 Word 应用程序
          app = ActiveXComponent("Word.Application")
          word2Pdf(app, source, target, password.orElse("bc"))
        }
        isExcel(type) -> {
          // 打开 Excel 应用程序
          app = ActiveXComponent("Excel.Application")
          excel2Pdf(app, source, target, password.orElse("bc"))
        }
        isPowerPoint(type) -> {
          // 打开 PowerPoint 应用程序
          app = ActiveXComponent("PowerPoint.Application")
          ppt2Pdf(app, source, target, password.orElse("bc"))
        }
        else -> {
          throw RemoteException("暂不支持将 .$type 格式的文档转换 PDF 文档！")
        }
      }

      return target
    } catch (e: Exception) {
      if (e.message!!.contains("密码不正确") || e.message!!.contains("重新输入打开权限密码")) {
        if (password.isPresent) {
          throw PasswordException("输入的密码错误！")
        } else {
          throw PasswordException("需输入文档密码！")
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
   * 使用 jacob 将 Word 文档转换为 PDF 文档。
   *
   * @param source 源文件资源路径
   * @param target 转换后的文件要保存到的绝对路径
   * @param password 打开来源文件所需的密码
   */
  private fun word2Pdf(
    wordApp: ActiveXComponent,
    source: String,
    target: String,
    password: String
  ) {
    // 设置 Word 不可见
    wordApp.setProperty("Visible", false)

    // 获 Documents 对象
    val documents = wordApp.getProperty("Documents").toDispatch()

    // 调用 Documents.Open 方法打开文档，并返回打开的文档对象 Document
    val document = Dispatch.call(documents, "Open", source
      , false     // ConfirmConversions：false-不显示转换文件对话框
      , true      // ReadOnly：true-只读
      , false     // AddToRecentFiles：将该文档添加到最近使用的文件列表底部
      , password  // PasswordDocument: 打开文档所需密码
      , password  // PasswordTemplate: 打开模板所需密码
      , true      // Revert: 如果已打开同名文件，强制关闭重新打开
      , password  // WritePasswordDocument: 保存文档要更改的密码
      , password  // WritePasswordTemplate: 保存模板更改的密码
    ).toDispatch()

    // 调用 Document.SaveAs2 方法，将 Word 文档另保存 PDF 文档
    Dispatch.call(document, "SaveAs2", target, 17)
    // 调用 Document.ExportAsFixedFormat 方法，将 Word 文档导出为 PDF 文档
    // Dispatch.call(document, "ExportAsFixedFormat", target, 17)

    // 调用 Document.Close 方法，关闭文档
    Dispatch.call(document, "Close", 0)
  }

  /**
   * 使用 jacob 将 Excel 文档转换为 PDF 文档。
   *
   * @param source 源文件资源路径
   * @param target 转换后的文件资源路径
   * @param password 打开来源文件所需的密码
   */
  private fun excel2Pdf(
    excelApp: ActiveXComponent,
    source: String,
    target: String,
    password: String
  ) {
    // 设置 Excel 不可见
    excelApp.setProperty("Visible", false)

    // 获取 Workbooks 对象
    val workbooks = excelApp.getProperty("Workbooks").toDispatch()

    // 调用 Workbooks.Open 方法打开文档，并返回打开的文档对象 Workbook
    val workbook = Dispatch.call(workbooks, "Open", source
      , false     // UpdateLinks
      , true      // ReadOnly：true-只读
      , 5         // Format: Nothing
      , password  // PasswordDocument: 打开文档所需密码
      , password  // PasswordTemplate: 打开模板所需密码
      , true      // Revert: 如果已打开同名文件，强制关闭重新打开
    ).toDispatch()

    // 调用 Workbook.ExportAsFixedFormat 方法，将 Excel 文档导出为 PDF 文档
    Dispatch.call(workbook, "ExportAsFixedFormat", 0, target)

    // 调用 Workbook.Close 方法，关闭文档
    Dispatch.call(workbook, "Close", false)
  }

  /**
   * 使用 jacob 调用 Office 服务执行 PowerPoint 文档格式转换
   *
   * @param source 源文件资源路径
   * @param target 转换后的文件要保存到资源路径
   * @param password 源文件加密密码
   */
  private fun ppt2Pdf(
    pptApp: ActiveXComponent,
    source: String,
    target: String,
    password: String
  ) {
    // 设置 PowerPoint 不可见：不支持
    // pptApp.setProperty("Visible", false);

    // 获取 Presentations 对象
    // val presentations = pptApp.getProperty("Presentations").toDispatch()
    // 调用 Presentations.Open 方法，打开文档，如果文档有密码，会弹出密码输入框，如果输入错误密码，就直接报错
    /*val presentation = Dispatch.call(presentations, "Open", source,
      true,  // ReadOnly
      true,  // Untitled指定文件是否有标题
      false  // WithWindow指定文件是否可见
    ).toDispatch()*/

    // 无论有无密码均可采用该方法打开文件并获取 Presentation 对象。
    val presentation = doSpecialPowerPointConvert(pptApp, source, password)

    // 调用 Presentation.SaveAs 方法，将 ppt 文档另保存为 PDF 文档
    Dispatch.call(presentation, "SaveAs", target, 32)

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
  private fun isWord(type: String): Boolean {
    return "doc" == type || "docx" == type || "docm" == type || "txt" == type || "rtf" == type
  }

  // 判断是否是 Excel 可以打开的文档格式
  private fun isExcel(type: String): Boolean {
    return "xls" == type || "xlsx" == type || "xlsm" == type || "csv" == type
  }

  // 判断是否是 PowerPoint 可以打开的文档格式
  private fun isPowerPoint(type: String): Boolean {
    return "ppt" == type || "pptx" == type || "pptm" == type
  }
}