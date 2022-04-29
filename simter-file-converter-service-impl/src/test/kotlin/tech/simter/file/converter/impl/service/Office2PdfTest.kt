package tech.simter.file.converter.impl.service

import org.junit.jupiter.api.Disabled
import org.junit.jupiter.api.Test
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.boot.test.context.SpringBootTest
import org.springframework.test.context.junit.jupiter.SpringJUnitConfig
import reactor.kotlin.test.test
import tech.simter.file.converter.core.FileConverterService
import tech.simter.util.RandomUtils.randomString
import java.io.FileNotFoundException
import java.rmi.RemoteException
import java.util.*

/**
 * 转换 Office 文档为 PDF 的测试用例。
 *
 * 执行前需将 docs/office 目录下的测试文件放到配置文件声明的 simter-file-converter.from-root-dir 目录下。
 */
@SpringJUnitConfig(UnitTestConfiguration::class)
@SpringBootTest
@Disabled
class Office2PdfTest @Autowired constructor(
  private val service: FileConverterService
) {
  val password = Optional.of("test")
  val errPassword = Optional.of(randomString(4))

  @Test
  fun word2Pdf() {
    val fromFile = "test-word.docx"
    val toFile = "test-word.pdf"

    // 正确输入密码转换成功
    service.convert(fromFile, toFile, password).test().verifyComplete()

    // 来源文件未找到
    service.convert("aaa.docx", toFile, errPassword).test().verifyError(FileNotFoundException::class.java)

    // 输入错误密码转换失败
    service.convert(fromFile, toFile, errPassword).test().verifyError(RemoteException::class.java)
  }

  @Test
  fun excel2Pdf() {
    val fromFile = "test-excel.xlsx"
    val toFile = "test-excel.pdf"

    // 正确输入密码转换成功
    service.convert(fromFile, toFile, password).test().verifyComplete()

    // 输入错误密码转换失败
    service.convert(fromFile, toFile, errPassword).test().verifyError(RemoteException::class.java)
  }

  @Test
  fun ppt2Pdf() {
    val fromFile = "test-ppt.pptx"
    val toFile = "test-ppt.pdf"
    val password = Optional.of("test")

    // 正确输入密码转换成功
    service.convert(fromFile, toFile, password).test().verifyComplete()

    // 输入错误密码转换失败
    service.convert(fromFile, toFile, errPassword).test().verifyError(RemoteException::class.java)
  }
}