package tech.simter.file.converter.core

import reactor.core.publisher.Mono
import java.nio.file.Path
import java.util.*

/**
 * 文件转换 Dao 接口。
 *
 * @author nb
 */
interface FileConverterDao {
  /**
   * 将其他类型文件转换为 pdf 文件，并返回。
   *
   * @param source 源文件资源路径
   * @param type 源文件的类型（后缀名）
   * @param password 源文件加密密码
   *
   * @return 转换后的 pdf 文件资源路径
   */
  fun convert2Pdf(source: Path, type: String, password: Optional<String>): Mono<Path>
}