package tech.simter.file.converter.core

import reactor.core.publisher.Mono
import java.util.*

/**
 * 文档转换 Service。
 */
interface FileConverterService {
  /**
   * 将文档转换为指定格式。
   *
   * @param fromFile 来源文件的相对路径，父路径由服务端的配置决定
   * @param toFile   存放转换后文件的相对路径，父路径由服务端的配置决定
   * @param password 打开来源文件所需的密码
   */
  fun convert(fromFile: String, toFile: String, password: Optional<String> = Optional.empty()): Mono<Void>
}