package tech.simter.file.converter.core

import kotlinx.serialization.Serializable

@Serializable
data class Conversion (
  /** 转换结果：true-转换成功、false-转换失败 */
  val result: Boolean,
  /** 说明，转换出错的具体描述 */
  val msg: String?
)