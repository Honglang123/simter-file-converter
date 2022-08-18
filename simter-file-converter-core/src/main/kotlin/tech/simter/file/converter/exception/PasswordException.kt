package tech.simter.file.converter.exception

/**
 * 密码异常。
 */
class PasswordException  : RuntimeException {
  constructor(message: String?) : super(message)
  constructor(message: String?, cause: Throwable?) : super(message, cause)
  constructor(cause: Throwable?) : super(cause)
}