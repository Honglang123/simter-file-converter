package tech.simter.file.converter.rest.webflux.handler

import org.springframework.beans.factory.annotation.Autowired
import org.springframework.http.HttpStatus
import org.springframework.stereotype.Component
import org.springframework.web.reactive.function.server.*
import reactor.core.publisher.Mono
import tech.simter.file.converter.core.FileConverterService
import tech.simter.reactive.web.Utils.responseGoneStatus
import tech.simter.reactive.web.Utils.responseSpecificStatus
import java.io.FileNotFoundException
import java.rmi.RemoteException

@Component
class FileConverterHandler @Autowired constructor(
  private val service: FileConverterService
) : HandlerFunction<ServerResponse> {
  override fun handle(request: ServerRequest): Mono<ServerResponse> {
    val fromFile = request.queryParam("from-file").get()
    val toFile = request.queryParam("to-file").get()
    val password = request.queryParam("password")

    return service.convert(fromFile, toFile, password)
      // 转换成功
      .then(ServerResponse.noContent().build())
      // 来源文件不存在
      .onErrorResume(FileNotFoundException::class.java, ::responseGoneStatus)
      // 内部转换出错
      .onErrorResume(RemoteException::class.java) {
        responseSpecificStatus(HttpStatus.INTERNAL_SERVER_ERROR, it)
      }
  }

  companion object {
    val REQUEST_PREDICATE: RequestPredicate = RequestPredicates.POST("").or(RequestPredicates.POST("/"))
  }
}