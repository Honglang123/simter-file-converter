package tech.simter.file.converter.rest.webflux.handler

import org.springframework.beans.factory.annotation.Autowired
import org.springframework.http.MediaType.APPLICATION_JSON
import org.springframework.stereotype.Component
import org.springframework.web.reactive.function.server.*
import org.springframework.web.reactive.function.server.ServerResponse.ok
import org.springframework.web.reactive.function.server.ServerResponse.temporaryRedirect
import reactor.core.publisher.Mono
import tech.simter.file.converter.core.FileConverterService
import java.net.URI

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
      .flatMap { ok().contentType(APPLICATION_JSON).bodyValue(it) }
  }

  companion object {
    val REQUEST_PREDICATE: RequestPredicate = RequestPredicates.POST("").or(RequestPredicates.POST("/"))
  }
}