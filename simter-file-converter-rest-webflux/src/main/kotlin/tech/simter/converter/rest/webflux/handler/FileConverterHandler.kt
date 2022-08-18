package tech.simter.converter.rest.webflux.handler

import org.springframework.beans.factory.annotation.Autowired
import org.springframework.core.io.FileSystemResource
import org.springframework.http.HttpStatus.INTERNAL_SERVER_ERROR
import org.springframework.http.HttpStatus.PRECONDITION_FAILED
import org.springframework.http.MediaType.APPLICATION_PDF
import org.springframework.stereotype.Component
import org.springframework.web.reactive.function.BodyExtractors
import org.springframework.web.reactive.function.BodyInserters
import org.springframework.web.reactive.function.server.*
import org.springframework.web.reactive.function.server.ServerResponse.badRequest
import org.springframework.web.reactive.function.server.ServerResponse.ok
import reactor.core.publisher.Mono
import tech.simter.file.converter.core.FileConverterService
import tech.simter.file.converter.exception.PasswordException
import tech.simter.reactive.web.Utils.buildContentDisposition
import tech.simter.reactive.web.Utils.responseSpecificStatus
import java.lang.Exception
import java.rmi.RemoteException
import java.util.*

@Component
class FileConverterHandler @Autowired constructor(
  private val service: FileConverterService
) : HandlerFunction<ServerResponse> {
  override fun handle(request: ServerRequest): Mono<ServerResponse> {
    val filename = request.queryParam("filename").orElse("unknown")
    val headers = request.headers()
    val contentType = headers.contentType()
    if (!contentType.isPresent) return badRequest().bodyValue("请设置请求头 content-type 值！")
    val source = request.body(BodyExtractors.toDataBuffers())
    val type = when (contentType.get().toString()) {
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document" -> {
        "docx"
      }
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" -> {
        "xlsx"
      }
      "application/vnd.openxmlformats-officedocument.presentationml.presentation" -> {
        "pptx"
      }
      else -> {
        throw Exception("不支持的文件类型！")
      }
    }
    val password = Optional.ofNullable(headers.firstHeader("password"))

    return service.convert2Pdf(source, type, password)
      .flatMap {
        ok().contentType(APPLICATION_PDF)
          .header("Content-Disposition", buildContentDisposition("attachment", "$filename.pdf"))
          .body(BodyInserters.fromResource(FileSystemResource(it)))
      }
      .onErrorResume(PasswordException::class.java) {
        responseSpecificStatus(PRECONDITION_FAILED, it)
      }
      .onErrorResume(RemoteException::class.java) {
        responseSpecificStatus(INTERNAL_SERVER_ERROR, it)
      }
  }

  companion object {
    val REQUEST_PREDICATE: RequestPredicate = RequestPredicates.POST("/convert2pdf")
  }
}