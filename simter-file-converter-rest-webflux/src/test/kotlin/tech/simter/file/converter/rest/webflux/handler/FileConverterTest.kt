package tech.simter.file.converter.rest.webflux.handler

import io.mockk.every
import org.junit.jupiter.api.Test
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.beans.factory.annotation.Value
import org.springframework.boot.test.autoconfigure.web.reactive.WebFluxTest
import org.springframework.test.context.junit.jupiter.SpringJUnitConfig
import org.springframework.test.web.reactive.server.WebTestClient
import org.springframework.test.web.reactive.server.expectBody
import reactor.core.publisher.Mono
import tech.simter.file.converter.core.FileConverterService
import tech.simter.file.converter.rest.webflux.UnitTestConfiguration
import java.rmi.RemoteException

@SpringJUnitConfig(UnitTestConfiguration::class)
@WebFluxTest
class FileConverterTest @Autowired constructor(
  @Value("\${simter-file-converter.rest-context-path}")
  private val contextPath: String,
  private val client: WebTestClient,
  private val service: FileConverterService
) {
  private val fromFile = "test.docx"
  private val toFile = "test.pdf"

  @Test
  fun `convert successful`() {
    // mock
    every { service.convert(fromFile, toFile) } returns Mono.empty()

    // verify
    client.post().uri("$contextPath?from-file={fromFile}&to-file={toFile}", fromFile, toFile)
      .exchange()
      .expectStatus().isNoContent
  }

  @Test
  fun `convert failed`() {
    // mock
    val msg = "需输入文档密码！"
    every { service.convert(fromFile, toFile) } returns Mono.error(RemoteException(msg))

    // verify
    client.post().uri("$contextPath?from-file={fromFile}&to-file={toFile}", fromFile, toFile)
      .exchange()
      .expectStatus().is5xxServerError
      .expectBody<String>().isEqualTo(msg)
  }
}