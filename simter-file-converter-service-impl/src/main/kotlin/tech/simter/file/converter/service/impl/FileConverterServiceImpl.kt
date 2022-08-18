package tech.simter.file.converter.service.impl

import org.reactivestreams.Publisher
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.beans.factory.annotation.Value
import org.springframework.core.io.buffer.DataBuffer
import org.springframework.core.io.buffer.DataBufferUtils
import org.springframework.stereotype.Service
import reactor.core.publisher.Mono
import tech.simter.file.converter.core.FileConverterDao
import tech.simter.file.converter.core.FileConverterService
import java.nio.channels.AsynchronousFileChannel
import java.nio.file.Path
import java.nio.file.Paths
import java.nio.file.StandardOpenOption
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter
import java.util.*

@Service
class FileConverterServiceImpl @Autowired constructor(
  @Value("\${simter-file-converter.base-dir}")
  private val baseDir: String,
  private val dao: FileConverterDao
) : FileConverterService {
  override fun convert2Pdf(source: Publisher<DataBuffer>, type: String, password: Optional<String>): Mono<Path> {
    val name = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddhhmmss"))
    val path = Paths.get(baseDir, "$name.$type")
    val channel = AsynchronousFileChannel.open(path, StandardOpenOption.CREATE_NEW, StandardOpenOption.WRITE)

    return DataBufferUtils.write(source, channel).map {
      DataBufferUtils.release(it)

      it
    }.doOnTerminate{ channel.close() }.then(Mono.defer {
      dao.convert2Pdf(path, type, password)
    })
  }
}