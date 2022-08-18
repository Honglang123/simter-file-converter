package tech.simter.file.converter.dao.jacob

import org.springframework.stereotype.Component
import reactor.core.publisher.Mono
import tech.simter.file.converter.core.FileConverterDao
import tech.simter.file.converter.exception.PasswordException
import java.nio.file.Path
import java.nio.file.Paths
import java.rmi.RemoteException
import java.util.*

@Component
class FileConverterDaoImpl : FileConverterDao {
  override fun convert2Pdf(source: Path, type: String, password: Optional<String>): Mono<Path> {
    return try {
      Mono.just(Paths.get(JacobUtils.convert2Pdf(source.toString(), type, password)))
    } catch (e: PasswordException) {
      Mono.error(PasswordException(e.message))
    } catch (e: Exception) {
      Mono.error(RemoteException(e.message))
    }
  }
}