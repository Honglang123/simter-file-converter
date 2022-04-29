package tech.simter.file.converter.rest.webflux

import org.slf4j.LoggerFactory
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.beans.factory.annotation.Value
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean
import org.springframework.context.annotation.Bean
import org.springframework.context.annotation.ComponentScan
import org.springframework.context.annotation.Configuration
import org.springframework.web.reactive.function.server.router
import tech.simter.file.converter.PACKAGE
import tech.simter.file.converter.rest.webflux.handler.FileConverterHandler

/**
 * All configuration for this module.
 *
 * Register a `RouterFunction<ServerResponse>` with all routers for this module.
 * The default context-path of this router is '/file-converter'.
 * And can be config by property `simter-file-converter.rest-context-path`.
 *
 * @author RJ
 */
@Configuration("$PACKAGE.rest.webflux.ModuleConfiguration")
@ComponentScan
class ModuleConfiguration @Autowired constructor(
  @Value("\${simter-file-converter.rest-context-path:/file-converter}")
  private val contextPath: String,
  private val fileConverterHandler: FileConverterHandler
) {
  private val logger = LoggerFactory.getLogger(ModuleConfiguration::class.java)

  init {
    logger.warn("simter-file-converter.rest-context-path='{}'", contextPath)
  }

  /** Register a `RouterFunction<ServerResponse>` with all routers for this module */
  @Bean("$PACKAGE.rest.webflux.Routes")
  @ConditionalOnMissingBean(name = ["$PACKAGE.rest.webflux.Routes"])
  fun operationRoutes() = router {
    contextPath.nest {
      // POST /file-converter?from-file=x&to-file=x&password=x
      FileConverterHandler.REQUEST_PREDICATE.invoke(fileConverterHandler::handle)
    }
  }
}