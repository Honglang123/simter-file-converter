package tech.simter.file.converter.rest.webflux

import com.ninjasquad.springmockk.MockkBean
import org.springframework.context.annotation.ComponentScan
import org.springframework.context.annotation.Configuration
import org.springframework.web.reactive.config.EnableWebFlux
import tech.simter.file.converter.core.FileConverterService

/**
 * All unit test config for this module.
 *
 * @author RJ
 */
@Configuration
@EnableWebFlux
@ComponentScan("tech.simter")
@MockkBean(FileConverterService::class)
class UnitTestConfiguration