package wise.co.kr.excel_processor

import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration
import org.springframework.boot.runApplication

@SpringBootApplication(exclude= [DataSourceAutoConfiguration::class])
class ExcelProcessorApplication

fun main(args: Array<String>) {
	runApplication<ExcelProcessorApplication>(*args)
}
