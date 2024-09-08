package wise.co.kr.excel_processor.controller

import jakarta.servlet.http.HttpServletResponse
import org.springframework.core.io.Resource
import org.springframework.http.HttpHeaders
import org.springframework.http.MediaType
import org.springframework.http.ResponseEntity
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RequestParam
import org.springframework.web.bind.annotation.RestController
import org.springframework.web.multipart.MultipartFile
import wise.co.kr.excel_processor.service.ExcelProcessService
import java.io.File
import java.net.URLDecoder
import java.nio.file.Files

@RestController
@RequestMapping("/excel")
class ExcelController(
    private val excelProcessService: ExcelProcessService
) {

    @GetMapping("/upload")
    fun uploadExcel(@RequestParam("path") path: String): ResponseEntity<ByteArray> {
        return try {
            val decodedURL = URLDecoder.decode(path, "UTF-8")
            val files = scanFile(decodedURL)
            val excelBytes = excelProcessService.processExcel(files)
            ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=resultExcel.xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(excelBytes)
        } catch (e: Exception) {
            ResponseEntity.badRequest().body("Error processing excel: ${e.message}".toByteArray())
        }
    }

    private fun scanFile (path: String) : List<File> {
        val directory = File(path)
        if (!directory.isDirectory) {
            throw IllegalArgumentException("Provided path is not a directory")
        }

        return directory.listFiles()?.toList() ?: emptyList()

    }


}