package wise.co.kr.excel_processor.controller

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

@RestController
@RequestMapping("/excel")
class ExcelController(
    private val excelProcessService: ExcelProcessService
) {

    @PostMapping("/upload")
    fun uploadExcel(@RequestParam("files") files: List<MultipartFile>): ResponseEntity<ByteArray> {
        return try {
            val excelBytes = excelProcessService.processExcel(files)
            ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=combined_excel.xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(excelBytes)
        } catch (e: Exception) {
            ResponseEntity.badRequest().body("Error processing excel: ${e.message}".toByteArray())
        }
    }

    /*
    @GetMapping("/download")
    fun downloadExcel(): ResponseEntity<Resource> {
        val resource = excelProcessService.generateExcel()
        return ResponseEntity.ok()
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=processed_data.xlsx")
            .body(resource)
    }
    */


}