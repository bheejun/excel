package wise.co.kr.excel_processor.controller

import org.springframework.core.io.ByteArrayResource
import org.springframework.http.HttpHeaders
import org.springframework.http.MediaType
import org.springframework.http.ResponseEntity
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RequestParam
import org.springframework.web.bind.annotation.RestController
import wise.co.kr.excel_processor.service.ExcelProcessService
import wise.co.kr.excel_processor.service.UnzipService
import java.io.ByteArrayOutputStream
import java.io.File
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

@RestController
@RequestMapping("/excel")
class ExcelController(
    private val excelProcessService: ExcelProcessService,
    private val unzipService: UnzipService
) {
    @GetMapping("/upload")
    fun uploadExcel(@RequestParam("path") path: String): ResponseEntity<ByteArrayResource> {
        val directory = File(path)


        val files = directory.listFiles() ?: throw IllegalArgumentException("No file")

        val excelFiles = mutableListOf<File>()
        val nonExcelFiles = mutableListOf<String>()

        files.forEach { file ->
            when {
                file.extension.lowercase() in listOf("xlsx", "xls") -> excelFiles.add(file)
                else -> nonExcelFiles.add(file.name)
            }
        }

        val processedExcel = excelProcessService.processExcel(excelFiles)


        // 비엑셀 파일 목록을 텍스트 파일로 저장
        saveListToFile(nonExcelFiles, File(directory, "non_excel_files.txt"))


        val zipBytes = createZipFile(processedExcel.processedExcel, processedExcel.errorList, nonExcelFiles)
        val resource = ByteArrayResource(zipBytes)

        return ResponseEntity.ok()
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=result.zip")
            .contentType(MediaType.APPLICATION_OCTET_STREAM)
            .contentLength(zipBytes.size.toLong())
            .body(resource)
    }

    @GetMapping("/unzip")
    fun unzipAndRename(@RequestParam("path") path: String): ResponseEntity<ByteArrayResource>{

        val directory = File(path)

        val resultZipByteArray = unzipService.unzipAndRenameService(directory)

        val resource = ByteArrayResource(resultZipByteArray)


        return ResponseEntity.ok()
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=result.zip")
            .contentType(MediaType.APPLICATION_OCTET_STREAM)
            .contentLength(resultZipByteArray.size.toLong())
            .body(resource)

    }

    @GetMapping("/scan")
    fun scanDirectoryAndGetReport(@RequestParam("path") path: String): ResponseEntity<ByteArrayResource>{
        val resultZip= unzipService.scanAndFindResultReport(path)
        val resource = ByteArrayResource(resultZip)

        return ResponseEntity.ok()
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=result.zip")
            .contentType(MediaType.APPLICATION_OCTET_STREAM)
            .contentLength(resultZip.size.toLong())
            .body(resource)

    }
}

private fun saveListToFile(list: List<String>, file: File) {
    file.bufferedWriter().use { writer ->
        list.forEach { writer.write(it + "\n") }
    }
}

private fun createZipFile(
    processedExcel: ByteArray,
    errorFiles: List<String>,
    nonExcelFiles: List<String>
): ByteArray {
    val baos = ByteArrayOutputStream()
    ZipOutputStream(baos).use { zos ->
        // 처리된 엑셀 파일 추가
        zos.putNextEntry(ZipEntry("processed_excel.xlsx"))
        zos.write(processedExcel)
        zos.closeEntry()

        // 에러 파일 목록 추가
        zos.putNextEntry(ZipEntry("error_files.txt"))
        zos.write(errorFiles.joinToString("\n").toByteArray())
        zos.closeEntry()

        // 비엑셀 파일 목록 추가
        zos.putNextEntry(ZipEntry("non_excel_files.txt"))
        zos.write(nonExcelFiles.joinToString("\n").toByteArray())
        zos.closeEntry()
    }
    return baos.toByteArray()
}



