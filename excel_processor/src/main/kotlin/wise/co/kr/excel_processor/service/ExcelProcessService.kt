package wise.co.kr.excel_processor.service

import org.springframework.core.io.Resource
import org.springframework.web.multipart.MultipartFile
import java.io.File


interface ExcelProcessService {

    fun processExcel(files: List<File>): ByteArray

}