package wise.co.kr.excel_processor.service

import org.springframework.core.io.Resource
import org.springframework.web.multipart.MultipartFile


interface ExcelProcessService {

    fun processExcel(files: List<MultipartFile>): Int
    fun mergeExcel()

}