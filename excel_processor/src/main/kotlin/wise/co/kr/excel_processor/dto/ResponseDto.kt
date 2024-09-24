package wise.co.kr.excel_processor.dto

import java.io.ByteArrayOutputStream

data class ResponseDto(
    val processedExcel : ByteArray,
    val errorList: MutableList<String>
)
