package wise.co.kr.excel_processor.service

import java.time.LocalDateTime
import java.time.ZoneId
import java.time.format.DateTimeFormatter


interface ProcessExcelService {
    fun getCurrentKoreanTime(): String {
        val now = LocalDateTime.now(ZoneId.of("Asia/Seoul"))
        val formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")
        return now.format(formatter)
    }
}