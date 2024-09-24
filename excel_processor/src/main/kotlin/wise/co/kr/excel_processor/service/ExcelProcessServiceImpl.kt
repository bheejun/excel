package wise.co.kr.excel_processor.service

import jakarta.transaction.Transactional
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service
import org.springframework.web.multipart.MultipartFile
import wise.co.kr.excel_processor.dto.ResponseDto
import java.io.ByteArrayOutputStream
import java.io.File
import kotlin.math.roundToInt


@Service
class ExcelProcessServiceImpl(
    private val generateExcelByVendorService: GenerateExcelByVendorService

) : ExcelProcessService {


    @Transactional
    override fun processExcel(files: List<File>): ResponseDto {
        println(files.size.toString())
        val targetWorkbook = XSSFWorkbook()
        //workbook sheet 생성
        //0, 1, 2, 3 으로 생성후 마지막에 rename 하는게 뭔가 더 좋을거같긴함
        targetWorkbook.createSheet("값진단_결과보고서")
        targetWorkbook.createSheet("값진단_결과보고서_상세")
        targetWorkbook.createSheet("진단대상테이블_목록")
        targetWorkbook.createSheet("진단대상컬럼_목록")

        val headers = listOf(
            listOf(
                "파일명",
                "기관명",
                "시스템명",
                "DB명",
                "DB서비스명",
                "DB종류",
                "버전",
                "업무규칙 수",
                "총 진단건수",
                "총 오류건수",
                "총 오류율",
                "출력일",
                "진단도구명",
                "작업시간"
            ),
            listOf("파일명", "기관명", "정보시스템명", "DBMS명", "품질지표명", "진단건수", "오류건수", "오류율"),
            listOf("DBMS명", "스키마명", "테이블명", "상태", "수집일자", "범위조건", "의견"),
            listOf("DBMS명", "스키마명", "테이블명", "컬럼명", "데이터타입", "검증룰명", "품질지표명", "검증룰", "오류제외데이터", "의견")
        )

        // 헤더 설정
        headers.forEachIndexed { sheetIndex, headerList ->
            val sheet = targetWorkbook.getSheetAt(sheetIndex)
            val headerRow = sheet.createRow(0)
            headerList.forEachIndexed { cellIndex, header ->
                headerRow.createCell(cellIndex).setCellValue(header)
            }
        }

        val notExcelFileList: MutableList<String> = mutableListOf()
        var count = 0

        files.forEach { file ->
            val excelName = file.name ?: throw IllegalArgumentException("The file name is Null.")
            try {
                val sourceWorkbook = WorkbookFactory.create(file.inputStream())
                generateExcelByVendorService.generateExcelByVendor(sourceWorkbook, targetWorkbook, excelName)
                count++
                val progress = ((count.toDouble() / files.size) * 100).roundToInt()

                println("$progress% ||| $excelName")

                sourceWorkbook.close()
            } catch (e: Exception) {
                notExcelFileList.add(excelName)
            }
        }


        val byteArrayOutputStream = ByteArrayOutputStream()
        targetWorkbook.write(byteArrayOutputStream)
        targetWorkbook.close()

        return ResponseDto(byteArrayOutputStream.toByteArray(),notExcelFileList)
    }
}