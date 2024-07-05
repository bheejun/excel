package wise.co.kr.excel_processor.service

import jakarta.transaction.Transactional
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.springframework.stereotype.Service
import org.springframework.web.multipart.MultipartFile
import java.io.ByteArrayOutputStream


@Service
class ExcelProcessServiceImpl(
    private val generateExcelByVendorService: GenerateExcelByVendorService

):ExcelProcessService {


    @Transactional
    override fun processExcel(files: List<MultipartFile>): ByteArray {
        val targetWorkbook = SXSSFWorkbook()
        //workbook sheet 생성
        //0, 1, 2, 3 으로 생성후 마지막에 rename 하는게 뭔가 더 좋을거같긴함
        targetWorkbook.createSheet("값진단_결과보고서")
        targetWorkbook.createSheet("값진단_결과보고서_상세")
        targetWorkbook.createSheet("진단대상테이블_목록")
        targetWorkbook.createSheet("진단대상컬럼_목록")

        //workbook 초기세팅
        //workbook 시트별 헤더 세팅
        val sheet0Headers = listOf(
            "파일명", "기관명", "정보시스템명", "DBMS명", "DBMS서비스명",
            "DBMS종류", "DBMS버전", "업무규칙 수", "총 진단건수", "총 오류건수",
            "총 오류율", "출력일", "진단도구명", "작업시간"
        )

        val sheet1Headers = listOf(
            "파일명", "기관명", "정보시스템명", "DBMS명", "품질지표명",
            "진단건수", "오류건수", "오류율"
        )

        val sheet2Headers = listOf(
            "DBMS명", "스키마명", "테이블명", "상태", "수집일자",
            "범위조건", "의견"
        )

        val sheet3Headers = listOf(
            "DBMS명", "스키마명", "테이블명", "컬럼명", "데이터타입",
            "검증룰명", "품질지표명(진단기준명)", "검증룰(진단기준)", "오류제외데이터", "의견"
        )

        //시트별 row 생성후 헤더 세팅
        sheet0Headers.forEachIndexed { index, header ->
            targetWorkbook.getSheetAt(0).createRow(0).createCell(index).setCellValue(header)

        }
        sheet1Headers.forEachIndexed { index, header ->
            targetWorkbook.getSheetAt(1).createRow(0).createCell(index).setCellValue(header)

        }
        sheet2Headers.forEachIndexed { index, header ->
            targetWorkbook.getSheetAt(2).createRow(0).createCell(index).setCellValue(header)

        }
        sheet3Headers.forEachIndexed { index, header ->
            targetWorkbook.getSheetAt(3).createRow(0).createCell(index).setCellValue(header)

        }

        //초기 세팅 완료 된 target workbook 에 파일별로 데이터 적재
        files.forEach { file ->

            val sourceWorkbook = WorkbookFactory.create(file.inputStream)
            val excelName = file.originalFilename

            if (excelName != null) {
                //source workbook 에서 데이터를 수집한 후, target workbook 에 재정렬 한 target workbook 이 출력으로 나옴
                generateExcelByVendorService.generateExcelByVendor(sourceWorkbook, targetWorkbook, excelName)
            }else{
                throw IllegalArgumentException("file name is null")
            }

            sourceWorkbook.close()


        }

        val byteArrayOutputStream = ByteArrayOutputStream()
        targetWorkbook.write(byteArrayOutputStream)
        targetWorkbook.dispose()

        return byteArrayOutputStream.toByteArray()
    }
}