package wise.co.kr.excel_processor.service

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.springframework.stereotype.Service
import java.time.LocalDateTime
import java.time.ZoneId
import java.time.format.DateTimeFormatter

@Service
class GenerateExcelByVendorServiceImpl : GenerateExcelByVendorService {


    override fun generateExcelByVendor(sourceWorkbook: Workbook, targetWorkbook: Workbook, excelName: String) :Workbook{
        val vendorName = getVendorName(sourceWorkbook)

        return when (vendorName) {
            "WDQ" -> {
                generateWDQExcel(sourceWorkbook,targetWorkbook, excelName)
            }
//            "SDQ" -> {
//                generateWDQExcel(sourceWorkbook,targetWorkbook, excelName)
//            }
//            "DQMINER" -> {
//                generateWDQExcel(sourceWorkbook,targetWorkbook, excelName)
//            }
//            "DQUBE" -> {
//                generateWDQExcel(sourceWorkbook,targetWorkbook, excelName)
//            }
            else -> {
                throw IllegalArgumentException("vendor not found")
            }
        }


    }

    //진단 도구 판별 함수
    private fun getVendorName(sourceWorkbook: Workbook): String {

        val sheet = sourceWorkbook.getSheetAt(0)
        val cell = sheet.getRow(0).getCell(1)

        return when {
            cell.stringCellValue.contains("WISE") -> "WDQ"
            cell.stringCellValue.contains("SDQ") -> "SDQ"
            cell.stringCellValue.contains("DQUBE") -> "DQUBE"
            cell.stringCellValue.equals("값진단 결과 보고서") -> "DQMINER"

            else -> throw IllegalArgumentException("vendor not found")
        }
    }

    //TODO:hashmap 만들고 생성된 hashmap 으로 엑셀 생성하는 로직 만들기
    private fun generateWDQExcel(
        sourceWorkbook: Workbook,targetWorkbook: Workbook, excelName: String
    ):Workbook{


        //값진단결과
        val sourceSheet0 = sourceWorkbook.getSheetAt(0)
        //진단대상테이블
        val sourceSheet1 = sourceWorkbook.getSheetAt(1)
        //도메인
        val sourceSheet2 = sourceWorkbook.getSheetAt(2)
        //업무규칙
        val sourceSheet4 = sourceWorkbook.getSheetAt(4)


        val dataHashMap0 = generateHashMap0(sourceSheet0)
        //targetSheet0
        dataHashMap0["파일명"] = excelName
        dataHashMap0["진단도구명"] = sourceSheet0.getRow(0).getCell(1).stringCellValue
        dataHashMap0["업무규칙 수"] = (sourceSheet4.physicalNumberOfRows - 1).toString()
        dataHashMap0["총 진단건수"] = sourceSheet0.getRow(28).getCell(3).stringCellValue
        dataHashMap0["총 오류건수"] = sourceSheet0.getRow(28).getCell(4).stringCellValue
        dataHashMap0["총 오류율"] = sourceSheet0.getRow(28).getCell(5).stringCellValue
        dataHashMap0["작업시간"] = getCurrentKoreanTime()


        //매번 행 개수 체크하는게 성능상 부담일 수 도 있으니 문제가 생기면 여기부터 개선
        val targetSheet0 = targetWorkbook.getSheetAt(0)
        val headerRow = targetSheet0.getRow(0)
        val newRow = targetSheet0.createRow(targetSheet0.lastRowNum+1)
        val colNum = headerRow.physicalNumberOfCells

        for(i in 0 until colNum) {
            newRow.createCell(i).setCellValue(
                dataHashMap0.getValue(targetSheet0.getRow(0).getCell(i).toString()).toString()
            )
        }

        //sheet 1
        dataHashMap0["품질지표명"] = sourceSheet0.getRow(28).getCell(5).stringCellValue

        //sheet 2
        //상태가 "대상"인 row 의 테이블명, 상태, 범위조건, 의견 가져오기
        //별도의 리스트는 필요하지 않을 듯

        val targetSheet2 = targetWorkbook.getSheetAt(2)
        var targetRowNum2 = targetSheet2.physicalNumberOfRows

        for (i in 1 until sourceSheet2.physicalNumberOfRows) {
            val sourceRow = sourceSheet2.getRow(i)
            if (sourceRow?.getCell(3)?.stringCellValue == "대상") {
                val targetRow = targetSheet2.createRow(targetRowNum2++)

                for (j in 0 until sourceRow.physicalNumberOfCells) {
                    targetRow.createCell(j).setCellValue(
                        sourceRow.getCell(j).toString()
                    )
                }
            }
        }
        //sheet 3
        //검증룰 명이 존재하는 row 의 테이블명, 컬럼명, 데이터타입, 검증룰명,품질지표명, 검증룰, 오류제외데이터, 의견 가져오기
        //source sheet 에서 row 를 기준으로 반복문을 돌면서 일치하는 단어가 있는 헤더 기준 셀을 복사해서 value 로 집어넣기



        return targetWorkbook

    }

//    private fun generateSDQExcel(
//        sourceWorkbook: Workbook,targetWorkbook: Workbook, excelName:
//    ):Workbook{
//
//    }
//
//    private fun generateDQUBExcel(
//        sourceWorkbook: Workbook,targetWorkbook: Workbook, excelName:
//    ):Workbook{
//
//    }
//
//    private fun generateDQMINERExcel(
//        sourceWorkbook: Workbook,targetWorkbook: Workbook, excelName: String
//    ):Workbook{
//
//    }

    private fun getCurrentKoreanTime(): String {
        val now = LocalDateTime.now(ZoneId.of("Asia/Seoul"))
        val formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")
        return now.format(formatter)
    }


    private fun generateHashMap0(sourceSheet: Sheet): HashMap<String, String> {
        val resultMap = HashMap<String, String>()


        for (row: Row in sourceSheet) {
            for (cell: Cell in row) {
                val cellValue = getCellValueAsString(cell)

                // 특정 단어와 일치하는지 확인
                if (isKeyword(cellValue)) {
                    val nextCell = row.getCell(cell.columnIndex + 1)
                    if (nextCell != null) {
                        val nextCellValue = getCellValueAsString(nextCell)
                        resultMap[cellValue] = nextCellValue
                    }
                }
            }
        }

        return resultMap
    }

    private fun getCellValueAsString(cell: Cell): String {
        return when (cell.cellType) {
            CellType.STRING -> cell.stringCellValue
            CellType.NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    cell.localDateTimeCellValue.toString()
                } else {
                    cell.numericCellValue.toString()
                }
            }
            CellType.BOOLEAN -> cell.booleanCellValue.toString()
            CellType.FORMULA -> cell.cellFormula
            else -> ""
        }
    }

    private fun isKeyword(value: String): Boolean {
        val keywords = listOf(
            "파일명", "기관명", "정보시스템명", "DBMS명", "DBMS서비스명",
            "DBMS종류", "DBMS버전", "업무규칙 수", "총 진단건수", "총 오류건수",
            "총 오류율", "출력일"
        )
        return keywords.contains(value)
    }




}