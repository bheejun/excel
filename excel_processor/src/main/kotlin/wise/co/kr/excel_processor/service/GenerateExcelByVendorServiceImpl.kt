package wise.co.kr.excel_processor.service

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
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
            "SDQ" -> {
                generateWDQExcel(sourceWorkbook,targetWorkbook, excelName)
            }
            "DQMINER" -> {
                generateWDQExcel(sourceWorkbook,targetWorkbook, excelName)
            }
            "DQUBE" -> {
                generateWDQExcel(sourceWorkbook,targetWorkbook, excelName)
            }
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
        //참조무결성
        val sourceSheet3 = sourceWorkbook.getSheetAt(3)
        //업무규칙
        val sourceSheet4 = sourceWorkbook.getSheetAt(4)


        val dataHashMap = HashMap<String, String>()
        //targetSheet0
        dataHashMap["파일명"] = excelName
        dataHashMap["진단도구명"] = sourceSheet0.getRow(0).getCell(1).stringCellValue
        dataHashMap["출력일"] = sourceSheet0.getRow(1).getCell(5).stringCellValue
        dataHashMap["기관명"] = sourceSheet0.getRow(3).getCell(2).stringCellValue
        dataHashMap["정보시스템명"] = sourceSheet0.getRow(4).getCell(2).stringCellValue
        dataHashMap["DBMS명"] = sourceSheet0.getRow(5).getCell(2).stringCellValue
        dataHashMap["DBMS서비스명"] = sourceSheet0.getRow(5).getCell(5).stringCellValue
        dataHashMap["DBMS종류"] = sourceSheet0.getRow(6).getCell(2).stringCellValue
        dataHashMap["DBMS버전"] = sourceSheet0.getRow(6).getCell(5).stringCellValue
        dataHashMap["DBMS서비스명"] = sourceSheet0.getRow(5).getCell(5).stringCellValue
        dataHashMap["업무규칙 수"] = (sourceSheet4.physicalNumberOfRows - 1).toString()
        dataHashMap["총 진단건수"] = sourceSheet0.getRow(28).getCell(3).stringCellValue
        dataHashMap["총 오류건수"] = sourceSheet0.getRow(28).getCell(4).stringCellValue
        dataHashMap["총 오류율"] = sourceSheet0.getRow(28).getCell(5).stringCellValue
        dataHashMap["작업시간"] = getCurrentKoreanTime()


        //매번 행 개수 체크하는게 성능상 부담일 수 도 있으니 문제가 생기면 여기부터 개선
        val targetSheet0 = targetWorkbook.getSheetAt(0)
        val headerRow = targetSheet0.getRow(0)
        val newRow = targetSheet0.createRow(targetSheet0.lastRowNum+1)
        val colNum = headerRow.lastCellNum

        for(i in 0 until colNum) {
            newRow.createCell(i).setCellValue(
                dataHashMap.getValue(targetSheet0.getRow(0).getCell(i).toString()).toString()
            )
        }
        //작업시간은 cell 에 집어넣을 떄 currentTime 으로 삽입

        //sheet 1
        dataHashMap["품질지표명"] = sourceSheet0.getRow(28).getCell(5).stringCellValue

        //sheet 2
        //상태가 "대상"인 row 의 테이블명, 상태, 범위조건, 의견 가져오기
        //별도의 리스트는 필요하지 않을 듯

        val targetSheet2 = targetWorkbook.getSheetAt(2)
        var targetRowNum = targetSheet2.physicalNumberOfRows

        for (i in 1 until sourceSheet2.physicalNumberOfRows) {
            val sourceRow = sourceSheet2.getRow(i)
            if (sourceRow?.getCell(3)?.stringCellValue == "대상") {
                val targetRow = targetSheet2.createRow(targetRowNum++)

                for (j in 0 until sourceRow.physicalNumberOfCells) {
                    val sourceCell = sourceRow.getCell(j)
                    val targetCell = targetRow.createCell(j)

                    when (sourceCell.cellType) {
                        CellType.STRING -> targetCell.setCellValue(sourceCell.stringCellValue)
                        CellType.NUMERIC -> {
                            if (DateUtil.isCellDateFormatted(sourceCell)) {
                                targetCell.setCellValue(sourceCell.dateCellValue)
                            } else {
                                targetCell.setCellValue(sourceCell.numericCellValue)
                            }
                        }
                        CellType.BOOLEAN -> targetCell.setCellValue(sourceCell.booleanCellValue)
                        CellType.FORMULA -> targetCell.cellFormula = sourceCell.cellFormula
                        else -> targetCell.setCellValue(sourceCell.toString())
                    }
                }
            }
        }



        //sheet 3
        //검증룰 명이 존재하는 row 의 테이블명, 컬럼명, 데이터타입, 검증룰명,품질지표명, 검증룰, 오류제외데이터, 의견 가져오기





        return targetWorkbook

    }

    private fun generateSDQExcel(
        sourceWorkbook: Workbook,targetWorkbook: Workbook, excelName:
    ):Workbook{

    }

    private fun generateDQUBExcel(
        sourceWorkbook: Workbook,targetWorkbook: Workbook, excelName:
    ):Workbook{

    }

    private fun generateDQMINERExcel(
        sourceWorkbook: Workbook,targetWorkbook: Workbook, excelName: String
    ):Workbook{

    }

    private fun getCurrentKoreanTime(): String {
        val now = LocalDateTime.now(ZoneId.of("Asia/Seoul"))
        val formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")
        return now.format(formatter)
    }



}