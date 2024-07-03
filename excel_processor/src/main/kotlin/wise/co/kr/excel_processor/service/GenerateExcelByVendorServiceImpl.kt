package wise.co.kr.excel_processor.service

import org.apache.poi.ss.usermodel.Workbook
import org.springframework.stereotype.Service

@Service
class GenerateExcelByVendorServiceImpl : GenerateExcelByVendorService {

    override fun generateExcelByVendor(workbook: Workbook, excelName: String) :Workbook{
        val vendorName = getVendorName(workbook)

        return when (vendorName) {
            "WDQ" -> {
                generateWDQExcel(workbook: Workbook)
            }
            "SDQ" -> {
                generateWDQExcel(workbook: Workbook)
            }
            "DQMINER" -> {
                generateWDQExcel(workbook: Workbook)
            }
            "DQUBE" -> {
                generateWDQExcel(workbook: Workbook)
            }
            else -> {
                throw IllegalArgumentException("vendor not found")
            }
        }


    }

    //진단 도구 판별 함수
    private fun getVendorName(workbook: Workbook): String {

        val sheet = workbook.getSheetAt(0)
        val cell = sheet.getRow(0).getCell(1)

        return when {
            cell.stringCellValue.contains("WISE") -> "WDQ"
            cell.stringCellValue.contains("SDQ") -> "SDQ"
            cell.stringCellValue.contains("DQUBE") -> "DQUBE"
            cell.stringCellValue.equals("값진단 결과 보고서") -> "DQMINER"

            else -> throw IllegalArgumentException("vendor not found")
        }
    }


    private fun generateWDQExcel(sourceWorkbook: Workbook):Workbook{
        //값진단결과
        val sheet0 = sourceWorkbook.getSheetAt(0)
        //진단대상테이블
        val sheet1 = sourceWorkbook.getSheetAt(1)
        //도메인
        val sheet2 = sourceWorkbook.getSheetAt(2)
        //참조무결성
        val sheet3 = sourceWorkbook.getSheetAt(3)
        //업무규칙
        val sheet4 = sourceWorkbook.getSheetAt(4)


        val dataHashMap = HashMap<String, Any>()
        //sheet 0
        dataHashMap["진단도구명"] = sheet0.getRow(0).getCell(1).stringCellValue
        dataHashMap["출력일"] = sheet0.getRow(1).getCell(5).stringCellValue
        dataHashMap["기관명"] = sheet0.getRow(3).getCell(2).stringCellValue
        dataHashMap["정보시스템명"] = sheet0.getRow(4).getCell(2).stringCellValue
        dataHashMap["DBMS명"] = sheet0.getRow(5).getCell(2).stringCellValue
        dataHashMap["DBMS서비스명"] = sheet0.getRow(5).getCell(5).stringCellValue
        dataHashMap["DBMS종류"] = sheet0.getRow(6).getCell(2).stringCellValue
        dataHashMap["DBMS버전"] = sheet0.getRow(6).getCell(5).stringCellValue
        dataHashMap["DBMS서비스명"] = sheet0.getRow(5).getCell(5).stringCellValue
        dataHashMap["업무규칙 수"] = (sheet4.physicalNumberOfRows - 1)
        dataHashMap["총 진단건수"] = sheet0.getRow(28).getCell(3).stringCellValue
        dataHashMap["총 오류건수"] = sheet0.getRow(28).getCell(4).stringCellValue
        dataHashMap["총 오류율"] = sheet0.getRow(28).getCell(5).stringCellValue

        //작업시간은 cell 에 집어넣을 떄 currentTime 으로 삽입

        //sheet 1







    }

    private fun generateSDQExcel(sourceWorkbook: Workbook):Workbook{

    }

    private fun generateDQUBExcel(sourceWorkbook: Workbook):Workbook{

    }

    private fun generateDQMINERExcel(sourceWorkbook: Workbook):Workbook{

    }



}