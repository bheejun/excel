package wise.co.kr.excel_processor.service

import org.apache.poi.ss.usermodel.Workbook
import org.springframework.stereotype.Service

@Service
class GenerateExcelByVendorServiceImpl(
    private val processWDQExcel: ProcessWDQExcel,
    private val processSDQExcel: ProcessSDQExcel,
    private val processDqminerExcel: ProcessDqminerExcel,
    private val processDqubeEXcel: ProcessDqubeEXcel
) : GenerateExcelByVendorService {


    override fun generateExcelByVendor(
        sourceWorkbook: Workbook,
        targetWorkbook: Workbook,
        excelName: String
    ): Workbook {

        return when (getVendorName(sourceWorkbook)) {
            "WDQ" -> {
                processWDQExcel.generateWDQExcel(sourceWorkbook, targetWorkbook, excelName)
            }
            "SDQ" -> {
                processSDQExcel.generateSDQExcel(sourceWorkbook,targetWorkbook, excelName)
            }
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

}

//    private fun generateSDQExcel(


