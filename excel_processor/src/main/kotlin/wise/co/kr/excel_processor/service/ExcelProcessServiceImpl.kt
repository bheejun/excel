package wise.co.kr.excel_processor.service

import jakarta.transaction.Transactional
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.core.io.ByteArrayResource
import org.springframework.core.io.Resource
import org.springframework.stereotype.Service
import org.springframework.web.multipart.MultipartFile
import java.io.ByteArrayOutputStream


@Service
class ExcelProcessServiceImpl(
    private val generateExcelByVendorService: GenerateExcelByVendorService

):ExcelProcessService {

    companion object {
        private val SHEET_TYPES = listOf("값진단_결과보고서", "값진단_결과보고서_상세", "진단대상테이블_목록", "진단대상컬럼_목록")
    }

    @Transactional
    override fun processExcel(files: List<MultipartFile>): Int {



        files.forEach { file ->

            val workbook = WorkbookFactory.create(file.inputStream)
            val excelName = file.originalFilename

            if (excelName != null) {
                generateExcelByVendorService.generateExcelByVendor(workbook, excelName)
            }else{
                throw IllegalArgumentException("file name is null")
            }


        }

        return 0

    }

    @Transactional
    override fun mergeExcel() {
        TODO("Not yet implemented")
    }

    private fun copySheetContent(sourceSheet: Sheet, destSheet: Sheet, startRowIndex: Int): Int {
        var currentRowIndex = startRowIndex
        for (sourceRow in sourceSheet) {
            val destRow = destSheet.createRow(currentRowIndex++)
            for (sourceCell in sourceRow) {
                val destCell = destRow.createCell(sourceCell.columnIndex)
                copyCell(sourceCell, destCell)
            }
        }
        return currentRowIndex
    }

    private fun copyCell(source: Cell, destination: Cell) {
        when (source.cellType) {
            CellType.STRING -> destination.setCellValue(source.stringCellValue)
            CellType.NUMERIC -> destination.setCellValue(source.numericCellValue)
            CellType.BOOLEAN -> destination.setCellValue(source.booleanCellValue)
            CellType.FORMULA -> destination.cellFormula = source.cellFormula
            else -> {}
        }
    }
}