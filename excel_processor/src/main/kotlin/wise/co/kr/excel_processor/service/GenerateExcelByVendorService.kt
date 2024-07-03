package wise.co.kr.excel_processor.service

import org.apache.poi.ss.usermodel.Workbook

interface GenerateExcelByVendorService {

    fun generateExcelByVendor(workbook: Workbook, excelName: String) : Workbook{

    }
}