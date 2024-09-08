package wise.co.kr.excel_processor.service

import org.springframework.stereotype.Service
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

@Service
class ProcessWDQExcel:ProcessExcelService {

    //TODO:hashmap 만들고 생성된 hashmap 으로 엑셀 생성하는 로직 만들기
    fun generateWDQExcel(
        sourceWorkbook: Workbook, targetWorkbook: Workbook, excelName: String
    ): Workbook {
        // 값진단결과
        val sourceSheet0 = sourceWorkbook.getSheet("(진단결과)값진단결과") ?: throw IllegalArgumentException("Sheet 0 not found")
        // 진단대상테이블
        val sourceSheet1 = sourceWorkbook.getSheet("(테이블선정)진단대상테이블") ?: throw IllegalArgumentException("Sheet 1 not found")
        // 도메인
        val sourceSheet2 = sourceWorkbook.getSheet("(룰설정)도메인") ?: throw IllegalArgumentException("Sheet 2 not found")
        // 업무규칙
        val sourceSheet4 = sourceWorkbook.getSheet("(룰설정)업무규칙") ?: throw IllegalArgumentException("Sheet 4 not found")


        val dataHashMap0 = generateHashMap(sourceSheet0)
        dataHashMap0["파일명"] = excelName
        dataHashMap0["진단도구명"] = sourceSheet0.getRow(0).getCell(1)
        dataHashMap0["업무규칙 수"] = (sourceSheet4.physicalNumberOfRows - 1).toString()
        dataHashMap0["작업시간"] = getCurrentKoreanTime()

        val targetSheet0 = targetWorkbook.getSheetAt(0) ?: throw IllegalArgumentException("Target Sheet 0 not found")
        val headerRowSheet0 = targetSheet0.getRow(0) ?: targetSheet0.createRow(0)
        val newRowSheet0 = targetSheet0.createRow(targetSheet0.lastRowNum + 1)
        val colNumSheet0 = headerRowSheet0.physicalNumberOfCells

        for (i in 0 until colNumSheet0) {
            val headerCell = headerRowSheet0.getCell(i) ?: continue
            val headerCellValue = headerCell.stringCellValue
            val value = dataHashMap0[headerCellValue]?.toString() ?: ""
            newRowSheet0.createCell(i).setCellValue(value)
        }

        // Sheet 1 처리
        val targetSheet1 = targetWorkbook.getSheetAt(1) ?: throw IllegalArgumentException("Target Sheet 1 not found")
        val headerRowSheet1 = targetSheet1.getRow(0) ?: targetSheet1.createRow(0)
        val rowNumSheet1 = targetSheet1.physicalNumberOfRows
        val colNumSheet1 = headerRowSheet1.physicalNumberOfCells

        //qualityIndicatorNameList 는 key 가 품질지표명이고 value 가 진단건수, 오류건수, 오류율을 갖고있는 List<Stirng> 인 HashMap의 List
        //
        val qualityIndicatorNameList = dataHashMap0["품질지표명"] as MutableList<*>

        for (i in rowNumSheet1 until rowNumSheet1 + qualityIndicatorNameList.size) {
            val newRowSheet1 = targetSheet1.createRow(i)
            val qualityHashMap = qualityIndicatorNameList[i - rowNumSheet1] as HashMap<String, List<String>>
            val key = qualityHashMap.keys.first().toString()
            val valueList = qualityHashMap.getValue(key)

            for (j in 0 until colNumSheet1) {
                val headerCell = headerRowSheet1.getCell(j) ?: continue
                when (val headerCellValue = headerCell.stringCellValue) {
                    "품질지표명" -> newRowSheet1.createCell(j).setCellValue(key)
                    "진단건수" -> newRowSheet1.createCell(j).setCellValue(valueList[0])
                    "오류건수" -> newRowSheet1.createCell(j).setCellValue(valueList[1])
                    "오류율" -> newRowSheet1.createCell(j).setCellValue(valueList[2])
                    else -> {
                        val value = dataHashMap0[headerCellValue]?.toString() ?: ""
                        newRowSheet1.createCell(j).setCellValue(value)
                    }
                }
            }

        }
        //sheet 2
        //상태가 "대상"인 row 의 테이블명, 상태, 범위조건, 의견 가져오기
        //별도의 리스트는 필요하지 않을 듯

        val targetSheet2 = targetWorkbook.getSheetAt(2) ?: throw IllegalArgumentException("Target Sheet 2 not found")

        for (i in 0 until sourceSheet2.physicalNumberOfRows) {
            val sourceRow = sourceSheet1.getRow(i) ?: continue
            val statusCell = sourceRow.getCell(3)
            if (statusCell.stringCellValue == "대상") {
                val targetRow = targetSheet2.createRow(targetSheet2.physicalNumberOfRows)
                for (j in 0 until targetSheet2.getRow(0).physicalNumberOfCells) {
                    val sourceCell = sourceRow.getCell(j)
                    val value = getCellValueAsString(sourceCell)
                    targetRow.createCell(j).setCellValue(value)
                }
            }
        }

        //sheet 3
        //검증룰 명이 존재하는 row 의 테이블명, 컬럼명, 데이터타입, 검증룰명,품질지표명, 검증룰, 오류제외데이터, 의견 가져오기
        //source sheet 에서 row 를 기준으로 반복문을 돌면서 일치하는 단어가 있는 헤더 기준 셀을 복사해서 value 로 집어넣기
        val targetSheet3 = targetWorkbook.getSheetAt(3) ?: throw IllegalArgumentException("Target Sheet 2 not found")

        for (i in 1 until sourceSheet2.physicalNumberOfRows) {
            val sourceHeaderRowList : MutableList<String> = mutableListOf()
            sourceSheet2.getRow(0).forEach { cell -> sourceHeaderRowList.add(cell.toString()) }
            val targetHeaderRow = targetSheet3.getRow(0) ?: throw IllegalArgumentException("TargetHeaderRow is not created")
            val sourceRow = sourceSheet2.getRow(i) ?: continue
            val elementHashMap: HashMap<String, String> = hashMapOf()

            if (sourceRow.getCell(6).toString().isNotBlank()) {

                for (j in 0 until targetHeaderRow.physicalNumberOfCells) {
                    if(targetHeaderRow.getCell(j).toString().contains("의견")){
                        elementHashMap["의견"] = sourceRow.getCell(j).toString()
                    }else if(sourceHeaderRowList.contains(targetHeaderRow.getCell(j).toString())){
                        //sourceSheet2의 헤더로우에 해당 key를 가진 cell 인덱스를 찾아서 넣는다
                        elementHashMap[targetHeaderRow.getCell(j).toString()] = sourceRow.getCell(
                            sourceHeaderRowList.indexOf(targetHeaderRow.getCell(j).toString())
                        ).toString()
                    }
                }

                val targetRow = targetSheet3.createRow(targetSheet3.physicalNumberOfRows)

                for (k in 0 until targetHeaderRow.physicalNumberOfCells) {
                    targetRow.createCell(k).setCellValue(
                            elementHashMap[targetHeaderRow.getCell(k).toString()]
                        )
                }

                elementHashMap.clear()

            } else {
                continue
            }
        }
        return targetWorkbook

    }




    private fun generateHashMap(sourceSheet: Sheet): HashMap<String, Any> {
        val resultMap = HashMap<String, Any>()
        val keywordsToRow = listOf("기관명", "정보시스템명", "DBMS명", "DBMS서비스(스키마)명", "DBMS종류", "DBMS버전")
        val keywordsToCol = listOf("진단건수", "오류건수", "오류율")

        for (rowIndex in 0..sourceSheet.lastRowNum) {
            val row = sourceSheet.getRow(rowIndex) ?: continue
            for (cellIndex in 0..row.lastCellNum) {
                val cell = row.getCell(cellIndex) ?: continue
                val cellValue = getCellValueAsString(cell)
                val qualityIndicatorHashMapList: MutableList<HashMap<String, List<String>>> = mutableListOf()

                when {
                    keywordsToRow.contains(cellValue) -> {
                        val nextCell = row.getCell(cellIndex + 1)
                        if (nextCell != null) {
                            when (cellValue) {
                                "정보시스템명" -> {
                                    resultMap["정보시스템명"] = getCellValueAsString(nextCell)
                                    resultMap["시스템명"] = getCellValueAsString(nextCell)
                                }

                                "DBMS명" -> {
                                    resultMap["DBMS명"] = getCellValueAsString(nextCell)
                                    resultMap["DB명"] = getCellValueAsString(nextCell)
                                }

                                "DBMS서비스(스키마)명" -> {
                                    resultMap["DB서비스명"] = getCellValueAsString(nextCell)
                                    resultMap["DB명"] = getCellValueAsString(nextCell)
                                }

                                "DBMS종류" -> {
                                    resultMap["DBMS종류"] = getCellValueAsString(nextCell)
                                    resultMap["DB종류"] = getCellValueAsString(nextCell)
                                }

                                "DBMS버전" -> {
                                    resultMap["버전"] = getCellValueAsString(nextCell)
                                }

                                else -> {
                                    resultMap[cellValue] = getCellValueAsString(nextCell)
                                }
                            }


                        }
                    }

                    keywordsToCol.contains(cellValue) -> {
                        for (i in sourceSheet.lastRowNum downTo rowIndex) {
                            val lastCell = sourceSheet.getRow(i).getCell(cellIndex)
                            if (lastCell != null && getCellValueAsString(lastCell).isNotBlank()) {
                                when (cellValue) {
                                    "진단건수" -> {
                                        resultMap["총 진단건수"] = getCellValueAsString(lastCell)
                                    }

                                    "오류건수" -> {
                                        resultMap["총 오류건수"] = getCellValueAsString(lastCell)
                                    }

                                    "오류율" -> {
                                        resultMap["총 오류율"] = getCellValueAsString(lastCell)
                                    }
                                }
                                break
                            }
                        }
                    }

                    cellValue.contains("출력일") -> {
                        resultMap["출력일"] = cellValue.substring(cellValue.indexOf(":") + 1).trim()
                    }

                    cellValue == "품질지표명" -> {
                        var newRowIndex = rowIndex + 1
                        var newRow = sourceSheet.getRow(newRowIndex)
                        var newCellValue = getCellValueAsString(newRow.getCell(cellIndex))

                        while (newCellValue.isNotBlank() && newCellValue != "합계") {
                            val qualityIndicatorHashMap: HashMap<String, List<String>> = hashMapOf()

                            val diagnosisCount = getCellValueAsString(newRow.getCell(cellIndex + 2))
                            val errorCount = getCellValueAsString(newRow.getCell(cellIndex + 3))
                            val errorRate = getCellValueAsString(newRow.getCell(cellIndex + 4))

                            qualityIndicatorHashMap[newCellValue] = listOf(diagnosisCount, errorCount, errorRate)
                            qualityIndicatorHashMapList.add(qualityIndicatorHashMap)

                            // 다음 행으로 이동
                            newRowIndex++
                            newRow = sourceSheet.getRow(newRowIndex)
                            newCellValue = getCellValueAsString(newRow.getCell(cellIndex))
                        }
                        resultMap["품질지표명"] = qualityIndicatorHashMapList
                    }

                    else -> {
                        continue
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


}