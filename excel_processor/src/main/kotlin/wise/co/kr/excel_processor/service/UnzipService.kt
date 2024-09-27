package wise.co.kr.excel_processor.service

import java.io.File

interface UnzipService {

    fun unzipAndRenameService(directory: File):ByteArray

    fun scanAndFindResultReport(path:String) :ByteArray
}