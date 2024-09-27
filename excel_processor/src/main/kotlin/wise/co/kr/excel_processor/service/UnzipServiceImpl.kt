package wise.co.kr.excel_processor.service


import org.springframework.stereotype.Service
import java.io.ByteArrayOutputStream
import java.io.File
import java.io.IOException
import java.nio.charset.Charset
import java.util.zip.ZipEntry
import java.util.zip.ZipFile
import java.util.zip.ZipOutputStream

@Service
class UnzipServiceImpl : UnzipService {

    private val charsets = listOf(Charset.forName("UTF-8"), Charset.forName("CP949"), Charset.forName("EUC-KR"))

    override fun unzipAndRenameService(directory: File): ByteArray {
        val errorList = mutableListOf<String>()
        val zipFiles = directory.listFiles { file -> file.extension.lowercase() == "zip" }?.toList() ?: emptyList()

        val processedZips = ByteArrayOutputStream()
        ZipOutputStream(processedZips).use { finalZipOut ->
            zipFiles.forEach { zipFile ->
                try {
                    val prefix = zipFile.nameWithoutExtension.split("_").first()
                    val renamedZipContent = ByteArrayOutputStream()

                    ZipOutputStream(renamedZipContent).use { renamedZipOut ->
                        processZipFile(zipFile, prefix, renamedZipOut, errorList)
                    }

                    // 처리된 ZIP 파일을 최종 ZIP에 추가
                    finalZipOut.putNextEntry(ZipEntry("${zipFile.nameWithoutExtension}_processed.zip"))
                    finalZipOut.write(renamedZipContent.toByteArray())
                    finalZipOut.closeEntry()

                } catch (e: IOException) {
                    val errorMessage = "ZIP 파일 처리 중 오류 발생: ${zipFile.name} - ${e.message}"
                    println(errorMessage)
                    errorList.add(errorMessage)
                    e.printStackTrace()
                }
            }

            // 에러 리스트를 텍스트 파일로 추가
            if (errorList.isNotEmpty()) {
                finalZipOut.putNextEntry(ZipEntry("error_list.txt"))
                finalZipOut.write(errorList.joinToString("\n").toByteArray())
                finalZipOut.closeEntry()
            }
        }

        println("모든 ZIP 파일 처리 완료 및 최종 ZIP 파일 생성")
        return processedZips.toByteArray()
    }

    private fun processZipFile(zipFile: File, prefix: String, renamedZipOut: ZipOutputStream, errorList: MutableList<String>) {
        for (charset in charsets) {
            try {
                ZipFile(zipFile, charset).use { zip ->
                    zip.entries().asSequence().forEach { entry ->
                        if (!entry.isDirectory) {
                            val originalName = entry.name
                            val newName = "${prefix}_${File(originalName).name}"

                            renamedZipOut.putNextEntry(ZipEntry(newName))
                            zip.getInputStream(entry).use { input ->
                                input.copyTo(renamedZipOut)
                            }
                            renamedZipOut.closeEntry()
                        }
                    }
                }
                return  // 성공적으로 처리되면 함수 종료
            } catch (e: IllegalArgumentException) {
                // 잘못된 인코딩으로 인한 예외. 다음 인코딩 시도
                continue
            } catch (e: IOException) {
                // 기타 IO 예외 처리
                val errorMessage = "ZIP 파일 처리 중 오류 발생 (${charset.name()}): ${zipFile.name} - ${e.message}"
                println(errorMessage)
                errorList.add(errorMessage)
            }
        }
        // 모든 인코딩 시도 실패
        errorList.add("모든 인코딩 시도 실패: ${zipFile.name}")
    }


    override fun scanAndFindResultReport(path: String): ByteArray {
        TODO("Not yet implemented")




    }
}