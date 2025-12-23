package com.excelsort.service

import com.excelsort.ESConstants
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.ByteArrayInputStream
import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths
import kotlin.io.path.extension
import kotlin.io.path.name
import kotlin.streams.asSequence

class ExcelDataLoaderService {
    private val validExtensions = setOf("xls", "xlsx")

    fun loadExcelWorkbooks(): Map<String, Workbook> {
        val rootPath: Path = Paths.get(ESConstants.ROOT_DIR)
        println("Searching for Excel files in: ${rootPath.toAbsolutePath()}")

        return Files.walk(rootPath, 1)
            .asSequence()
            .filter { path ->
                Files.isRegularFile(path) && path.extension.lowercase() in validExtensions
            }
            .mapNotNull { path ->
                val filename = getFileNameWithoutSuffix(path.name)
                println("Loading file and creating Workbook: $filename")
                try {
                    val fileBytes = Files.readAllBytes(path)
                    ByteArrayInputStream(fileBytes).use { inputStream ->
                        val workbook = WorkbookFactory.create(inputStream)
                        filename to workbook
                    }
                } catch (e: Exception) {
                    System.err.println("Error processing file '$filename': ${e.message}")
                    null
                }
            }
            .toMap()
    }


    fun getFileNameWithoutSuffix(fullName: String): String {
        val dotIndex = fullName.lastIndexOf('.')
        if (dotIndex > 0) {
            return fullName.substring(0, dotIndex)
        }
        return fullName
    }
}