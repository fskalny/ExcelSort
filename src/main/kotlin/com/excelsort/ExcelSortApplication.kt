package com.excelsort

import com.excelsort.service.WorkbookWriterService
import org.springframework.boot.CommandLineRunner
import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.runApplication

@SpringBootApplication
class ExcelSortApplication(val workbookWriterService: WorkbookWriterService) : CommandLineRunner {

    override fun run(vararg args: String) {
        workbookWriterService.write()
    }
}

fun main(args: Array<String>) {
    runApplication<ExcelSortApplication>(*args)
}
