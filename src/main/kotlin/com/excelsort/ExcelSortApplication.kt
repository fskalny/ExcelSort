package com.excelsort

import com.excelsort.service.WorkbookWriterService

fun main(args: Array<String>) {
    val workbookWriterService = WorkbookWriterService()
    workbookWriterService.write()
}
