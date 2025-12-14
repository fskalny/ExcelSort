package com.excelsort.utils

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook


class CellStyling {
    companion object {

        fun getNazwaProduktuStyle(workbook: XSSFWorkbook): XSSFCellStyle {
            val style = getRegularStyle(workbook)
            style.alignment = HorizontalAlignment.LEFT
            return style
        }

        fun getRegularStyle(workbook: XSSFWorkbook): XSSFCellStyle {
            val createHelper: CreationHelper = workbook.creationHelper
            val format = createHelper.createDataFormat()


            val style: XSSFCellStyle = workbook.createCellStyle()
            style.setFont(getRegularFont(workbook))
            style.borderLeft = BorderStyle.NONE
            style.borderRight = BorderStyle.NONE
            style.borderTop = BorderStyle.HAIR
            style.borderBottom = BorderStyle.HAIR
            style.wrapText = true
            style.alignment = HorizontalAlignment.CENTER
            style.verticalAlignment = VerticalAlignment.CENTER
            style.dataFormat = format.getFormat("0")
            return style
        }

        private fun getRegularFont(workbook: XSSFWorkbook): Font {
            val font: XSSFFont = workbook.createFont()
            font.fontHeightInPoints = 11.toShort()
            font.fontName = "Calibri"
            font.color = IndexedColors.BLACK.index
            return font
        }


    }
}