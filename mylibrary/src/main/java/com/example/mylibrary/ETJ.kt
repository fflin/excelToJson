package com.example.mylibrary

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.util.*

class ETJ {

    companion object {
        @JvmStatic
        fun main(args: Array<String>) {
            pp("main log start")
            val languageNeed = listOf("zh", "en", "vi")

            val wb = readExcel("resource.xlsx")
            if (wb != null) {

            } else {
                pp("wb = null  over")
            }

        }

        private fun pp(s: String) {
            println(s)
        }

        private fun readExcel(filePath: String) : Workbook? {
            var wb: Workbook? = null
            val substring = filePath.substring(filePath.lastIndexOf("."))
            val inputStream: FileInputStream = FileInputStream(filePath)

            if (substring == ".xls") {
                wb = HSSFWorkbook(inputStream)
            } else if (substring == ".xlsx") {
                wb = XSSFWorkbook(inputStream)
            }

            return wb
        }
    }



}