package au.com.tradefloor.qldjvm

import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.openxml4j.opc.PackageAccess
import org.apache.poi.ss.usermodel.*
import java.io.File
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import java.util.regex.Matcher
import java.util.regex.Pattern

fun main(args: Array<String>) {

    val reportFile = File(args[0])

    // ARC replacement
    WorkbookFactory.create(OPCPackage.open(reportFile, PackageAccess.READ)).use {

        val overviewSheet = it.getSheet("Overview")!!
        parseOverviewSheet(overviewSheet)

        val daybookSheet = it.getSheet("Daybook")!!
        parseDaybookSheet(daybookSheet)


    }
}

private fun parseDaybookSheet(daybookSheet: Sheet) {

    Sequence(daybookSheet::rowIterator)
            .drop(2)
            .filter(Row::hasData)
            .map { DaybookEntry(it.getCell(0).stringCellValue!!, it.getCell(1).stringCellValue!!, it.getCell(2).numericCellValue.toInt()) }
            .forEachIndexed { index, daybookEntry ->  println("${index + 1}, ${daybookEntry.accountNumber}, ${daybookEntry.accountName}, ${daybookEntry.cnoteNumber}") }
}

private fun parseOverviewSheet(overviewSheet: Sheet) {
    val reportDatePattern = Pattern.compile("Equity Advisor Reports for (.+?, .+? \\d{2}, \\d{4})")
    val reportDateFormatter = DateTimeFormatter.ofPattern("EEEE, MMMM dd, yyyy")

    val reportDate: LocalDate? = Sequence(overviewSheet::rowIterator)
            .take(4)
            .map { it.cell(reportDatePattern) }
            .filterNotNull()
            .map(Cell::getStringCellValue)
            .map {reportDatePattern.matcher(it)}
            .filter(Matcher::matches)
            .map { it.group(1) }
            .map { LocalDate.parse(it, reportDateFormatter) }
            .first()

    // fancy if/then/else and case together
    when {
        reportDate != null -> println("Date in report was $reportDate")
        else -> println("Unable to find date in overview sheet")
    }
}