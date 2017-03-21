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

    // Types are implied (String in this case)
    val reportFile = File(args[0])

    // ARC replacement
    WorkbookFactory.create(OPCPackage.open(reportFile, PackageAccess.READ)).use {

        // unsafe cast to non null type
        // Implicit 'it' variable from 'use' function
        val overviewSheet = it.getSheet("Overview")!!
        parseOverviewSheet(overviewSheet)

        val daybookSheet = it.getSheet("Daybook")!!
        parseDaybookSheet(daybookSheet)

    }
}

private fun parseDaybookSheet(daybookSheet: Sheet) {

    println(
"""
    |Row, Account Number, Account Name, Cnote Number
""".trimMargin()) // Trim margin takes an optional margin with a default of '|'

    // Helper to convert from ()->Iterator<T> to Kotlin specific Sequence (similar to a Java Stream)
    Sequence(daybookSheet::rowIterator)
            .drop(2)
            .filter(Row::hasData)
            .map(Row::toDaybookEntry)
            // Copy is provided by data classes to help write immutable objects
            // All params are options, so using named parameters helps reduce number of overloaded methods
            .map { it.copy(accountName = "*** MASKED ***") }
            // String template/interpolation
            .forEachIndexed { index, daybookEntry -> println("""${index + 1}, "${daybookEntry.accountNumber}", "${daybookEntry.accountName}", ${daybookEntry.cnoteNumber}""") }
}

private fun parseOverviewSheet(overviewSheet: Sheet) {
    val reportDatePattern = Pattern.compile("""Equity Advisor Reports for (.+?, .+? \d{2}, \d{4})""")
    val reportDateFormatter = DateTimeFormatter.ofPattern("EEEE, MMMM dd, yyyy")

    // Implied type here is LocalDate? as firstOrNull could return null
    val reportDate: LocalDate? = Sequence(overviewSheet::rowIterator)
            // Only process at most 4 elements from source iterator
            .take(4)
            .map { it.cell(reportDatePattern) }
            // Removes null values so type can be Sequence<Cell> instead of Sequence<Cell?>
            .filterNotNull()
            .map(Cell::getStringCellValue)
            .map { reportDatePattern.matcher(it) }
            .filter(Matcher::matches)
            .map { it.group(1) }
            .map { LocalDate.parse(it, reportDateFormatter) }
            .firstOrNull()

    // fancy if/then/else and case together
    when(reportDate) {
        null -> println("Unable to find date in overview sheet")
        LocalDate.now() -> println("Report date is today")
        else -> println("Date in report was $reportDate")
    }
}

// This extension function is internal to hide the conversion from other modules
internal fun Row.toDaybookEntry() : DaybookEntry {
    // getCell is not null safe, but not picked up by type checking
    return DaybookEntry(this.getCell(0).stringCellValue,
                 this.getCell(1).stringCellValue,
                 this.getCell(2).numericCellValue.toInt())
}