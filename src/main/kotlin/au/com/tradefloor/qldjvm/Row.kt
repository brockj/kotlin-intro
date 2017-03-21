package au.com.tradefloor.qldjvm

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import java.util.regex.Pattern

fun Row.cell(value: String): Cell? = cell { value == it }

fun Row.cell(pattern: Pattern): Cell? = cell { pattern.matcher(it).matches() }

fun Row.cell(filter: (String) -> Boolean): Cell? = this.filter(Cell::isString)
        .firstOrNull { filter(it.stringCellValue) }


fun Row.hasData(): Boolean = this.firstCellNum >= 0
