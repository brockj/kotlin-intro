package au.com.tradefloor.qldjvm

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType

fun Cell.isString(): Boolean = this.cellTypeEnum === CellType.STRING
