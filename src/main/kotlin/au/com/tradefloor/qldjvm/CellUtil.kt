package au.com.tradefloor.qldjvm

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType

// Simple expression function that is also an extension to Cell
fun Cell.isString(): Boolean = this.cellTypeEnum === CellType.STRING
