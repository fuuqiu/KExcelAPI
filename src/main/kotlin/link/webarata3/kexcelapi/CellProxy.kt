/*
The MIT License (MIT)

Copyright (c) 2017 Shinichi ARATA.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
 */
package link.webarata3.kexcelapi

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.CellValue
import org.apache.poi.ss.usermodel.DateUtil
import java.util.*

class CellProxy(private val cell: Cell) {
    private var cellValue: CellValue? = null

    init {
        if (cell.cellTypeEnum == CellType.FORMULA) {
            cellValue = getFomulaCellValue(cell)
        }
    }

    private fun getCellTypeEnum(): CellType {
        return if (cellValue == null) {
            cell.cellTypeEnum
        } else {
            (cellValue as CellValue).cellTypeEnum
        }
    }

    private fun getStringCellValue(): String {
        return if (cellValue == null) cell.stringCellValue else (cellValue as CellValue).stringValue
    }

    private fun getNumericCellValue(): Double {
        return if (cellValue == null) cell.numericCellValue else (cellValue as CellValue).numberValue
    }

    private fun getBooleanCellValue(): Boolean {
        return if (cellValue == null) cell.booleanCellValue else (cellValue as CellValue).booleanValue
    }

    private fun isDateType(): Boolean {
        return if (cellValue == null) {
            if (cell.cellTypeEnum == CellType.NUMERIC) DateUtil.isCellDateFormatted(cell)
            else false
        } else {
            if ((cellValue as CellValue).cellTypeEnum == CellType.NUMERIC) DateUtil.isCellDateFormatted(cell)
            else false
        }
    }

    private fun normalizeNumericString(numeric: Double): String {
        // 为了获得诸如44.0之类的数值为44，输入数值和小数点后四舍五入的数值为
        // 如果匹配，则转换为int，以便不显示小数部分
        return if (numeric == Math.ceil(numeric)) {
            numeric.toInt().toString()
        } else numeric.toString()
    }

    private fun stringToInt(value: String): Int {
        try {
            return java.lang.Double.parseDouble(value).toInt()
        } catch (e: NumberFormatException) {
            throw IllegalAccessException("cell无法转换为Int")
        }
    }

    private fun stringToDouble(value: String): Double {
        try {
            return java.lang.Double.parseDouble(value)
        } catch (e: NumberFormatException) {
            throw IllegalAccessException("cell无法转换为Double")
        }
    }

    private fun getFomulaCellValue(cell: Cell): CellValue {
        val wb = cell.sheet.workbook
        val helper = wb.creationHelper
        val evaluator = helper.createFormulaEvaluator()
        return evaluator.evaluate(cell)
    }

    fun toStr(): String {
        when (getCellTypeEnum()) {
            CellType.STRING -> return getStringCellValue()
            CellType.NUMERIC -> return if (isDateType()) {
                throw UnsupportedOperationException("现在不支持")
            } else {
                normalizeNumericString(getNumericCellValue())
            }
            CellType.BOOLEAN -> return getBooleanCellValue().toString()
            CellType.BLANK -> return ""
            else // _NONE, ERROR
            -> throw IllegalAccessException("cell无法转换为String")
        }
    }

    fun toInt(): Int {
        when (getCellTypeEnum()) {
            CellType.STRING -> return stringToInt(getStringCellValue())
            CellType.NUMERIC -> return if (isDateType()) {
                throw IllegalAccessException("cell无法转换为Int")
            } else {
                getNumericCellValue().toInt()
            }
            else -> throw IllegalAccessException("cell无法转换为Int")
        }
    }

    fun toDouble(): Double {
        when (getCellTypeEnum()) {
            CellType.STRING -> return stringToDouble(getStringCellValue())
            CellType.NUMERIC -> return if (isDateType()) {
                throw IllegalAccessException("cell无法转换为Double")
            } else {
                getNumericCellValue()
            }
            else -> throw IllegalAccessException("cell无法转换为Double")
        }
    }

    fun toBoolean(): Boolean {
        when (getCellTypeEnum()) {
            CellType.BOOLEAN -> return getBooleanCellValue()
            else -> throw IllegalAccessException("cell无法转换为布尔值")
        }
    }

    fun toDate(): Date {
        when {
            isDateType() -> return cell.dateCellValue
            else -> throw IllegalAccessException("cell无法转换为日期")
        }
    }
}
