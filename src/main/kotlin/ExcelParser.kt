import org.apache.poi.hssf.usermodel.*
import org.apache.poi.hssf.util.*
import org.apache.poi.ss.usermodel.*
import java.io.*
import org.apache.poi.ss.usermodel.*
import org.apache.poi.hssf.usermodel.*

open class ExcelParser {

    fun parse(fileName: String, cellChange: String ): String {
        var result = ""
        var inputStream: InputStream?
        var outputStream: OutputStream? = null
        var workBook: HSSFWorkbook? = null
        try {
            inputStream = FileInputStream(fileName)
            workBook = HSSFWorkbook(inputStream)
            outputStream = FileOutputStream(fileName)

        } catch (e: IOException) {
            e.printStackTrace()
        }

        val colorStyleRed = workBook?.createCellStyle()
        colorStyleRed?.fillPattern = HSSFCellStyle.SOLID_FOREGROUND
        colorStyleRed?.fillForegroundColor = HSSFColor.RED.index

        val colorStyleBlue = workBook?.createCellStyle()
        colorStyleBlue?.fillPattern = HSSFCellStyle.SOLID_FOREGROUND
        colorStyleBlue?.fillForegroundColor = HSSFColor.BLUE.index

        val colorStyleYellow = workBook?.createCellStyle()
        colorStyleYellow?.fillPattern = HSSFCellStyle.SOLID_FOREGROUND
        colorStyleYellow?.fillForegroundColor = HSSFColor.YELLOW.index

        val colorStyleGreen = workBook?.createCellStyle()
        colorStyleGreen?.fillPattern = HSSFCellStyle.SOLID_FOREGROUND
        colorStyleGreen?.fillForegroundColor = HSSFColor.GREEN.index

        //разбираем первый лист входного файла на объектную модель
        val createHelper = workBook!!.creationHelper
        val sheet = workBook.getSheetAt(0)
        val it = sheet.iterator()
//        val cellReference = CellReference(cellChange)
        //проходим по всему листу
        while (it.hasNext()) {
            val row = it.next()
            val cells = row.iterator()
            while (cells.hasNext()) {
                val cell = cells.next()
//                if(cell.columnIndex  == (cellReference.col.toInt()) && row.rowNum == cellReference.row ) {
//                    cell.cellStyle = colorStyleRed
//                }
                val cellType = cell.cellType

                //перебираем возможные типы ячеек
                when (cellType) {
                    Cell.CELL_TYPE_STRING ->{
                        result += cell.stringCellValue + "="

                        val fontHeader = workBook.createFont()
                        fontHeader.boldweight = HSSFFont.BOLDWEIGHT_BOLD
                        fontHeader.fontName = "Arial"
                        colorStyleGreen?.setFont(fontHeader)
                        cell.cellStyle = colorStyleGreen
                    }

                    Cell.CELL_TYPE_NUMERIC -> {
                        if (DateUtil.isCellDateFormatted(cell)) {
                            result += "[" + cell.dateCellValue + "]"
                            colorStyleYellow?.dataFormat = createHelper.createDataFormat().getFormat("dd/MM/yyyy")
                            cell.cellStyle = colorStyleYellow
                        } else {
                            result += "[" + cell.numericCellValue + "]"
                            cell.cellStyle = colorStyleRed
                        }
                    }

                    Cell.CELL_TYPE_FORMULA -> {
                        result += "[" + cell.numericCellValue + "]"
                        cell.cellStyle = colorStyleBlue
                    }
                    else -> result += "|"
                }
            }
            result += "\n"
        }
        try {
            workBook.write(outputStream)
        }catch (e:IOException){
            e.printStackTrace()
        }
        return result
    }

}