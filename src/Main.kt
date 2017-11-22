import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.model.StylesTable
import org.apache.poi.xssf.usermodel.*
import java.io.FileInputStream
import java.io.FileNotFoundException
import java.io.FileOutputStream
import java.util.*
import javax.print.DocFlavor
import java.text.SimpleDateFormat
import java.time.Instant


fun main(args: Array<String>) {
    var path : String = "hoatdong.xlsx"

//    val excelFileName = "hoatdong.xlsx"//name of excel file
//
//    val sheetName = "Sheet1"//name of sheet
//
//    val wb = XSSFWorkbook()
//    val sheet = wb.createSheet(sheetName)
//
//    //iterating r number of rows
//    for (r in 0..4) {
//        val row = sheet.createRow(r)
//
//        //iterating c number of columns
//        for (c in 0..4) {
//            val cell = row.createCell(c)
//
//            cell.setCellValue("Cell $r $c")
//        }
//    }
//
//    val fileOut = FileOutputStream(excelFileName)
//
//    //write this workbook to an Outputstream.
//    wb.write(fileOut)
//    fileOut.flush()
//    fileOut.close()



    try {
        var fis : FileInputStream = FileInputStream(path)
        fis.close()
    }catch (fnf : FileNotFoundException){
        var fileOutputStream : FileOutputStream = FileOutputStream(path)
        var wb : XSSFWorkbook = XSSFWorkbook()
        var st = wb.createSheet("hoatdong")
        var row = st.createRow(0)
        st.addMergedRegion(CellRangeAddress(0,0,0,3))
        var cell = row.createCell(0)
        cell.setCellValue("Hoạt động xuất nhập sản phẩm")
        wb.write(fileOutputStream)
        fileOutputStream.flush()
        wb.close()
        fileOutputStream.close()
    }

    var fis : FileInputStream = FileInputStream(path)
    var workBook : XSSFWorkbook = XSSFWorkbook(fis)
    var sheeta : XSSFSheet = workBook.getSheet("hoatdong")
//    var day = Date.from(Instant.now())
    var sdf : SimpleDateFormat = SimpleDateFormat("dd/MM/yyyy")
    var date = sdf.format(Date.from(Instant.now()))
    var hoatdong = "Abcdef"
    var soluong = 100
    var rowCount : Int = sheeta.lastRowNum
    var columtCout : Int = 0
    var row = sheeta.createRow(++rowCount)
    var cell = row.createCell(columtCout)
    cell.setCellValue(rowCount.toDouble()+1)
    for(index in 1..3){
        var cell = row.createCell(++columtCout)
        when(index){
            1->{
                cell.setCellValue(date)
                // chuyển date từ format String sang format Date
                var cellStyle : XSSFCellStyle = workBook.createCellStyle()
                var createHelper : CreationHelper = workBook.creationHelper
                cellStyle.dataFormat = createHelper.createDataFormat().getFormat("dd/MM/yyyy")
                cell.cellStyle = cellStyle
            }
            2->cell.setCellValue(hoatdong)
            3->{
                cell.setCellValue(soluong.toDouble())
            }
        }
    }
    fis.close()
    sheeta.autoSizeColumn(1)
    var fos : FileOutputStream = FileOutputStream(path)
    workBook.write(fos)
    workBook.close()
    fos.close()


}