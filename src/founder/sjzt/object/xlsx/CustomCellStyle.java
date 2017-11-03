package founder.sjzt.object.xlsx;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CustomCellStyle {

  //单元格格式：货币（人民币￥）
  public static XSSFCellStyle currency(XSSFWorkbook workbook) {
    XSSFCellStyle cellStyle = workbook.createCellStyle();
    XSSFDataFormat xssfDataFormat = workbook.createDataFormat();
    cellStyle.setDataFormat(xssfDataFormat.getFormat("¥#,##0.00"));
    return cellStyle;
  }

  //单元格格式：日期（yyyy-mm-dd hh:mm:ss）
  public static XSSFCellStyle dateAndTime(XSSFWorkbook workbook) {
    XSSFCellStyle cellStyle = workbook.createCellStyle();
    XSSFDataFormat xssfDataFormat = workbook.createDataFormat();
    cellStyle.setDataFormat(xssfDataFormat.getFormat("yyyy-mm-dd hh:mm:ss"));
    return cellStyle;
  }

  //单元格格式：日期（yyyy-mm-dd hh:mm:ss）
  public static XSSFCellStyle date(XSSFWorkbook workbook) {
    XSSFCellStyle cellStyle = workbook.createCellStyle();
    XSSFDataFormat xssfDataFormat = workbook.createDataFormat();
    cellStyle.setDataFormat(xssfDataFormat.getFormat("yyyy-mm-dd"));
    return cellStyle;
  }

  //单元格格式：水平居中
  public static XSSFCellStyle verticalCenter(XSSFWorkbook workbook) {
    XSSFCellStyle cellStyle = workbook.createCellStyle();
    cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    cellStyle.setAlignment(HorizontalAlignment.CENTER);
    return cellStyle;
  }

  //单元格格式：水平右对齐
  public static XSSFCellStyle verticalRight(XSSFWorkbook workbook) {
    XSSFCellStyle cellStyle = workbook.createCellStyle();
    cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    cellStyle.setAlignment(HorizontalAlignment.RIGHT);
    return cellStyle;
  }
}
