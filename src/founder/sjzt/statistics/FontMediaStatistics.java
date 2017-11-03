package founder.sjzt.statistics;

import static founder.sjzt.Constant.FONTSHOP_DATABASE_URL;
import static founder.sjzt.Constant.PASSWORD;
import static founder.sjzt.Constant.USER_NAME;
import founder.sjzt.object.FontMediaRecord;
import founder.sjzt.object.xlsx.CustomCellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

/**
 * 字媒体发布数据统计
 */
public class FontMediaStatistics {
  /**
   * 获取字媒体发布数据
   */
  public static List<FontMediaRecord> getFontMediaStatistics() {
    Statement statement;
    List<FontMediaRecord> result = new ArrayList<>();
    Connection connection = null;
    try {
      connection = DriverManager.getConnection(FONTSHOP_DATABASE_URL, USER_NAME, PASSWORD);
      statement = connection.createStatement();
      ResultSet resultSet = statement.executeQuery(sql());
      while (resultSet.next()) {
        FontMediaRecord fontMediaRecord = new FontMediaRecord(
                resultSet.getTimestamp("date"),
                resultSet.getInt("number")
        );
        result.add(fontMediaRecord);
      }
    } catch (Exception e) {
      e.printStackTrace();
    } finally {
      try {
        connection.close();
      } catch (SQLException e) {
        e.printStackTrace();
      }
    }
    return result;
  }

  /**
   * 生成字媒体发布查询sql
   *
   * @return 返回sql语句
   */
  private static String sql() {
    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
    Calendar endTime = Calendar.getInstance();
    Calendar startTime = Calendar.getInstance();
    startTime.add(Calendar.DATE, -7);
    return "SELECT DATE_FORMAT(CREATE_DATE, '%Y-%m-%d') date, count(*) number FROM " +
            "fs_font_media WHERE IS_DELETED = '0' AND create_date > '" +
            sdf.format(startTime.getTime()) + "'  AND create_date < '" +
            sdf.format(endTime.getTime()) + "' GROUP BY date ORDER BY date DESC";
  }

  /**
   * 字媒体发布数据写入excel
   */
  public static void writeFontMediaStatistics(XSSFWorkbook workbook,
                                              List<FontMediaRecord> fontMediaStatisticsList) throws IOException {
    XSSFSheet sheet = workbook.createSheet("每周字媒体统计");
    //设置特定列的宽度（1像素=32，所以宽度设为100像素）
    //日期列
    sheet.setColumnWidth(1, 100 * 32);
    XSSFRow row = sheet.createRow(0);
    XSSFCell cell = row.createCell(1);
    cell.setCellStyle(CustomCellStyle.verticalCenter(workbook));
    cell.setCellValue("每日发布数量");
    //合并单元格
    sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 2));

    row = sheet.createRow(1);
    //从第二格开始，写入列标题
    cell = row.createCell(1);
    cell.setCellValue("日期");
    cell.setCellStyle(CustomCellStyle.verticalCenter(workbook));
    cell = row.createCell(2);
    cell.setCellValue("数量");
    cell.setCellStyle(CustomCellStyle.verticalRight(workbook));

    for (int i = 0; i < fontMediaStatisticsList.size(); i++) {
      XSSFRow contentRow = sheet.createRow(i + 2);
      XSSFCell contentCell;

      //写入日期
      contentCell = contentRow.createCell(1);
      contentCell.setCellValue(fontMediaStatisticsList.get(i).getDate());
      contentCell.setCellStyle(CustomCellStyle.date(workbook));
      //写入数量
      contentCell = contentRow.createCell(2);
      contentCell.setCellValue(fontMediaStatisticsList.get(i).getNumber());
    }
  }
}
