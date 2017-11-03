package founder.sjzt.statistics;

import static founder.sjzt.Constant.CMS_DATABASE_URL;
import static founder.sjzt.Constant.PASSWORD;
import static founder.sjzt.Constant.USER_NAME;
import static founder.sjzt.MainEntrance.sdf;
import founder.sjzt.object.DownloadRecord;
import founder.sjzt.object.xlsx.CustomCellStyle;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.ChartLegend;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 下载数据统计
 */
public class DownloadsStatistics {
  /**
   * 获取月度下载数据
   */
  public static Map<String, List<List<DownloadRecord>>> getMonthlyDownloadStatistics(Map<String, String> DownloadsAppKey) {
    Map<String, List<List<DownloadRecord>>> result = new HashMap<>();
    Statement statement;
    Connection connection = null;


    return null;
  }

  /**
   * 获取下载数据
   */
  public static Map<String, List<List<DownloadRecord>>> getDownloadStatistics(Map<String, String> DownloadsAppKey, Calendar startTime, Calendar endTime) {
    Statement statement;
    Map<String, List<List<DownloadRecord>>> result = new HashMap<>();
    Connection connection = null;

    try {
      connection = DriverManager.getConnection(CMS_DATABASE_URL, USER_NAME, PASSWORD);
      statement = connection.createStatement();
      for (String appKey : DownloadsAppKey.keySet()) {
        List<List<DownloadRecord>> lists = new ArrayList<>();
        ResultSet resultSet1 = statement.executeQuery(sqlByFont(appKey, startTime, endTime));
        List<DownloadRecord> dataList1 = new ArrayList<>();
        while (resultSet1.next()) {
          DownloadRecord downloadRecord = new DownloadRecord(
                  resultSet1.getString("fontName"),
                  resultSet1.getInt("downloads")
          );
          dataList1.add(downloadRecord);
        }
        lists.add(dataList1);

        statement = connection.createStatement();
        ResultSet resultSet2 = statement.executeQuery(sqlByDate(appKey, startTime, endTime));
        List<DownloadRecord> dataList2 = new ArrayList<>();
        while (resultSet2.next()) {
          DownloadRecord downloadRecord = new DownloadRecord(
                  resultSet2.getTimestamp("date"),
                  resultSet2.getInt("downloads")
          );
          dataList2.add(downloadRecord);
        }
        lists.add(dataList2);

        result.put(appKey, lists);
      }
    } catch (SQLException e) {
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
   * 生成所需sql
   *
   * @param appKey 商户唯一标识
   * @return 返回sql语句
   */
  private static String sqlByFont(String appKey, Calendar startTime, Calendar endTime) {
    return "SELECT ffp.font_name AS fontName, count(*) AS downloads FROM " +
            "fs_customer_font_downloads fcfd LEFT JOIN fs_font_pool ffp " +
            "ON ffp.id = fcfd.font_id WHERE fcfd.app_key = '" + appKey +
            "' AND fcfd.is_deleted = '0' AND fcfd.create_date > '" +
            sdf.format(startTime.getTime()) + "' AND fcfd.create_date < '" +
            sdf.format(endTime.getTime()) + "' AND ffp.IS_DELETED = '0' " +
            "GROUP BY fcfd.font_id ORDER BY downloads DESC";
  }

  /**
   * @param appKey 商户唯一标识
   * @return 返回sql语句
   */
  private static String sqlByDate(String appKey, Calendar startTime, Calendar endTime) {
    return "SELECT DATE_FORMAT(create_date,'%Y-%m-%d') date, count(*) AS downloads " +
            "FROM fs_customer_font_downloads WHERE app_key = '" + appKey +
            "' AND is_deleted = '0' AND create_date > '" + sdf.format(startTime.getTime()) +
            "' AND create_date < '" + sdf.format(endTime.getTime()) + "' GROUP BY date " +
            "ORDER BY date";
  }

  /**
   * 下载数据写入excel
   */
  public static void writeDownloadStatistics(XSSFWorkbook workbook,
                                             Map<String, List<List<DownloadRecord>>> downloadStatisticsMap,
                                             Map<String, String> appKeyMap, boolean monthly) {
    for (String appKey : downloadStatisticsMap.keySet()) {
      List<DownloadRecord> listByFont = downloadStatisticsMap.get(appKey).get(0);
      List<DownloadRecord> listByDate = downloadStatisticsMap.get(appKey).get(1);
      XSSFSheet sheet = workbook.createSheet(appKeyMap.get(appKey) + "字体下载情况");
      //设置特定列的宽度（1像素=32，所以宽度设为160/100像素）
      //字体名称列
      sheet.setColumnWidth(1, 160 * 32);
      //日期列
      sheet.setColumnWidth(6, 100 * 32);

      XSSFRow row = sheet.createRow(0);
      XSSFCell cell = row.createCell(1);
      cell.setCellStyle(CustomCellStyle.verticalCenter(workbook));
      cell.setCellValue("字体下载量（按字体）");
      cell = row.createCell(6);
      cell.setCellStyle(CustomCellStyle.verticalCenter(workbook));
      cell.setCellValue("字体下载量（按日期）");
      //合并单元格
      sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 2));
      sheet.addMergedRegion(new CellRangeAddress(0, 0, 6, 7));

      //写入列标题
      row = sheet.createRow(1);
      cell = row.createCell(1);
      cell.setCellValue("字体名称");
      cell.setCellStyle(CustomCellStyle.verticalCenter(workbook));
      cell = row.createCell(2);
      cell.setCellValue("下载量");
      cell.setCellStyle(CustomCellStyle.verticalRight(workbook));
      cell = row.createCell(6);
      cell.setCellValue("日期");
      cell.setCellStyle(CustomCellStyle.verticalCenter(workbook));
      cell = row.createCell(7);
      cell.setCellValue("下载量");
      cell.setCellStyle(CustomCellStyle.verticalRight(workbook));

      int fontRow = 0;

      //写入按字体名统计的字体下载量
      for (int i = 0; i < listByFont.size(); i++) {
        XSSFRow contentRow = sheet.createRow(i + 2);
        fontRow = i + 2;
        XSSFCell contentCell;

        //写入字体名
        contentCell = contentRow.createCell(1);
        contentCell.setCellValue(listByFont.get(i).getFontName());
        //写入下载量
        contentCell = contentRow.createCell(2);
        contentCell.setCellValue(listByFont.get(i).getDownloads());
      }

      int dateRow = 0;
      //写入按日期统计的字体下载量
      for (int i = 0; i < listByDate.size(); i++) {
        dateRow = i + 2;
        XSSFRow contentRow;
        if (i + 2 <= fontRow) {
          contentRow = sheet.getRow(i + 2);
        } else {
          contentRow = sheet.createRow(i + 2);
        }
        XSSFCell contentCell;

        //写入日期
        contentCell = contentRow.createCell(6);
        contentCell.setCellStyle(CustomCellStyle.date(workbook));
        contentCell.setCellValue(listByDate.get(i).getDate());
        //写入下载量
        contentCell = contentRow.createCell(7);
        contentCell.setCellValue(listByDate.get(i).getDownloads());
      }
      if (monthly) {
        Drawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor clientAnchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 10, 15);
        Chart chart = drawing.createChart(clientAnchor);

        ChartLegend legend = chart.getOrCreateLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);

        LineChartData data = chart.getChartDataFactory().createLineChartData();

        ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
        ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        ChartDataSource<Number> xs = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(2, dateRow, 6, 6));
        ChartDataSource<Number> ys1 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(2, dateRow, 6, 6));
        ChartDataSource<Number> ys2 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(2, dateRow, 7, 7));

        data.addSeries(xs, ys1);
        data.addSeries(xs, ys2);

        chart.plot(data, bottomAxis, leftAxis);
      }
    }
  }
}
