package founder.sjzt.statistics;

import static founder.sjzt.Constant.CMS_DATABASE_URL;
import static founder.sjzt.Constant.PASSWORD;
import static founder.sjzt.Constant.USER_NAME;
import static founder.sjzt.MainEntrance.sdf;
import founder.sjzt.object.PurchaseData;
import founder.sjzt.object.PurchaseRecord;
import founder.sjzt.object.xlsx.CustomCellStyle;
import org.apache.poi.ss.usermodel.CellType;
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
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 销售数据统计
 */
public class SalesStatistics {
  public static void main(String[] args) {

  }

  /**
   * 获取月度销售数据
   */
  public static Map<String, List<PurchaseData>> getMonthlySalesStatistics(Map<String, String> appKeyMap, Calendar startTime, Calendar endTime) {
    Map<String, List<PurchaseData>> result = new HashMap<>();
    Statement statement;
    Connection connection = null;

    try {
      connection = DriverManager.getConnection(CMS_DATABASE_URL, USER_NAME, PASSWORD);
      statement = connection.createStatement();
      for (String appKey : appKeyMap.keySet()) {
        ResultSet resultSet = statement.executeQuery(sqlByMonth(appKey, startTime, endTime));
        List<PurchaseData> dataList = new ArrayList<>();
        while (resultSet.next()) {
          PurchaseData purchaseData = new PurchaseData(
                  resultSet.getTimestamp("date"),
                  resultSet.getInt("number")
          );
          dataList.add(purchaseData);
        }
        result.put(appKey, dataList);
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
   * 获取销售数据
   */
  public static Map<String, List<PurchaseRecord>> getSalesStatistics(Map<String, String> appKeyMap) {
    Statement statement;
    Map<String, List<PurchaseRecord>> result = new HashMap<>();
    Connection connection = null;

    Calendar endTime = Calendar.getInstance();
    Calendar startTime = Calendar.getInstance();
    startTime.add(Calendar.DATE, -7);

    try {
      connection = DriverManager.getConnection(CMS_DATABASE_URL, USER_NAME, PASSWORD);
      statement = connection.createStatement();
      for (String appKey : appKeyMap.keySet()) {
        ResultSet resultSet = statement.executeQuery(sql(appKey, startTime, endTime));
        List<PurchaseRecord> dataList = new ArrayList<>();
        while (resultSet.next()) {
          PurchaseRecord purchaseRecord = new PurchaseRecord(
                  resultSet.getString("appName"),
                  resultSet.getString("userId"),
                  resultSet.getString("fontName"),
                  resultSet.getInt("price"),
                  resultSet.getTimestamp("date")
          );
          dataList.add(purchaseRecord);
        }
        result.put(appKey, dataList);
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
   * 生成订单查询sql
   *
   * @param appKey 商户唯一标识
   * @return 返回sql语句
   */
  private static String sql(String appKey, Calendar startTime, Calendar endTime) {
    return "SELECT fca.USER_NAME AS appName, prr.buyer AS userId, " +
            "ffp.FONT_NAME AS fontName, fo.FINAL_FEE AS price, " +
            "prr.create_date AS date FROM pay_result_record prr " +
            "LEFT JOIN fs_order fo ON fo.ORDER_ID = prr.ORDERID " +
            "LEFT JOIN fs_customer_app fca ON fca.APP_KEY = fo.APP_KEY " +
            "LEFT JOIN fs_order_item foi ON foi.ORDER_ID = prr.ORDERID " +
            "LEFT JOIN fs_font_pool ffp ON ffp.id = foi.ITEM_ID " +
            "WHERE prr.create_date > '" + sdf.format(startTime.getTime()) + "' " +
            "AND prr.create_date < '" + sdf.format(endTime.getTime()) + "' " +
            "AND fo.final_fee > 10 " +
            "AND fca.APP_KEY = '" + appKey + "' ORDER BY prr.CREATE_DATE DESC;";
  }

  /**
   * 生成每月销售查询sql
   */
  private static String sqlByMonth(String appKey, Calendar startTime, Calendar endTime) {
    return "SELECT DATE_FORMAT(prr.CREATE_DATE, '%Y-%m-%d') date, count(*) number FROM " +
            "pay_result_record prr LEFT JOIN fs_order fo ON fo.ORDER_ID = prr.ORDERID WHERE " +
            "prr.create_date > '" + sdf.format(startTime.getTime()) + "' " +
            "AND prr.create_date < '" + sdf.format(endTime.getTime()) + "' " +
            "AND fo.final_fee > 10 AND fo.APP_KEY = '" + appKey + "' GROUP BY date " +
            "ORDER BY date DESC;";
  }

  /**
   * 销售数据写入excel
   */
  public static void writeSalesStatistics(XSSFWorkbook workbook,
                                          Map<String, List<PurchaseRecord>> salesStatisticsMap,
                                          Map<String, String> appKeyMap) throws IOException {
    for (String appKey : salesStatisticsMap.keySet()) {
      List<PurchaseRecord> list = salesStatisticsMap.get(appKey);
      XSSFSheet sheet = workbook.createSheet(appKeyMap.get(appKey) + "字体销售情况");
      //设置特定列的宽度（1像素=32，所以宽度设为160像素）
      //用户ID列
      sheet.setColumnWidth(1, 160 * 32);
      //字体名称列
      sheet.setColumnWidth(2, 160 * 32);
      //购买日期列
      sheet.setColumnWidth(4, 160 * 32);
      XSSFRow row = sheet.createRow(0);
      //从第二格开始，写入列标题
      XSSFCell cell = row.createCell(1);
      cell.setCellValue("用户ID");
      cell.setCellStyle(CustomCellStyle.verticalCenter(workbook));
      cell = row.createCell(2);
      cell.setCellValue("字体名称");
      cell.setCellStyle(CustomCellStyle.verticalCenter(workbook));
      cell = row.createCell(3);
      cell.setCellValue("价格");
      cell.setCellStyle(CustomCellStyle.verticalRight(workbook));
      cell = row.createCell(4);
      cell.setCellValue("购买日期");
      cell.setCellStyle(CustomCellStyle.verticalCenter(workbook));
      for (int i = 0; i < list.size(); i++) {
        PurchaseRecord purchaseRecord = list.get(i);
        XSSFRow contentRow = sheet.createRow(i + 1);
        XSSFCell contentCell;

        //写入用户ID
        contentCell = contentRow.createCell(1);
        contentCell.setCellValue(purchaseRecord.getUserId());

        //写入字体名称
        contentCell = contentRow.createCell(2);
        contentCell.setCellValue(purchaseRecord.getFontName());

        //写入价格
        contentCell = contentRow.createCell(3);
        contentCell.setCellStyle(CustomCellStyle.currency(workbook));
        contentCell.setCellValue(((double) purchaseRecord.getPrice()) / 100);

        //写入购买日期
        contentCell = contentRow.createCell(4);
        contentCell.setCellStyle(CustomCellStyle.dateAndTime(workbook));
        contentCell.setCellValue(purchaseRecord.getDate());
      }

      //插入总计金额
      row = sheet.createRow(list.size() + 1);
      cell = row.createCell(0);
      cell.setCellValue("总计金额");
      cell = row.createCell(3);
      cell.setCellType(CellType.FORMULA);
      cell.setCellFormula("SUM(D2:D" + list.size() + ")");
      cell.setCellStyle(CustomCellStyle.currency(workbook));
    }
  }
}
