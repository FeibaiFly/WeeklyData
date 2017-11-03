package founder.sjzt;

import static founder.sjzt.Constant.DOWNLOADS_APPKEY;
import static founder.sjzt.Constant.SALES_APPKEY;
import static founder.sjzt.statistics.DownloadsStatistics.getDownloadStatistics;
import static founder.sjzt.statistics.DownloadsStatistics.writeDownloadStatistics;
import static founder.sjzt.statistics.FontMediaStatistics.getFontMediaStatistics;
import static founder.sjzt.statistics.FontMediaStatistics.writeFontMediaStatistics;
import static founder.sjzt.statistics.SalesStatistics.getMonthlySalesStatistics;
import static founder.sjzt.statistics.SalesStatistics.getSalesStatistics;
import static founder.sjzt.statistics.SalesStatistics.writeSalesStatistics;
import founder.sjzt.object.DownloadRecord;
import founder.sjzt.object.FontMediaRecord;
import founder.sjzt.object.PurchaseRecord;
import founder.sjzt.util.JarTool;
import jcifs.UniAddress;
import jcifs.smb.NtlmPasswordAuthentication;
import jcifs.smb.SmbFile;
import jcifs.smb.SmbFileOutputStream;
import jcifs.smb.SmbSession;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MainEntrance {
  public static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

  public static void main(String[] args) {
    Calendar endTime = Calendar.getInstance();
    endTime.add(Calendar.DAY_OF_WEEK, -1);
    Calendar startTime = Calendar.getInstance();
    startTime.add(Calendar.DAY_OF_WEEK, -7);
    XSSFWorkbook workbook = new XSSFWorkbook();
    FileOutputStream fos = null;

    //查看是否需要月度数据
    boolean monthly = endTime.get(Calendar.MONTH) != startTime.get(Calendar.MONTH) || startTime.get(Calendar.DAY_OF_MONTH) == 1;

    try {
      //获取需要查询销售量的appKey
      Map<String, String> SalesAppKey = new HashMap<>();
      getAppKey(SalesAppKey, SALES_APPKEY);

      //获取字体销售数据
      Map<String, List<PurchaseRecord>> salesStatisticsMap = getSalesStatistics(SalesAppKey);
      System.out.println("==============   已取得销售量数据   ==============");

      String fileName = JarTool.getJarDir() + "/" + sdf.format(startTime.getTime()) + "-" + sdf.format(endTime.getTime()) + ".xlsx";

      fos = new FileOutputStream(fileName);

      //将销售统计数据信息写入excel
      writeSalesStatistics(workbook, salesStatisticsMap, SalesAppKey);

      //获取需要查询下载量的appKey
      Map<String, String> DownloadsAppKey = new HashMap<>();
      getAppKey(DownloadsAppKey, DOWNLOADS_APPKEY);

      Map<String, List<List<DownloadRecord>>> downloadStatisticsMap = getDownloadStatistics(DownloadsAppKey, startTime, endTime);
      System.out.println("==============   已取得下载量数据   ==============");

      //将下载统计信息写入excel
      writeDownloadStatistics(workbook, downloadStatisticsMap, DownloadsAppKey, false);

      System.out.println("==============   已取得字媒体数据   ==============");
      List<FontMediaRecord> fontMediaList = getFontMediaStatistics();

      //将字媒体统计信息写入excel
      writeFontMediaStatistics(workbook, fontMediaList);

      //将excel写入文件
      System.out.println("==============   正在准备写入excel  ==============");
      workbook.write(fos);
      System.out.println("==============       写入成功       ==============");

      if (monthly) {
        System.out.println("==============   开始处理月度数据   ==============");
        if (startTime.get(Calendar.DAY_OF_MONTH) == 1) {
          startTime.set(startTime.get(Calendar.YEAR), startTime.get(Calendar.MONTH) - 1, 1);
          endTime.set(startTime.get(Calendar.YEAR), endTime.get(Calendar.MONTH), 1);
        } else {
          startTime.set(startTime.get(Calendar.YEAR), startTime.get(Calendar.MONTH), 1);
          endTime.set(startTime.get(Calendar.YEAR), endTime.get(Calendar.MONTH), 1);
        }
        getMonthlySalesStatistics(SalesAppKey, startTime, endTime);

        String fileNameByMonth = JarTool.getJarDir() + "/" + (startTime.get(Calendar.MONTH) + 1) + "月数据统计.xlsx";

        workbook.close();
        workbook = new XSSFWorkbook();

        writeDownloadStatistics(workbook, getDownloadStatistics(DownloadsAppKey, startTime, endTime), DownloadsAppKey, true);

        fos.close();
        fos = new FileOutputStream(fileNameByMonth);


        workbook.write(fos);
      }

      //System.out.println("==============正在准备上传共享文件夹==============");
      //将excel文件上传共享文件夹
      //uploadFile(fileName);
      //System.out.println("==============       上传成功       ==============");

    } catch (IOException e) {
      e.printStackTrace();
    } finally {
      try {
        workbook.close();
        fos.close();
      } catch (IOException e) {
        e.printStackTrace();
      }
    }
  }

  /**
   * 获取需要查询的appKey
   */
  public static void getAppKey(Map<String, String> appKeyMap, String fileName) {
    FileInputStream fis = null;
    InputStreamReader isr = null;
    BufferedReader br = null;
    try {
      fis = new FileInputStream(JarTool.getJarDir() + "/" + fileName);
      isr = new InputStreamReader(fis);
      br = new BufferedReader(isr);
      String appKey;
      while ((appKey = br.readLine()) != null) {
        String[] app = appKey.split("/");
        appKeyMap.put(app[0], app[1]);
      }
    } catch (IOException e) {
      e.printStackTrace();
    } finally {
      try {
        br.close();
        isr.close();
        fis.close();
        // 关闭的时候最好按照先后顺序关闭最后开的先关闭
      } catch (IOException e) {
        e.printStackTrace();
      }
    }
  }

  //上传共享文件夹所需的参数
  private static final String URL = "file://fontfileshare/fontfilesv/手持设备项目/方正字酷/字体云周报/字体云统计数据/";
  private static final String DOMAIN_IP = "172.18.113.3";
  private static final String DOMAIN_NAME = "HOLD";
  private static final String USER_NAME = "cao.zm";
  private static final String USER_PASSWORD = "Founder@2011!";

  /**
   * 文件上传共享文件夹
   */
  private static void uploadFile(String fileName) {
    BufferedInputStream bis = null;
    BufferedOutputStream bos = null;
    try {
      File file = new File(fileName);
      UniAddress dc = UniAddress.getByName(DOMAIN_IP);
      NtlmPasswordAuthentication authentication = new NtlmPasswordAuthentication(DOMAIN_NAME, USER_NAME, USER_PASSWORD);
      SmbSession.logon(dc, authentication);
      SmbFile smbFile = new SmbFile(URL + file.getName(), authentication);
      bis = new BufferedInputStream(new FileInputStream(file));
      bos = new BufferedOutputStream(new SmbFileOutputStream(smbFile));
      byte[] buffer = new byte[1024];
      while (bis.read(buffer) != -1) {
        bos.write(buffer);
        buffer = new byte[1024];
      }
    } catch (IOException e) {
      e.printStackTrace();
    } finally {
      try {
        bos.close();
        bis.close();
      } catch (Exception e) {
        e.printStackTrace();
      }

    }
  }
}
