package founder.sjzt.object;

import java.sql.Timestamp;
import java.util.Date;

public class DownloadRecord {
  private String fontName;
  private Date date;
  private Integer downloads;

  public DownloadRecord() {
    super();
  }

  public DownloadRecord(String fontName, Integer downloads) {
    super();
    this.fontName = fontName;
    this.downloads = downloads;
  }

  public DownloadRecord(Timestamp date, Integer downloads) {
    super();
    this.downloads = downloads;
    this.date = new Date(date.getTime());
  }

  public String getFontName() {
    return fontName;
  }

  public void setFontName(String fontName) {
    this.fontName = fontName;
  }

  public Date getDate() {
    return date;
  }

  public void setDate(Date date) {
    this.date = date;
  }

  public Integer getDownloads() {
    return downloads;
  }

  public void setDownloads(Integer downloads) {
    this.downloads = downloads;
  }
}
