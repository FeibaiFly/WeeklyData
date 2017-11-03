package founder.sjzt.object;

import java.sql.Timestamp;
import java.util.Date;

public class PurchaseRecord {
  private String appName;
  private String userId;
  private String fontName;
  private Integer price;
  private Date date;

  public PurchaseRecord() {
    super();
  }

  public PurchaseRecord(String appName, String userId, String fontName, Integer price, Timestamp date) {
    super();
    this.appName = appName;
    this.userId = userId;
    this.fontName = fontName;
    this.price = price;
    this.date = new Date(date.getTime());
  }

  public String getAppName() {
    return appName;
  }

  public void setAppName(String appName) {
    this.appName = appName;
  }

  public String getUserId() {
    return userId;
  }

  public void setUserId(String userId) {
    this.userId = userId;
  }

  public String getFontName() {
    return fontName;
  }

  public void setFontName(String fontName) {
    this.fontName = fontName;
  }

  public Integer getPrice() {
    return price;
  }

  public void setPrice(Integer price) {
    this.price = price;
  }

  public Date getDate() {
    return date;
  }

  public void setDate(Date date) {
    this.date = date;
  }
}
