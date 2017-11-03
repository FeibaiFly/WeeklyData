package founder.sjzt.object;

import java.sql.Timestamp;
import java.util.Date;

public class FontMediaRecord {
  private Date date;
  private Integer number;

  public FontMediaRecord() {
  }

  public FontMediaRecord(Timestamp date, Integer number) {
    this.date = new Date(date.getTime());
    this.number = number;
  }

  public Date getDate() {
    return date;
  }

  public void setDate(Date date) {
    this.date = date;
  }

  public Integer getNumber() {
    return number;
  }

  public void setNumber(Integer number) {
    this.number = number;
  }
}
