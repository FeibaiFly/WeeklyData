package founder.sjzt.object;

import java.sql.Timestamp;
import java.util.Date;

public class PurchaseData {
  private Date date;
  private Integer number;


  public PurchaseData() {
    super();
  }

  public PurchaseData(Timestamp data, int number) {
    super();
    this.date = new Date(data.getTime());
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
