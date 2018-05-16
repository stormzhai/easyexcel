package excel;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

import java.util.Date;

/**
 * Created by jipengfei on 17/3/15.
 *
 * @author jipengfei
 * @date 2017/03/15
 */
public class ExcelRowJavaModel extends BaseRowModel {
    @ExcelProperty(index = 0,value = "银行放款编号")
    private int num;

    @ExcelProperty(index = 1,value = "code")
    private Long code;

    @ExcelProperty(index = 2,value = "银行存放期期",format = "yyyy-MM-dd")
    private Date endTime;

    @ExcelProperty(index = 3,value = "测试1")
    private Double money;

    @ExcelProperty(index = 4,value = "测试2")
    private String times;

    @ExcelProperty(index = 5,value = "测试3")
    private int activityCode;

    @ExcelProperty(index = 6,value = "测试4")
    private Date date;

    @ExcelProperty(index = 7,value = "测试5")
    private Double lx;

    @ExcelProperty(index = 8,value = "测试6")

    private String name;
    @ExcelProperty(index = 9,value = "性别",replace = "男_1,女_2")
    private int sex;

    public int getNum() {
        return num;
    }

    public void setNum(int num) {
        this.num = num;
    }

    public Long getCode() {
        return code;
    }

    public void setCode(Long code) {
        this.code = code;
    }

    public Date getEndTime() {
        return endTime;
    }

    public void setEndTime(Date endTime) {
        this.endTime = endTime;
    }

    public Double getMoney() {
        return money;
    }

    public void setMoney(Double money) {
        this.money = money;
    }

    public String getTimes() {
        return times;
    }

    public void setTimes(String times) {
        this.times = times;
    }

    public int getActivityCode() {
        return activityCode;
    }

    public void setActivityCode(int activityCode) {
        this.activityCode = activityCode;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public Double getLx() {
        return lx;
    }

    public void setLx(Double lx) {
        this.lx = lx;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getSex() {
        return sex;
    }

    public void setSex(int sex) {
        this.sex = sex;
    }

    @Override
    public String toString() {
        return "ExcelRowJavaModel{" +
                "num=" + num +
                ", code=" + code +
                ", endTime=" + endTime +
                ", money=" + money +
                ", times='" + times + '\'' +
                ", activityCode=" + activityCode +
                ", date=" + date +
                ", lx=" + lx +
                ", name='" + name + '\'' +
                ", sex=" + sex +
                '}';
    }
}
