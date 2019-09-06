package com.iris.excelfile.model;

import com.iris.excelfile.annotation.ExcelReadProperty;
import com.iris.excelfile.annotation.ExcelReadSheet;

import java.math.BigDecimal;
import java.util.Date;

/**
 * @author liu_wp
 * @date 2019/9/6 15:38
 * @see
 */
@ExcelReadSheet
public class ReadExcelModel1 {
    @ExcelReadProperty(cellIndex = 0)
    private String userName;
    @ExcelReadProperty(cellIndex = 1)
    private String phone;
    @ExcelReadProperty(cellIndex = 2)
    private BigDecimal money;
    @ExcelReadProperty(cellIndex = 3, dateFormat = "YYYY/MM/dd HH:mm:ss")
    private Date createDate;

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public BigDecimal getMoney() {
        return money;
    }

    public void setMoney(BigDecimal money) {
        this.money = money;
    }

    public Date getCreateDate() {
        return createDate;
    }

    public void setCreateDate(Date createDate) {
        this.createDate = createDate;
    }
}
