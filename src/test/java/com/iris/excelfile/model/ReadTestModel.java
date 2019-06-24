package com.iris.excelfile.model;

import com.iris.excelfile.metadata.BaseRowModel;

import java.math.BigDecimal;
import java.util.Date;

/**
 * @author liu_wp
 * @date Created in 2019/3/25 10:42
 * @see
 */
public class ReadTestModel extends BaseRowModel {
    private String a;
    private BigDecimal d;
    private int c;
    private Double b;
    private Integer e;
    private Date date;
    private String phone;

    public String getA() {
        return a;
    }

    public void setA(String a) {
        this.a = a;
    }

    public BigDecimal getD() {
        return d;
    }

    public void setD(BigDecimal d) {
        this.d = d;
    }

    public int getC() {
        return c;
    }

    public void setC(int c) {
        this.c = c;
    }

    public Double getB() {
        return b;
    }

    public void setB(Double b) {
        this.b = b;
    }

    public Integer getE() {
        return e;
    }

    public void setE(Integer e) {
        this.e = e;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }
}
