package com.iris.excelfile.utils;


import com.iris.excelfile.constant.DateTypeEnum;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

/**
 * @author liu_wp
 * @date Created in 2019/3/25 11:51
 * @see
 */
public class DateUtil {
    /**
     * @param dateStr
     * @param dateFormat
     * @return
     */
    public static String formatDate(String dateStr, String dateFormat) {
        Date date = covertStrToDate(dateStr);
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(dateFormat, Locale.SIMPLIFIED_CHINESE);
        return simpleDateFormat.format(date);
    }

    /**
     * @param date
     * @param dateFormat
     * @return
     */
    public static String formatDate(Date date, String dateFormat) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(dateFormat, Locale.SIMPLIFIED_CHINESE);
        return simpleDateFormat.format(date);
    }

    /**
     * @param dateStr
     * @return
     */
    public static Date covertStrToDate(String dateStr) {
        Date date = null;
        for (DateTypeEnum dateTypeEnum : DateTypeEnum.values()) {
            try {
                SimpleDateFormat simpleDateFormat = new SimpleDateFormat(dateTypeEnum.getFormat());
                date = simpleDateFormat.parse(dateStr);
                if (date != null) {
                    return date;
                }
            } catch (Exception e) {
            }
        }
        return date;
    }

}
