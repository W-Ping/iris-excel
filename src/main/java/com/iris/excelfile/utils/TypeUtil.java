package com.iris.excelfile.utils;

import com.iris.excelfile.exception.ExcelParseException;
import jdk.nashorn.internal.runtime.ParserException;
import lombok.extern.slf4j.Slf4j;
import net.sf.cglib.beans.BeanMap;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author liu_wp
 * @date Created in 2019/3/8 10:29
 * @see
 */
@Slf4j
public class TypeUtil {

    private static List<String> DATE_FORMAT_LIST = new ArrayList<>(4);
    public static final Pattern pattern = Pattern.compile("[\\+\\-]?[\\d]+([\\.][\\d]*)?([Ee][+-]?[\\d]+)?$");

    static {
        DATE_FORMAT_LIST.add("yyyy/MM/dd HH:mm:ss");
        DATE_FORMAT_LIST.add("yyyy-MM-dd HH:mm:ss");
        DATE_FORMAT_LIST.add("yyyyMMdd HH:mm:ss");
    }

    private static int getCountOfChar(String value, char c) {
        int n = 0;
        if (value == null) {
            return 0;
        }
        char[] chars = value.toCharArray();
        for (char cc : chars) {
            if (cc == c) {
                n++;
            }
        }
        return n;
    }

    public static Object convert(String value, Field field, String format, boolean us) {
        if (!StringUtils.isEmpty(value)) {
            if (Float.class.equals(field.getType())) {
                return Float.parseFloat(value);
            }
            if (Integer.class.equals(field.getType()) || int.class.equals(field.getType())) {
                return Integer.parseInt(value);
            }
            if (Double.class.equals(field.getType()) || double.class.equals(field.getType())) {
                if (null != format && !"".equals(format)) {
                    int n = getCountOfChar(value, '0');
                    return Double.parseDouble(TypeUtil.formatFloat0(value, n));
                } else {
                    return Double.parseDouble(TypeUtil.formatFloat(value));
                }
            }
            if (Boolean.class.equals(field.getType()) || boolean.class.equals(field.getType())) {
                String valueLower = value.toLowerCase();
                if (valueLower.equals("true") || valueLower.equals("false")) {
                    return Boolean.parseBoolean(value.toLowerCase());
                }
                Integer integer = Integer.parseInt(value);
                if (integer == 0) {
                    return false;
                } else {
                    return true;
                }
            }
            if (Long.class.equals(field.getType()) || long.class.equals(field.getType())) {
                return Long.parseLong(value);
            }
            if (Date.class.equals(field.getType())) {
                if (value.contains("-") || value.contains("/") || value.contains(":")) {
                    return getSimpleDateFormatDate(value, format);
                } else {
                    Double d = Double.parseDouble(value);
                    return HSSFDateUtil.getJavaDate(d, us);
                }
            }
            if (BigDecimal.class.equals(field.getType())) {
                return new BigDecimal(value);
            }
            if (String.class.equals(field.getType())) {
                return formatFloat(value);
            }

        }
        return null;
    }

    /**
     * @param field
     * @return
     */
    public static boolean isIntNum(Field field) {
        if (Integer.class.equals(field.getType()) || int.class.equals(field.getType())) {
            return true;
        }
        return false;
    }

    public static boolean isNum(Field field) {
        if (field == null) {
            return false;
        }
        if (Integer.class.equals(field.getType()) || int.class.equals(field.getType())) {
            return true;
        }
        if (Double.class.equals(field.getType()) || double.class.equals(field.getType())) {
            return true;
        }
        if (Float.class.equals(field.getType()) || float.class.equals(field.getType())) {
            return true;
        }
        if (Long.class.equals(field.getType()) || long.class.equals(field.getType())) {
            return true;
        }
        if (BigDecimal.class.equals(field.getType())) {
            return true;
        }
        return false;
    }

    public static boolean isNum(Object cellValue) {
        if (cellValue instanceof Integer
                || cellValue instanceof Double
                || cellValue instanceof Short
                || cellValue instanceof Long
                || cellValue instanceof Float
                || cellValue instanceof BigDecimal) {
            return true;
        }
        return false;
    }

    /**
     * @param date
     * @return
     */
    public static String getDefaultDateString(Date date) {
        Instant instant = date.toInstant();
        ZoneId zoneId = ZoneId.systemDefault();
        LocalDateTime localDateTime = instant.atZone(zoneId).toLocalDateTime();
        return localDateTime.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
    }

    /**
     * @param value
     * @param format
     * @return
     */
    public static Date getSimpleDateFormatDate(String value, String format) {
        if (!StringUtils.isEmpty(value)) {
            Date date = null;
            if (!StringUtils.isEmpty(format)) {
                SimpleDateFormat simpleDateFormat = new SimpleDateFormat(format);
                try {
                    date = simpleDateFormat.parse(value);
                    return date;
                } catch (ParseException e) {
                }
            }
            for (String dateFormat : DATE_FORMAT_LIST) {
                try {
                    SimpleDateFormat simpleDateFormat = new SimpleDateFormat(dateFormat);
                    date = simpleDateFormat.parse(value);
                } catch (ParseException e) {
                }
                if (date != null) {
                    break;
                }
            }
            return date;

        }
        return null;

    }


    /**
     * @param value
     * @return
     */
    public static String formatFloat(String value) {
        if (null != value && value.contains(".")) {
            if (isNumeric(value)) {
                try {
                    BigDecimal decimal = new BigDecimal(value);
                    BigDecimal setScale = decimal.setScale(10, BigDecimal.ROUND_HALF_DOWN).stripTrailingZeros();
                    return setScale.toPlainString();
                } catch (Exception e) {
                    log.error("{} string cast to float error", value);
                    throw new ExcelParseException(value + " string cast to float error");
                }
            }
        }
        return value;
    }

    public static String formatFloat0(String value, int n) {
        if (null != value && value.contains(".")) {
            if (isNumeric(value)) {
                try {
                    BigDecimal decimal = new BigDecimal(value);
                    BigDecimal setScale = decimal.setScale(n, BigDecimal.ROUND_HALF_DOWN);
                    return setScale.toPlainString();
                } catch (Exception e) {
                    log.error("{} string cast to float error", value);
                    throw new ExcelParseException(value + " string cast to float error");
                }
            }
        }
        return value;
    }


    private static boolean isNumeric(String str) {
        Matcher isNum = pattern.matcher(str);
        if (!isNum.matches()) {
            return false;
        }
        return true;
    }


    public static String formatDate(Date cellValue, String format) {
        try {
            Instant instant = cellValue.toInstant();
            ZoneId zone = ZoneId.systemDefault();
            LocalDateTime localDateTime = LocalDateTime.ofInstant(instant, zone);
            return formatLocalDateTime(localDateTime, format);
        } catch (Exception e) {
            log.error("{} date format error", cellValue);
            throw new ExcelParseException(cellValue + " date format error");
        }
    }

    public static String formatDate(String cellValue, String format) {
        try {
            return formatLocalDateTime(dateStrToLocalDateTime(cellValue), format);
        } catch (Exception e) {
            log.error("{} date format error", cellValue);
            throw new ExcelParseException(cellValue + " date format error");
        }
    }

    public static LocalDateTime dateStrToLocalDateTime(String dateStr) {
        Timestamp timestamp = Timestamp.valueOf(dateStr);
        return timestamp.toLocalDateTime();
    }

    public static String formatLocalDateTime(LocalDateTime localDateTime, String format) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(format);
        return localDateTime.format(formatter);
    }

    public static String getFieldStringValue(BeanMap beanMap, String fieldName, String format) {
        String cellValue = null;
        Object value = beanMap.get(fieldName);
        if (value != null) {
            if (value instanceof Date) {
                cellValue = TypeUtil.formatDate((Date) value, format);
            } else if (value instanceof String && StringUtils.isNotBlank(format)) {
                cellValue = TypeUtil.formatDate(value.toString(), format);
            } else {
                cellValue = value.toString();
            }
        }
        return cellValue;
    }

}
