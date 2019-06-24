package com.iris.excelfile.constant;

/**
 * @author liu_wp
 * @date Created in 2019/3/1 11:30
 * @see
 */
public class FileConstant {
    /**
     * 默认字体
     */
    public static String DEFAULT_FONT_NAME = "CorpoS";
    /**
     * 默认列宽
     */
    public static int DEFAULT_COLUMN_WIDTH = 10;
    /**
     * 默认内容字体大小
     */
    public static short DEFAULT_CONTENT_FONT_SIZE = 10;
    /**
     * 默认表头字体大小
     */
    public static short DEFAULT_HEAD_FONT_SIZE = 12;
    /**
     * 默认表尾字体大小（暂没有用）
     */
    public static short DEFAULT_FOOT_FONT_SIZE = 11;
    /**
     * 默认大于1000条数 多线程处理
     */
    public static int QUEUE_MAX_SIZE = 1000;
    /**
     * 最大线程数
     */
    public static int MAX_THREAD_COUNT = 5;
    /**
     * 线程数量计算因子 表数量 * 计算因子
     */
    public static float DEFAULT_QUEUE_FACTOR = 0.5f;

    /**
     * 默认保护密码
     */
    public static String PROTECT_SHEET_PASSWORD = "123";
    /**
     * 成功
     */
    public static String SUCCESS_CODE = "success";
    /**
     * 失败
     */
    public static String FAIL_CODE = "fail";
}
