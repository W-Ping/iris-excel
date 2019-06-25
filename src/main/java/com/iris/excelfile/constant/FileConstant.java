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
     * 导出多线程名称前缀 前缀唯一性避免被误杀
     */
    public static String EXPORT_THREAD_PREFIX_NAME = "IRIS-EXPORT-EXCEL-";
    /**
     * 默认数据量最小值-触发多线程处理
     */
    public static int QUEUE_MIN_DATA_SIZE = 1000;
    /**
     * 默认最小队列数-触发多线程处理
     */
    public static int QUEUE_MIN_SIZE = 5;
    /**
     * 多线程处理等待最大时间（秒）
     */
    public static long QUEUE_AWAIT_TIME = 300;
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
