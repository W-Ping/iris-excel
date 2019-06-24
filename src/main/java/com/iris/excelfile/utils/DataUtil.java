package com.iris.excelfile.utils;

import java.util.Arrays;

/**
 * @author liu_wp
 * @date Created in 2019/3/6 19:35
 * @see
 */
public class DataUtil {
    /**
     * 合并数组
     *
     * @param first
     * @param rest
     * @param <T>
     * @return
     */
    public static <T> T[] concatAll(T[] first, T[]... rest) {
        if (rest == null || rest.length == 0) {
            return first;
        }
        int totalLength = first.length;
        for (T[] array : rest) {
            totalLength += array.length;
        }
        T[] result = Arrays.copyOf(first, totalLength);
        int offset = first.length;
        for (T[] array : rest) {
            System.arraycopy(array, 0, result, offset, array.length);
            offset += array.length;
        }
        return result;
    }

    /**
     * 合并数组
     *
     * @param first
     * @param rest
     * @param <T>
     * @return
     */
    public static <T> T[] concatAll(T[] first, T... rest) {
        if (rest == null || rest.length == 0) {
            return first;
        }
        T[] result = Arrays.copyOf(first, rest.length + first.length);
        int offset = first.length;
        System.arraycopy(rest, 0, result, offset, rest.length);
        return result;
    }
}
