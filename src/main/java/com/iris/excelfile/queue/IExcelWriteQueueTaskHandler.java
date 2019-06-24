package com.iris.excelfile.queue;

import java.util.List;

/**
 * @param <T>
 * @author lwp
 */
public interface IExcelWriteQueueTaskHandler<T> {

    /**
     * 加入队列
     *
     * @param t
     * @return
     */
    boolean putQueue(T t);

    /**
     * 批量加入队列
     *
     * @param list
     * @return
     */
    boolean putQueue(List<T> list);

    /**
     * 消费队列信息
     */
    void consumeQueue();
}
