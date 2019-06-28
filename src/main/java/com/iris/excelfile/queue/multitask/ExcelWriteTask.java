package com.iris.excelfile.queue.multitask;

import com.iris.excelfile.constant.FileConstant;
import com.iris.excelfile.core.builder.IWriteBuilder;
import com.iris.excelfile.metadata.ExcelTable;
import com.iris.excelfile.response.BaseResponse;
import lombok.extern.slf4j.Slf4j;

import java.util.concurrent.Callable;
import java.util.concurrent.ConcurrentLinkedQueue;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.locks.ReentrantLock;

/**
 * @author liu_wp
 * @date Created in 2019/6/25 9:13
 * @see
 */
@Slf4j
public class ExcelWriteTask implements Callable<BaseResponse> {
    private ConcurrentLinkedQueue<ExcelTable> queue;
    private CountDownLatch countDownLatch;
    private IWriteBuilder iWriteBuilder;
    private BaseResponse baseResponse;

    public ExcelWriteTask(BaseResponse baseResponse, ConcurrentLinkedQueue<ExcelTable> queue, CountDownLatch countDownLatch, IWriteBuilder iWriteBuilder) {
        this.baseResponse = baseResponse;
        this.queue = queue;
        this.countDownLatch = countDownLatch;
        this.iWriteBuilder = iWriteBuilder;
    }

    @Override
    public BaseResponse call() {
        long l = System.currentTimeMillis();
        log.info("当前执行线程：{}，开始时间：{}", Thread.currentThread().getName(), l);
        try {
            while (!queue.isEmpty()) {
                iWriteBuilder.addContent(queue.poll());
            }
            log.info("当前执行线程：{}，耗时：{}秒", Thread.currentThread().getName(), (System.currentTimeMillis() - l) / 1000);
            this.countDownLatch.countDown();
            baseResponse.setCode(FileConstant.SUCCESS_CODE);
        } catch (Exception e) {
            e.printStackTrace();
            baseResponse.setCode(FileConstant.FAIL_CODE);
        }
        return baseResponse;
    }
}
