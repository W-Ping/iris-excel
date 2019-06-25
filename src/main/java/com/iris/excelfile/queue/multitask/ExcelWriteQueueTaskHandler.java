package com.iris.excelfile.queue.multitask;

import com.iris.excelfile.constant.FileConstant;
import com.iris.excelfile.core.builder.IWriteBuilder;
import com.iris.excelfile.exception.ExcelParseException;
import com.iris.excelfile.metadata.ExcelTable;
import com.iris.excelfile.queue.IExcelWriteQueueTaskHandler;
import com.iris.excelfile.queue.executor.LogRejectedExecutionHandler;
import com.iris.excelfile.queue.executor.NameThreadFactory;
import com.iris.excelfile.response.BaseResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections.CollectionUtils;

import java.util.List;
import java.util.concurrent.*;

/**
 * @author liu_wp
 * @date Created in 2019/6/25 9:13
 * @see
 */
@Slf4j
public class ExcelWriteQueueTaskHandler implements IExcelWriteQueueTaskHandler<ExcelTable> {
    private static final ConcurrentLinkedQueue<ExcelTable> EXCELTABLE_QUEUE = new ConcurrentLinkedQueue();
    private ThreadPoolExecutor threadPoolExecutor;
    private CountDownLatch countDownLatch;
    private int threadMaxCount;
    private IWriteBuilder iWriteBuilder;

    public ExcelWriteQueueTaskHandler(int tableSize, IWriteBuilder iWriteBuilder) {
        this.threadMaxCount = calculateTaskCount(tableSize);
        this.iWriteBuilder = iWriteBuilder;
        this.countDownLatch = new CountDownLatch(this.threadMaxCount);
        //线程池初始化
        threadPoolExecutor = new ThreadPoolExecutor(this.threadMaxCount + FileConstant.MAX_THREAD_COUNT, this.threadMaxCount + FileConstant.MAX_THREAD_COUNT * 2, 10, TimeUnit.SECONDS,
                new LinkedBlockingQueue<>(10), new NameThreadFactory(FileConstant.EXPORT_THREAD_PREFIX_NAME), new LogRejectedExecutionHandler(iWriteBuilder));
    }

    /**
     * 计算最大线程数
     *
     * @param tableCount
     * @return
     */
    public static int calculateTaskCount(int tableCount) {
        int floor = (int) Math.floor(tableCount * FileConstant.DEFAULT_QUEUE_FACTOR);
        return floor == 0 ? 1 : (floor < FileConstant.MAX_THREAD_COUNT ? floor : FileConstant.MAX_THREAD_COUNT);
    }

    @Override
    public boolean putQueue(ExcelTable table) {
        if (table != null) {
            return EXCELTABLE_QUEUE.offer(table);
        }
        return true;
    }

    @Override
    public boolean putQueue(List<ExcelTable> list) {
        if (CollectionUtils.isEmpty(list)) {
            return EXCELTABLE_QUEUE.addAll(list);
        }
        return true;
    }

    @Override
    public void consumeQueue() {
        try {
            long start = System.currentTimeMillis();
            log.info("开始执行多线程导出{}", start);
            BaseResponse response = new BaseResponse();
            for (int i = 0; i < threadMaxCount; i++) {
                ExcelWriteTask excelWriteTask = new ExcelWriteTask(response, EXCELTABLE_QUEUE, countDownLatch, iWriteBuilder);
                FutureTask<BaseResponse> futureTask = new FutureTask<>(excelWriteTask);
                threadPoolExecutor.submit(futureTask, response);
            }
            countDownLatch.await(FileConstant.QUEUE_AWAIT_TIME, TimeUnit.SECONDS);
            long time = (System.currentTimeMillis() - start) / 1000;
            log.info("结束执行多线程导出{}，多线程总耗时：{}秒", System.currentTimeMillis(), time);
            threadPoolExecutor.shutdown();
            if (time >= FileConstant.QUEUE_AWAIT_TIME) {
                throw new ExcelParseException("多线程导出文档超时");
            }
        } catch (InterruptedException e) {
            e.printStackTrace();
            throw new ExcelParseException("多线程导出文档失败", e);
        }
    }
}
