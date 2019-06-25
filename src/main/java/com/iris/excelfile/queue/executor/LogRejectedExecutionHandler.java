package com.iris.excelfile.queue.executor;

import com.iris.excelfile.core.builder.IWriteBuilder;
import lombok.extern.slf4j.Slf4j;

import java.util.concurrent.RejectedExecutionException;
import java.util.concurrent.RejectedExecutionHandler;
import java.util.concurrent.ThreadPoolExecutor;

/**
 * @author liu_wp
 * @date Created in 2019/6/25 10:06
 * @see
 */
@Slf4j
public class LogRejectedExecutionHandler implements RejectedExecutionHandler {

    private IWriteBuilder iWriteBuilder;

    public LogRejectedExecutionHandler(IWriteBuilder iWriteBuilder) {
        this.iWriteBuilder = iWriteBuilder;
    }

    @Override
    public void rejectedExecution(Runnable r, ThreadPoolExecutor e) {
        writeLog(r, e);
        throw new RejectedExecutionException("Task " + r.toString() +
                " rejected from " +
                e.toString());
    }

    private void writeLog(Runnable r, ThreadPoolExecutor e) {
        String outFilePath = "";
        if (iWriteBuilder != null) {
            outFilePath = iWriteBuilder.getOutFilePath();
        }
        log.info("导出{}，任务【" + r.toString() + "】执行被拒绝", outFilePath);
    }
}
