package com.iris.excelfile.core.handler.extend;

import java.util.Objects;
import java.util.function.BiFunction;
import java.util.function.Function;

/**
 * @author lwp
 */
public abstract class AbstractDictionaryRefHandler implements BiFunction<Integer, Integer, String> {

    @Override
    public String apply(Integer t, Integer u) {
        return getDicValue(t, u);
    }

    @Override
    public <V> BiFunction<Integer, Integer, V> andThen(Function<? super String, ? extends V> after) {
        Objects.requireNonNull(after);
        return (Integer t, Integer u) -> after.apply(apply(t, u));
    }

    /**
     * 获取数据字典值
     *
     * @param cellIndex
     * @param code
     * @return
     */
    protected abstract String getDicValue(Integer cellIndex, Integer code);
}
