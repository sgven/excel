package demo.excel.easyexcel.realcase.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.List;

public abstract class AbstractReadListener<T> extends AnalysisEventListener<T> {

    @Override
    public void invoke(T data, AnalysisContext context) {

    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

    }

    public abstract List getDatas();
}
