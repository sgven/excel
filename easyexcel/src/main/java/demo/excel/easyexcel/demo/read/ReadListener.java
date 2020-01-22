package demo.excel.easyexcel.demo.read;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import demo.excel.easyexcel.realcase.listener.AbstractReadListener;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

//public class ReadListener extends AnalysisEventListener<ReadDemoData> {
public class ReadListener extends AbstractReadListener<ReadDemoData> {
    private static final Logger LOGGER = LoggerFactory.getLogger(ReadListener.class);
    /**
     * 每隔5条存储数据库，实际使用中可以3000条，然后清理list ，方便内存回收
     */
    private static final int BATCH_COUNT = 5;
    List<ReadDemoData> list = new ArrayList<>();

    /**
     * 解析的过程，就读取出来，观察者模式一行一行读取
     * @param data
     * @param context
     */
    @Override
    public void invoke(ReadDemoData data, AnalysisContext context) {
        list.add(data);
        if (list.size() >= BATCH_COUNT) {
            saveData();
//            list.clear();
        }
    }

    /**
     * 全部解析完的回调
     * @param context
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        saveData();
    }

    @Override
    public List getDatas() {
        return list;
    }

    /**
     * 加上存储数据库
     */
    private void saveData() {
        LOGGER.info("{}条数据，开始存储数据库！", list.size());
        LOGGER.info("存储数据库成功！");
    }
}
