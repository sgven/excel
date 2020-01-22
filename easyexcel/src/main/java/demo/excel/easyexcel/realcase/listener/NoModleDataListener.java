package demo.excel.easyexcel.realcase.listener;

import com.alibaba.excel.context.AnalysisContext;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 异步处理监听器
 */
public class NoModleDataListener extends AbstractReadListener<Map<Integer, String>> {
    private static final Logger LOGGER = LoggerFactory.getLogger(NoModleDataListener.class);

    private int skipRowCount;// 除第一个文件外，后合并的文件需要跳过的行数
    private int currentRowIndex=0;// 记录每个文件读取时，当前所在行数，如1，表示当前读取到第一行

    public NoModleDataListener(int skipRowCount) {
        this.skipRowCount = skipRowCount;
    }

    List<Map<Integer, String>> list = new ArrayList<Map<Integer, String>>();

    @Override
    public void invoke(Map<Integer, String> data, AnalysisContext context) {
//        LOGGER.info("解析到一条数据:{}", JSON.toJSONString(data));
        // 除第一个excel，后面合并进来的excel需要跳过表头
        currentRowIndex++;
        if (currentRowIndex <= skipRowCount) {
            return;
        }

        list.add(data);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        LOGGER.info("所有数据解析完成！");
    }

    @Override
    public List getDatas() {
        return list;
    }

}
