package demo.excel.easyexcel.realcase.handler;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.handler.CellWriteHandler;
import demo.excel.easyexcel.realcase.param.EasyExcelProperty;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * 默认表头的处理
 */
public class DefaultHeadHandler {

    private CellWriteHandler styleStrategy = new CustomCellWriteHandler();

    public void handle(ExcelWriterBuilder excelWriterBuilder, EasyExcelProperty param) {
        // 默认使用 “第一行：大标题titleName + 第二行：注解的列标题” 作为表头
        String titleName = StringUtils.isBlank(param.getTitleName()) ? param.getFileName() : param.getTitleName();
        List<List<String>> head = getDefaultDynamicHead(titleName, param.getClazz());
        // 设置动态表头
        excelWriterBuilder.head(head);
        // 设置样式，使用拦截器来自定义样式
        excelWriterBuilder.useDefaultStyle(false);
        excelWriterBuilder.registerWriteHandler(styleStrategy);
    }

    private List<List<String>> getDefaultDynamicHead(String titleName, Class clazz) {
        List<List<String>> head = new ArrayList();
        Field[] fields = clazz.getDeclaredFields();
        if (fields != null && fields.length > 0) {
            List<ExcelProperty> headAnnotations = new ArrayList<>();
            Arrays.stream(fields).forEach(field -> {
                ExcelProperty columnProperty = field.getAnnotation(ExcelProperty.class);
                /*
                // 使用动态表头后，easyExcel的ExcelProperty中的index属性、ExcelIgnore等注解无效，建议：
                // 如果要排序，则直接改变类中的字段顺序即可
                // 如果要隐藏、不显示某一列，则直接在类中删除这个字段
                ExcelIgnore columnIgnore = field.getAnnotation(ExcelIgnore.class);
                if (columnIgnore == null && columnProperty != null) {
                    headAnnotations.add(columnProperty);
                }*/
                headAnnotations.add(columnProperty);
            });
            /*// 表头根据index升序排序
            Comparator<ExcelProperty> comparator = new Comparator<ExcelProperty>() {
                @Override
                public int compare(ExcelProperty o1, ExcelProperty o2) {
                    return (o1.index() < o2.index()) ? -1 : ((o1.index() == o2.index()) ? 0 : 1);
                }
            };
            Collections.sort(headAnnotations, comparator);*/
            // 添加动态表头
            List<String> columnHead = null;
            for (ExcelProperty column : headAnnotations) {
                columnHead = new ArrayList<>();
                columnHead.add(titleName);
                columnHead.add(column.value()[0]);
                head.add(columnHead);
            }
        }
        return head;
    }
}
