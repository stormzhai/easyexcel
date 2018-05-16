package excel;

import com.alibaba.excel.read.context.AnalysisContext;
import com.alibaba.excel.read.event.AnalysisEventListener;

/**
 * Created by daojia on 2018-5-14.
 */
public class ExcelListener extends AnalysisEventListener {

    @Override
    public void invoke(Object object, AnalysisContext context) {
        System.out.println("sheet:" + context.getCurrentSheet().getSheetNo() + ",row："
                + context.getCurrentRowNum() + ",data:" + object);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

    }


}
