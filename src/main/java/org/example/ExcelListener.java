package org.example;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import lombok.Data;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Data
public class ExcelListener extends AnalysisEventListener {

    private List<ReserveInfo> haiGuan  = new ArrayList<>();
    @Override
    public void invoke(Object object, AnalysisContext context) {
        System.out.println("当前行：" + context.getCurrentRowNum());
        ReserveInfo reserveInfo = (ReserveInfo) object;
        String zone = reserveInfo.getZone();
        switch(zone) {
            case "海关":
                haiGuan.add(reserveInfo);
            default:
                break;
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
    }

}
