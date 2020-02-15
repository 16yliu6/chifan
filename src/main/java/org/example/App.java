package org.example;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;

import java.io.*;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Hello world!
 *
 */
public class App 
{
    private List<ReserveInfo> haiGuan;
    public static void main( String[] args ) throws FileNotFoundException {
        System.out.println( "开始!" );
        String readPath = "/Users/liuyu/Downloads/data.xls";
        String name = LocalDateTime.now().getMonthValue() + "-" + LocalDateTime.now().getDayOfMonth() + "-";
        if (LocalDateTime.now().getHour() < 12) {
            name = name + "午餐";
        }  else {
            name = name + "晚餐";
        }
        String writepath = "/Users/liuyu/Desktop/chifan/"+ name +".xls";
        List<Object> objects = ExcelUtil.readLessThan1000RowBySheet(readPath, null);

        ArrayList<ReserveInfo> reserveInfos = new ArrayList<>();
        objects.stream().forEach(it -> reserveInfos.add((ReserveInfo)it));
        List<ReserveInfo> haiGuan = reserveInfos.stream().filter(it -> "海关".equals(it.getZone())).collect(Collectors.toList());
        List<ReserveInfo> xiangFu = reserveInfos.stream().filter(it -> "翔福".equals(it.getZone())).collect(Collectors.toList());
        List<ReserveInfo> anXin = reserveInfos.stream().filter(it -> "安歆".equals(it.getZone())).collect(Collectors.toList());
        List<ReserveInfo> renCai = reserveInfos.stream().filter(it -> "人才公寓".equals(it.getZone())).collect(Collectors.toList());
        List<ReserveInfo> YHW = reserveInfos.stream().filter(it -> "悦海湾".equals(it.getZone())).collect(Collectors.toList());
        List<ReserveInfo> YJY = reserveInfos.stream().filter(it -> "研究院".equals(it.getZone())).collect(Collectors.toList());

        System.out.println(objects);

        ArrayList<ExcelUtil.MultipleSheelPropety> multipleSheelPropeties = new ArrayList<>();
        multipleSheelPropeties.add(new ExcelUtil.MultipleSheelPropety(reserveInfos, new Sheet(1, 1, ReserveInfo.class, "总表(" + reserveInfos.size() + ")", null)));
        multipleSheelPropeties.add(new ExcelUtil.MultipleSheelPropety(haiGuan, new Sheet(2, 1, ReserveInfo.class, "海关(" + haiGuan.size() + ")", null)));
        multipleSheelPropeties.add(new ExcelUtil.MultipleSheelPropety(xiangFu, new Sheet(3, 1, ReserveInfo.class, "翔福(" + xiangFu.size() + ")", null)));
        multipleSheelPropeties.add(new ExcelUtil.MultipleSheelPropety(xiangFu, new Sheet(4, 1, ReserveInfo.class, "安歆(" + anXin.size() + ")", null)));
        multipleSheelPropeties.add(new ExcelUtil.MultipleSheelPropety(xiangFu, new Sheet(5, 1, ReserveInfo.class, "人才公寓(" + renCai.size() + ")", null)));
        multipleSheelPropeties.add(new ExcelUtil.MultipleSheelPropety(xiangFu, new Sheet(6, 1, ReserveInfo.class, "悦海湾(" + YHW.size() + ")", null)));
        multipleSheelPropeties.add(new ExcelUtil.MultipleSheelPropety(xiangFu, new Sheet(7, 1, ReserveInfo.class, "研究院(" + YJY.size() + ")", null)));
        ExcelUtil.writeWithMultipleSheel(writepath, multipleSheelPropeties);

        System.out.println("处理完毕");

        new File(readPath).delete();
        System.out.println("删除data.xls");


    }

}
