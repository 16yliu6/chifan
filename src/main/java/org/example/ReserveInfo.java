package org.example;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Data;

import java.util.Date;
@Data
public class ReserveInfo extends BaseRowModel {
    @ExcelProperty(value = "序号", index = 0)
    private String num;

    @ExcelProperty(value = "微信昵称", index = 1)
    private String nickname;

    @ExcelProperty(value = "姓名", index = 2)
    private String name;

    @ExcelProperty(value = "手机号", index = 3)
    private String phone;

    @ExcelProperty(value = "区域", index = 4)
    private String zone;

    @ExcelProperty(value = "报名状态", index = 5)
    private String status;

    @ExcelProperty(value = "是否核销", index = 6)
    private String writeOff;

    @ExcelProperty(value = "备注说明", index = 7)
    private String remark;

    @ExcelProperty(value = "报名时间", index = 8)
    private Date time;
}
