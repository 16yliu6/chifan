package org.example;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Data;

import java.util.Date;
@Data
public class ReserveInfo extends BaseRowModel {
    @ExcelProperty(value = {"序号"})
    private String num;

    @ExcelProperty(value = {"微信昵称"})
    private String nickname;

    @ExcelProperty(value = {"姓名"})
    private String name;

    @ExcelProperty(value = {"手机号"})
    private String phone;

    @ExcelProperty(value = {"区域"})
    private String zone;

    @ExcelProperty(value = {"报名状态"})
    private String status;

    @ExcelProperty(value = {"是否核销"})
    private boolean writeOff;

    @ExcelProperty(value = {"备注说明"})
    private String remark;

    @ExcelProperty(value = {"报名时间"})
    private Date time;
}
