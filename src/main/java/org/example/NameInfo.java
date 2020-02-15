package org.example;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class NameInfo extends BaseRowModel {
    @ExcelProperty(value = "序号", index = 0)
    private String num;

    @ExcelProperty(value = "姓名", index = 2)
    private String name;
}
