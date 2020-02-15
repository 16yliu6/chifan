package org.example;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.github.Dorae132.easyutil.easyexcel.ExcelUtils;

import java.io.*;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws FileNotFoundException {
        System.out.println( "Hello World!" );
        if(args.length == 0) {
            System.out.println("请输入文件名");
            System.exit(1);
        }
//        String path = args[0];
        String path = "C:\\Users\\liuy655\\Downloads\\data.xls";
        testExcel2003NoModel(path);
        writeExcel(path);
        System.out.println("处理完毕");
    }
    public static void testExcel2003NoModel(String fileName) throws FileNotFoundException {
        File file = new File(fileName);
        InputStream fileInputStream = new FileInputStream(file.getPath());
        try {
            // 解析每行结果在listener中处理
            ExcelListener listener = new ExcelListener();

            ExcelReader excelReader = new ExcelReader(fileInputStream, ExcelTypeEnum.XLS, null, listener, true);
            excelReader.read(new Sheet(1, 2, ReserveInfo.class));
        } catch (Exception e) {

        } finally {
            try {
                fileInputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    public static void writeExcel(String fileName) throws FileNotFoundException {
        OutputStream out = new FileOutputStream(fileName);
        try {
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX, true);
            ExcelListener excelListener = new ExcelListener();
            //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
            Sheet sheet1 = new Sheet(2, 0, ReserveInfo.class, "海关（" +excelListener.getHaiGuan().size() + ")", null);
//            sheet1.setSheetName("海关（" +excelListener.getHaiGuan().size() + ")");
            writer.write(excelListener.getHaiGuan(), sheet1);
            System.out.println("海关写入完毕");
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
