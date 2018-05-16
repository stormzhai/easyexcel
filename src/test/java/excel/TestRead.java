package excel;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by daojia on 2018-5-14.
 */
public class TestRead {
    public static void main(String[] args) {
        //readExcel();
        //writeExcel();
        //read03Excel();
        write03Excel();
    }


    private static void read03Excel() {
        InputStream inputStream = null;
        try {
            inputStream = getInputStream("aa.xls");
            ExcelListener excelListener = new ExcelListener();
            ExcelReader reader = new ExcelReader(inputStream, ExcelTypeEnum.XLS, null, excelListener);
            reader.read(new Sheet(1, 1, ExcelRowJavaModel.class));

        } catch (Exception e) {
            e.printStackTrace();

        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static void readExcel() {
        InputStream inputStream = null;
        try {
            inputStream = getInputStream("bb.xlsx");
            ExcelListener excelListener = new ExcelListener();
            ExcelReader reader = new ExcelReader(inputStream, ExcelTypeEnum.XLSX, null, excelListener);
            reader.read(new Sheet(1, 1, ExcelRowJavaModel.class));

        } catch (Exception e) {
            e.printStackTrace();

        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static void write03Excel() {
        OutputStream out = null;
        try {
            out = new FileOutputStream("c:/78.xls");
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLS);
            //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
            Sheet sheet1 = new Sheet(1, 0,ExcelRowJavaModel.class);
            writer.write(getData(), sheet1);
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

    private static void writeExcel() {
        OutputStream out = null;
        try {
            out = new FileOutputStream("c:/78.xlsx");
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
            //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
            Sheet sheet1 = new Sheet(1, 0,ExcelRowJavaModel.class);
            writer.write(getData(), sheet1);
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


    public static List<ExcelRowJavaModel> getData() {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        List<ExcelRowJavaModel> datas = new ArrayList<ExcelRowJavaModel>();
        try {
            ExcelRowJavaModel model = new ExcelRowJavaModel();
            model.setNum(1);
            model.setCode(1L);
            model.setEndTime(simpleDateFormat.parse("2018-05-15"));
            model.setMoney(1.0d);
            model.setTimes("1");
            model.setActivityCode(1);
            model.setDate(new Date());
            model.setLx(0.0);
            model.setName("测试111");
            model.setSex(1);

            datas.add(model);
        } catch (Exception e) {
            e.printStackTrace();
        }


        return datas;
    }
    private static InputStream getInputStream(String fileName) {
        return Thread.currentThread().getContextClassLoader().getResourceAsStream("" + fileName);

    }
}
