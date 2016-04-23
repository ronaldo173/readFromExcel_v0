import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

/**
 * Created by Developer on 23.04.2016.
 */
public class WriteBookDemo {

    public static void main(String[] args) {
        HSSFWorkbook workbook = new HSSFWorkbook();

        HSSFSheet sheet = workbook.createSheet("Employee data test");

        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[]{"ID", "Имя", "Фамилия"});
        data.put("2", new Object[]{1, "Alex", "Pendergast"});
        data.put("3", new Object[]{2, "Sergey", "Nem4inskiy"});
        data.put("4", new Object[]{3, "Andrey", "Shev4enko"});

        Set<String> keySet = data.keySet();
        int rowNum = 0;

        for (String key : keySet) {
            Row row = sheet.createRow(rowNum++);
            Object[] objArr = data.get(key);
            int cellNum = 0;

            for (Object obj : objArr) {
                Cell cell = row.createCell(cellNum++);
                if (obj instanceof String) {
                    cell.setCellValue((String) obj);
                } else if (obj instanceof Integer) {
                    cell.setCellValue((Integer) obj);
                }
            }
        }


        FileOutputStream out = null;
        try {
            out = new FileOutputStream(new File("test_demo_excel.xls"));
            workbook.write(out);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        finally {
           if (out!=null){
               try {
                   out.close();
               } catch (IOException e) {
                   e.printStackTrace();
               }
           }
        }

    }
}
