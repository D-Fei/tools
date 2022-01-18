package com.dfei.tools.tools.exportUtils.excelUtils.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

import static com.dfei.tools.tools.exportUtils.excelUtils.utils.Util.*;
/**
 * @author dfei
 * @ClassName XSSFUtils.java
 * @createTime 2022/1/18 11:22
 * @Description TODO
 */
public class XSSFUtils {

    public static Workbook run(List<Map<String, Object>> list, Map<String, Object> title) {
        if (list.isEmpty()) {
            return null;
        }
        Workbook wb = new XSSFWorkbook();
        createWBHead(wb, list, title);
        createWBBody(wb, list);
        return wb;
    }

    protected static Workbook createWBHead(Workbook wb, List<Map<String, Object>> list, Map<String, Object> titleMapping) {

        if (list.isEmpty()) {
            return null;
        }
        Sheet sheet = wb.createSheet("0");
        Row row = sheet.createRow(0);
        int columnNum = list.get(0).size();
        List<String> keys = getMapKeys(list.get(0));
        Map<String, Object> replaceMap = new HashMap<>();
        if (titleMapping == null || titleMapping.isEmpty()) {
            keys.stream().forEach(e -> replaceMap.put(e, e.toUpperCase()));
            titleMapping = replaceMap;
        }
        for (int i = 0; i < columnNum; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(String.valueOf(titleMapping.get(keys.get(i))));
        }

        return wb;
    }

    protected static Workbook createWBBody(Workbook wb, List<Map<String, Object>> list) {
        int rowNum = list.size();
        int columnNum = list.get(0).size();

        List<String> columns = getMapKeys(list.get(0));

        if (rowNum == 0 || columnNum == 0) {
            return null;
        }

        Sheet sheet = wb.getSheet("0");

        int firstRowNum = sheet.getLastRowNum() + 1;

        for (int i = 0; i < rowNum; i++) {
            Row row = sheet.createRow(firstRowNum + i);
            for (int j = 0; j < columnNum; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(String.valueOf(list.get(i).get(columns.get(j))));
            }
        }
        return wb;
    }



    public static void main(String[] args) {
        List<Map<String, Object>> list = new ArrayList<>() {{
            add(new HashMap<>() {{
                put("name", "Job");
                put("age", 18);
                put("sex", "male");
            }});
            add(new HashMap<>() {{
                put("name", "Candy");
                put("age", 20);
                put("sex", "female");
            }});
            add(new HashMap<>() {{
                put("name", "Kan");
                put("age", 19);
                put("sex", "male");
            }});
        }};
        Workbook wb = run(list, null);
        try (OutputStream out = new FileOutputStream("workbook.xlsx")) {
            wb.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
