package com.vbn.taobaokeMoney.excelOperation;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Array;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by vbn on 2018/12/20.
 */

public class OperationExcel {

    private static final String FILE_NAME = "classpath:datafile/test.xls";
    private static  final String ZTF_AD = "65802000112";
    private static  final String ZC_AD = "63841000166";

    public static void readExcel(String path) {
        if (path == null) {
            path = FILE_NAME;
        }
        List<Map> list = new ArrayList<Map>();
        DateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        try {
            FileInputStream excelFile = new FileInputStream(ResourceUtils.getFile(path));
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet dataSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = dataSheet.iterator();
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                if (currentRow.getRowNum() == 0) {
                    continue;
                }
                Map<String, Object> rowMap = new HashMap<String, Object>();
                Cell cell = currentRow.getCell(2);
                if (cell.getCellTypeEnum() == CellType.STRING) {
                    rowMap.put("title", cell.getStringCellValue());
                }
                cell = currentRow.getCell(0);
                if (cell.getCellTypeEnum() == CellType.STRING) {
                    String dateStr = cell.getStringCellValue();
                    try {
                        rowMap.put("date", format.parse(dateStr));
                    }
                    catch (Exception e) {
                        e.printStackTrace();
                    }
                }
                cell = currentRow.getCell(18);
                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                    rowMap.put("money", cell.getNumericCellValue());
                }
                cell = currentRow.getCell(29);
                if (cell.getCellTypeEnum() == CellType.STRING) {
                    rowMap.put("adId", cell.getStringCellValue());
                }
                list.add(rowMap);
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        Map<String,Double> adMap = new HashMap<String,Double>();
        Date lastDate = null;
        Date beforDate = null;
        if (list.size() != 0) {
            for(Map map: list) {
                Date date = (Date) map.get("date");
                if (lastDate == null) {
                    lastDate = (Date) map.get("date");
                } else {
                    lastDate = lastDate.before(date)? date : lastDate;
                }
                if (beforDate == null) {
                    beforDate = (Date) map.get("date");
                } else {
                    beforDate = beforDate.after(date)? date : beforDate;
                }

                if (adMap.get(map.get("adId").toString()) == null) {
                    Double money = (Double) map.get("money");
                    adMap.put(map.get("adId").toString(),money * 0.9);
                } else {
                    Double money = (Double) map.get("money") *0.9;
                    Double allMoney = (Double)adMap.get(map.get("adId").toString());
                    allMoney = allMoney + money;
                    adMap.put(map.get("adId").toString(),allMoney);
                }
            }
        }
        System.out.println("********************按广告位结算**********************");
        System.out.println("********************最早时间**********************");
        System.out.println("********************" + format.format(beforDate) + "**********************");
        System.out.println("********************最晚时间**********************");
        System.out.println("********************" + format.format(lastDate) + "**********************");
        for (String key: adMap.keySet()) {
            System.out.println("********************" + key + "**********************");
            System.out.println("********************" + adMap.get(key) + "**********************");
        }

    }
}
