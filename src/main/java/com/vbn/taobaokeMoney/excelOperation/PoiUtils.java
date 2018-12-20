package com.vbn.taobaokeMoney.excelOperation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
//import org.springframework.web.multipart.MultipartFile;

//import com.porui.hjl.commons.utils.DateUtils;

public class PoiUtils {
	
	

	private static Logger logger= LoggerFactory.getLogger(PoiUtils.class);
	
	/**
	 * 读取xls MultipartFile
	 * @param file
	 * @return
	 * @throws IOException
	 */
//	public static List<Map<String,String>> readXls(MultipartFile file) throws IOException {
//		InputStream is = file.getInputStream();
//		return readXls(is);
//	}
	
	/**
	 * 读取xls File
	 * @param file
	 * @return
	 * @throws IOException
	 */
	public static List<Map<String,String>> readXls(File file) throws IOException {
		InputStream is = new FileInputStream(file);
		return readXls(is);
	}
	
	/**
	 * 读取xls InputStream
	 * @param is
	 * @return
	 * @throws IOException
	 */
	public static List<Map<String,String>> readXls(InputStream is) throws IOException {
		
		List<String> list = new ArrayList<String>();
		List<Map<String,String>> entityMap = new ArrayList<Map<String,String>>();
		try {
			Workbook wb = WorkbookFactory.create(is);
			
			// 循环工作表Sheet
	         for (int numSheet = 0; numSheet < wb.getNumberOfSheets(); numSheet++) {
	        	 Sheet hssfSheet = wb.getSheetAt(numSheet);
	             if (hssfSheet == null) {
	                 continue;
	             }
	             // 循环行Row
	             for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
	                 Row hssfRow = hssfSheet.getRow(rowNum);
	                 if (hssfRow != null) {
	                	 Map<String,String> map = new HashMap<String,String>();
	                	 for(int columnNum = 0; columnNum < hssfRow.getLastCellNum(); columnNum++){
		                	 if(rowNum == 0){
		                		 list.add(getValue(hssfRow.getCell(columnNum)));
		                	 }else{
		                		 if(list!=null&&list.size()!=0){
		                			 try{
		                				 map.put(list.get(columnNum),getValue(hssfRow.getCell(columnNum)));
		                			 }catch(Exception e){
		                				 logger.info(e.getMessage());
		                			 }
		                		 }
		                	 }
	                	 }
	                	 if(rowNum != 0){
	                		 entityMap.add(map);
	                 	 }
	                 }
	              }
	         }
			
		} catch (EncryptedDocumentException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (InvalidFormatException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
         
         return entityMap;
     }
	 
	 /**
	  * 读取单元格内容
	  * @param hssfCell
	  * @return
	  */
     @SuppressWarnings("static-access")
     private static String getValue(Cell hssfCell) {
    	 String result = "";
    	 if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
             // 返回布尔类型的值
        	 DecimalFormat df = new DecimalFormat("#");
        	 result = df.format(hssfCell.getBooleanCellValue());
         } else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
        	 if (HSSFDateUtil.isCellDateFormatted(hssfCell)) {
                 // 返回日期类型的值
//            	 result = DateUtils.format(hssfCell.getDateCellValue(), "yyyy-MM-dd");
        	 } else{
        		 // 返回数值类型的值
	        	 DecimalFormat df = new DecimalFormat("#");
	        	 result = df.format(hssfCell.getNumericCellValue());
        	 }
         } else {
        	 try{
//        		 result = DateUtils.format(DateUtils.parseDate(hssfCell.getStringCellValue(), "yyyy-MM-dd"), "yyyy-MM-dd");
        	 }catch(Exception e){
	             // 返回字符串类型的值
	        	 result = String.valueOf(hssfCell.getStringCellValue());
        	 }
         }
         return result.trim();
     }
     
}
