import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * Created by 热带雨林 on 2018/9/16.
 */
public class test3 {
    public static void main(String[] args){
        OutputStream out =null;
        try{
            String filePath="d:/test.xls";
            Workbook wb = null;
            String extString = filePath.substring(filePath.lastIndexOf("."));
            InputStream is = null;
            try {
                is = new FileInputStream(filePath);
                if(".xls".equals(extString)){
                    wb = new HSSFWorkbook(is);
                }else if(".xlsx".equals(extString)){
                    wb = new XSSFWorkbook(is);
                }else{
                    wb = null;
                }
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
            Sheet sheet = wb.getSheetAt(0);
            //System.out.println("总行数："+sheet.getLastRowNum());

            Row row = null;
            Map<Integer,Map> outMap= new HashMap<Integer,Map>();
            Integer rowCount=null;
            for (Integer i = 0; i <= sheet.getLastRowNum(); i++) {
                //System.out.println("当前行："+i);
                row = sheet.getRow(i);
                if(row==null){
                    System.out.println("当前行是空行");
                }else{
                    if(rowCount==null){
                        rowCount=0;
                        //System.out.println("输出rowCount："+rowCount);
                    }else{
                        rowCount=rowCount+1;
                        //System.out.println("输出rowCount："+rowCount);
                    }
                    Map<Integer,String> innerMap= new HashMap<Integer,String>();
                    for (Integer j = 0; j <row.getLastCellNum(); j++) {
                        String cellinfo = (String) getCellFormatValue(row.getCell(j));
                        innerMap.put(j,cellinfo);
                        System.out.println("  第"+j+"列："+cellinfo);
                    }
                    outMap.put(rowCount,innerMap);
                }
            }
            is.close();

            // 声明一个工作薄
            HSSFWorkbook newWB = new HSSFWorkbook();
            //XSSFWorkbook newWB=new XSSFWorkbook();
            // 生成一个表格
            HSSFSheet newSheet = newWB.createSheet("sheet1");
            //XSSFSheet newSheet = newWB.createSheet("sheet1");
            //将输出流和HSSFWorkbook关联
            out = new FileOutputStream("d:/test1.xls");

            Set<Integer> outKeySet = outMap.keySet();
            Cell cell =null;
            for(Integer okey : outKeySet){
                Row row2 = newSheet.createRow(okey);
                System.out.println("写入，当前行：:"+okey);
                Map<Integer,String> writeInnerMap=outMap.get(okey);
                Set<Integer> innerKeySet = writeInnerMap.keySet();
                for(Integer ikey : innerKeySet){
                    System.out.println("   写入，当前列：:"+writeInnerMap.get(ikey));
                    cell = row2.createCell(ikey);
                    cell.setCellValue(writeInnerMap.get(ikey));
                }
            }
            newWB.write(out);
        }catch (Exception e){
            e.printStackTrace();
        }finally{
            try {
                if(out != null){
                    out.flush();
                    out.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static Workbook getWorkbok(File file) throws IOException {
        Workbook hwb = null;
        FileInputStream in = new FileInputStream(file);
        if(file.getName().endsWith("xls")){     //Excel&nbsp;2003
            hwb = new HSSFWorkbook(in);
        }else if(file.getName().endsWith("xlsx")){    // Excel 2007/2010
            hwb = new XSSFWorkbook(in);
        }
        return hwb;
    }

    //读取excel
    public static Workbook readExcel(String filePath){
        Workbook wb = null;
        if(filePath==null){
            return null;
        }
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if(".xls".equals(extString)){
                return wb = new HSSFWorkbook(is);
            }else if(".xlsx".equals(extString)){
                return wb = new XSSFWorkbook(is);
            }else{
                return wb = null;
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }

    public static Object getCellFormatValue(Cell cell){
        Object cellValue = null;
        if(cell!=null){
            //判断cell类型
            switch(cell.getCellType()){
                case Cell.CELL_TYPE_NUMERIC:{
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                }
                case Cell.CELL_TYPE_FORMULA:{
                    //判断cell是否为日期格式
                    if(DateUtil.isCellDateFormatted(cell)){
                        //转换为日期格式YYYY-mm-dd
                        cellValue = cell.getDateCellValue();
                    }else{
                        //数字
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case Cell.CELL_TYPE_STRING:{
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                default:
                    cellValue = "";
            }
        }else{
            cellValue = "";
        }
        return cellValue;
    }
}
