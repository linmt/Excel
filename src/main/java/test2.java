import jxl.Sheet;
import jxl.Workbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * Created by 热带雨林 on 2018/9/16.
 */
public class test2 {
    public static void main(String[] args){
        try{
            FileInputStream is = new FileInputStream("d:/test.xls");
            jxl.Workbook wb = jxl.Workbook.getWorkbook(is);
            Sheet sheet = wb.getSheet(0);
//            System.out.println("总行数："+sheet.getRows()+"，总列数："+sheet.getColumns());
            Map<Integer,Map> outMap= new HashMap<Integer,Map>();
            for (Integer i = 0; i < sheet.getRows(); i++) {
                System.out.println("当前行："+i);
                Map<Integer,String> innerMap= new HashMap<Integer,String>();
                for (Integer j = 0; j < sheet.getColumns(); j++) {
                    String cellinfo = sheet.getCell(j, i).getContents();
                    innerMap.put(j,cellinfo);
                    System.out.println("  当前列："+j+","+cellinfo);
                }
                outMap.put(i,innerMap);
            }
            is.close();

            OutputStream os = new FileOutputStream("d:/test1.xls");
            jxl.write.WritableWorkbook wwb = Workbook.createWorkbook(os);
            jxl.write.WritableSheet ws = wwb.createSheet("Sheet1", 0);

            jxl.write.Label labelCFC =null;
            Set<Integer> outKeySet = outMap.keySet();
            for(Integer okey : outKeySet){
                System.out.println("写入，当前行：:"+okey);
                Map<Integer,String> writeInnerMap=outMap.get(okey);
                Set<Integer> innerKeySet = writeInnerMap.keySet();
                for(Integer ikey : innerKeySet){
                    System.out.println("写入，当前列：:"+writeInnerMap.get(ikey));
                    labelCFC = new jxl.write.Label(ikey, okey, writeInnerMap.get(ikey));
                    ws.addCell(labelCFC);
                }
            }
            wwb.write();
            wwb.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
