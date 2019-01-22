/**
 * @Author: EgbertW
 * @Description:
 * @Date: Created in 21:17 2019/1/9
 */
import java.awt.List;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

    String path;
    public String getPath() {
        return path;
    }
    public void setPath(String path) {
        this.path = path;
    }




    //默认单元格内容为数字时格式
    private static DecimalFormat df = new DecimalFormat("0");
    // 默认单元格格式化日期字符串
    private static SimpleDateFormat sdf = new SimpleDateFormat(  "yyyy-MM-dd HH:mm:ss");
    // 格式化数字
    private static DecimalFormat nf = new DecimalFormat("0.00");
    public static ArrayList<ArrayList<Object>> readExcel(File file){
        if(file == null){
            return null;
        }
        if(file.getName().endsWith("xlsx")){
            //处理ecxel2007
            return readExcel2007(file);
        }else{
            //处理ecxel2003
            return readExcel2003(file);
        }
    }
    /*
     * @return 将返回结果存储在ArrayList内，存储结构与二位数组类似
     * lists.get(0).get(0)表示过去Excel中0行0列单元格
     */
    public static ArrayList<ArrayList<Object>> readExcel2003(File file){
        try{
            ArrayList<ArrayList<Object>> rowList = new ArrayList<ArrayList<Object>>();
            ArrayList<Object> colList;
            HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row;
            HSSFCell cell;
            Object value;
            for(int i = sheet.getFirstRowNum() , rowCount = 0; rowCount < sheet.getPhysicalNumberOfRows() ; i++ ){
                row = sheet.getRow(i);
                colList = new ArrayList<Object>();
                if(row == null){
                    //当读取行为空时
                    if(i != sheet.getPhysicalNumberOfRows()){//判断是否是最后一行
                        rowList.add(colList);
                    }
                    continue;
                }else{
                    rowCount++;
                }
                for( int j = row.getFirstCellNum() ; j <= row.getLastCellNum() ;j++){
                    cell = row.getCell(j);
                    if(cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK){
                        //当该单元格为空
                        if(j != row.getLastCellNum()){//判断是否是该行中最后一个单元格
                            colList.add("");
                        }
                        continue;
                    }
                    switch(cell.getCellType()){
                        case XSSFCell.CELL_TYPE_STRING:
                            //System.out.println(i + "行" + j + " 列 is String type");
                            value = cell.getStringCellValue();
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC:
                            if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                                value = df.format(cell.getNumericCellValue());
                            } else if ("General".equals(cell.getCellStyle()
                                    .getDataFormatString())) {
                                value = nf.format(cell.getNumericCellValue());
                            } else {
                                value = sdf.format(HSSFDateUtil.getJavaDate(cell
                                        .getNumericCellValue()));
                            }
                            //                                System.out.println(i + "行" + j
                            //                                        + " 列 is Number type ; DateFormt:"
                            //                                        + value.toString());
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN:
                            //System.out.println(i + "行" + j + " 列 is Boolean type");
                            value = Boolean.valueOf(cell.getBooleanCellValue());
                            break;
                        case XSSFCell.CELL_TYPE_BLANK:
                            //System.out.println(i + "行" + j + " 列 is Blank type");
                            value = "";
                            break;
                        default:
                            //System.out.println(i + "行" + j + " 列 is default type");
                            value = cell.toString();
                    }// end switch
                    colList.add(value);
                }//end for j
                rowList.add(colList);
            }//end for i

            return rowList;
        }catch(Exception e){
            return null;
        }
    }

    public static ArrayList<ArrayList<Object>> readExcel2007(File file){
        try{
            ArrayList<ArrayList<Object>> rowList = new ArrayList<ArrayList<Object>>();
            ArrayList<Object> colList;
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row;
            XSSFCell cell;
            Object value;
            for(int i = sheet.getFirstRowNum() , rowCount = 0; rowCount < sheet.getPhysicalNumberOfRows() ; i++ ){
                row = sheet.getRow(i);
                colList = new ArrayList<Object>();
                if(row == null){
                    //当读取行为空时
                    if(i != sheet.getPhysicalNumberOfRows()){//判断是否是最后一行
                        rowList.add(colList);
                    }
                    continue;
                }else{
                    rowCount++;
                }
                for( int j = row.getFirstCellNum() ; j <= row.getLastCellNum() ;j++){
                    cell = row.getCell(j);
                    if(cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK){
                        //当该单元格为空
                        if(j != row.getLastCellNum()){//判断是否是该行中最后一个单元格
                            colList.add("");
                        }
                        continue;
                    }
                    switch(cell.getCellType()){
                        case XSSFCell.CELL_TYPE_STRING:
                            //System.out.println(i + "行" + j + " 列 is String type");
                            value = cell.getStringCellValue();
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC:
                            if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                                value = df.format(cell.getNumericCellValue());
                            } else if ("General".equals(cell.getCellStyle()
                                    .getDataFormatString())) {
                                value = nf.format(cell.getNumericCellValue());
                            } else {
                                value = sdf.format(HSSFDateUtil.getJavaDate(cell
                                        .getNumericCellValue()));
                            }
                            //                                System.out.println(i + "行" + j
                            //                                        + " 列 is Number type ; DateFormt:"
                            //                                        + value.toString());
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN:
                            //System.out.println(i + "行" + j + " 列 is Boolean type");
                            value = Boolean.valueOf(cell.getBooleanCellValue());
                            break;
                        case XSSFCell.CELL_TYPE_BLANK:
                            //System.out.println(i + "行" + j + " 列 is Blank type");
                            value = "";
                            break;
                        default:
                            //System.out.println(i + "行" + j + " 列 is default type");
                            value = cell.toString();
                    }// end switch
                    colList.add(value);
                }//end for j
                rowList.add(colList);
            }//end for i

            return rowList;
        }catch(Exception e){
            System.out.println("exception");
            return null;
        }
    }
    public static ArrayList getFiles(String filePath){
        File root = new File(filePath);
        File[]files = root.listFiles();
        ArrayList filelist = new ArrayList();
        for(File file:files){

            if(file.isDirectory()){
                filelist.addAll(getFiles(file.getAbsolutePath()));
            }else{
                String newpath = file.getAbsolutePath();
                if(newpath.contains("交易记录")){

                    filelist.add(newpath);
                }
            }
        }
        return filelist;
    }
    public  void readBook(String path3) {
        String filePath = path3;
        ArrayList filelist = getFiles(filePath);
        ArrayList<ArrayList>resultAll = new ArrayList<ArrayList>();
        for(int i = 0;i<filelist.size();i++){
            String path = (String) filelist.get(i);
            System.out.println(path);
            ArrayList<ArrayList>result = Graph(path);

            String[] path2 = path.split("\\\\");
            int num = result.get(0).size();
            ArrayList result2 = new ArrayList();
            for(int j = 0;j<num;j++){
                result2.add(path2[path2.length-2]);
            }
            ArrayList result3 = new ArrayList();
            for(int j = 0;j<num;j++){
                result3.add(path2[path2.length-3]);
            }
            result.add(result2);
            result.add(result3);
            if(resultAll.size()==0){
                resultAll = result;
            }else{
                for(int j = 0;j<result.size();j++){
                    for(int k = 0;k<result.get(j).size();k++){
                        resultAll.get(j).add(result.get(j).get(k));
                    }
                }
            }
        }

        writeExcel(resultAll,"D:/a.xls");
    }
    public static void writeExcel(ArrayList<ArrayList> result,String path){
        if(result == null){
            return;
        }
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("sheet1");
        for(int i = 0 ;i < result.get(0).size() ; i++){
            HSSFRow row = sheet.createRow(i);
            for(int j = 0; j < result.size() ; j ++){
                HSSFCell cell = row.createCell((short)j);
                cell.setCellValue(result.get(j).get(i).toString());
            }
        }
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try
        {
            wb.write(os);
        } catch (IOException e){
            e.printStackTrace();
        }
        byte[] content = os.toByteArray();
        File file = new File(path);//Excel文件生成后存储的位置。
        OutputStream fos  = null;
        try
        {
            fos = new FileOutputStream(file);
            wb.write(fos);
            os.close();
            fos.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public static DecimalFormat getDf() {
        return df;
    }
    public static void setDf(DecimalFormat df) {
        ExcelRead.df = df;
    }
    public static SimpleDateFormat getSdf() {
        return sdf;
    }
    public static void setSdf(SimpleDateFormat sdf) {
        ExcelRead.sdf = sdf;
    }
    public static DecimalFormat getNf() {
        return nf;
    }
    public static void setNf(DecimalFormat nf) {
        ExcelRead.nf = nf;
    }




    public static ArrayList<ArrayList> Graph(String path){
        File file = new File(path);
        ArrayList<ArrayList<Object>> result = ExcelRead.readExcel(file);
        ArrayList<Double>price = new ArrayList<Double>();//价格序列
        ArrayList<String>time = new ArrayList<String>();//时间序列
        ArrayList<String>buyList = new ArrayList<String>();//买方序列
        ArrayList<String>sellList = new ArrayList<String>();//卖方序列
        ArrayList<Double>vol = new ArrayList<Double>();//成交量
        ArrayList<String>Share = new ArrayList<String>();//股票名字
        ArrayList<String>id = new ArrayList<String>();
        ArrayList<String>Shareid = new ArrayList<String>();
        for(int i = 2 ;i < result.size() ;i++){
            for(int j = 0;j<result.get(i).size(); j++){
                //第5列表示价格，第8列表示时间
                if(j==0){
                    String temp = (String) result.get(i).get(j);
                    id.add(temp);
                }
                if(j==3){
                    String temp = (String) result.get(i).get(j);
                    Shareid.add(temp);
                }
                if(j==5){
                    //price.add((String) result.get(i).get(j));
                    String temp = (String) result.get(i).get(j);
                    String[] units = temp.split("￥");
                    price.add(Double.valueOf(units[1]));
                }
                if(j==7){
                    String temp = (String) result.get(i).get(j);
                    time.add(temp);
                    //                  time.add((String) result.get(i).get(j));
                }
                if(j==1){
                    buyList.add((String) result.get(i).get(j));
                }
                if(j==2){
                    sellList.add((String) result.get(i).get(j));
                }
                if(j==6){
                    vol.add(Double.valueOf((String)result.get(i).get(j)));
                }
                if(j==4){
                    Share.add((String)result.get(i).get(j));
                }
            }
        }
        ArrayList<ArrayList>resultList = new ArrayList<ArrayList>();
        resultList.add(Shareid);
        resultList.add(id);
        resultList.add(buyList);
        resultList.add(sellList);
        resultList.add(Share);
        resultList.add(price);
        resultList.add(vol);
        resultList.add(time);
        return resultList;
    }




}