import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtil
{
    /**
     * 导出excel文件
     */
    public HSSFWorkbook exprotExcel(List<Person> list, String headInfo, String filepath){
        // 第一步、创建一个workbook对象，对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();
        // 第二步、在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet("sheet_1");
        sheet.setDefaultColumnWidth(20);
        sheet.setDefaultRowHeight((short) 30);
        sheet.setHorizontallyCenter(true);
        //上下左右内边距
        sheet.setMargin(HSSFSheet.BottomMargin, (double)1.0);
        sheet.setMargin(HSSFSheet.LeftMargin, (double)0.7);
        sheet.setMargin(HSSFSheet.RightMargin, (double)0.7);
        sheet.setMargin(HSSFSheet.TopMargin, (double)1.0);
        // 第三步，在sheet中添加表头第0行需合并单元格
        CellRangeAddress region = new CellRangeAddress(0,0,0,3);
        sheet.addMergedRegion(region);
        HSSFRow row0 = sheet.createRow((int) 0);
        // 在sheet中添加表头第1行
        HSSFRow row = sheet.createRow((int) 1);
        row0.setHeightInPoints((float) 18);
        row.setHeightInPoints((float) 18);
        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        //设置居中格式--水平居中且垂直居中
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER_SELECTION);
        HSSFFont font = wb.createFont();
        //设置字体,9号加粗
        font.setFontHeightInPoints((short) 12);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        style.setFont(font);
        //设置表头行各个列的名字
        setSheetHeader0(row0, style, headInfo);
        setSheetHeader(row, style);
        //第五步，添加导出数据到表中
        insertDatasToSheet(sheet, list);
        //第六步，将excel文件存到指定位置
//        writeExcelToDisk(filepath, wb);
        return wb;
    }

    /**
     * 将excel文件存到指定位置
     * @param filePath
     * @param wb
     */
    private void writeExcelToDisk(String filePath, HSSFWorkbook wb) {
        try {
            FileOutputStream fos = new FileOutputStream(filePath);
            wb.write(fos);
            fos.close();
            System.out.println("excel已经导出到:" + filePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 添加导出数据到表中
     * @param sheet
     * @param list
     */
    private void insertDatasToSheet(HSSFSheet sheet, List<Person> list) {
        HSSFCell cell = null;
        HSSFRow row = null;
        for (int i = 0; i < list.size(); i++) {
            row = sheet.createRow((int) i + 2);
            Person person = list.get(i);
            // 创建单元格，并设置各个列中实际数据的值
            cell = row.createCell(0);
            cell.setCellValue(i+1);
            cell = row.createCell(1);
            cell.setCellValue(person.getId());
            cell = row.createCell(2);
            cell.setCellValue(person.getName());
            cell = row.createCell(3);
            cell.setCellValue(person.getAge());
        }
    }
    /**
     * 设置表头行各个列的名字
     * @param row
     * @param style
     */
    private void setSheetHeader0(HSSFRow row, HSSFCellStyle style, String headInfo) {
        HSSFCell cell = row.createCell(0);
        cell.setCellStyle(style);
        cell.setCellValue(headInfo);
        cell = row.createCell(1);
    }
    private void setSheetHeader(HSSFRow row, HSSFCellStyle style) {
        HSSFCell cell = row.createCell(0);
        cell.setCellStyle(style);
        cell.setCellValue("序号");
        cell = row.createCell(1);
        cell.setCellStyle(style);
        cell.setCellValue("学号");
        cell = row.createCell(2);
        cell.setCellStyle(style);
        cell.setCellValue("姓名");
        cell = row.createCell(3);
        cell.setCellStyle(style);
        cell.setCellValue("年龄");
    }

    /**
     * 创造数据
     * @return
     */
    private static List<Person> getStudentData() {
        List<Person> list = new ArrayList<Person>();
        for (int i = 1; i <= 3; i++) {
            String name = "学生";
            Person stu = new Person();
            stu.setId(i);
            stu.setName(name);
            stu.setAge(10+i);
            list.add(stu);
        }
        return list;
    }

    public static void main(String[] args){
        ExcelUtil excelUtil = new ExcelUtil();
        String filepath = "D:/app/ExportExcel" + ".xls";
        List<Person> list = ExcelUtil.getStudentData();
        excelUtil.exprotExcel(list,"导出excel文件", filepath);
    }
}
