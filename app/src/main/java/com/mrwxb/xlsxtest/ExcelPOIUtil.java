package com.mrwxb.xlsxtest;

import android.os.Environment;
import android.util.Log;
import android.widget.EditText;
import android.widget.TextView;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Locale;

/**
 * ProjectName:ExcelPoi
 * Description:
 * Created by zyhang on 2017/5/26.下午4:42
 * Modify by:
 * Modify time:
 * Modify remark:
 */

public class ExcelPOIUtil {

    /**
     * 读Excel
     *
     * @param filePath 文件路径
     */
    public static void read(EditText output, String filePath) {
        try {
            InputStream stream = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(stream);
            XSSFSheet sheet = workbook.getSheetAt(0);
            int rowsCount = sheet.getPhysicalNumberOfRows();
            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            for (int r = 0; r < rowsCount; r++) {
                Row row = sheet.getRow(r);
                int cellsCount = row.getPhysicalNumberOfCells();
                for (int c = 0; c < cellsCount; c++) {
                    String value = getCellAsString(row, c, formulaEvaluator);
                    String cellInfo = "r:" + r + "; c:" + c + "; v:" + value;
                    printlnToUser(output, cellInfo);
                }
            }
        } catch (Exception e) {
            /* proper exception handling to be here */
            e.printStackTrace();
        }
    }

    //全部用字符显示
    protected static String getCellAsString(Row row, int c, FormulaEvaluator formulaEvaluator) {
        String value = "";
        try {
            Cell cell = row.getCell(c);
            CellValue cellValue = formulaEvaluator.evaluate(cell);
            switch (cellValue.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    value = "" + cellValue.getBooleanValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    double numericValue = cellValue.getNumberValue();
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        double date = cellValue.getNumberValue();
                        SimpleDateFormat formatter =
                                new SimpleDateFormat("dd/MM/yy");
                        value = formatter.format(HSSFDateUtil.getJavaDate(date));
                    } else {
                        value = "" + numericValue;
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    value = "" + cellValue.getStringValue();
                    break;
                default:
            }
        } catch (NullPointerException e) {
            /* proper error handling should be here */
            e.printStackTrace();
        }
        return value;
    }

    /**
     * print line to the output TextView
     *
     * @param str
     */
    private static void printlnToUser(EditText output, String str) {
        final String string = str;
        if (output.length() > 8000) {
            CharSequence fullOutput = output.getText();
            fullOutput = fullOutput.subSequence(5000, fullOutput.length());
            output.setText(fullOutput);
            output.setSelection(fullOutput.length());
        }
        output.append(string + "\n");
    }

    /**
     * 获取单元格内容
     *
     * @param cell 单元格
     * @return content
     */
    private static String getCellFormatValue(Cell cell) throws Exception {
        String value = "";
        // 判断当前Cell的Type
        switch (cell.getCellType()) {
            // 如果当前Cell的Type为NUMERIC
            case Cell.CELL_TYPE_NUMERIC:
                // 判断当前的cell是否为Date
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    // 方法2：这样子的data格式是不带带时分秒的：2011-10-12
                    double date = cell.getNumericCellValue();
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm", Locale.CHINA);
                    value = sdf.format(HSSFDateUtil.getJavaDate(date));
                } else {
                    // 如果是纯数字通过NumberToTextConverter.toText(double)将double转成string
                    value = NumberToTextConverter.toText(cell.getNumericCellValue());
                }
                break;
            // 如果当前Cell的Type为STRING
            case Cell.CELL_TYPE_STRING:
                // 取得当前的Cell字符串
                value = cell.getStringCellValue();
                break;
            // 如果当前Cell的Type为BOOLEAN
            case Cell.CELL_TYPE_BOOLEAN:
                value = String.valueOf(cell.getBooleanCellValue());
                break;
        }
        return value;
    }

    /**
     * 改Excel xlsx
     *
     * @param filePath 文件路径
     * @param s        标签
     * @param r        行
     * @param c        列
     * @param content  内容
     */
    //这里是更新数据，用的是sheet.getRow(r);得到行，然后通过row.createCell(c);更新单元格
    //一行中间不能有空白，否则空白后的单元就读不出来
    //  1 2 3
    // 3    5 这个5就读不出来，这样其实在建立表格的时候直接把每个单元格都初始化，要建立单元格，这里的第二行第二列就是没有创建，所以读不出来，只有
    //row.createCell后的单元格才能读出来，或则该单元格本来就有内容
    public static void update(String filePath, int s, int r, int c, String content) {
        try {
            FileInputStream fis = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(s);
            XSSFRow row = sheet.getRow(r);
            XSSFCell cell = row.createCell(c);//这里是创建
            cell.setCellValue(content);//设置内容

            //设置颜色
            //XSSFCellStyle cellStyle = workbook.createCellStyle();
            //cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
           // cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
            //cell.setCellStyle(cellStyle);
            fis.close();//记得关闭输入流
            FileOutputStream fos = new FileOutputStream(filePath);
            workbook.write(fos);
            fos.flush();
            fos.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 写Excel xlsx
     */
    //此方法会覆盖以前的内容，因为用的是Row row = sheet.createRow(i);创建行，所以是新的一行
    public static void write(String filePath) {
        //写之前先用文件建立输入流，让excel工作，然后再建立输出流
        //XXX: Using blank template file as a workaround to make it work
        //Original library contained something like 80K methods and I chopped it to 60k methods
        //so, some classes are missing, and some things not working properly
        try {
            InputStream stream = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(stream);
            XSSFSheet sheet = workbook.getSheetAt(0);
            for (int i=0;i<10;i++) {
                Row row = sheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(i);
            }
            stream.close();//记得关闭输入流
           // String outFileName = "file.xlsx";
           // File outFile = new File(Environment.getExternalStorageDirectory(), outFileName);
            //OutputStream outputStream = new FileOutputStream(outFile.getAbsolutePath());//如果文件不存在，则会新建
            OutputStream outputStream = new FileOutputStream(filePath);//如果文件不存在，则会新建
            workbook.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (Exception e) {
            /* proper exception handling to be here */
            e.printStackTrace();
        }
    }
}

