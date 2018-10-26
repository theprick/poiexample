package com.padrian;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.Date;

public class PoiExample {
    public static void main(String[] args) throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("MySheet1");
        sheet.setDefaultColumnWidth(30);
        Row row = sheet.createRow(0);

        Cell cell1 = row.createCell(0);
        cell1.setCellValue(34537.98282);

        Cell cell2 = row.createCell(1);
        cell2.setCellValue(13.90);

        Cell cell3 = row.createCell(2);
        // set formula
        cell3.setCellFormula("A1+B1");

        CellStyle cell3Style = wb.createCellStyle();
        // set font cell
        Font boldFont = wb.createFont();
        boldFont.setBold(true);
        boldFont.setColor(IndexedColors.DARK_RED.getIndex());
        cell3Style.setFont(boldFont);
        // set border cell
        cell3Style.setBorderBottom(BorderStyle.THIN);
        cell3Style.setBorderTop(BorderStyle.THIN);
        cell3Style.setBorderLeft(BorderStyle.THIN);
        cell3Style.setBorderRight(BorderStyle.THIN);
        // add a number formatter
        CreationHelper createHelper = wb.getCreationHelper();
        // cell3Style.setDataFormat(createHelper.createDataFormat().getFormat("#.##"));

        // use a builtin instead: see http://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html
        // https://stackoverflow.com/questions/41439591/number-and-cell-formatting-in-apache-poi
        cell3Style.setDataFormat(createHelper.createDataFormat().getFormat(BuiltinFormats.getBuiltinFormat(4)));
        // apply the style
        cell3.setCellStyle(cell3Style);
        cell3.setAsActiveCell();

        // add a date type
        Row row2 = sheet.createRow(1);
        Cell dateCell = row2.createCell(0);
        CellStyle dateCellStyle = wb.createCellStyle();
        // add date formatter
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));
        dateCell.setCellStyle(dateCellStyle);
        dateCell.setCellStyle(dateCellStyle);
        // set current date time
        dateCell.setCellValue(new Date());

        FileOutputStream outputStream = new FileOutputStream("C:\\Temp\\test1.xlsx");
        wb.write(outputStream);
        wb.close();
        outputStream.close();
    }
}
