package com.example.demo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;

@Controller
class FileDownloadController {

    private static final Logger log = LoggerFactory.getLogger(FileDownloadController.class);

    @GetMapping
    public String download() {
        return "download";
    }

    private static final java.util.Base64.Encoder ENCODER = java.util.Base64.getEncoder();

    @ResponseBody
    @GetMapping("/getDataUrl")
    public String downloadFile() throws IOException {
        String xlsxMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        byte[] xlsxBytes = buildXlsxBytes();
        String base64Content = ENCODER.encodeToString(xlsxBytes);
        return "data:" + xlsxMimeType + ";base64," + base64Content;
    }

    public byte[] buildXlsxBytes() {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {

            // 创建工作表
            Sheet sheet = workbook.createSheet("Sheet1");

            // 创建样式
            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle rowStyle = createRowStyle(workbook);

            // 创建表头
            Row headerRow = sheet.createRow(0);
            Cell headerCell1 = headerRow.createCell(0);
            headerCell1.setCellValue("ID");
            headerCell1.setCellStyle(headerStyle);
            Cell headerCell2 = headerRow.createCell(1);
            headerCell2.setCellValue("Name");
            headerCell2.setCellStyle(headerStyle);

            // 写入数据行
            for (int i = 1; i <= 10; i++) {
                Row row = sheet.createRow(i);
                row.createCell(0).setCellValue(i);
                row.createCell(1).setCellValue("Name " + i);

                // 设置数据行样式
                row.getCell(0).setCellStyle(rowStyle);
                row.getCell(1).setCellStyle(rowStyle);
            }

            // 设置自动列宽
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);

            // 写入字节数组
            workbook.write(byteArrayOutputStream);
            return byteArrayOutputStream.toByteArray();

        } catch (IOException e) {
            throw new IllegalStateException("Unable to build XLSX bytes", e);
        }
    }

    // 创建表头样式
    private CellStyle createHeaderStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // 设置字体
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);

        // 设置背景色为深蓝色
        style.setFillForegroundColor(new XSSFColor(new byte[]{0, 51, (byte) 102})); // RGB(0, 51, 102)
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 设置边框
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    // 创建数据行样式
    private CellStyle createRowStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // 设置背景色为浅蓝色
        style.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 204, (byte) 229, (byte) 255})); // RGB(204, 229, 255)
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 设置边框
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }
}
