package com.example.demo;


import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

import org.apache.poi.xssf.usermodel.*;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.stream.IntStream;

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
            XSSFSheet sheet = workbook.createSheet("Sheet1");

            // 创建样式
            XSSFCellStyle headerStyle = createHeaderStyle(workbook);

            XSSFCellStyle oddStyle = createRowStyle(workbook);
            XSSFCellStyle evenStyle = createRowStyle(workbook);

            // 创建表头, 第0行, 假设有两列
            List<String> headers = Arrays.asList("ID", "Name");
            XSSFRow headerRow = sheet.createRow(0);
            IntStream.range(0, headers.size()).forEach(index -> {
                XSSFCell cell = headerRow.createCell(index);
                cell.setCellValue(headers.get(index));
                cell.setCellStyle(headerStyle);
            });

            // 写入数据行, 第1行开始, 有两个cell
            for (int i = 1; i <= 10; i++) {
                XSSFRow row = sheet.createRow(i);
                row.createCell(0).setCellValue(i);
                row.createCell(1).setCellValue("Name " + i);

                // 设置数据行样式
                row.getCell(0).setCellStyle(oddStyle);
                row.getCell(1).setCellStyle(oddStyle);
            }

            // 为表头列添加筛选
            //sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, 1)); // 设置筛选区域从A1到B1

            // 创建表格区域
            // 设置表格的范围（例如：A1:B11）, topLeft(row0, cell0), botRight(row10, cell1)
            AreaReference areaReference = new AreaReference(new CellReference(0, 0), new CellReference(10, 1), SpreadsheetVersion.EXCEL2007);
            // 方式2, 直接指定区域AreaReference areaReference = new AreaReference("A1:B11", SpreadsheetVersion.EXCEL2007);
            sheet.createTable(areaReference);

            // 设置自动列宽, 有几列设置几列的
            IntStream.range(0, headers.size()).forEach(sheet::autoSizeColumn);

            // 写入字节数组
            workbook.write(byteArrayOutputStream);
            return byteArrayOutputStream.toByteArray();

        } catch (IOException e) {
            throw new IllegalStateException("Unable to build XLSX bytes", e);
        }
    }

    // 创建表头样式
    private XSSFCellStyle createHeaderStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();

        // 设置字体

        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setColor(new XSSFColor(new byte[]{(byte) 255, (byte) 255, (byte) 255}));
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
    private XSSFCellStyle createRowStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();

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
