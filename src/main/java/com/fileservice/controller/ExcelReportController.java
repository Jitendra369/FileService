package com.fileservice.controller;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.ss.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.STShapeType;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTHoleSize;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;


import java.io.File;

@RestController
@RequestMapping("/api/excel")
public class ExcelReportController {

    @GetMapping("/donut2")
    public void downloadExcelV1() throws Exception {

        File file = new File("D:\\excels\\dashboard.xlsx");
        file.getParentFile().mkdirs();

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Dashboard");

        int row = 0;

        // Header
        XSSFRow h = sheet.createRow(row++);
        h.createCell(0).setCellValue("Status");
        h.createCell(1).setCellValue("Count");

        // Data
        createRow(sheet, row++, "Posted", 10);
        createRow(sheet, row++, "In Progress", 5);
        createRow(sheet, row++, "Reviewed", 7);
        createRow(sheet, row++, "Revoked", 2);

        // Drawing
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor =
                drawing.createAnchor(0, 0, 0, 0, 3, 1, 10, 15);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("Drawings / Documents");

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);

        XDDFDataSource<String> categories =
                XDDFDataSourcesFactory.fromStringCellRange(sheet,
                        new CellRangeAddress(1, 4, 0, 0));

        XDDFNumericalDataSource<Double> values =
                XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                        new CellRangeAddress(1, 4, 1, 1));

        XDDFPieChartData pieData =
                (XDDFPieChartData) chart.createData(ChartTypes.PIE, null, null);

        XDDFPieChartData.Series series =
                (XDDFPieChartData.Series) pieData.addSeries(categories, values);

        chart.plot(pieData);


        String xml = chart.getCTChart().xmlText();

// inject holeSize manually
        xml = xml.replace("<c:pieChart>",
                "<c:pieChart><c:holeSize val=\"60\"/>");

        chart.getCTChart().set(
                org.openxmlformats.schemas.drawingml.x2006.chart.CTChart.Factory.parse(xml)
        );


        // ===== DONUT MAGIC =====
//        CTPieChart pie =
//                chart.getCTChart().getPlotArea().getPieChartArray(0);
//
//        CTHoleSize hole = pie.addNewHoleSize();
//        hole.setVal(BigInteger.valueOf(60));   // donut hole size
        // ======================

        // Write file
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);

        fos.close();
        workbook.close();
    }

    private void createRow(XSSFSheet sheet, int rowIndex, String name, int value) {
        XSSFRow row = sheet.createRow(rowIndex);
        row.createCell(0).setCellValue(name);
        row.createCell(1).setCellValue(value);
    }


    @GetMapping("/donut")
    public void downloadExcel(HttpServletResponse response) throws Exception {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Dashboard");

        // Data
        sheet.createRow(0).createCell(0).setCellValue("Submitted");
        sheet.getRow(0).createCell(1).setCellValue(10);

        sheet.createRow(1).createCell(0).setCellValue("Cancelled");
        sheet.getRow(1).createCell(1).setCellValue(5);

        // Drawing
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

//        XSSFClientAnchor textAnchor =
//                drawing.createAnchor(0, 0, 0, 0,
//                        6, 7,   // column, row START (below chart)
//                        7, 9);  // column, row END
//
//        XSSFSimpleShape textBox = drawing.createSimpleShape(textAnchor);
//        XSSFRichTextString text =
//                new XSSFRichTextString("Total : 15");
//
//        XSSFFont font = workbook.createFont();
//        font.setBold(true);
//        font.setFontHeight(14);
//
//        text.applyFont(font);
////        textBox.setShapeType(XSSFSimpleShape.c);
//        textBox.setText(text);


        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 3, 1, 10, 15);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("Status");

<<<<<<< HEAD
        // Create DONUT chart ONCE
        XDDFChartData chartData = chart.createData(ChartTypes.DOUGHNUT, null, null);

        XDDFDataSource<String> labels = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, 1, 0, 0));

        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(0, 1, 1, 1));

        chartData.addSeries(labels, values);
        chart.plot(chartData);
        chart.getCTChart()
                .getPlotArea()
                .getDoughnutChartArray(0)
                .getSerArray(0)
                .addNewDLbls()
                .addNewShowCatName()
                .setVal(true);

        chart.getCTChart()
                .getPlotArea()
                .getDoughnutChartArray(0)
                .getSerArray(0)
                .getDLbls()
                .addNewShowVal()
                .setVal(true);

        // Set hole size AFTER plot
        chart.getCTChart()
                .getPlotArea()
                .getDoughnutChartArray(0)
                .addNewHoleSize()
                .setVal((short) 70);

//        ###################### Adding Text , center Number ######################
// ================== HYPERLINK CELL ==================
//        Row linkRow = sheet.createRow(20);
//        Cell linkCell = linkRow.createCell(6);
//        linkCell.setCellValue("Total : 15");
//
//        CreationHelper helper = workbook.getCreationHelper();
//        Hyperlink link = helper.createHyperlink(HyperlinkType.URL);
//        link.setAddress("https://example.com"); // your URL
//
//        linkCell.setHyperlink(link);
//
//// hide row
//        linkRow.setZeroHeight(true);


//        ###################### Adding Text , center Number ######################
//        int total = 15;
//
//// Approximate donut center (based on chart anchor 3,1 → 10,15)
//        int centerRowIndex = 8;
//        int centerColIndex = 6;
//
//        Row centerRow = sheet.getRow(centerRowIndex);
//        if (centerRow == null) {
//            centerRow = sheet.createRow(centerRowIndex);
//        }
//
//        Cell centerCell = centerRow.createCell(centerColIndex);
//        centerCell.setCellValue(total);
//
//// Hyperlink
//        CreationHelper helper = workbook.getCreationHelper();
//        Hyperlink hyperlink = helper.createHyperlink(HyperlinkType.URL);
//        hyperlink.setAddress("https://example.com");
//        centerCell.setHyperlink(hyperlink);
//
//// Styling to look like center text
//        CellStyle style = workbook.createCellStyle();
//        Font font = workbook.createFont();
//        font.setBold(true);
//        font.setFontHeightInPoints((short) 18);
//
//        style.setFont(font);
//        style.setAlignment(HorizontalAlignment.CENTER);
//        style.setVerticalAlignment(VerticalAlignment.CENTER);
//
//        centerCell.setCellStyle(style);
//
//// Increase row height so it looks centered
//        centerRow.setHeightInPoints(35);

        // ================= CENTER TEXT =================
        int total = 15;

// Center of your chart (chart anchor = 3,1 → 10,15)
        XSSFClientAnchor textAnchor =
                drawing.createAnchor(0, 0, 0, 0,
                        6, 7,
                        7, 9);

        XSSFSimpleShape textBox = drawing.createSimpleShape(textAnchor);

// Important: use OBJECT_TYPE_TEXT (not STShapeType)
//        textBox.setShapeType(XSSFSimpleShape.OBJECT_TYPE_TEXT);

        XSSFRichTextString centerText =
                new XSSFRichTextString(String.valueOf(total));

        XSSFFont centerFont = workbook.createFont();
        centerFont.setBold(true);
        centerFont.setFontHeight(18);

        centerText.applyFont(centerFont);

        textBox.setText(centerText);
// =============================================

        // ================= LINK BELOW CHART =================
        int linkRowIndex = 16;
        int linkColIndex = 6;

        Row linkRow = sheet.getRow(linkRowIndex);
        if (linkRow == null) {
            linkRow = sheet.createRow(linkRowIndex);
        }

        Cell linkCell = linkRow.createCell(linkColIndex);
        linkCell.setCellValue("View details");

// Create hyperlink
        CreationHelper helper = workbook.getCreationHelper();
        Hyperlink link = helper.createHyperlink(HyperlinkType.URL);
        link.setAddress("https://example.com");   // your target URL
        linkCell.setHyperlink(link);

// Style like a hyperlink
        CellStyle linkStyle = workbook.createCellStyle();
        Font linkFont = workbook.createFont();
        linkFont.setUnderline(Font.U_SINGLE);
        linkFont.setColor(IndexedColors.BLUE.getIndex());
        linkFont.setBold(true);

        linkStyle.setFont(linkFont);
        linkStyle.setAlignment(HorizontalAlignment.CENTER);

        linkCell.setCellStyle(linkStyle);

// Optional: increase row height
        linkRow.setHeightInPoints(22);
// ===================================================



=======
        // Old POI chart API
        CTChart ctChart = chart.getCTChart();
        CTPlotArea plotArea = ctChart.getPlotArea();
        CTPieChart pieChart = plotArea.addNewPieChart();
        pieChart.addNewVaryColors().setVal(true);

        CTPieSer ser = pieChart.addNewSer();
        ser.addNewIdx().setVal(0);

        // Categories
        CTAxDataSource cat = ser.addNewCat();
        cat.addNewStrRef().setF("Dashboard!$A$1:$A$2");

        // Values
        CTNumDataSource val = ser.addNewVal();
        val.addNewNumRef().setF("Dashboard!$B$1:$B$2");

        // Response
>>>>>>> 92b823f208e96779067eca56de9a4f23b7a4f304
        response.setContentType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition",
                "attachment; filename=dashboard.xlsx");

        workbook.write(response.getOutputStream());
        workbook.close();
    }
}
