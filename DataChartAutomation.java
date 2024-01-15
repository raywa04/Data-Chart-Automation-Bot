package com.epsoftinc.datachartautomation;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFGraphicFrame;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class DataChartAutomation {

    public static void copyWorkbook(XSLFChart chart, XSSFWorkbook source) throws IOException, InvalidFormatException {
        XSSFWorkbook target = chart.getWorkbook();
        XSSFSheet sheet = source.getSheet("Sheet1");

        // delete sheet1 if it already exists in target workbook
        if(target.getSheet("Sheet1") != null){
            target.removeSheetAt(target.getSheetIndex("Sheet1"));
        }

        target.createSheet("Sheet1");

        for(Row row : sheet){
            // copy each row to the target sheet
            Row destRow = target.getSheet("Sheet1").createRow(row.getRowNum());
            for(Cell cell : row){
                // copy each cell to the target cell
                Cell destCell = destRow.createCell(cell.getColumnIndex());
                if(cell.getCellType() == CellType.STRING) {
                    destCell.setCellValue(cell.getStringCellValue());
                } else if(cell.getCellType() == CellType.NUMERIC) {
                    destCell.setCellValue(cell.getNumericCellValue());
                } else if(cell.getCellType() == CellType.BOOLEAN) {
                    destCell.setCellValue(cell.getBooleanCellValue());
                } else if(cell.getCellType() == CellType.BLANK) {
                    destCell.setCellValue("");
                }
            }
        }
        // save the updated embedded workbook
        chart.setWorkbook(target);
    }

    public static void copyWorkbook(XSLFChart chart, XSSFWorkbook source, String sheetname) throws IOException, InvalidFormatException {
        XSSFWorkbook target = chart.getWorkbook();
        XSSFSheet sheet = source.getSheet(sheetname);

        // delete sheet1 if it already exists in target workbook
        if(target.getSheet(sheetname) != null){
            target.removeSheetAt(target.getSheetIndex(sheetname));
        }

        target.createSheet(sheetname);

        for(Row row : sheet){
            // copy each row to the target sheet
            Row destRow = target.getSheet(sheetname).createRow(row.getRowNum());
            for(Cell cell : row){
                // copy each cell to the target cell
                Cell destCell = destRow.createCell(cell.getColumnIndex());
                if(cell.getCellType() == CellType.STRING) {
                    destCell.setCellValue(cell.getStringCellValue());
                } else if(cell.getCellType() == CellType.NUMERIC) {
                    destCell.setCellValue(cell.getNumericCellValue());
                } else if(cell.getCellType() == CellType.BOOLEAN) {
                    destCell.setCellValue(cell.getBooleanCellValue());
                } else if(cell.getCellType() == CellType.BLANK) {
                    destCell.setCellValue("");
                }
            }
        }
        // save the updated embedded workbook
        chart.setWorkbook(target);
    }

    private static void updateChart(XSSFSheet sheet, XSLFChart chart){
        XDDFDataSource<Double> x = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, sheet.getLastRowNum() - 1, 0, 0));
        XDDFNumericalDataSource<Double> y = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, sheet.getLastRowNum() - 1, 1, 1));

        XDDFScatterChartData scatterData = (XDDFScatterChartData) chart.getChartSeries().get(0);

        // Get the first series of the chart
        XDDFChartData.Series oldSeries = scatterData.getSeries(0);

        // Replace the old data with the new data
        oldSeries.replaceData(x, y);

        // Plot the new data
        chart.plot(scatterData);
    }


    private static XSLFShape findShapeByName(XSLFShape shape, String name) {
        if (shape.getShapeName().equals(name)) {
            return shape;
        } else if (shape instanceof XSLFGroupShape) {
            XSLFGroupShape group = (XSLFGroupShape) shape;
            for (XSLFShape subShape : group.getShapes()) {
                XSLFShape found = findShapeByName(subShape, name);
                if (found != null) {
                    return found;
                }
            }
        }
        return null;
    }

    private static XSLFChart findChartInSlide(XSLFSlide slide, String name) {
        for (XSLFShape shape : slide.getShapes()) {
            XSLFShape found = findShapeByName(shape, name);
            if (found != null) {
                if(found instanceof XSLFGraphicFrame) {
                    XSLFGraphicFrame frame = (XSLFGraphicFrame) found;
                    XSLFChart chart = frame.getChart();
                    return chart;
                }
            }
        }
        return null;
    }

    public static void main(String[] args) throws Exception {
        String excelFilePath = "/Users/rayyanwaris/Downloads/DataChartAutomation/data.xlsx";
        try {
            FileInputStream inputStream = new FileInputStream(excelFilePath);

            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

            inputStream.close();

            // Open the PowerPoint file
            FileInputStream pptStream = new FileInputStream("/Users/rayyanwaris/Downloads/DataChartAutomation/test.pptx");
            XMLSlideShow ppt = new XMLSlideShow(pptStream);

            Scanner sc = new Scanner(System.in);
            do {
                // Get the slide to update
                System.out.println("Slide number to be updated (slides start at 0): ");
                int slideNumber = sc.nextInt();

                // Consume the newline character
                sc.nextLine();

                // Get the chart to update
                System.out.println("Chart name: ");
                String chartName = sc.nextLine();

                XSLFChart chart = findChartInSlide(ppt.getSlides().get(slideNumber), chartName);

                // Get the sheet name of the embedded workbook to update
                System.out.println("Sheet name: ");
                String sheetName = sc.nextLine();

                // replace the data in the old embedded workbook with the new data from sheet
                copyWorkbook(chart, workbook, sheetName);

                updateChart(workbook.getSheet(sheetName), chart);

                System.out.println("Enter q to quit or any other key to continue: ");
            } while (!sc.nextLine().equals("q"));

            FileOutputStream out = new FileOutputStream("/Users/rayyanwaris/Downloads/DataChartAutomation/output.pptx");
            ppt.write(out);
            out.close();
            pptStream.close();
            ppt.close();
            sc.close();
            System.out.println("New PPT @ /Users/rayyanwaris/Downloads/DataChartAutomation/output.pptx");




        } catch (FileNotFoundException e) {
            System.out.println(e);
        }
    }
}