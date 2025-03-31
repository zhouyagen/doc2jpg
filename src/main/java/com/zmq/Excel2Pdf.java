package com.zmq;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.Color;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.properties.TransparentColor;
import com.itextpdf.layout.properties.UnitValue;

import cn.hutool.core.util.NumberUtil;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
/**
 * MainTest
 */
public class Excel2Pdf {

    public static void main(String[] args) {

        String dest = "hello_world1.pdf"; // 输出 PDF 文件路径


        xlsx2pdf("test.xlsx",dest);
    }

    public static void xlsx2pdf(String fileName,String dest ){
        
        
        
        try {
            PdfFont font = PdfFontFactory.createFont("STSong-Light","UniGB-UCS2-H");
            // 读取Excel文件
            FileInputStream file = new FileInputStream(new File(fileName));
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);
            
            JPGRange range = getJPGRange(sheet);

            
            // 高度
            // 创建图像
            int x = 25;
            int y = 30;
            // 创建 PdfWriter 实例
            PdfWriter writer = new PdfWriter(dest);
            // 创建 PdfDocument 实例
            PdfDocument pdfDoc = new PdfDocument(writer);
            // 创建 Document 实例
            
            int width = 0;
            for (int c = range.getMin();c < range.getMax(); c++) {
                width += Math.round(sheet.getColumnWidthInPixels(c));
            }
            int height = 842;
            Document document = new Document(pdfDoc,new PageSize(width +50, height));
            // Document document = new Document(pdfDoc);
            int size = range.getMax() - range.getMin();
            Table table = drawTable(sheet, range, width, document, size);

            drawImage(sheet, table, range);
            // 添加必须放在不在变更后，添加后再对cell变更是无效的
            document.add(table);
            // 释放资源
            document.close();
            workbook.close();
            file.close();
            System.out.println("PDF 已生成！");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static Table drawTable(Sheet sheet, JPGRange range, int width, Document document, int size) throws IOException {
        PdfFont font = PdfFontFactory.createFont("STSong-Light","UniGB-UCS2-H");
        float[] array = new float[size];
        Row row = sheet.getRow(sheet.getFirstRowNum());
        int i = 0;
        for (int ci = range.getMin();ci < range.getMax(); ci++) {
            Cell cell = row.getCell(ci);
            if(cell == null){                    
                cell = row.createCell(ci);
            }
            float percent = cell.getSheet().getColumnWidthInPixels(cell.getColumnIndex()) / width;
            array[i] = percent;
            i++;
        }
        Table pdftable = new Table(UnitValue.createPercentArray(array)).useAllAvailableWidth();
        pdftable.setFixedLayout();
        
        for (Row r : sheet) {  
            for (int ci = range.getMin();ci < range.getMax(); ci++) {
                Cell cell = r.getCell(ci);
                if(cell == null){
                    continue;
                }
                String cellValue = cell.getStringCellValue();
                CellRangeAddress mergAddress = getMergedCell(sheet,cell);
                int rowspan = 1;
                int columnspan = 1;
                
                if(mergAddress != null 
                && mergAddress.getFirstRow() == cell.getRowIndex() 
                && mergAddress.getFirstColumn() == cell.getColumnIndex()){
                    rowspan = mergAddress.getLastRow() - mergAddress.getFirstRow() + 1;
                    columnspan = mergAddress.getLastColumn() - mergAddress.getFirstColumn() +1;
                } else if (mergAddress != null){
                    continue;
                }
                
                
                com.itextpdf.layout.element.Cell c = new com.itextpdf.layout.element.Cell(rowspan, columnspan);
                CellStyle cellStyle = cell.getCellStyle();
                if(cellStyle.getFillForegroundColorColor() != null){
                    byte[] rgb = ((XSSFColor)cell.getCellStyle().getFillForegroundColorColor()).getRGB();
                    c.setBackgroundColor(getColor(rgb));
                    // c.setNextRenderer(new TransparentCellRenderer(cell, 0.5f));
                }
                Paragraph textArea = new Paragraph().add(cellValue);
                textArea.setFont(font).setFontSize(12);
                    // c.setProperty(Property.FLEX_WRAP, FlexWrapPropertyValue.WRAP);
                c.add(textArea);
                c.setMinHeight(UnitValue.createPointValue(15));
                if(StringUtils.isBlank(cellValue) && (cell.getColumnIndex() == (int)(range.getMax() -1) || mergAddress != null && range.getMax()-1 == mergAddress.getLastColumn())){
                    c.setBorderLeft(Border.NO_BORDER);
                } else if(StringUtils.isBlank(cellValue) && (cell.getColumnIndex() == range.getMin() || mergAddress != null && range.getMin() == mergAddress.getFirstColumn())){
                    c.setBorderRight(Border.NO_BORDER);
                } else if (StringUtils.isBlank(cellValue)){
                    c.setBorderLeft(Border.NO_BORDER);
                    c.setBorderRight(Border.NO_BORDER);
                }
                pdftable.addCell(c);
            }        
        }
        pdftable.setBorder(new SolidBorder(1));

        // document.add(pdftable);
        return pdftable;
    }


    public static DeviceRgb getColor(byte[] rgb){
        java.awt.Color awtColor = new java.awt.Color(Byte.toUnsignedInt(rgb[0]), Byte.toUnsignedInt(rgb[1]), Byte.toUnsignedInt(rgb[2]));
        // float alpha = awtColor.getAlpha() / 255f; // 透明度转换

        DeviceRgb baseColor = new DeviceRgb(awtColor.getRed(), awtColor.getGreen(), awtColor.getBlue());

        return baseColor;
    }


    private static void drawImage(Sheet sheet, Table pdftable,JPGRange range)
            throws IOException {
        if (sheet instanceof XSSFSheet) {
            XSSFSheet xssfSheet = (XSSFSheet) sheet;
            XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
            if (drawing != null) {
                for (XSSFShape shape : drawing.getShapes()) {
                    if (shape instanceof XSSFPicture) {
                        XSSFPicture picture = (XSSFPicture) shape;
                        XSSFClientAnchor anchor = picture.getPreferredSize();

                        // 获取图片数据
                        byte[] pictureData = picture.getPictureData().getData();
                        // Image img = ImageIO.read(new ByteArrayInputStream(pictureData));
                        // XWPFPictureData pictureData = picture.getPictureData();
                        // byte[] imageBytes = pictureData.getData();
                        Image img = new Image(ImageDataFactory.create(pictureData));
                        // img.setWidth(UnitValue.createPercentValue(100));

                        CellRangeAddress address = getMergedCell(sheet,sheet.getRow(anchor.getRow1()).getCell(anchor.getCol1()));

                        System.out.println(address.getFirstRow());
                        System.out.println(address.getFirstColumn());
                        com.itextpdf.layout.element.Cell c = pdftable.getCell(address.getFirstRow() - 1, address.getFirstColumn() - range.getMin());
                        c.getChildren().clear();
                        c.add(img);
                        
                    }
                }
            }
        }
    }



    private static JPGRange getJPGRange(Sheet sheet){
        int minRow = Integer.MAX_VALUE;
        int maxRow = 0;
        for (Row row : sheet) {
            minRow = NumberUtil.min(minRow,row.getFirstCellNum());
            maxRow = NumberUtil.max(maxRow,row.getLastCellNum());
        }
        // getLastCellNum 方法返回值不是基于零  Gets the index of the last cell contained in this row PLUS ONE
        return JPGRange.builder().max(maxRow-1).min(minRow).build();
            
    }


    private static CellRangeAddress getMergedCell(Sheet sheet, Cell cell) {
        int numMergedRegions = sheet.getNumMergedRegions();
        for (int i = 0; i < numMergedRegions; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                return range;
            }
        }
        return null;
    }

    @Data
    @Builder
    @AllArgsConstructor
    @NoArgsConstructor
    public static class XyPoint {
        private Integer x;
        private Integer y;

    }

    @Data
    @Builder
    @AllArgsConstructor
    @NoArgsConstructor
    public static class JPGRange {
        private Integer min;
        private Integer max;

    }
}