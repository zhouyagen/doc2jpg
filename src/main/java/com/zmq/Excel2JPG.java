package com.zmq;

import java.awt.Color;
import java.awt.FontMetrics;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
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

import cn.hutool.core.util.NumberUtil;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
/**
 * MainTest
 */
public class Excel2JPG {

    public static void main(String[] args) {
        xlsx2jpg("test.xlsx");
    }

    public static void xlsx2jpg(String fileName){
        try {
            // 读取Excel文件
            FileInputStream file = new FileInputStream(new File(fileName));
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);

            
            JPGRange range = getJPGRange(sheet);

            // 高度
            // 创建图像
            int x = 25;
            int y = 30;
            // 创建图片
            Graphics2D graphics = new BufferedImage(1,1, BufferedImage.TYPE_INT_RGB).createGraphics();
            Map<Integer,Integer> rowHeigh = getrowHeigh(workbook, sheet, graphics);
            BufferedImage image = createImage(sheet, rowHeigh,x,y,range);
            graphics = image.createGraphics();
            graphics.setColor(Color.WHITE);
            graphics.fillRect(0, 0, image.getWidth(), image.getHeight());            
            graphics.setColor(Color.BLACK);
            // graphics.setFont(font);

            // 画网格
            Map<Integer,Map<Integer,XyPoint>> cellXy = drawLine(sheet,graphics,x,y,rowHeigh,range);

            // 画文本内容
            drawString(sheet, graphics, cellXy);


            // 渲染Excel中的图片
            drawImage(sheet, graphics, cellXy);

            // 保存为JPG图片
            File outputFile = new File("output_with_images2.jpg");
            ImageIO.write(image, "jpg", outputFile);

            // 释放资源
            graphics.dispose();
            workbook.close();
            file.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void drawImage(Sheet sheet, Graphics2D graphics, Map<Integer, Map<Integer, XyPoint>> cellXy)
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
                        Image img = ImageIO.read(new ByteArrayInputStream(pictureData));

                        // 绘制图片到BufferedImage
                        XyPoint xyPoint = cellXy.get(anchor.getRow1()).get(new Integer(anchor.getCol1()));
                        int imgX = xyPoint.getX() + 10;
                        int imgY = xyPoint.getY() +  10;

                        int imgWidth = 150;
                        int imgHeight = 150;
                        graphics.drawImage(img,  imgX, imgY,imgWidth,imgHeight,null);
                        // graphics.drawImage(img, imgX, imgY, imgWidth * colWidth, imgHeight * rowHeight,null);
                    }
                }
            }
        }
    }

    private static BufferedImage createImage(Sheet sheet,Map<Integer,Integer> heightMap,int startX,int startY,JPGRange range){
        Integer height = heightMap.values().stream().mapToInt( e -> e).sum();
        Integer width = 0;
        for (int c = range.getMin();c < range.getMax(); c++) {
            width += Math.round(sheet.getColumnWidthInPixels(c));
        }
        BufferedImage image = new BufferedImage(width + startX*2, height+startY * 2, BufferedImage.TYPE_INT_RGB);
        return image;
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

    private static Map<Integer,Map<Integer,XyPoint>> drawLine(Sheet sheet, Graphics2D graphics, int startX, int startY, Map<Integer,Integer> rowHeigh,JPGRange range) {
        
        int x = startX;
        int y = startY;
        Map<Integer,Map<Integer,XyPoint>> cellXy = new HashMap<>();
        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            int rowHeight1 = rowHeigh.get(rowNum);
            if(cellXy.get(rowNum) == null){
                HashMap<Integer,XyPoint> culumn = new HashMap<>();
                cellXy.put(rowNum, culumn);
            }
            for (int ci = range.getMin();ci < range.getMax(); ci++) {
                Cell cell = row.getCell(ci);
                if(cell == null){                    
                    cell = row.createCell(ci);
                }
                int cwidth = Math.round( cell.getSheet().getColumnWidthInPixels(cell.getColumnIndex()));                     
                // 获取单元格格式
                CellStyle cellStyle = cell.getCellStyle();
                graphics.setFont(convertPoiFontToAwtFont(sheet.getWorkbook().getFontAt(cellStyle.getFontIndexAsInt())));
              
                CellRangeAddress getMergedCell = getMergedCell(cell.getSheet(), cell);
                if(cell.getCellStyle().getFillForegroundColorColor() != null){
                    byte[] rgb = ((XSSFColor)cell.getCellStyle().getFillForegroundColorColor()).getARGB();
                    Color c = new java.awt.Color(Byte.toUnsignedInt(rgb[1]), Byte.toUnsignedInt(rgb[2]), Byte.toUnsignedInt(rgb[3]));
                    graphics.setColor(c);
                    graphics.fillRect(x, y, cwidth, rowHeight1);
                    graphics.setColor(Color.BLACK);
                }
                BorderStyle top = cell.getCellStyle().getBorderTop();
                if(getMergedCell == null || cell.getRowIndex() == getMergedCell.getFirstRow()){
                    if(!top.equals(BorderStyle.NONE)){
                        graphics.drawLine(x, y, x + cwidth, y);
                    }
                }
                
                if(getMergedCell == null || cell.getRowIndex() == getMergedCell.getLastRow()){
                    BorderStyle button = cell.getCellStyle().getBorderBottom();
                    if(!button.equals(BorderStyle.NONE)){
                        graphics.drawLine(x, y + rowHeight1 , x +cwidth , y + rowHeight1);
                    }
                }
                if(getMergedCell == null || cell.getColumnIndex() == getMergedCell.getFirstColumn()){
                    // left
                    BorderStyle left = cell.getCellStyle().getBorderLeft();
                    if(!left.equals(BorderStyle.NONE)){
                        graphics.drawLine(x, y, x , y + rowHeight1);
                    }
                }
                if(getMergedCell == null || cell.getColumnIndex() == getMergedCell.getLastColumn()){
                    BorderStyle right = cell.getCellStyle().getBorderRight();
                    if(!right.equals(BorderStyle.NONE)){
                        graphics.drawLine(x+cwidth, y, x + cwidth, y + rowHeight1);
                    }
                }
                cellXy.get(rowNum).put(cell.getColumnIndex(), XyPoint.builder().x(x).y(y).build());
                x += cwidth;
            }
            y += rowHeight1;
            x = startX;
        }
        return cellXy;
    }


    private static void drawString(Sheet sheet, Graphics2D graphics,  Map<Integer,Map<Integer,XyPoint>> cellXy) {

        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            if(cellXy.get(rowNum) == null){
                HashMap<Integer,XyPoint> culumn = new HashMap<>();
                cellXy.put(rowNum, culumn);
            }
            for (Cell cell : row) {
                // 获取单元格格式
                CellStyle cellStyle = cell.getCellStyle();
                graphics.setFont(convertPoiFontToAwtFont(sheet.getWorkbook().getFontAt(cellStyle.getFontIndexAsInt())));
                
                FontMetrics fontMetrics = graphics.getFontMetrics();
                int fonth = fontMetrics.getHeight();
                List<String> cellValues = getCellValues(cell, fontMetrics);
                CellRangeAddress getMergedCell = getMergedCell(cell.getSheet(), cell);
                if(getMergedCell == null || cell.getRowIndex() == getMergedCell.getFirstRow()){
                    XyPoint xyPoint = cellXy.get(rowNum).get(cell.getColumnIndex());
                    for(int i = 0;i < cellValues.size(); i++){
                        // y -3 避免与线条重合
                        graphics.drawString(cellValues.get(i), xyPoint.getX()+5, xyPoint.getY() -3 + (fonth * (i+1)));
                    }
                }
            }
        }
    }

    private static List<String> getCellValues(Cell cell, FontMetrics fontMetrics) {
        String cellValue = cell.getStringCellValue();
        List<String> cellValues = new ArrayList<>();
        int widthincharacter = getWidth(cell);
        
        int len = cellValue.length();
        int rowWidth = 0;
        int fromIndex = 0;
        for(int i = 0; i< len ; i ++){
            rowWidth += fontMetrics.charWidth(cellValue.charAt(i));
            if(rowWidth > widthincharacter){
                // 换行，
                rowWidth = 0;
                cellValues.add(cellValue.substring(fromIndex,i));
                fromIndex = i;
            }
        }
        // 最后文本添加
        if(fromIndex > 0 && fromIndex <= len){
            cellValues.add(cellValue.substring(fromIndex,len));
        }
        if(cellValues.size() == 0 && len > 0){
            cellValues.add(cellValue);
        }
        return cellValues;
    }

    private static int getWidth(Cell cell) {
        CellRangeAddress getMergedCell = getMergedCell(cell.getSheet(), cell);
        int cwidth = (int)cell.getSheet().getColumnWidthInPixels(cell.getColumnIndex());                     
        int widthincharacter = cwidth - 15;
        if(getMergedCell != null && getMergedCell.getFirstColumn() == cell.getColumnIndex()){
            widthincharacter = 0;
            for ( int c = getMergedCell.getFirstColumn(); c <= getMergedCell.getLastColumn(); c ++){
                widthincharacter += cell.getSheet().getColumnWidthInPixels(c);
            }
            widthincharacter -= 15;
        }
        return widthincharacter;
    }

    private static Map<Integer,Integer> getrowHeigh(Workbook workbook, Sheet sheet, Graphics2D graphics) {
        Map<Integer,Integer> rowHeigh = new HashMap<>();
        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            for (Cell cell : row) {
                int rowHeight1 = (int)row.getHeightInPoints();        
                int margeHeight = 0;    
                CellRangeAddress getMergedCell = getMergedCell(cell.getSheet(), cell);
                if(getMergedCell != null){
                    for(int r = getMergedCell.getFirstRow() ; r <= getMergedCell.getLastRow(); r++){
                        margeHeight += cell.getSheet().getRow(r).getHeightInPoints();
                    }
                }
                // 获取单元格格式
                CellStyle cellStyle = cell.getCellStyle();
                graphics.setFont(convertPoiFontToAwtFont(workbook.getFontAt(cellStyle.getFontIndexAsInt())));
                
                // 获取字符宽度（以像素为单位）
                FontMetrics fontMetrics = graphics.getFontMetrics();
                int fonth = fontMetrics.getHeight();
   
                // 计算高度，超长必须换行
                String cellValue = cell.toString();
   
                int widthincharacter = getWidth(cell);
                int len = cellValue.length();
                int rowWidth = 0;
                int fontHeight = fonth;
                for(int i = 0; i< len ; i ++){
                    rowWidth += fontMetrics.charWidth(cellValue.charAt(i));
                    if(rowWidth > widthincharacter){
                        // 换行，
                        rowWidth = 0;
                        fontHeight += fonth;
                    }
                }

                if(fontHeight > margeHeight){
                    rowHeight1 = fontHeight;
                } 

                Integer exists = Optional.ofNullable(rowHeigh.get(rowNum)).orElse(0);

                rowHeigh.put(rowNum, NumberUtil.max(exists,rowHeight1));
            }
        }
        return rowHeigh;
    }

    public static java.awt.Font convertPoiFontToAwtFont(Font poiFont) {
        // 获取字体名称
        String fontName = poiFont.getFontName();

        // 获取字体大小
        int fontSize = poiFont.getFontHeightInPoints();

        // 获取字体样式
        int fontStyle = java.awt.Font.PLAIN;
        if (poiFont.getBold()) {
            fontStyle |= java.awt.Font.BOLD; // 加粗
        }
        if (poiFont.getItalic()) {
            fontStyle |= java.awt.Font.ITALIC; // 斜体
        }

        // 创建 AWT Font 对象
        return new java.awt.Font(fontName, fontStyle, fontSize);
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