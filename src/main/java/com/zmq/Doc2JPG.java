package com.zmq;
import java.awt.Color;
import java.awt.FontMetrics;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;

import cn.hutool.core.util.NumberUtil;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

public class Doc2JPG {
    private static double DPI = 96D;
    public static void main(String[] args) throws FileNotFoundException, IOException, XmlException {
    // 加载 DOCX 文件

    int dpi = Toolkit.getDefaultToolkit().getScreenResolution();

    FileInputStream fis = new FileInputStream("test.docx");
    XWPFDocument document = new XWPFDocument(fis);
    document.getSettings().setZoomPercent(100);
    
    CTFonts ctFonts =  document.getStyle().getDocDefaults().getRPrDefault().getRPr().getRFontsList().get(0);

    List<IBodyElement> elements = document.getBodyElements();
    int y = 30;
    int width = 800;
    int contentWidth = 670;
    int height = 1850;
    BufferedImage image = new BufferedImage(width,height, BufferedImage.TYPE_INT_RGB);
    Graphics2D graphics = image.createGraphics();
    graphics.setColor(Color.WHITE);
    graphics.fillRect(0, 0, image.getWidth(), image.getHeight());            
    graphics.setColor(Color.BLACK);
    // graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
    graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
    graphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_LCD_HRGB);
    
    int x = 85;
    int lastSpace = 0;
    for (IBodyElement element : elements) {
        if (element instanceof XWPFParagraph) {
            // 处理段落
            XWPFParagraph paragraph = (XWPFParagraph) element;
            // paragraph.getRuns().get(0).getFontFamily();
            FontInfo getFont = getFont(document,paragraph.getStyleID(), ctFonts);
            ParagraphAlignment paragraphAlignment = paragraph.getAlignment();
            // graphics.setFont(convertPoiFontToAwtFont(sheet.getWorkbook().getFontAt(cellStyle.getFontIndexAsInt())));
            // 获取字体名称
            // 获取字体样式
            
        // 创建 AWT Font 对象
            graphics.setFont(new java.awt.Font(getFont.getFontName(), getFont.getFontStyle(), getFont.getSize()));


            FontMetrics fontMetrics = graphics.getFontMetrics();
            int fonth = fontMetrics.getHeight();

            if(CollectionUtils.isEmpty(paragraph.getRuns())){
                y += fonth + lastSpace;
                continue;
            }
            
            int lastx = x;
            for(XWPFRun run : paragraph.getRuns()){
                
                String value = run.getText(run.getTextPosition());
                if(value == null){
                    List<XWPFPicture> pictures = run.getEmbeddedPictures();
                    if(CollectionUtils.isEmpty(pictures)){
                        continue;
                    }
                    for (XWPFPicture picture : pictures) {
                        XWPFPictureData pictureData = picture.getPictureData();
                        String extension = pictureData.suggestFileExtension();
                        byte[] imageBytes = pictureData.getData();
                        Image img = ImageIO.read(new ByteArrayInputStream(imageBytes));


                        int imgHeight = (int)(picture.getDepth() * contentWidth / picture.getWidth());
                        int imagex = x;
                        int imagey = y + fonth;
                        graphics.drawImage(img,  imagex, imagey,contentWidth,imgHeight,null);
                        y += imgHeight + fonth;
                    }


                    continue;
                }
                boolean br = false;
                if(CollectionUtils.isNotEmpty( run.getCTR().getBrList())||value.contains("\n") || run == paragraph.getRuns().get(0)){
                    y += fonth + lastSpace;
                    lastx = x;
                    br = true;
                } 
                // 
                
                // 判断是否自动换行
                
                int lineWidth = fontMetrics.stringWidth(value);
                if(lastx + lineWidth > contentWidth){
                    // 自动换行
                    int currentX = lastx;
                    int fromIndex = 0;
                    for(int i = 0; i< value.length(); i++){
                        currentX = currentX + fontMetrics.charWidth(value.charAt(i)) ;
                        if(currentX * 0.95 > contentWidth){
                            graphics.drawString(value.substring(fromIndex,i), lastx, y);            
                            fromIndex = i;
                            lastx = x;
                            currentX = lastx;
                            y += fonth ;
                        }
                    }
                    if(fromIndex < value.length()){
                        lastx = x;
                        value = value.substring(fromIndex);
                        graphics.drawString(value, lastx, y);
                        lastx += fontMetrics.stringWidth(value);
                        // y += fonth ;
                    }
                }else{
                    if(paragraphAlignment != null && br == true && paragraphAlignment.equals(ParagraphAlignment.CENTER)){
                        lastx = lastx + (contentWidth - fontMetrics.stringWidth(paragraph.getText())) / 2;
                        graphics.drawString(value, lastx, y);
                        lastx += lineWidth; 
                    } else {
                        graphics.drawString(value, lastx, y);
                        lastx += lineWidth;
                    }
                }                
            }
            lastSpace = getFont.getSpace();
        } else if (element instanceof XWPFTable) {
            // 处理表格
            XWPFTable table = (XWPFTable) element;
            int lastx = x;
            FontInfo getFont = getFont(document,table.getStyleID(), ctFonts);
            graphics.setFont(new java.awt.Font(getFont.getFontName(), getFont.getFontStyle(), getFont.getSize()));
            FontMetrics fontMetrics = graphics.getFontMetrics();
            Map<XWPFTableRow,Integer> hm = getRowHeigh(graphics,table,fontMetrics);
            for (XWPFTableRow row : table.getRows()) {
                int fonth = hm.get(row);
                for (XWPFTableCell cell : row.getTableCells()) {
                    int cellWidth = ((BigInteger)cell.getCTTc().getTcPr().getTcW().getW()).intValue();
                    cellWidth = (int)((cellWidth/20) * (DPI/72));

                    // 获取高度
                    String value = cell.getText();

                    // top
                    graphics.drawLine(lastx, y, lastx + cellWidth, y);
                    // left
                    graphics.drawLine(lastx, y, lastx, y + fonth);
                    // right
                    graphics.drawLine(lastx+cellWidth, y, lastx + cellWidth, y + fonth);
                    // botton
                    graphics.drawLine(lastx, y + fonth, lastx + cellWidth, y + fonth);
                    
                    int lineWidth = fontMetrics.stringWidth(value);
                    int lineRow = (int)Math.ceil((double)lineWidth / cellWidth );
                    if(lineRow > 1){
                        // 自动换行
                        int currentX = lastx;
                        int fromIndex = 0;
                        int fontH = fontMetrics.getHeight();
                        int currentY = y + fontH;
                        for(int i = 0; i< value.length(); i++){
                            currentX = currentX + fontMetrics.charWidth(value.charAt(i));
                            if(currentX   > (lastx + cellWidth - 6)){
                                graphics.drawString(value.substring(fromIndex,i), lastx, currentY );            
                                fromIndex = i;
                                currentX = lastx;
                                // lastx = x;
                                currentY += fontH;
                            }
                        }
                        if(fromIndex < value.length()){
                                                    
                            value = value.substring(fromIndex);
                            graphics.drawString(value, lastx, currentY);
                            // lastx += fontMetrics.stringWidth(value);
                        }
                    }else {
                        graphics.drawString(value, lastx + 3, y + fonth - 3);
                    }
                    lastx += cellWidth;
                }
                lastx = x;
                y += fonth;
                System.out.println(); // 换行
            } 
        }  else {
            System.out.println(element);
        }
    }

                // 保存为JPG图片
                File outputFile = new File("output_with_images.jpg");
                ImageIO.write(image, "jpg", outputFile);
    
                // 释放资源
                graphics.dispose();
                document.close();
      
    fis.close();
    }

    private static Map<XWPFTableRow,Integer> getRowHeigh(Graphics2D graphics, XWPFTable table,FontMetrics fontMetrics) {
        Map<XWPFTableRow,Integer> map = new HashMap<>();
        for (XWPFTableRow row : table.getRows()) {
            int fonth = fontMetrics.getHeight();
            map.put(row, fonth);
            for (XWPFTableCell cell : row.getTableCells()) {
                int cellWidth = ((BigInteger)cell.getCTTc().getTcPr().getTcW().getW()).intValue();
                cellWidth = (int) ((cellWidth/20) * (DPI/72));

                // 获取高度
                String value = cell.getText();
                int lineWidth = fontMetrics.stringWidth(value) + 6;
                int currentFonth = (int)Math.ceil((double)lineWidth / (cellWidth )) * fonth;

                map.put(row, NumberUtil.max(fonth,currentFonth));
            }
        }
        return map;
    }

    @Data
    @lombok.Builder
    @AllArgsConstructor
    @NoArgsConstructor
    public static class FontInfo{
        // 只去中文字体
        String fontName;
        Integer size;
        Integer fontStyle;
        Integer space;

    }

    public static FontInfo getFont(XWPFDocument document,String id,CTFonts defaultFont) {
        int fontStyle = java.awt.Font.PLAIN;
        FontInfo fontInfo = FontInfo.builder().fontName(defaultFont.getEastAsia()).size(14).fontStyle(fontStyle).space(0).build();
        int size = -1;
        String fontName = null;
        boolean bold = false;
        boolean italic = false;
        int space = -1;
        if(StringUtils.isBlank(id)){
            return fontInfo;
        }
        // 获取字体
        String currentId = id;
        do{
            XWPFStyle style = document.getStyles().getStyle(currentId);
            if(style == null){
                continue;
            }
            currentId = style.getBasisStyleID();
            if(style.getCTStyle().getRPr() == null){
                continue;
            }
            List<CTFonts> fonts = style.getCTStyle().getRPr().getRFontsList();
            if(CollectionUtils.isNotEmpty( fonts) && StringUtils.isBlank( fontName)){
                CTFonts curreFonts = fonts.get(0);
                if(StringUtils.isBlank(curreFonts.getEastAsia())){
                    fontName = curreFonts.getEastAsia();
                }
            }
            
            List<CTHpsMeasure> measures = style.getCTStyle().getRPr().getSzList();
            if(CollectionUtils.isNotEmpty( measures) && size < 0){
                CTHpsMeasure curreFonts = measures.get(0);
                size = (int)(((BigInteger)curreFonts.getVal()).intValue() * 0.5 * 1.5)  ;
            }
            List<CTOnOff> ctOnOffs = style.getCTStyle().getRPr().getBList();
            if(CollectionUtils.isNotEmpty(ctOnOffs) && !bold){
                bold = true;
            }

            ctOnOffs = style.getCTStyle().getRPr().getIList();
            if(CollectionUtils.isNotEmpty(ctOnOffs) && !italic){
                italic = true;
            }

            CTSpacing ctSpacing = style.getCTStyle().getPPr().getSpacing();
            if(ctSpacing != null && space < 0){
                 space = ((BigInteger)ctSpacing.getLine()).intValue();
                 space = (int)((space/20) * (DPI/72));
            }

            break;
        }while(StringUtils.isNotBlank(currentId));

        if(StringUtils.isNotBlank(fontName)){
            fontInfo.setFontName(fontName);
        }
        if(size > 0){
            fontInfo.setSize(size);
        }

        if (bold) {
            fontStyle |= java.awt.Font.BOLD; // 加粗
        }
        if (italic) {
            fontStyle |= java.awt.Font.ITALIC; // 斜体
        }
        if(space > 0 ){
            fontInfo.setSpace(space);
        }
        fontInfo.setFontStyle(fontStyle);

        return fontInfo;
    }

    public static int emuToPixels(long emu) {
        return (int) (emu / 9525.0); // 1 EMU = 1/9525 px
    }
}


