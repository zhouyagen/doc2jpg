package com.zmq;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

import com.itextpdf.forms.PdfAcroForm;
import com.itextpdf.forms.fields.CheckBoxFormFieldBuilder;
import com.itextpdf.forms.fields.PdfButtonFormField;
import com.itextpdf.forms.fields.properties.CheckBoxType;
import com.itextpdf.forms.form.element.Button;
import com.itextpdf.forms.form.element.CheckBox;
import com.itextpdf.forms.form.element.Radio;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.ColorConstants;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.properties.FlexWrapPropertyValue;
import com.itextpdf.layout.properties.Property;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;



public class Doc2Pdf {

    private static double DPI = 96D;

    public static void main(String[] args) throws IOException, XmlException {
        String dest = "hello_world.pdf"; // 输出 PDF 文件路径
        // 创建 PdfWriter 实例
        PdfWriter writer = new PdfWriter(dest);
        // 创建 PdfDocument 实例
        PdfDocument pdfDoc = new PdfDocument(writer);
        // 创建 Document 实例
        Document document = new Document(pdfDoc);
        // 向文档中添加段落
        
        
        FileInputStream fis = new FileInputStream("test.docx");
        XWPFDocument xwpdocument = new XWPFDocument(fis);
        xwpdocument.getSettings().setZoomPercent(100);
        
        CTFonts ctFonts =  xwpdocument.getStyle().getDocDefaults().getRPrDefault().getRPr().getRFontsList().get(0);
        
        List<IBodyElement> elements = xwpdocument.getBodyElements();
        

        for (IBodyElement element : elements) {

            if (element instanceof XWPFParagraph) {
                // 处理段落
                XWPFParagraph paragraph = (XWPFParagraph) element;
                // paragraph.getRuns().get(0).getFontFamily();
                
                FontInfo getFont = getFont(xwpdocument,paragraph.getStyleID(), ctFonts);
                ParagraphAlignment paragraphAlignment = paragraph.getAlignment();
                
                PdfFont font = PdfFontFactory.createFont("STSong-Light","UniGB-UCS2-H");
                try{
                    font = PdfFontFactory.createFont(getFont.getFontName());
                }catch(Exception e){}
                Paragraph p  = new Paragraph(paragraph.getText());
                p.setFontSize(getFont.getSize());
                p.setFont(font);
                TextAlignment alignment = coverAlign(paragraphAlignment);
                p.setTextAlignment(alignment);
                document.add(p);
                
                for(XWPFRun run : paragraph.getRuns()){
                    List<XWPFPicture> pictures = run.getEmbeddedPictures();
                    if(CollectionUtils.isEmpty(pictures)){
                        continue;
                    }
                    for (XWPFPicture picture : pictures) {
                        XWPFPictureData pictureData = picture.getPictureData();
                        byte[] imageBytes = pictureData.getData();
                        Image img = new Image(ImageDataFactory.create(imageBytes));
                        img.setWidth(UnitValue.createPercentValue(100));
                        document.add(img);
                    }
                }

            } else if (element instanceof XWPFTable) {
                // 处理表格
                XWPFTable table = (XWPFTable) element;
                FontInfo getFont = getFont(xwpdocument,table.getStyleID(), ctFonts);
                int cellSize = 0;
                float[] array = new float[]{};
                for (XWPFTableRow row : table.getRows()) {
                    if(row.getTableICells().size() > cellSize){
                        cellSize = row.getTableICells().size();
                        array = new float[cellSize];
                        int total = 0;
                        for (XWPFTableCell cell : row.getTableCells()) {
                            total += ((BigInteger)cell.getCTTc().getTcPr().getTcW().getW()).intValue();
                        } 

                        int i = 0;
                        for (XWPFTableCell cell : row.getTableCells()) {
                             float v = ((BigInteger)cell.getCTTc().getTcPr().getTcW().getW()).floatValue() / total * 100;
                             array[i] = v;
                             i++;
                        } 
                    }
                } 

                Table pdftable = new Table(UnitValue.createPercentArray(array)).useAllAvailableWidth();
                pdftable.setFixedLayout();
                for (XWPFTableRow row : table.getRows()) {
                    int cellIndex = 0;
                    for (XWPFTableCell cell : row.getTableCells()) {
                        CTTcPr ctTcPr = cell.getCTTc().getTcPr();
                        
                        int mergeculumn = 1;
                        if(ctTcPr.isSetGridSpan()){
                             mergeculumn = ctTcPr.getGridSpan().getVal().intValue();
                        }

                        // int mergeculumn = 1;
                        if(ctTcPr.isSetVMerge()){
                             mergeculumn = ctTcPr.getVMerge().getVal().intValue();
                        }
                        // 获取高度
                        String value = cell.getText();
                        Cell c = new Cell(1, mergeculumn);
                        Paragraph textArea = new Paragraph().add(value);
                        PdfFont font = PdfFontFactory.createFont("STSong-Light","UniGB-UCS2-H");
                        try{
                            font = PdfFontFactory.createFont(getFont.getFontName());
                        }catch(Exception e){}
                        textArea.setFont(font).setFontSize(getFont.getSize());
                        c.setProperty(Property.FLEX_WRAP, FlexWrapPropertyValue.WRAP);
                        c.add(textArea);
                        pdftable.addCell(c);
                        cellIndex ++;
                    } 
                    
                }
                document.add(pdftable);
            }
                
        }
        xwpdocument.close();    

        Button button = new Button("submit");
        button.setValue("Submit");
        button.setInteractive(true);
        button.setBorder(new SolidBorder(2));
        button.setWidth(50);
        button.setBackgroundColor(ColorConstants.LIGHT_GRAY);
        
        document.add(button);
        Radio male = new Radio("male", "radioGroup");
        male.setChecked(false);
        male.setInteractive(true);
        male.setBorder(new SolidBorder(1));
        document.add(male);

        CheckBox checkBox2 = new CheckBox("i");
        checkBox2.setChecked(false);
        checkBox2.setInteractive(true);
        checkBox2.setBorder(new SolidBorder(1));
        checkBox2.setCheckBoxType(CheckBoxType.CHECK);
        checkBox2.setDestination("destination");
        checkBox2.setInteractive(false);
        document.add(checkBox2);
        // 关闭文档
        document.close();

        System.out.println("PDF 已生成！");
    }

    private static TextAlignment coverAlign(ParagraphAlignment paragraphAlignment) {
        TextAlignment alignment = TextAlignment.JUSTIFIED;
        switch (paragraphAlignment) {
            case CENTER:
                alignment = TextAlignment.CENTER;
                break;
            case LEFT:
                alignment = TextAlignment.LEFT;
                break;
            case RIGHT:
                alignment = TextAlignment.RIGHT;
                break;
            default:
                break;
        }
        return alignment;
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
