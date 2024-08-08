package com.li.utils.word;

import com.li.utils.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.InputStream;
import java.util.List;
import java.util.Objects;

/**
 * POI word工具类
 * @author longxiong
 * @date 2024/8/7 14:22:53
 */
public class PoiWordUtils {

    private static Logger logger = LoggerFactory.getLogger(PoiWordUtils.class);

    // 图片显示最大宽度
    private static int PIC_WIDTH_MAX = 454;

    /**
     * 替换word中占位符内容
     * 内容中如需换行使用 \n
     * @param document word文档
     * @param placeholders 占位符信息集合
     * @return 替换数据后的word
     */
    public static XWPFDocument replaceByPlaceholder(XWPFDocument document, List<Placeholder> placeholders){
        replaceParagraph(document, placeholders);
        replaceTable(document, placeholders);
        return  document;
    }

    /**
     * 替换段落内容
     *
     * @param document word文档
     * @param placeholders 占位符信息集合
     */
    public static void replaceParagraph(XWPFDocument document, List<Placeholder> placeholders){
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            String paragraphText = paragraph.getText();

            if (StringUtils.isEmpty(paragraphText)) continue;

            for (Placeholder placeholder : placeholders) {
                String key = placeholder.getKey();

                if (!paragraphText.contains(key)) continue;

                for (XWPFRun cellRun : paragraph.getRuns()) {
                    replaceRun(cellRun, placeholder);

                    // 判定段落是否有图片，处理图片显示问题
                    if (cellRun.getEmbeddedPictures().size() > 0) {
                        int rule = paragraph.getSpacingLineRule().getValue();
                        // 如果段落行距为固定值，会导致图片显示不全
                        if (LineSpacingRule.EXACT.getValue() == rule) {
                            // 设置段落行距为单倍行距
                            paragraph.setSpacingBetween(1);
                        }
                    }
                }
            }
        }
    }

    /**
     * 替换表格中的数据
     * @param document word文档
     * @param placeholders 占位符信息集合
     */
    public static void replaceTable(XWPFDocument document, List<Placeholder> placeholders) {
        // 替换表格中的数据
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell tableCell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
                        String paragraphText = paragraph.getText();
                        if (StringUtils.isEmpty(paragraphText)) continue;

                        for (XWPFRun cellRun : paragraph.getRuns()) {
                            for(Placeholder placeholder : placeholders){
                                replaceRun(cellRun, placeholder, tableCell);
                            }

                        }
                    }
                }
            }
        }
    }

    /**
     * 替换XWPFRun内容
     * @param run XWPFRun对象
     * @param placeholder 占位数据
     */
    private static void replaceRun(XWPFRun run, Placeholder placeholder) {
        replaceRun(run, placeholder, null);
    }

    /**
     * 替换XWPFRun内容
     * @param run XWPFRun对象
     * @param placeholder 占位数据
     */
    private static void replaceRun(XWPFRun run, Placeholder placeholder, XWPFTableCell tableCell) {
        String text = run.getText(0);
        //获取占位符key值
        String key = placeholder.getKey();

        if (text != null && text.contains(key)) {
            //获取占位符类型
            int type = placeholder.getType();
            //获取对应key的value
            String value = placeholder.getValue();
            if (type == 1) { // 处理文本
                //把文本的内容，key替换为value
                text = text.replace(key, value);
                //把替换好的文本内容，保存到当前这个文本对象
                run.setText(text, 0);
            } else if (type == 2) { // 处理图片
                text = text.replace(key, "");
                run.setText(text, 0);

                if (StringUtils.isEmpty(value)) return;

                try {
                    InputStream imageStream = FileUtils.getInputStream(value);
                    if (imageStream == null) return;

                    // 通过BufferedImage获取图片信息
                    BufferedImage bufferedImage = ImageIO.read(imageStream);
                    int height = bufferedImage.getHeight();
                    int width = bufferedImage.getWidth();

                    // 表格插入图片时根据表格宽度缩放
                    if (tableCell != null) {
                        height = height * tableCell.getWidth() / (width * 20);
                        width = tableCell.getWidth() / 20;
                    } else {
                        CTSectPr sectPr = run.getParent().getDocument().getDocument().getBody().getSectPr();
                        // 页面宽度
                        int pgWith = Integer.parseInt(sectPr.getPgSz().getW().toString());
                        // 左边距
                        int left = Integer.parseInt(sectPr.getPgMar().getLeft().toString());
                        // 右边距
                        int right = Integer.parseInt(sectPr.getPgMar().getRight().toString());
                        // 内容宽度（px）
                        int w = (pgWith - left - right) / 20;
                        height = height * w / width;
                        width = w;
                    }

                    String fileName = value.substring(value.lastIndexOf("/") + 1);
                    // 重新获取流，之前的流已经被BufferedImage使用掉了
                    run.addPicture(FileUtils.getInputStream(value), XWPFDocument.PICTURE_TYPE_JPEG, fileName, Units.toEMU(width), Units.toEMU(height));
                    run.setStyle(run.getStyle());
                    run.addCarriageReturn();
                } catch (Exception e) {
                    logger.error("插入图片异常", e);
                }
            }
        }
    }

}
