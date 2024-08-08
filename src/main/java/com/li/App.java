package com.li;

import com.li.utils.word.Placeholder;
import com.li.utils.word.PoiWordUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws IOException, InvalidFormatException {
        testWord();
    }

    private static void testWord() throws IOException, InvalidFormatException {
        XWPFDocument word = new XWPFDocument(new FileInputStream("E:\\home\\双抢气象服务专报.docx"));
        List<Placeholder> list = new ArrayList<>();

        Placeholder p = new Placeholder();
        p.setKey("${title}");
        p.setValue("江西省双抢气象服务专报");
        p.setType(1);
        list.add(p);

        Placeholder p2 = new Placeholder();
        p2.setKey("${pic1}");
        p2.setValue("https://www.toopic.cn/public/uploads/small/170426433974170426433973.jpg");
        p2.setType(2);
        list.add(p2);

        Placeholder p3 = new Placeholder();
        p3.setKey("${part1_content1}");
        p3.setValue("7月1~7日，全省平均气温29.8°C,较常年同期明显偏高1.6°C; 日照时数全省平均57.7h,较常年同期偏多14.7h。" +
                "\n降水量全省平均 53.5毫米，接近常年同期；但降水时空分布不均，北多南少，贛北北 部偏多5成-2倍，贛中、赣南普遍偏少5成以上。（各气象要素分布 见图1~6）。");
        p3.setType(1);
        list.add(p3);

        Placeholder p4 = new Placeholder();
        p4.setKey("${table1_level11}");
        p4.setValue("全省1-2级");
        p4.setType(1);
        list.add(p4);

        Placeholder p5 = new Placeholder();
        p5.setKey("${table1_level12}");
        p5.setValue("https://www.toopic.cn/public/uploads/small/170426433974170426433973.jpg");
        p5.setType(2);
        list.add(p5);

        XWPFDocument word1 = PoiWordUtils.replaceByPlaceholder(word, list);
        File file = new File("E:\\home\\1.docx");
        word1.write(new FileOutputStream(file));
    }
}
