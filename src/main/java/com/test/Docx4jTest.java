package com.test;

import com.plutext.merge.pptx.MergePptxException;
import com.plutext.merge.pptx.PresentationBuilder;
import com.plutext.merge.pptx.SlideRange;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.packages.PresentationMLPackage;

import java.io.File;
import java.util.Date;

/**
 * Demo class
 *
 * @author junqiang.xiao@hand-china.com
 * @date 2018/7/13
 */
public class Docx4jTest {
    public static void main(String[] args) throws Docx4JException, MergePptxException {
        String dir = "/Users/xjune/Downloads/阅读DIY-活动资料/";
        String[] deck = {
                "Ox 7BU8 活动1.pptx", "Ox 7BU8 活动2.pptx", "Ox 7BU8 活动3.pptx"
        };

        PresentationBuilder builder = new PresentationBuilder();

        // Uncomment to retain look/feel of each presentation
        //builder.setThemeTreatment(ThemeTreatment.RESPECT);

        for (int i = 0; i < deck.length; i++) {

            // Create a SlideRange representing the slides in this pptx
            SlideRange sr = new SlideRange(
                    (PresentationMLPackage) OpcPackage.load(
                            new File(dir + deck[i])));

            // Add the slide range to the output
            builder.addSlideRange(sr);
        }

        String file3 = "merge" + (new Date()).getTime() + ".pptx";
        builder.getResult().save(
                new File(file3));
    }
}
