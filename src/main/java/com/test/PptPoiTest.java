package com.test;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * PPT 测试
 *
 * @author junqiang.xiao@hand-china.com
 * @date 2018/7/11
 */
public class PptPoiTest {
    public static void main(String[] args) throws IOException {
        // creating empty presentation
        XMLSlideShow ppt = new XMLSlideShow();

        // taking the two presentations that are to be merged
        String file1 = "/Users/xjune/Downloads/阅读DIY-活动资料/Ox 7BU8 活动1-1.pptx";
        String file2 = "/Users/xjune/Downloads/阅读DIY-活动资料/Ox 7BU8 活动2-1.pptx";
        String[] inputs = { file1, file2 };

        for (String arg : inputs) {
            FileInputStream inputstream = new FileInputStream(arg);
            XMLSlideShow src = new XMLSlideShow(inputstream);

            for (XSLFSlide srcSlide : src.getSlides()) {

                // merging the contents
                ppt.createSlide().importContent(srcSlide);
            }
        }

        String file3 = "merge"+(new Date()).getTime()+".pptx";

        // creating the file object
        FileOutputStream out = new FileOutputStream(file3);

        // saving the changes to a file
        ppt.write(out);
        System.out.println("Merging done successfully");
        out.close();
    }
}
