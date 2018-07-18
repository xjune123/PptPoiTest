package com.test.jdo;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.xmlbeans.XmlException;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class TestPPTMain {
    public static void main(String[] args) throws OpenXML4JException, IOException, XmlException {
        String outPutPPTPath = "merge" + (new Date()).getTime() + ".pptx";
        String file1 = "/Users/xjune/Downloads/阅读DIY-活动资料/Ox 7BU8 活动1-1.pptx";
        String file2 = "/Users/xjune/Downloads/阅读DIY-活动资料/Ox 7BU8 活动2-1.pptx";

        List<String> mergePPTPathList = new ArrayList<String>();
        mergePPTPathList.add(file1);
        mergePPTPathList.add(file2);

        MergePPT.merge(outPutPPTPath, mergePPTPathList);
        ;
    }
}