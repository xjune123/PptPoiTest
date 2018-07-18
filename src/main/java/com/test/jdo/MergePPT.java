package com.test.jdo;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.List;

/**
 * 通过jacob组件调用COM接口完成PPT文件的合并。合并后图表数据不丢失，用户可正常手工修改。
 * 调用函数前将jacob.jar添加到项目中，同时将jcaob.dll拷贝到path路径下。
 * @author Elon
 * 
 */
public class MergePPT
{
    /**
     * 合并多个PPT文件。要求输出文件和合并文件均已存在，不创建新文件。
     * @param outPutPPTPath 合并后输出的文件路径。
     * @param mergePPTPathList 依次追加合并的文件。
     */
    public synchronized static void merge(String outPutPPTPath, List<String> mergePPTPathList)
    {
        // 启动 office PowerPoint程序
        ActiveXComponent pptApp = new ActiveXComponent("PowerPoint.Application");
        Dispatch.put(pptApp, "Visible", new Variant(true));   

        Dispatch presentations = pptApp.getProperty("Presentations").toDispatch();  

        // 打开输出文件
        Dispatch outputPresentation = Dispatch.call(presentations, "Open", outPutPPTPath, false,  
                false, true).toDispatch();

        // 循环添加合并文件
        for (String mergeFile : mergePPTPathList)
        {
            Dispatch mergePresentation = Dispatch.call(presentations, "Open", mergeFile, false, 
                false, true).toDispatch();

            Dispatch mergeSildes = Dispatch.get(mergePresentation, "Slides").toDispatch();
            @SuppressWarnings("deprecation")
            int mergePageNum = Dispatch.get(mergeSildes, "Count").toInt();

            // 关闭合并文件
            Dispatch.call(mergePresentation, "Close");

            Dispatch outputSlides = Dispatch.call(outputPresentation, "Slides").toDispatch();
            @SuppressWarnings("deprecation")
            int outputPageNum = Dispatch.get(outputSlides, "Count").toInt();

            // 追加待合并文件内容到输出文件末尾
            Dispatch.call(outputSlides, "InsertFromFile", mergeFile, outputPageNum, 1, mergePageNum);
        }

        // 保存输出文件，关闭退出PowerPonit.
        Dispatch.call(outputPresentation, "Save");
        Dispatch.call(outputPresentation, "Close");
        Dispatch.call(pptApp, "Quit");
    }
}