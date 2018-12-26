package com.littlezheng.officeconverter;

import java.io.File;

import org.apache.commons.lang3.StringUtils;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;

/**
 * office文档转换
 * 
 * @author shoppingzh
 *
 */
public class OfficeConverter {
    
    /**
     * word文档转pdf
     * 
     * @param word      word文件路径及名称
     * @param pdf       pdf文件路径及名称
     * @return          是否转换成功
     */
    public static boolean word2Pdf(String word, String pdf) {
        doCheck(word, pdf);

        ComThread.InitSTA();
        ActiveXComponent app = null;
        try {
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", true);
            Dispatch docs = app.getProperty("Documents").toDispatch();
            Dispatch doc = Dispatch.call(docs, "Open", word, null, true).toDispatch();
            Dispatch.call(doc, "SaveAs", pdf, WdSaveFormat.PDF);
            Dispatch.call(doc, "Close", 0); // 0表示不保存修改
            return true;
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (app != null) {
                 Dispatch.call(app, "Quit");
            }
            ComThread.Release();
        }

        return false;
    }

    /**
     * excel转pdf
     * 
     * @param excel     excel源文件位置
     * @param pdf       pdf目标文件位置
     * @return          是否转换成功
     */
    public static boolean excel2Pdf(String excel, String pdf) {
        doCheck(excel, pdf);
        ComThread.InitSTA();
        ActiveXComponent app = null;
        try {
            app = new ActiveXComponent("Excel.Application");
            app.setProperty("Visible", true);
            Dispatch wbs = app.getProperty("Workbooks").toDispatch();
            Dispatch wb = Dispatch.call(wbs, "Open", excel, true).toDispatch();
            Dispatch.call(wb, "ExportAsFixedFormat", XLFixedFormatType.PDF, pdf, 0);
            Dispatch.call(wb, "Close");
            return true;
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (app != null) {
                Dispatch.call(app, "Quit");
            }
            ComThread.Release();
        }
        return false;
    }
    
    /**
     * ppt转pdf
     * 
     * @param ppt       ppt源文件位置
     * @param pdf       pdf目标文件位置
     * @return          是否转换成功
     */
    public static boolean ppt2Pdf(String ppt, String pdf) {
        doCheck(ppt, pdf);
        ComThread.InitSTA();
        ActiveXComponent app = null;
        try {
            app = new ActiveXComponent("PowerPoint.Application");
            app.setProperty("Visible", true);
            Dispatch psts = app.getProperty("Presentations").toDispatch();
            Dispatch pst = Dispatch.call(psts, "Open", ppt, true).toDispatch();
            Dispatch.call(pst, "SaveAs", pdf, PpSaveAsFileType.PDF);
            Dispatch.call(pst, "Close");
            return true;
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (app != null) {
                Dispatch.call(app, "Quit");
            }
            ComThread.Release();
        }

        return false;
    }

    private static void doCheck(String src, String dst) {
        if (src == null) {
            throw new RuntimeException("源文件名为空!");
        }
        File srcFile = new File(src);
        if (!srcFile.exists()) {
            throw new RuntimeException("源文件不存在!");
        }
        if (StringUtils.isBlank(dst)) {
            throw new RuntimeException("目标文件名为空!");
        }
        File dstFile = new File(dst);
        if (dstFile.exists()) {
            throw new RuntimeException("目标文件已存在!");
        }
    }

}
