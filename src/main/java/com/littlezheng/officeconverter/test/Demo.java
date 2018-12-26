package com.littlezheng.officeconverter.test;

import java.util.Scanner;

import com.littlezheng.officeconverter.OfficeConverter;

public class Demo {

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        System.out.print("请选择: 1.word转pdf 2.excel转pdf 3.ppt转pdf 0.退出\n");
        String cmd = scanner.nextLine();
        if("1".equals(cmd) || "2".equals(cmd) || "3".equals(cmd)) {
            System.out.print("请输入源文件路径: \t");
            String src = scanner.nextLine();
            System.out.print("请输入目标文件路径: \t");
            String dest = scanner.nextLine();
            convert(src, dest, cmd);
        }else {
            scanner.close();
            System.exit(0);
        }
    }

    private static void convert(String src, String dest, String cmd) {
        switch (cmd) {
        case "1":
            OfficeConverter.word2Pdf(src, dest);
            break;
        case "2":
            OfficeConverter.excel2Pdf(src, dest);
            break;
        case "3":
            OfficeConverter.ppt2Pdf(src, dest);
            break;
        default:
            break;
        }
    }
    
}
