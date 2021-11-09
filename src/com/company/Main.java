package com.company;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class Main {

    public static void main(String[] args) throws BiffException, IOException, WriteException {
        // 我目前持有的低溢价可转债列表
        String[][] strMyLowPremium = new String[100][2];
        // 最新低溢价可转债列表
        String[][] strLastestLowPremium = new String[300][2];
        // 上次的VIP可转债列表
        String[][] strVipOld = new String[100][2];
        // 最新的VIP可转债列表
        String[][] strVipNew = new String[100][2];
        // 我目前持有的双低可转债列表
        String[][] strMyDoubleLow = new String[100][2];
        // 最新双低可转债列表
        String[][] strLastestDoubleLow = new String[300][2];

        ExcelTools excelTools = new ExcelTools();
        //先删除我的低溢价可转债持仓的其他品种，只保留可转债
        String[][] strMyTemp = new String[100][2];
        excelTools.readExcel(strMyTemp, "我的低溢价可转债持仓", 1, 0, 2);
        excelTools.DeleteNotConvertibleBond(strMyTemp);
        //然后得到了纯粹的可转债持仓列表
        excelTools.readExcel(strMyLowPremium, "我的低溢价可转债持仓",  1, 0, 2);
        //excelTools.PrintData(strMyLowPremium, "最终strMy", 0 , strMyLowPremium.length, 0, 2);

        excelTools.readExcel(strLastestLowPremium, "最新低溢价可转债排名",  1, 0, 2);
        //excelTools.PrintData(strLastestLowPremium, "strLastestLowPremium", 0 , strLastestLowPremium.length, 0, 2);

        excelTools.readExcel(strVipOld, "VIP轮动old",  1, 0, 2);
        //excelTools.PrintData(strVipOld, "strVipOld", 0 , strVipOld.length, 0, 2);
        excelTools.readExcel(strVipNew, "VIP轮动new",  1, 0, 2);
        //excelTools.PrintData(strVipNew, "strVipNew", 0 , strVipNew.length, 0, 2);

        excelTools.readExcel(strMyDoubleLow, "我的双低可转债持仓",  1, 0, 2);
        //excelTools.PrintData(strMyDoubleLow, "strMyDoubleLow", 0 , strMyDoubleLow.length, 0, 2);
        excelTools.readExcel(strLastestDoubleLow, "最新双低可转债排名",  2, 0, 2);
        //excelTools.PrintData(strLastestDoubleLow, "strLastestDoubleLow", 0 , strLastestDoubleLow.length, 0, 2);

        System.out.println("我的低溢价可转债持仓在最新的低溢价可转债列表里的排名:");
        for (int i = 0; i < strMyLowPremium.length; i++) {
            int rank = -1;
            int row = i+1;
            if (strMyLowPremium[i][1] != null) {
                for (int j = 0; j < strLastestLowPremium.length; j++) {
                    if (strLastestLowPremium[j][1] != null) {
                        if (strMyLowPremium[i][0].contains(strLastestLowPremium[j][0]))
                        {
                            rank = j+1;
                        }
                    }
                }
                //注意：raw是行，实际排名需要-1
                System.out.println(strMyLowPremium[i][1]+ "["+row+"]" + "在最新低溢价可转债的排名是" +"["+rank+"];");
            }
        }

        System.out.println("最新低溢价可转债排名前20里，我的低溢价可转债持仓未买入的");
        for (int i = 0; i < 20; i++) {
            int isExist = 0;
            int row = i+1;
            if (strLastestLowPremium[i][1] != null) {
                for (int j = 0; j < strMyLowPremium.length; j++) {
                    if (strMyLowPremium[j][1] != null) {
                        if (strMyLowPremium[j][0].contains(strLastestLowPremium[i][0]))
                        {
                            isExist = 1;
                        }
                    }
                }
                if (isExist != 0) {
                    //System.out.println("最新低溢价可转债排名前20已买:" + strLastestLowPremium[i][1]+ "["+row+"]");
                } else {
                    System.out.println(strLastestLowPremium[i][0] + strLastestLowPremium[i][1]+ "["+row+"];");
                }
            }
        }

        System.out.println("我的低溢价可转债持仓在最新的VIP可转债列表里的排名:");
        for (int i = 0; i < strMyLowPremium.length; i++) {
            int rank = -1;
            int row = i+1;
            if (strMyLowPremium[i][1] != null) {
                for (int j = 0; j < strVipNew.length; j++) {
                    if (strVipNew[j][1] != null) {
                        if (strMyLowPremium[i][0].contains(strVipNew[j][0]))
                        {
                            rank = j+1;
                        }
                    }
                }
                //注意：raw是行，实际排名需要-1
                System.out.println(strMyLowPremium[i][1]+ "["+row+"]" + "在最新VIP可转债的排名是" +"["+rank+"];");
            }
        }

        System.out.println("VIP轮动列表前20里，我的低溢价可转债持仓未买入的:");
        for (int i = 0; i < 20; i++) {
            int isExist = 0;
            int row = i+1;
            if (strVipNew[i][1] != null) {
                for (int j = 0; j < strMyLowPremium.length; j++) {
                    if (strMyLowPremium[j][1] != null) {
                        if (strMyLowPremium[j][0].contains(strVipNew[i][0]))
                        {
                            isExist = 1;
                        }
                    }
                }
                if (isExist != 0) {
                    //System.out.println("VIP轮动前20已买:" + strVipNew[i][1]+ "["+row+"]");
                } else {
                    System.out.println(strVipNew[i][0] + strVipNew[i][1]+ "["+row+"];");
                }
            }
        }

        System.out.println("VIP轮动列表前30里，我的低溢价可转债持仓未买入的:");
        for (int i = 0; i < 30; i++) {
            int isExist = 0;
            int row = i+1;
            if (strVipNew[i][1] != null) {
                for (int j = 0; j < strMyLowPremium.length; j++) {
                    if (strMyLowPremium[j][1] != null) {
                        if (strMyLowPremium[j][0].contains(strVipNew[i][0]))
                        {
                            isExist = 1;
                        }
                    }
                }
                if (isExist != 0) {
                    //System.out.println("VIP轮动前30已买:" + strVipNew[i][1]+ "["+row+"]");
                } else {
                    System.out.println(strVipNew[i][0] + strVipNew[i][1]+ "["+row+"];");
                }
            }
        }

        System.out.println("VIP轮动列表前40里，我的低溢价可转债持仓未买入的:");
        for (int i = 0; i < 40; i++) {
            int isExist = 0;
            int row = i+1;
            if (strVipNew[i][1] != null) {
                for (int j = 0; j < strMyLowPremium.length; j++) {
                    if (strMyLowPremium[j][1] != null) {
                        if (strMyLowPremium[j][0].contains(strVipNew[i][0]))
                        {
                            isExist = 1;
                        }
                    }
                }
                if (isExist != 0) {
                    //System.out.println("VIP轮动前40已买:" + strVipNew[i][1]+ "["+row+"]");
                } else {
                    System.out.println(strVipNew[i][0] + strVipNew[i][1]+ "["+row+"];");
                }
            }
        }

        System.out.println("VIP轮动列表前50里，我的低溢价可转债持仓未买入的:");
        for (int i = 0; i < 50; i++) {
            int isExist = 0;
            int row = i+1;
            if (strVipNew[i][1] != null) {
                for (int j = 0; j < strMyLowPremium.length; j++) {
                    if (strMyLowPremium[j][1] != null) {
                        if (strMyLowPremium[j][0].contains(strVipNew[i][0]))
                        {
                            isExist = 1;
                        }
                    }
                }
                if (isExist != 0) {
                    //System.out.println("VIP轮动前50已买:" + strVipNew[i][1]+ "["+row+"]");
                } else {
                    System.out.println(strVipNew[i][0] + strVipNew[i][1]+ "["+row+"];");
                }
            }
        }

        System.out.println("我的双低可转债持仓在最新的双低可转债列表里的排名:");
        for (int i = 0; i < strMyDoubleLow.length; i++) {
            int rank = -1;
            int row = i+1;
            if (strMyDoubleLow[i][1] != null) {
                for (int j = 0; j < strLastestDoubleLow.length; j++) {
                    if (strLastestDoubleLow[j][1] != null) {
                        if (strMyDoubleLow[i][0].contains(strLastestDoubleLow[j][0]))
                        {
                            rank = j+1;
                        }
                    }
                }
                //注意：raw是行，实际排名需要-2
                System.out.println(strMyDoubleLow[i][1]+ "["+row+"]" + "在最新双低可转债的排名是" +"["+rank+"];");
            }
        }

        System.out.println("最新双低可转债排名前20里，我的双低可转债持仓未买入的:");
        for (int i = 0; i < 20; i++) {
            int isExist = 0;
            int row = i+1;
            if (strLastestDoubleLow[i][1] != null) {
                for (int j = 0; j < strMyDoubleLow.length; j++) {
                    if (strMyDoubleLow[j][1] != null) {
                        if (strMyDoubleLow[j][0].contains(strLastestDoubleLow[i][0]))
                        {
                            isExist = 1;
                        }
                    }
                }
                if (isExist != 0) {
                    //System.out.println("最新双低可转债排名前20已买:" + strLastestDoubleLow[i][1]+ "["+row+"]");
                } else {
                    System.out.println(strLastestDoubleLow[i][1]+ "["+row+"];");
                }
            }
        }

    }

}
