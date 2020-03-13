package com.learn.poi.handler;

import com.learn.poi.PoiEntity;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * @author: webin
 * @date: 2020/3/12 22:16
 * @description: 事件处理器
 * @version: 0.0.1
 */
public class SheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

    private PoiEntity poiEntity;
    /**
     *  开始解析某一行的时候解析的方法
     * @param i 索引行
     */
    public void startRow(int i) {
        if (i > 0) {
            poiEntity = new PoiEntity();
        }
    }

    /**
     * 结束执行某一行执行的方法
     * @param i
     */
    public void endRow(int i) {
        System.out.println(poiEntity);
    }

    /**
     * 对行中的某一个表格进行处理
     * @param s             单元格名称
     * @param s1            当前单元格数据
     * @param xssfComment   批注
     */
    public void cell(String s, String s1, XSSFComment xssfComment) {
        if (poiEntity != null) {
            String pix = s.substring(0, 1);
            switch (pix) {
                case "A":
                    poiEntity.setId(s1);
                    break;
                case "B":
                    poiEntity.setBreast(s1);
                    break;
                case "C":
                    poiEntity.setAdipocytes(s1);
                    break;
                case "D":
                    poiEntity.setNegative(s1);
                    break;
                case "E":
                    poiEntity.setStaining(s1);
                    break;
                case "F":
                    poiEntity.setSupportive(s1);
                    break;
                default:
                    break;
            }
        }
    }
}
