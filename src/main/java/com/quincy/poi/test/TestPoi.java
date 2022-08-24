package com.quincy.poi.test;


import java.io.*;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import javax.imageio.stream.FileImageInputStream;

public class TestPoi {
    /**
     * 读取word文件内容
     * @param path
     * @return buffer
     */
    public String readWord(String path) {
        String buffer = "";
        try {
            if (path.endsWith(".doc")) {
                XWPFDocument xdoc = new XWPFDocument(new FileInputStream(new File(path)));
                List<XWPFTable> tables = xdoc.getTables();
                for (XWPFTable table : tables) {
                    // 获取表格的行
                    List<XWPFTableRow> rows = table.getRows();
                    for (XWPFTableRow row : rows) {
                        // 获取表格的每个单元格
                        List<XWPFTableCell> tableCells = row.getTableCells();
                        for (XWPFTableCell cell : tableCells) {
                            // 获取单元格的内容
                            String text1 = cell.getText();
                            System.out.println(text1);
                        }
                    }
                }
            } else if (path.endsWith("docx")) {
            } else {
                System.out.println("此文件不是word文件！");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        return buffer;
    }
    public static void main(String[] args) throws IOException {
        FileImageInputStream stream=new FileImageInputStream(new File(""));

    }

}

