package com.quincy.poi.util;

import com.quincy.poi.dto.*;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PAPBinTable;
import org.apache.poi.hwpf.model.PAPX;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.*;

/*1.每个表格一个对象 每个表格跳过第一行不获取值（表头）针对*/
public class PoiUtils {
    //Docx表格及文本数据提取处理
    public static  List<TableNotKeyValue> getTablesTQDocx(XWPFDocument doc) {
        //处理word表格内容
        Iterator<XWPFTable> it = doc.getTablesIterator();
        List<TableNotKeyValue> tableNotKeyValues=new ArrayList<>();
        int tableIndex=0;
        HashMap<String, Object> map = new HashMap<String, Object>();
        while (it.hasNext()) {
            TableNotKeyValue tableNotKeyValue=new TableNotKeyValue();
            XWPFTable xwpfTable = it.next();
            // 读取每一行
            List<TableRowNotKeyValue> rowValue=new ArrayList<>();
            for (int i = 0; i < xwpfTable.getRows().size(); i++) {
                TableRowNotKeyValue tableRowNotKeyValue=new TableRowNotKeyValue();
                List<String> rowCellValues = tableRowNotKeyValue.getRowCellValues();
                XWPFTableRow row = xwpfTable.getRow(i);
                if (row != null) {
                    for (int j = 0; j < row.getTableCells().size(); j++) {
                        XWPFTableCell cell = row.getCell(j);
                        String cellText = cell.getText().replaceAll("\r","");;
                        rowCellValues.add(cellText);
                    }
                }
                tableRowNotKeyValue.setRowCellValues(rowCellValues);
                rowValue.add(tableRowNotKeyValue);
                tableNotKeyValue.setRows(rowValue);
            }
            tableNotKeyValue.setTableName(tableIndex+"");
            tableIndex++;
            tableNotKeyValues.add(tableNotKeyValue);
        }
        return tableNotKeyValues;
    }
    //Doc表格内容及文本内容提取  一个表格里面有很多行 ，每行有很多个cell
    public static  List<TableNotKeyValue> getTablesTQDoc(HWPFDocument document) {
        Range r = document.getRange();//区间
        HashMap<String,Object> map = new HashMap<String,Object>();
        //开始提取表格--每个表格的获取都不一样：所以需要区分
        TableIterator it=new TableIterator(r);
        List<TableNotKeyValue> tableNotKeyValues=new ArrayList<>();
        int tableIndex=0;
        while(it.hasNext()){
            TableNotKeyValue tableNotKeyValue=new TableNotKeyValue();
            Table tb=(Table)it.next();
            List<TableRowNotKeyValue> rowValue=new ArrayList<>();
            for(int i=0;i<tb.numRows();i++){
                TableRow tr=tb.getRow(i);
                TableRowNotKeyValue tableRowNotKeyValue=new TableRowNotKeyValue();
                List<String> rowCellValues = tableRowNotKeyValue.getRowCellValues();
                for(int j=0;j<tr.numCells();j++){
                    TableCell td=tr.getCell(j);
                    String cellText = td.text().
                            replaceAll("\u0007", "").
                            replaceAll("\u0001", "").
                            replaceAll("\r","");
                    rowCellValues.add(cellText);
                }
                tableRowNotKeyValue.setRowCellValues(rowCellValues);
                rowValue.add(tableRowNotKeyValue);
                tableNotKeyValue.setRows(rowValue);
            }
            tableNotKeyValue.setTableName(tableIndex+"");
            tableIndex++;
            tableNotKeyValues.add(tableNotKeyValue);
        }
        return tableNotKeyValues;
    }

    public static WordDocument getWordDocument(String fileName, InputStream inputStream) throws IOException {
        String fileSuffix = fileName.substring(fileName.lastIndexOf(".") + 1);
        if("doc".equals(fileSuffix)){
            HWPFDocument document = new HWPFDocument(inputStream);
           return getWordDocumentDoc(fileName, document);
        }else if("docx".equals(fileSuffix)){
            XWPFDocument document= new XWPFDocument(inputStream);
           return  getWordDocumentDocx(fileName, document);
        }else{
            throw new RuntimeException("此文件不是word文件!");
        }

    }
    public static WordDocument getWordDocumentDoc(String fileName, HWPFDocument document) throws IOException {
        WordDocument wordDocument=new WordDocument();
        /*1.获取word中的表格*/
        List<TableNotKeyValue> tableNotKeyValue = getTablesTQDoc(document);
        wordDocument.setTableNotKeyValueList(tableNotKeyValue);
        /*2.获取word中的段落*/
        PAPBinTable paragraphTable = document.getParagraphTable();
        ArrayList<PAPX> paragraphs = paragraphTable.getParagraphs();
        Range range = document.getRange();
        int numP = range.numParagraphs();
        List<ParagraphValue> paragraphValues=new ArrayList<>();
        for (int i = 0; i < numP; ++i) {
            Paragraph p = range.getParagraph(i);
            ParagraphValue paragraphValue=new ParagraphValue();
            String text = p .text();
            paragraphValue.setParagraphText(text);
            paragraphValues.add(paragraphValue);
        }
        wordDocument.setParagraphValues(paragraphValues);
        return wordDocument;
    }
    public static WordDocument getWordDocumentDocx(String fileName, XWPFDocument document) throws IOException {
        WordDocument wordDocument=new WordDocument();
        /*1.获取word中的表格*/
        List<TableNotKeyValue> tableNotKeyValue = getTablesTQDocx(document);
        wordDocument.setTableNotKeyValueList(tableNotKeyValue);
        /*2.获取word中的段落*/
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        List<ParagraphValue> paragraphValues=new ArrayList<>();
        if(paragraphs!=null){
            /*需要获取到批次号：就是第一个段落的后面的数字    System.out.println(text.split("\\s")[text.split("\\s").length-1]); */
            for (XWPFParagraph paragraph : paragraphs) {
                ParagraphValue paragraphValue=new ParagraphValue();
                String text = paragraph.getText();
                paragraphValue.setParagraphText(text);
                paragraphValues.add(paragraphValue);
            }
            wordDocument.setParagraphValues(paragraphValues);
        }
        return wordDocument;
    }

    public static void main(String[] args) throws  Exception {
    /*    String fileName="F:\\test.doc";
        File file = new File(fileName);
        FileInputStream fileInputStream = new FileInputStream(file);
        WordDocument wordDocument = getWordDocument(fileName, fileInputStream);*/
        String fileName="F:\\test1.doc";
        File file = new File(fileName);
        FileInputStream fileInputStream = new FileInputStream(file);
        /*解析 word文档的数据 */
        WordDocument wordDocument = PoiUtils.getWordDocument(fileName, fileInputStream);
        /*解析 颗粒度分析仪--零件清洁度报告.word-解析表格1 表格2 表格4 */
        CleanlinessAnalysisData cleanlinessAnalysisData = CleanLinessParse.parseCleanLinessAnalysisData(wordDocument);
        /*解析 颗粒度分析仪--零件清洁度报告.word-DetailedResults 表格7 */
        List<DetailedResultsData> resultsData = CleanLinessParse.parseDetailedResults(wordDocument);
        System.out.println(cleanlinessAnalysisData);
        System.out.println(resultsData);
    }
}
