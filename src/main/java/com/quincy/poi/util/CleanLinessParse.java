package com.quincy.poi.util;

import com.quincy.poi.dto.*;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
/*解析 颗粒度分析仪--零件清洁度报告.word*/
public class CleanLinessParse {
    /*解析表格模板 每个表格都是固定的*/
   public static CleanlinessAnalysisData  parseCleanLinessAnalysisData(WordDocument wordDocument){
       CleanlinessAnalysisData cleanlinessAnalysisData=new CleanlinessAnalysisData();
       List<ParagraphValue> paragraphValues = wordDocument.getParagraphValues();
       /*1.解析批次号*/
       if(paragraphValues!=null&&paragraphValues.size()>0){
           for (ParagraphValue paragraphValue : paragraphValues) {
               if(paragraphValue.getParagraphText().contains("清洁度分析")){
                   String paragraphText = paragraphValue.getParagraphText();
                   String batchNo= paragraphText.split("\\s")[paragraphText.split("\\s").length-1];
                   cleanlinessAnalysisData.setBatchNo(batchNo);
                   break;
               }
           }
       }
       if(cleanlinessAnalysisData.getBatchNo()==null){
           throw new RuntimeException("批次号不能为空");
       }       /*2. 获取第1个表格的一些数据*/
       if(wordDocument.getTableNotKeyValueList()==null||wordDocument.getTableNotKeyValueList().size()==0){
           return cleanlinessAnalysisData;
       }
       TableNotKeyValue table1 = wordDocument.getTableNotKeyValueList().get(0);
       if(table1.getRows()!=null&&table1.getRows().size()>0&&table1.getRows().get(0).getRowCellValues().get(0).contains("Description of sample")){
           List<TableRowNotKeyValue> rows = table1.getRows();
           /*获取零件*/
           TableRowNotKeyValue row1 = rows.get(1);/*获取第一行*/
           String com = row1.getRowCellValues().get(1);/*获取第一行第二个cell的值*/
           cleanlinessAnalysisData.setComponent(com);
           /*获取零件号*/
           TableRowNotKeyValue row2 = rows.get(2);
           String comNo = row2.getRowCellValues().get(1);
           cleanlinessAnalysisData.setCompNo(comNo);
           /*获取制样编号*/
           TableRowNotKeyValue row3 = rows.get(3);
           String sampleNo = row3.getRowCellValues().get(1);
           cleanlinessAnalysisData.setSampleNo(sampleNo);
           /*分析日期*/
           String analysisDate= row3.getRowCellValues().get(3);
           cleanlinessAnalysisData.setAnalysisDate(analysisDate);
       }
       if(wordDocument.getTableNotKeyValueList().size()==1){
           return cleanlinessAnalysisData;
       }
        /*3.获取第2个表格的数据*/
        TableNotKeyValue table2 = wordDocument.getTableNotKeyValueList().get(1);
        if(table2.getRows()!=null&&table2.getRows().size()>0&&table2.getRows().get(0).getRowCellValues().get(0).contains("Extraction") ){
            /*滤膜上零件数量*/
            List<TableRowNotKeyValue> rows1 = table2.getRows();
            String filterComponents = rows1.get(1).getRowCellValues().get(3);
            /*获取重量*/
            cleanlinessAnalysisData.setFilterComponents(filterComponents);
            String weight = rows1.get(4).getRowCellValues().get(1);
            cleanlinessAnalysisData.setWeight(weight);
        }
       if(wordDocument.getTableNotKeyValueList().size()<=3){
           return cleanlinessAnalysisData;
       }
       /*获取第四个表格的数据*/
       TableNotKeyValue table4 = wordDocument.getTableNotKeyValueList().get(3);
       List<TableRowNotKeyValue> rows2 = table4.getRows();
       /* 最大金属*/
       String largestMetallicParticle = rows2.get(0).getRowCellValues().get(2);
       /*最大非金属*/
       String largestNonmetallicParticle = rows2.get(1).getRowCellValues().get(2);
       cleanlinessAnalysisData.setLargestMetallicParticle(largestMetallicParticle);
       cleanlinessAnalysisData.setLargestNonmetallicParticle(largestNonmetallicParticle);
       return cleanlinessAnalysisData;
   }
   /*解析 Detailed results明细。*/
    public static List<DetailedResultsData>  parseDetailedResults(WordDocument wordDocument){
        List<DetailedResultsData> resultsData=new ArrayList<>();
        TableNotKeyValue table =null;
        /*检查哪个表格里面有Detailed results 的表头  或者直接获取第七个表格直接解析值。 */
        if(wordDocument.getTableNotKeyValueList()!=null&&wordDocument.getTableNotKeyValueList().size()>0){
            tableNotKeyValue:for (TableNotKeyValue tableNotKeyValue : wordDocument.getTableNotKeyValueList()) {
                if(tableNotKeyValue.getRows().get(0).getRowCellValues().get(0).contains("Detailed results")){
                    table=tableNotKeyValue;
                    break tableNotKeyValue;
                }
            }
             /*获取表格的所有的行数据*/
            if(table!=null&&table.getRows()!=null){
                List<TableRowNotKeyValue> rows = table.getRows();
                if(rows.size()>1){
                    for (int i = 1; i < rows.size(); i++) {
                        DetailedResultsData detailedResultsData=new DetailedResultsData();
                        TableRowNotKeyValue tableRowNotKeyValue = rows.get(i);
                        List<String> rowCellValues = tableRowNotKeyValue.getRowCellValues();
                        detailedResultsData.setCode(rowCellValues.get(1));
                        detailedResultsData.setMetallic1(rowCellValues.get(3));
                        detailedResultsData.setMetallic2(rowCellValues.get(5));
                        detailedResultsData.setParticleSize(rowCellValues.get(0));
                        detailedResultsData.setTotal11(rowCellValues.get(2));
                        detailedResultsData.setTotal12(rowCellValues.get(4));
                        resultsData.add(detailedResultsData);
                    }
                }
            }

        }
        return  resultsData;
    }

    public static void main(String[] args) throws Exception {
        String fileName="F:\\test1.doc";
        File file = new File(fileName);
        FileInputStream fileInputStream = new FileInputStream(file);
        /*解析 word文档的数据 */
        WordDocument wordDocument = PoiUtils.getWordDocument(fileName, fileInputStream);
        /*解析 颗粒度分析仪--零件清洁度报告.word-解析表格1 表格2 表格4 */
        CleanlinessAnalysisData cleanlinessAnalysisData = parseCleanLinessAnalysisData(wordDocument);
        /*解析 颗粒度分析仪--零件清洁度报告.word-DetailedResults 表格7 */
        List<DetailedResultsData> resultsData = parseDetailedResults(wordDocument);
        System.out.println(cleanlinessAnalysisData);
        System.out.println(resultsData);
    }
}
