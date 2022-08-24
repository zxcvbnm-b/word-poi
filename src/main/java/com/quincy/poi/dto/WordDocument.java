package com.quincy.poi.dto;

import java.util.List;
/*word文档DTO */
public class WordDocument {
   private List<TableNotKeyValue> tableNotKeyValueList;/*word文档所有表格数据*/
   private List<ParagraphValue> paragraphValues;/*word文档所有段落数据*/

   public List<TableNotKeyValue> getTableNotKeyValueList() {
      return tableNotKeyValueList;
   }

   public void setTableNotKeyValueList(List<TableNotKeyValue> tableNotKeyValueList) {
      this.tableNotKeyValueList = tableNotKeyValueList;
   }

   public List<ParagraphValue> getParagraphValues() {
      return paragraphValues;
   }

   public void setParagraphValues(List<ParagraphValue> paragraphValues) {
      this.paragraphValues = paragraphValues;
   }
}
