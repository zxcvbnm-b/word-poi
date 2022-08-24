package com.quincy.poi.dto;
/*解析word得到的DTO 一个word文件一个这DTO*/
public class CleanlinessAnalysisData {
    private String component;/*配件号*/
    private String compNo;/*零件号*/
    private String batchNo;/*批次号*/
    private String sampleNo;/*制样编号*/
    private String AnalysisDate;/*分析日期*/
    private String weight;/*重量*/
    private String largestNonmetallicParticle;/*最大非金属颗粒*/
    private String largestMetallicParticle;/*最大金属颗粒*/
    private String filterComponents;/*滤膜上零件数量*/

    public String getComponent() {
        return component;
    }

    public void setComponent(String component) {
        this.component = component;
    }

    public String getCompNo() {
        return compNo;
    }

    public void setCompNo(String compNo) {
        this.compNo = compNo;
    }

    public String getBatchNo() {
        return batchNo;
    }

    public void setBatchNo(String batchNo) {
        this.batchNo = batchNo;
    }

    public String getSampleNo() {
        return sampleNo;
    }

    public void setSampleNo(String sampleNo) {
        this.sampleNo = sampleNo;
    }

    public String getAnalysisDate() {
        return AnalysisDate;
    }

    public void setAnalysisDate(String analysisDate) {
        AnalysisDate = analysisDate;
    }

    public String getWeight() {
        return weight;
    }

    public void setWeight(String weight) {
        this.weight = weight;
    }

    public String getLargestNonmetallicParticle() {
        return largestNonmetallicParticle;
    }

    public void setLargestNonmetallicParticle(String largestNonmetallicParticle) {
        this.largestNonmetallicParticle = largestNonmetallicParticle;
    }

    public String getLargestMetallicParticle() {
        return largestMetallicParticle;
    }

    public void setLargestMetallicParticle(String largestMetallicParticle) {
        this.largestMetallicParticle = largestMetallicParticle;
    }

    public String getFilterComponents() {
        return filterComponents;
    }

    public void setFilterComponents(String filterComponents) {
        this.filterComponents = filterComponents;
    }
}
