package com.quincy.poi.dto;
/*解析 DetailedResults的值*/
public class DetailedResultsData {
    private String particleSize;/*等级 */
    private String code;/*代码*/
    private String total11;/*每滤膜总数*/
    private String metallic1;/*每滤膜金属*/
    private String total12;/*每零件总数*/
    private String metallic2;/*每零件金属*/

    public String getParticleSize() {
        return particleSize;
    }

    public void setParticleSize(String particleSize) {
        this.particleSize = particleSize;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getTotal11() {
        return total11;
    }

    public void setTotal11(String total11) {
        this.total11 = total11;
    }

    public String getMetallic1() {
        return metallic1;
    }

    public void setMetallic1(String metallic1) {
        this.metallic1 = metallic1;
    }

    public String getTotal12() {
        return total12;
    }

    public void setTotal12(String total12) {
        this.total12 = total12;
    }

    public String getMetallic2() {
        return metallic2;
    }

    public void setMetallic2(String metallic2) {
        this.metallic2 = metallic2;
    }
}
