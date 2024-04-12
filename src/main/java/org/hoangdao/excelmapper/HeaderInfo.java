package org.hoangdao.excelmapper;

public class HeaderInfo {
    private String name;
    private Integer index;

    public HeaderInfo(String name, Integer index) {
        this.name = name;
        this.index = index;
    }

    public HeaderInfo() {
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getIndex() {
        return index;
    }

    public void setIndex(Integer index) {
        this.index = index;
    }
}
