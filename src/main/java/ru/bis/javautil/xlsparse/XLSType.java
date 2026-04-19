package ru.bis.javautil.xlsparse;

public enum XLSType {
    XLS ("Old Excel format"),
    XLSX ("New Excel format");
    private final String name;
    XLSType(String name) {
        this.name = name;
    }
    public String getName() {
        return this.name;
    }
}
