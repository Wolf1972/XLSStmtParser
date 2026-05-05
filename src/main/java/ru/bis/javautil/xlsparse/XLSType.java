package ru.bis.javautil.xlsparse;

public enum XLSType {
    AUTO ("Auto defining Excel format (XLS | XLSX)"),
    XLS ("Old Excel format (XLS)"),
    XLSX ("New Excel format (XLSX)");
    private final String name;
    XLSType(String name) {
        this.name = name;
    }
    public String getName() {
        return this.name;
    }
}
