package ru.bis.javautil.xlsparse;

public enum StatementType {
    BTB ("BTB Bank"),
    UNKNOWN ("Unknown");
    private final String name;
    StatementType(String name) {
        this.name = name;
    }
    public String getName() {
        return this.name;
    }
}
