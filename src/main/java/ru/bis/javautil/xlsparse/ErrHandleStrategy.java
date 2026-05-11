package ru.bis.javautil.xlsparse;

public enum ErrHandleStrategy {
    ALL ("All errors will be terminated process"),
    FORMAT ("Only checked format errors will be terminated process"),
    NONE ("Process won't be terminated on any errors");
    private final String name;
    ErrHandleStrategy (String name) {
        this.name = name;
    }
    public String getName() {
        return this.name;
    }
}
