package ru.bis.javautil.xlsparse;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class Operation {
    static DateTimeFormatter formatterDt = DateTimeFormatter.ofPattern(Util.outDateFormat);

    String id; // Operation id (statement line number)
    String opNum; // Operation number
    LocalDate opDate; // Operation date
    long dtAmount; // Debit amount
    long crAmount; // Credit amount
    String ctrPartName; // Counterparty name
    String ctrPartAccount; // Counterparty account
    String purpose; // Purpose
    String ourAccount; // Our account

    String getCSVString() {
        return id + Util.fSep +
                "\"" + (opNum == null ? "" : Util.leftStr(opNum, 100)) + "\"" + Util.fSep +
                (opDate == null ? "" : opDate.format(formatterDt)) + Util.fSep +
                Util.long2str(dtAmount) + Util.fSep +
                Util.long2str(crAmount) + Util.fSep +
                "\"" + (ctrPartAccount == null ? "" : Util.leftStr(ctrPartAccount, 35)) + "\"" + Util.fSep +
                "\"" + (ctrPartName == null ? "" : Util.str2CSV(Util.leftStr(ctrPartName, 300))) + "\"" + Util.fSep +
                "\"" + (purpose == null ? "" : Util.str2CSV(Util.leftStr(purpose, 600))) + "\"" + Util.fSep +
                "\"" + (ourAccount == null ? "" : Util.leftStr(ourAccount, 35)) + "\"";
    }
}
