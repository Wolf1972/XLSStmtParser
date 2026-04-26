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

    String getCSVString() {
        StringBuilder sb = new StringBuilder();
        sb.append(id); sb.append(Util.fSep);
        sb.append("\""); sb.append(opNum); sb.append("\""); sb.append(Util.fSep);
        sb.append(opDate.format(formatterDt)); sb.append(Util.fSep);
        sb.append(Util.long2str(dtAmount)); sb.append(Util.fSep);
        sb.append(Util.long2str(crAmount)); sb.append(Util.fSep);
        sb.append("\""); sb.append(ctrPartAccount); sb.append("\""); sb.append(Util.fSep);
        sb.append("\""); sb.append(Util.str2CSV(ctrPartName)); sb.append("\""); sb.append(Util.fSep);
        sb.append("\""); sb.append(Util.str2CSV(purpose)); sb.append("\"");
        return sb.toString();
    }
}
