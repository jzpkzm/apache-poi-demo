package com.jizp.apachepoidemo.excel;

import java.util.Calendar;
import java.util.Date;

public class CommonUtils {
    /**
     * 日期转星期
     * @param date
     * @return
     */
    public static String dateToWeek(Date date) {
        if (date == null){
            return "";
        }
        // 获得一个日历
        Calendar cal = Calendar.getInstance();
        cal.setTime(date);
        // 指示一个星期中的某天。
        int w = cal.get(Calendar.DAY_OF_WEEK) - 1;
        if (w < 0)
            w = 0;
        return Constants.WEEK_DAYS[w];
    }

}
