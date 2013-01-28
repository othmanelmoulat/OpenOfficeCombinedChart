/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */



package com.sun.star.openofficecombinedchart.chart;

//~--- JDK imports ------------------------------------------------------------

/**
 *
 * @author othman
 */
import com.sun.star.container.NoSuchElementException;
import com.sun.star.openofficecombinedchart.exceptions.RangeException;
import com.sun.star.uno.Exception;

public abstract class BaseChart {
    protected ChartDoc chart;

    public BaseChart() {}

    public BaseChart(ChartDoc chart) throws NoSuchElementException, RangeException {
        this.chart = chart;
    }

    public abstract void createChart() throws Exception;
}
