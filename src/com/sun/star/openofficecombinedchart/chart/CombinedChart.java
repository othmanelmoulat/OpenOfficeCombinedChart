/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */



package com.sun.star.openofficecombinedchart.chart;

//~--- JDK imports ------------------------------------------------------------

import com.sun.star.container.NoSuchElementException;
import com.sun.star.openofficecombinedchart.exceptions.RangeException;
import com.sun.star.openofficecombinedchart.spreadsheet.SpreadsheetDoc;
import com.sun.star.openofficecombinedchart.spreadsheet.SpreadsheetInfo;
import com.sun.star.uno.Exception;
import com.sun.star.uno.XComponentContext;

/**
 *
 * @author othman
 */
public class CombinedChart {
    private SpreadsheetInfo   msheetInfo;
    private XComponentContext mxContext;
    private SpreadsheetDoc    mxDoc;

    public CombinedChart(XComponentContext xContext, SpreadsheetDoc xDoc, SpreadsheetInfo sheetInfo)
            throws NoSuchElementException, RangeException {
        this.mxContext  = xContext;
        this.mxDoc      = xDoc;
        this.msheetInfo = sheetInfo;
    }

    public void createChart() throws Exception {
        try {
            ChartDoc         chart  = new ChartDoc(this.msheetInfo, this.mxContext, this.mxDoc);
            CandleStickChart candle = new CandleStickChart(chart);

            candle.createChart();

            LineChart line = new LineChart(chart);

            for (int i = 0; i < chart.LINE_SERIES.length; i++) {
                chart.setCurrent_series(chart.LINE_SERIES[i]);
                line.createChart();
            }

            chart.setChartGrid();
            chart.raiseChartSheet();
        } catch (RangeException e) {
            throw new Exception(e.toString());
        } catch (Exception e) {
            throw new Exception(e.toString());
        }
    }
}
