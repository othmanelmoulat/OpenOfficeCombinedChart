/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */



package com.sun.star.openofficecombinedchart.chart;

//~--- JDK imports ------------------------------------------------------------

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.drawing.FillStyle;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.openofficecombinedchart.exceptions.RangeException;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author othman
 */
public class CandleStickChart extends BaseChart {
    public CandleStickChart(ChartDoc chart) throws NoSuchElementException, RangeException {
        super(chart);
    }

    @Override
    public void createChart() throws Exception {
        com.sun.star.chart2.XChartType stockChartType = chart.getChartType("com.sun.star.chart2.CandleStickChartType");

        // add candlestick chart type
        chart.getChartTypeContainer().addChartType(stockChartType);
        setCandleStickChartProps(stockChartType);

        com.sun.star.chart2.XDataSeriesContainer dataSeriesCnt =
            (com.sun.star.chart2.XDataSeriesContainer) UnoRuntime.queryInterface(
                com.sun.star.chart2.XDataSeriesContainer.class, stockChartType);
        com.sun.star.chart2.XDataSeries oSeries = chart.getDataSeries();

        dataSeriesCnt.addDataSeries(oSeries);

        com.sun.star.chart2.data.XDataSink dataSink =
            (com.sun.star.chart2.data.XDataSink) UnoRuntime.queryInterface(com.sun.star.chart2.data.XDataSink.class,
                oSeries);

        // data sequences
        com.sun.star.chart2.data.XDataProvider        oDataProv        = chart.getXChartDoc().getDataProvider();
        com.sun.star.chart2.data.XLabeledDataSequence dLabeledSequence = chart.createLabeledSequence(oDataProv,
                                                                             chart.DATE_SERIES, "categories");
        com.sun.star.chart2.data.XLabeledDataSequence oLabeledSequence = chart.createLabeledSequence(oDataProv,
                                                                             chart.OPEN_SERIES, "values-first");
        com.sun.star.chart2.data.XLabeledDataSequence hLabeledSequence = chart.createLabeledSequence(oDataProv,
                                                                             chart.HIGH_SERIES, "values-max");
        com.sun.star.chart2.data.XLabeledDataSequence lLabeledSequence = chart.createLabeledSequence(oDataProv,
                                                                             chart.LOW_SERIES, "values-min");
        com.sun.star.chart2.data.XLabeledDataSequence cLabeledSequence = chart.createLabeledSequence(oDataProv,
                                                                             chart.CLOSE_SERIES, "values-last");
        com.sun.star.chart2.data.XLabeledDataSequence[] aLabeledSequence = { dLabeledSequence, oLabeledSequence,
                hLabeledSequence, lLabeledSequence, cLabeledSequence };

        dataSink.setData(aLabeledSequence);
    }

    private void setCandleStickChartProps(com.sun.star.chart2.XChartType type)
            throws UnknownPropertyException, PropertyVetoException, IllegalArgumentException, WrappedTargetException,
                   Exception {
        XPropertySet cProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, type);

        cProp.setPropertyValue("Japanese", true);
        cProp.setPropertyValue("ShowFirst", true);
        cProp.setPropertyValue("ShowHighLow", true);

        Object fillProps = chart.getXMCF().createInstanceWithContext("com.sun.star.drawing.FillProperties",
                               chart.getXContext());
        XPropertySet whiteDayProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, fillProps);

        whiteDayProp.setPropertyValue("FillStyle", FillStyle.SOLID);
        whiteDayProp.setPropertyValue("FillColor", new Integer(0x008000));
        cProp.setPropertyValue("WhiteDay", whiteDayProp);

        Object fillProps2 = chart.getXMCF().createInstanceWithContext("com.sun.star.drawing.FillProperties",
                                chart.getXContext());
        XPropertySet blackDayProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, fillProps2);

        blackDayProp.setPropertyValue("FillStyle", FillStyle.SOLID);
        blackDayProp.setPropertyValue("FillColor", new Integer(0xFF0000));
        cProp.setPropertyValue("BlackDay", blackDayProp);
    }
}
