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
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.openofficecombinedchart.chart.ChartDoc;
import com.sun.star.openofficecombinedchart.exceptions.RangeException;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author othman
 */
public class LineChart extends BaseChart {
    public LineChart(ChartDoc chart) throws NoSuchElementException, RangeException {
        super(chart);
    }

    @Override
    public void createChart() throws Exception {
        com.sun.star.chart2.XChartType lineChartType = chart.getChartType("com.sun.star.chart2.LineChartType");

        // add line chart type
        chart.getXChartTypeContainer().addChartType(lineChartType);

        com.sun.star.chart2.XDataSeriesContainer dataSeriesCnt =
            (com.sun.star.chart2.XDataSeriesContainer) UnoRuntime.queryInterface(
                com.sun.star.chart2.XDataSeriesContainer.class, lineChartType);
        com.sun.star.chart2.XDataSeries oSeries = chart.getDataSeries();

        dataSeriesCnt.addDataSeries(oSeries);
        setLineChartProps(oSeries);

        com.sun.star.chart2.data.XDataSink dataSink =
            (com.sun.star.chart2.data.XDataSink) UnoRuntime.queryInterface(com.sun.star.chart2.data.XDataSink.class,
                oSeries);

        // data sequences
        com.sun.star.chart2.data.XDataProvider        oDataProv        = chart.getXChartDoc().getDataProvider();
        com.sun.star.chart2.data.XLabeledDataSequence dLabeledSequence = chart.createLabeledSequence(oDataProv,
                                                                             chart.getCurrent_series(), "values-y");
        com.sun.star.chart2.data.XLabeledDataSequence oLabeledSequence = chart.createLegendLabel(oDataProv,
                                                                             chart.OPEN_SERIES, "values-y");
        com.sun.star.chart2.data.XLabeledDataSequence hLabeledSequence = chart.createLegendLabel(oDataProv,
                                                                             chart.HIGH_SERIES, "values-y");
        com.sun.star.chart2.data.XLabeledDataSequence lLabeledSequence = chart.createLegendLabel(oDataProv,
                                                                             chart.LOW_SERIES, "values-y");
        com.sun.star.chart2.data.XLabeledDataSequence[] aLabeledSequence = { dLabeledSequence, oLabeledSequence,
                hLabeledSequence, lLabeledSequence };

        dataSink.setData(aLabeledSequence);
    }

    private void setLineChartProps(com.sun.star.chart2.XDataSeries oSeries)
            throws UnknownPropertyException, PropertyVetoException, IllegalArgumentException, WrappedTargetException,
                   IllegalArgumentException {
        XPropertySet cProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, oSeries);

        cProp.setPropertyValue("Color", new Integer(0x333333));
    }
}
