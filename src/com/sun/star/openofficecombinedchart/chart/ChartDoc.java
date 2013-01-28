/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */



package com.sun.star.openofficecombinedchart.chart;

//~--- JDK imports ------------------------------------------------------------

import com.sun.star.awt.Rectangle;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.chart.XAxisXSupplier;
import com.sun.star.chart.XDiagram;
import com.sun.star.chart2.XChartDocument;
import com.sun.star.chart2.XChartTypeContainer;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.document.XEmbeddedObjectSupplier;
import com.sun.star.drawing.LineStyle;
import com.sun.star.frame.XModel;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.openofficecombinedchart.exceptions.RangeException;
import com.sun.star.openofficecombinedchart.spreadsheet.*;
import com.sun.star.openofficecombinedchart.util.Utils;
import com.sun.star.sheet.XCellRangeAddressable;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetView;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCellRange;
import com.sun.star.table.XTableChart;
import com.sun.star.table.XTableCharts;
import com.sun.star.table.XTableChartsSupplier;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XInterface;

import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author othman
 */
public class ChartDoc {
    public int                CHART_HEIGHT = 15000;
    public int                CHART_WIDTH  = 25000;
    public int                CHART_X      = 500;
    public int                CHART_Y      = 1000;
    protected int             SIZE_LIMIT   = 500;
    public String             CLOSE_SERIES;
    public String             DATE_SERIES;
    public String             HIGH_SERIES;
    public String[]           LINE_SERIES;
    public String             LOW_SERIES;
    public int                MAX_ROW;
    public int                MIN_ROW;
    public String             OPEN_SERIES;
    private String            current_series;
    protected SpreadsheetInfo sheetInfo;

    // protected XTableChart xTableChart;
    protected com.sun.star.chart2.XChartDocument xChartDoc;

    // private ArrayList chartSheetList = new ArrayList();//
    private XSpreadsheet                              xChartSheet;
    protected com.sun.star.chart2.XChartTypeContainer xChartTypeContainer;
    protected XComponentContext                       xContext;
    protected SpreadsheetDoc                          xDoc;
    protected XMultiComponentFactory                  xMCF;

    public ChartDoc(SpreadsheetInfo sheetInfo, XComponentContext xContext, SpreadsheetDoc xDoc)
            throws RangeException, NoSuchElementException, WrappedTargetException {
        this.sheetInfo = sheetInfo;
        this.xContext  = xContext;
        this.xDoc      = xDoc;
        this.init();
    }

    private void init() throws RangeException, NoSuchElementException, WrappedTargetException {

        // load selected data series
        loadDataSheetInfo(sheetInfo);

        // get service manager
        xMCF = xContext.getServiceManager();

        // create or get a chart document
        xChartDoc = this.getChartDocument();

        // get chart type container
        xChartTypeContainer = this.getChartTypeContainer();

        // remove default bar chart type
        removeChartType(xChartTypeContainer);
    }

    protected void loadDataSheetInfo(SpreadsheetInfo sheetInfo) throws RangeException {
        if (sheetInfo == null) {
            throw new RangeException("Fatal Error!");
        }

        String[] seriesHeaders = sheetInfo.getColumnHeader();

        if ((seriesHeaders == null) || (seriesHeaders.length < 5)) {
            throw new RangeException(Utils.ERROR_MSG);
        }

        this.MIN_ROW      = sheetInfo.getMin_row();
        this.MAX_ROW      = sheetInfo.getMax_row();
        this.DATE_SERIES  = seriesHeaders[0];
        this.OPEN_SERIES  = seriesHeaders[1];
        this.HIGH_SERIES  = seriesHeaders[2];
        this.LOW_SERIES   = seriesHeaders[3];
        this.CLOSE_SERIES = seriesHeaders[4];

        int j = 0;

        this.LINE_SERIES = new String[seriesHeaders.length - 5];

        for (int i = 5; i < seriesHeaders.length; i++) {
            this.LINE_SERIES[j++] = seriesHeaders[i];
        }
    }

    public com.sun.star.chart2.XChartDocument getChartDocument() {

        // if (this.xChartDoc == null) {
        this.xChartDoc = this.insertChart(xContext, xDoc);

        // }
        return xChartDoc;
    }

    public com.sun.star.chart2.XChartType getChartType(String type) throws Exception {
        Object                         object    = xMCF.createInstanceWithContext(type, xContext);
        com.sun.star.chart2.XChartType chartType =
            (com.sun.star.chart2.XChartType) UnoRuntime.queryInterface(com.sun.star.chart2.XChartType.class, object);

        return chartType;
    }

    public com.sun.star.chart2.XChartTypeContainer getChartTypeContainer() {
        if (this.xChartDoc == null) {
            return null;
        }

        // if (this.xChartTypeContainer == null) {

        /* let the Calc document create a data provider, set it at the chart */
        com.sun.star.chart2.data.XDataProvider oDataProv = xChartDoc.getDataProvider();
        com.sun.star.chart2.XDiagram           oDiagram  = xChartDoc.getFirstDiagram();

        // insert a coordinate system into the diagram
        com.sun.star.chart2.XCoordinateSystemContainer oCoordSysCnt =
            (com.sun.star.chart2.XCoordinateSystemContainer) UnoRuntime.queryInterface(
                com.sun.star.chart2.XCoordinateSystemContainer.class, oDiagram);
        com.sun.star.chart2.XCoordinateSystem[] oCoordSys = oCoordSysCnt.getCoordinateSystems();

        // Query the coordinate system for interface com::sun::star::chart2::XChartTypeContainer
        xChartTypeContainer = (com.sun.star.chart2.XChartTypeContainer) UnoRuntime.queryInterface(
            com.sun.star.chart2.XChartTypeContainer.class, oCoordSys[0]);

        // }
        return xChartTypeContainer;
    }

    public com.sun.star.chart2.XDataSeries getDataSeries() throws Exception {
        Object                          object  = xMCF.createInstanceWithContext("com.sun.star.chart2.DataSeries",
                                                      xContext);
        com.sun.star.chart2.XDataSeries oSeries =
            (com.sun.star.chart2.XDataSeries) UnoRuntime.queryInterface(com.sun.star.chart2.XDataSeries.class, object);

        return oSeries;
    }

    public com.sun.star.chart2.data.XLabeledDataSequence createLabeledSequence(
            com.sun.star.chart2.data.XDataProvider oDataProv, String columnH, String role_values)
            throws Exception {
        com.sun.star.chart2.data.XDataSequence oSequence = oDataProv.createDataSequenceByRangeRepresentation("$"
                                                               + sheetInfo.getSheetName() + ".$" + columnH + "$"
                                                               + MIN_ROW + ":$" + columnH + "$" + MAX_ROW);
        XPropertySet cProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, oSequence);

        cProp.setPropertyValue("Role", role_values);

        Object                                        object           =
            xMCF.createInstanceWithContext("com.sun.star.chart2.data.LabeledDataSequence", xContext);
        com.sun.star.chart2.data.XLabeledDataSequence oLabeledSequence =
            (com.sun.star.chart2.data.XLabeledDataSequence) UnoRuntime.queryInterface(
                com.sun.star.chart2.data.XLabeledDataSequence.class, object);

        oLabeledSequence.setValues(oSequence);

        com.sun.star.chart2.data.XDataSequence labelSequence = oDataProv.createDataSequenceByRangeRepresentation("$"
                                                                   + sheetInfo.getSheetName() + ".$" + columnH + "$"
                                                                   + 1);

        oLabeledSequence.setLabel(labelSequence);

        return oLabeledSequence;
    }

    public com.sun.star.chart2.data.XLabeledDataSequence createLegendLabel(
            com.sun.star.chart2.data.XDataProvider oDataProv, String columnH, String role_values)
            throws Exception {

        // com.sun.star.chart2.data.XDataSequence oSequence = oDataProv.createDataSequenceByRangeRepresentation("$" + sheetInfo.getSpreadsheetName() + ".$" + columnH + "$" + MIN_ROW + ":$" + columnH + "$" + MAX_ROW);
        com.sun.star.chart2.data.XDataSequence labelSequence = oDataProv.createDataSequenceByRangeRepresentation("$"
                                                                   + sheetInfo.getSheetName() + ".$" + columnH + "$"
                                                                   + 1);
        XPropertySet cProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, labelSequence);

        cProp.setPropertyValue("Role", role_values);

        Object                                        object           =
            xMCF.createInstanceWithContext("com.sun.star.chart2.data.LabeledDataSequence", xContext);
        com.sun.star.chart2.data.XLabeledDataSequence oLabeledSequence =
            (com.sun.star.chart2.data.XLabeledDataSequence) UnoRuntime.queryInterface(
                com.sun.star.chart2.data.XLabeledDataSequence.class, object);

        oLabeledSequence.setValues(labelSequence);
        oLabeledSequence.setLabel(labelSequence);

        return oLabeledSequence;
    }

    /**
     * Loading an OpenOffice.org Calc document and getting a chart by name.
     * @param stringFileName Name of the OpenOffice.org Calc document which should
     *                       be loaded.
     * @param stringChartName Name of the chart which should get a new chart type.
     */
    public XTableChart getTableChart(SpreadsheetDoc xDoc, XSpreadsheet xSheet) {
        XTableChart xTableChart = null;

        try {
            XTableChartsSupplier xTableChartsSupplier =
                (XTableChartsSupplier) UnoRuntime.queryInterface(XTableChartsSupplier.class, xSheet);
            XIndexAccess xIndexAccess = (XIndexAccess) UnoRuntime.queryInterface(XIndexAccess.class,
                                            xTableChartsSupplier.getCharts());

            xTableChart = (XTableChart) UnoRuntime.queryInterface(XTableChart.class, xIndexAccess.getByIndex(0));
        } catch (Exception exception) {
            exception.printStackTrace();
        }

        // }
        return xTableChart;
    }

    /**
     *
     * initialize a new chart
     */
    public com.sun.star.chart2.XChartDocument insertChart(XComponentContext xContext, SpreadsheetDoc xDoc) {
        xChartSheet = null;

        try {
            xChartSheet = xDoc.insertSheet();
        } catch (NoSuchElementException ex) {
            Logger.getLogger(ChartDoc.class.getName()).log(Level.SEVERE, null, ex);

            return null;
        } catch (WrappedTargetException ex) {
            Logger.getLogger(ChartDoc.class.getName()).log(Level.SEVERE, null, ex);

            return null;
        }

        XSpreadsheet xSheet = xDoc.getActiveSheet();

        // insert new chart.
        // get the CellRange which holds the data for the chart and its RangeAddress
        // get the TableChartSupplier from the sheet and then the TableCharts from it.
        // add a new chart based on the data to the TableCharts.
        Rectangle oRect = new Rectangle();

        oRect.X = this.CHART_X;
        oRect.Y = this.CHART_Y;

        // adjustChartSize();
        oRect.Width  = this.CHART_WIDTH;
        oRect.Height = this.CHART_HEIGHT;

        XCellRange oRange = (XCellRange) UnoRuntime.queryInterface(XCellRange.class, xSheet);

        // System.out.println(sheetInfo.getRangeString());
        XCellRange            myRange    = oRange.getCellRangeByName(sheetInfo.getRangeString());
        XCellRangeAddressable oRangeAddr =
            (XCellRangeAddressable) UnoRuntime.queryInterface(XCellRangeAddressable.class, myRange);
        CellRangeAddress   myAddr = oRangeAddr.getRangeAddress();
        CellRangeAddress[] oAddr  = new CellRangeAddress[1];

        oAddr[0] = myAddr;

        XTableChartsSupplier oSupp = (XTableChartsSupplier) UnoRuntime.queryInterface(XTableChartsSupplier.class,
                                         xChartSheet);
        XTableCharts oCharts = oSupp.getCharts();

        System.out.println("charts list before insert : ");
        Utils.print(this.getChartElementNames(oCharts));

        String _sName = xDoc.getSheetName();

        System.out.println(" chart name=" + _sName);
        oCharts.addNewByName(_sName, oRect, oAddr, true, true);
        System.out.println(_sName + " Chart Inserted successfully");
        System.out.println("charts list after insert : ");
        Utils.print(this.getChartElementNames(oCharts));

        XTableChart             xTableChart          = getTableChart(xDoc, xChartSheet);
        XEmbeddedObjectSupplier xEmbeddedObjSupplier =
            (XEmbeddedObjectSupplier) UnoRuntime.queryInterface(XEmbeddedObjectSupplier.class, xTableChart);
        XInterface                         xInterface = xEmbeddedObjSupplier.getEmbeddedObject();
        com.sun.star.chart2.XChartDocument oChartDoc  =
            (com.sun.star.chart2.XChartDocument) UnoRuntime.queryInterface(com.sun.star.chart2.XChartDocument.class,
                xInterface);

        return oChartDoc;
    }

    private String[] getChartElementNames(XTableCharts oCharts) {
        if (oCharts == null) {
            return null;
        }

        return oCharts.getElementNames();
    }

    public boolean removeChartType(com.sun.star.chart2.XChartTypeContainer chartTypeContainer)
            throws NoSuchElementException {
        com.sun.star.chart2.XChartType[] chartType = chartTypeContainer.getChartTypes();

        for (int i = 0; i < chartType.length; i++) {
            if (chartType[i].getChartType().equals("com.sun.star.chart2.ColumnChartType")) {
                chartTypeContainer.removeChartType(chartType[i]);
                System.out.println("removed chart type=com.sun.star.chart2.ColumnChartType");

                return true;
            }
        }

        return false;
    }

    public void adjustChartSize() {
        int size = sheetInfo.getMax_row() - sheetInfo.getMin_row();

        if (size >= this.SIZE_LIMIT) {
            this.CHART_WIDTH  *= 10;
            this.CHART_HEIGHT *= 10;
        }
    }

    public void setChartGrid()
            throws RuntimeException, UnknownPropertyException, PropertyVetoException,
                   com.sun.star.lang.IllegalArgumentException, WrappedTargetException, Exception {
        com.sun.star.chart.XChartDocument oChartDoc =
            (com.sun.star.chart.XChartDocument) UnoRuntime.queryInterface(com.sun.star.chart.XChartDocument.class,
                xChartDoc);
        XDiagram maDiagram = oChartDoc.getDiagram();

        // show x major grid
        XPropertySet aGridProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class,
                                     ((XAxisXSupplier) UnoRuntime.queryInterface(XAxisXSupplier.class, maDiagram)));

        if (aGridProp != null) {
            aGridProp.setPropertyValue("HasXAxisGrid", true);
        }

        // set x major grid props
        XPropertySet xGridProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class,
                                     ((XAxisXSupplier) UnoRuntime.queryInterface(XAxisXSupplier.class,
                                         maDiagram)).getXMainGrid());

        if (xGridProp != null) {

            /*
             *  LineDash aDash = new LineDash();
             *
             * aDash.Style = DashStyle.ROUND;
             * aDash.Dots = 0;
             * aDash.DotLen = 10;
             * aDash.Dashes = 1;
             * aDash.DashLen = 50;
             * aDash.Distance = 10;
             * String dashName=Utils.createUniqueName(xDoc);
             * dashTable.insertByName(dashName, aDash);
             */
            xGridProp.setPropertyValue("LineColor", new Integer(0x999999));
            xGridProp.setPropertyValue("LineStyle", LineStyle.DASH);

            // xGridProp.setPropertyValue("LineDashName", dashName);
            xGridProp.setPropertyValue("LineWidth", new Integer(30));
        }
    }

    /**
     * Bring the sheet containing charts visually to the foreground
     */
    public void raiseChartSheet() {
        ((XSpreadsheetView) UnoRuntime.queryInterface(XSpreadsheetView.class,
                ((XModel) UnoRuntime.queryInterface(XModel.class,
                    xDoc.getSpreadSheetDocument())).getCurrentController())).setActiveSheet(xChartSheet);
    }

    public void lockControllers() throws RuntimeException {
        ((XModel) UnoRuntime.queryInterface(XModel.class, this.xChartDoc)).lockControllers();
    }

    // ____________________
    public void unlockControllers() throws RuntimeException {
        ((XModel) UnoRuntime.queryInterface(XModel.class, this.xChartDoc)).unlockControllers();
    }

    /*
     *  public void scaleChartXAxis() throws Exception {
     *
     *        Object object = xMCF.createInstanceWithContext(
     *      "com.sun.star.chart2.Axis", xContext);
     *      com.sun.star.chart2.XAxis xaxis = (com.sun.star.chart2.XAxis) UnoRuntime.queryInterface(
     *      com.sun.star.chart2.XAxis.class, object);
     *     // System.out.println(xaxis);
     *      ScaleData scaledata = new ScaleData();
     *      IncrementData incData=new IncrementData();
     *      incData.Distance=30;
     *
     *      scaledata.IncrementData=incData;
     *      xaxis.setScaleData(scaledata);
     *
     *     // System.out.println(scaledata);
     *     // System.out.println(xaxis);
     *      com.sun.star.chart.XChartDocument oChartDoc = (com.sun.star.chart.XChartDocument) UnoRuntime.queryInterface(
     *              com.sun.star.chart.XChartDocument.class, xChartDoc);
     *
     *      // x major grid
     *    /*  XPropertySet aGridProp = (XPropertySet) UnoRuntime.queryInterface(
     *      XPropertySet.class,
     *      ((XAxisXSupplier) UnoRuntime.queryInterface(
     *      XAxisXSupplier.class, oChartDoc.getDiagram())).getXAxis());
     */

    // check whether the current chart supports a y-axis

    /*
     *   XAxisXSupplier aYAxisSupplier = (XAxisXSupplier) UnoRuntime.queryInterface(
     *         XAxisXSupplier.class, oChartDoc.getDiagram());
     *
     * if (aYAxisSupplier != null) {
     *     XPropertySet aAxisProp = aYAxisSupplier.getXAxis();
     *     System.out.println(aAxisProp);
     *     if (aAxisProp != null) {
     *         aAxisProp.setPropertyValue("Origin", origin);
     *         aAxisProp.setPropertyValue("Max", max);
     *         aAxisProp.setPropertyValue("Min", min);
     *         aAxisProp.setPropertyValue("StepMain", step);
     *     }
     *
     *
     * }
     * }
     */
    public SpreadsheetInfo getSheetInfo() {
        return sheetInfo;
    }

    public void setSheetInfo(SpreadsheetInfo sheetInfo) {
        this.sheetInfo = sheetInfo;
    }

    public XChartDocument getXChartDoc() {
        return xChartDoc;
    }

    public void setXChartDoc(XChartDocument xChartDoc) {
        this.xChartDoc = xChartDoc;
    }

    public XChartTypeContainer getXChartTypeContainer() {
        return xChartTypeContainer;
    }

    public void setXChartTypeContainer(XChartTypeContainer xChartTypeContainer) {
        this.xChartTypeContainer = xChartTypeContainer;
    }

    public XComponentContext getXContext() {
        return xContext;
    }

    public void setXContext(XComponentContext xContext) {
        this.xContext = xContext;
    }

    public SpreadsheetDoc getXDoc() {
        return xDoc;
    }

    public void setXDoc(SpreadsheetDoc xDoc) {
        this.xDoc = xDoc;
    }

    public XMultiComponentFactory getXMCF() {
        return xMCF;
    }

    public void setXMCF(XMultiComponentFactory xMCF) {
        this.xMCF = xMCF;
    }

//  public XTableChart getXTableChart() {
//      return xTableChart;
//  }
//
//  public void setXTableChart(XTableChart xTableChart) {
//      this.xTableChart = xTableChart;
//  }
    public String getCurrent_series() {
        return current_series;
    }

    public void setCurrent_series(String current_series) {
        this.current_series = current_series;
    }
}
