/*
 * SpreadsheetDoc.java
 *
 * Created on 23 janvier 2008, 10:47
 * */



package com.sun.star.openofficecombinedchart.spreadsheet;

//~--- JDK imports ------------------------------------------------------------

import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.openofficecombinedchart.OOoCombinedChartController;
import com.sun.star.openofficecombinedchart.actions.CalcRangeListener;
import com.sun.star.openofficecombinedchart.exceptions.RangeException;
import com.sun.star.openofficecombinedchart.util.Utils;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.uno.*;

import java.io.File;

import java.lang.Exception;

import java.net.MalformedURLException;
import java.net.URISyntaxException;
import java.net.URL;

import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * <PRE>This class stores Info of openoffice Document</PRE>
 * <PRE>UNO API is used to retrive info on Document </PRE>
 * <PRE>implements  methods for processing & manipulating openoffice document</PRE>
 * @author othman EL Moulat
 */
public class SpreadsheetDoc {

    // private static SugarLogger     _logger          = new SugarLogger(SpreadsheetDoc.class);
    private String                                  application;
    private XComponent                              document;
    private XComponentContext                       m_xContext;
    private XFrame                                  m_xFrame;
    private XMultiComponentFactory                  m_xMCF;
    private String                                  sheetName;
    private XSpreadsheet                            xActiveSheet;
    private XController                             xController;
    private XDesktop                                xDesktop;
    private com.sun.star.sheet.XSpreadsheetDocument xSpreadsheetDocument;

    /**
     * Creates a new instance of SpreadsheetDoc
     */
    public SpreadsheetDoc() {}

    /**
     * Creates a new instance of SpreadsheetDoc with 2 arguments
     * @param ctx reference to XComponentContext
     * @param xFrame reference to XFrame
     */
    public SpreadsheetDoc(XComponentContext ctx, XFrame xFrame) {
        this.m_xContext  = ctx;
        this.m_xFrame    = xFrame;
        this.xController = this.m_xFrame.getController();
        initialize();
    }

    /**
     * <PRE>This method initialize a OpenOffice </PRE>
     * <PRE>Gets references to it's main components : Desktop, Application name & it's current Document</PRE>
     */
    private void initialize() {
        try {
            m_xMCF = m_xContext.getServiceManager();

            Object desktop = m_xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", m_xContext);

            xDesktop    = (XDesktop) UnoRuntime.queryInterface(com.sun.star.frame.XDesktop.class, desktop);
            document    = xDesktop.getCurrentComponent();
            application = getApplicationName(document);
            initSpreadSheetDocument();
        } catch (com.sun.star.uno.Exception ex) {
            ex.printStackTrace();
        }
    }

    /**
     * Get the application name of a document. E.g. for a writer document
     * "com.sun.star.text.TextDocument"
     *
     * @param myXComponent UNO Representativ of the opened document.
     * @return The OpenOffice application name. It looks like
     * "com.sun.star.text.TextDocument".
     */
    public String getApplicationName(XComponent myXComponent) {
        XModuleManager xMM = null;

        try {
            xMM = (XModuleManager) UnoRuntime.queryInterface(XModuleManager.class,
                    m_xMCF.createInstanceWithContext("com.sun.star.frame.ModuleManager", m_xContext));
        } catch (com.sun.star.uno.Exception e) {
            return null;
        }

        String sOOoApp = null;

        try {

            // Getting the application name of the document,
            // e.g. "com.sun.star.text.TextDocument" for writer
            sOOoApp = xMM.identify(myXComponent);
        } catch (com.sun.star.uno.Exception e) {
            return null;
        }

        return sOOoApp;
    }

    /**
     * creates a new service with the default remote context
     * @param sService service name
     * @return  reference to openoffice XService Object
     * @throws java.lang.Exception thrown if XService creation fails
     */
    public Object getService(String sService) throws java.lang.Exception {
        Object xService = null;

        xService = m_xMCF.createInstanceWithContext(sService, m_xContext);

        if (xService == null) {
            throw new java.lang.Exception(sService + " could not be established");
        }

        return xService;
    }

    public void initSpreadSheetDocument() {
        xSpreadsheetDocument = (com.sun.star.sheet.XSpreadsheetDocument) UnoRuntime.queryInterface(
            com.sun.star.sheet.XSpreadsheetDocument.class, document);
    }

    public XSpreadsheetDocument getSpreadSheetDocument() {
        return xSpreadsheetDocument;
    }

    public String[] getSheetsNames() {
        com.sun.star.sheet.XSpreadsheets xSheets = xSpreadsheetDocument.getSheets();

        return xSheets.getElementNames();
    }

    public String getSheetsType() {
        com.sun.star.sheet.XSpreadsheets xSheets = xSpreadsheetDocument.getSheets();
        Type                             xType   = xSheets.getElementType();

        return xType.getTypeName();
    }

    /**
     * insert new unique name sheet into calc document
     */
    public XSpreadsheet insertSheet() throws NoSuchElementException, WrappedTargetException {
        XSpreadsheet aSheet = null;
        String[]     sheets = getSheetsNames();

        if ((sheets == null) || (sheets.length <= 0)) {
            return null;
        }

        System.out.println("sheet list before insert :");
        Utils.print(sheets);

        int index = sheets.length;

        sheetName = Utils.createUniqueName(this);

        XSpreadsheets aSheets = getSpreadSheetDocument().getSheets();

        aSheets.insertNewByName(sheetName, (short) (index + 1));
        aSheet = getSpreadsheetByName(sheetName);
        System.out.println(sheetName + " sheet inserted successfully ");
        System.out.println("sheet list after insert :");
        Utils.print(getSheetsNames());

        return aSheet;
    }

    /**
     * Returns the spreadsheet with the specified index (0-based).
     * @param nIndex  The index of the sheet.
     * @return  XSpreadsheet interface of the sheet.
     *
     */
    public com.sun.star.sheet.XSpreadsheet getSpreadsheet(int nIndex) {

        // Collection of sheets
        com.sun.star.sheet.XSpreadsheets xSheets = xSpreadsheetDocument.getSheets();
        com.sun.star.sheet.XSpreadsheet  xSheet  = null;

        try {
            com.sun.star.container.XIndexAccess xSheetsIA =
                (com.sun.star.container.XIndexAccess) UnoRuntime.queryInterface(
                    com.sun.star.container.XIndexAccess.class, xSheets);

            xSheet = (com.sun.star.sheet.XSpreadsheet) UnoRuntime.queryInterface(com.sun.star.sheet.XSpreadsheet.class,
                    xSheetsIA.getByIndex(nIndex));
        } catch (java.lang.Exception ex) {
            System.err.println("Error: caught exception in getSpreadsheet()!\nException Message = " + ex.getMessage());
            ex.printStackTrace();
        }

        return xSheet;
    }

    public XSpreadsheet getSpreadsheetByName(String msChartSheetName) {
        XSpreadsheet   aChartSheet = null;
        XSpreadsheets  aSheets     = xSpreadsheetDocument.getSheets();
        XNameContainer aSheetsNC   = (XNameContainer) UnoRuntime.queryInterface(XNameContainer.class, aSheets);
        XIndexAccess   aSheetsIA   = (XIndexAccess) UnoRuntime.queryInterface(XIndexAccess.class, aSheets);

        if ((aSheets != null) && (aSheetsNC != null) && (aSheetsIA != null)) {
            try {
                aChartSheet = (XSpreadsheet) UnoRuntime.queryInterface(XSpreadsheet.class,
                        aSheetsNC.getByName(msChartSheetName));
            } catch (NoSuchElementException noSuchElementException) {
                return null;
            } catch (WrappedTargetException wrappedTargetException) {
                return null;
            }
        }

        return aChartSheet;
    }

    public boolean hasSheetByname(String name) {
        XSpreadsheets  aSheets   = xSpreadsheetDocument.getSheets();
        XNameContainer aSheetsNC = (XNameContainer) UnoRuntime.queryInterface(XNameContainer.class, aSheets);

        return aSheetsNC.hasByName(name);
    }

    public XController getXController() {
        return xController;
    }

    public SpreadsheetInfo getRangeSelection() throws com.sun.star.uno.Exception, RangeException {

        // let the user select a range and use it as the view's selection
        com.sun.star.sheet.XRangeSelection xRngSel =
            (com.sun.star.sheet.XRangeSelection) UnoRuntime.queryInterface(com.sun.star.sheet.XRangeSelection.class,
                xController);
        CalcRangeListener aListener = new CalcRangeListener();
        SpreadsheetInfo   sheetInfo = null;

        xRngSel.addRangeSelectionListener(aListener);

        com.sun.star.beans.PropertyValue[] aArguments = new com.sun.star.beans.PropertyValue[2];

        aArguments[0]       = new com.sun.star.beans.PropertyValue();
        aArguments[0].Name  = "Title";
        aArguments[0].Value = "Please select a range";
        aArguments[1]       = new com.sun.star.beans.PropertyValue();
        aArguments[1].Name  = "CloseOnMouseRelease";
        aArguments[1].Value = new Boolean(false);
        xRngSel.startRangeSelection(aArguments);

        synchronized (aListener) {
            try {
                aListener.wait();    // wait until the selection is done
            } catch (InterruptedException ex) {
                Logger.getLogger(OOoCombinedChartController.class.getName()).log(Level.SEVERE, null, ex);
            }
        }

        xRngSel.removeRangeSelectionListener(aListener);

        if ((aListener.aResult != null) && (aListener.aResult.length() != 0)) {
            com.sun.star.view.XSelectionSupplier xSel =
                (com.sun.star.view.XSelectionSupplier) UnoRuntime.queryInterface(
                    com.sun.star.view.XSelectionSupplier.class, xController);
            com.sun.star.table.XCellRange xResultRange = this.getActiveSheet().getCellRangeByName(aListener.aResult);

            xSel.select(xResultRange);

            com.sun.star.table.XColumnRowRange xColRowRange =
                (com.sun.star.table.XColumnRowRange) UnoRuntime.queryInterface(
                    com.sun.star.table.XColumnRowRange.class, xResultRange);
            com.sun.star.table.XTableColumns xColumns     = xColRowRange.getColumns();
            String[]                         columnHeader = new String[xColumns.getCount()];

            for (int i = 0; i < xColumns.getCount(); i++) {
                Object                        aColumnObj = xColumns.getByIndex(i);
                com.sun.star.container.XNamed xNamed     =
                    (com.sun.star.container.XNamed) UnoRuntime.queryInterface(com.sun.star.container.XNamed.class,
                        aColumnObj);

                columnHeader[i] = xNamed.getName();
            }

            sheetInfo = new SpreadsheetInfo(columnHeader, aListener.aResult);
        } else {
            throw new RangeException(Utils.ERROR_MSG);
        }

        return sheetInfo;
    }

    public com.sun.star.sheet.XSpreadsheet getActiveSheet() {
        if (this.xActiveSheet == null) {
            com.sun.star.sheet.XSpreadsheetView xView = (com.sun.star.sheet.XSpreadsheetView) UnoRuntime.queryInterface(
                                                            com.sun.star.sheet.XSpreadsheetView.class, xController);

            this.xActiveSheet = xView.getActiveSheet();
        }

        return this.xActiveSheet;
    }

    /**
     *
     * getters and setters methods
     *
     * @return
     */
    public XComponentContext getM_xContext() {
        return m_xContext;
    }

    /**
     * Method description
     *
     *
     * @param m_xContext
     */
    public void setM_xContext(XComponentContext m_xContext) {
        this.m_xContext = m_xContext;
    }

    /**
     * Method description
     *
     *
     * @return
     */
    public XMultiComponentFactory getM_xMCF() {
        return m_xMCF;
    }

    /**
     * Method description
     *
     *
     * @param m_xMCF
     */
    public void setM_xMCF(XMultiComponentFactory m_xMCF) {
        this.m_xMCF = m_xMCF;
    }

    /**
     * Method description
     *
     *
     * @return
     */
    public XFrame getM_xFrame() {
        return m_xFrame;
    }

    /**
     * Method description
     *
     *
     * @param m_xFrame
     */
    public void setM_xFrame(XFrame m_xFrame) {
        this.m_xFrame = m_xFrame;
    }

    /**
     * Method description
     *
     *
     * @return
     */
    public XDesktop getXDesktop() {
        return xDesktop;
    }

    /**
     * Method description
     *
     *
     * @param xDesktop
     */
    public void setXDesktop(XDesktop xDesktop) {
        this.xDesktop = xDesktop;
    }

    /**
     * Method description
     *
     *
     * @return
     */
    public XComponent getDocument() {
        return document;
    }

    public String getSheetName() {
        return sheetName;
    }

    /**
     * Method description
     *
     *
     * @param document
     */
    public void setDocument(XComponent document) {
        this.document = document;
    }
}
