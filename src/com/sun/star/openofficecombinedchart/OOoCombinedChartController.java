package com.sun.star.openofficecombinedchart;

//~--- JDK imports ------------------------------------------------------------

import com.sun.star.frame.*;
import com.sun.star.openofficecombinedchart.chart.CombinedChart;
import com.sun.star.openofficecombinedchart.dialogs.MessageBox;
import com.sun.star.openofficecombinedchart.exceptions.RangeException;
import com.sun.star.openofficecombinedchart.spreadsheet.SpreadsheetDoc;
import com.sun.star.openofficecombinedchart.spreadsheet.SpreadsheetInfo;
import com.sun.star.uno.*;

import java.util.logging.Level;
import java.util.logging.Logger;

import javax.swing.SwingUtilities;

/**
 * <PRE> Controller class </PRE>
 * <PRE> delegates execution paths to specific methods</PRE>
 * @author othman EL Moulat
 */
public class OOoCombinedChartController {
    private static OOoCombinedChartController        instance = new OOoCombinedChartController();
    private SpreadsheetDoc                           document;
    private XComponentContext                        m_xContext;
    private XFrame                                   m_xFrame;
    private com.sun.star.lang.XMultiComponentFactory m_xMCF;
    private XController                              xController;
    private XDesktop                                 xDesktop;

    /**
     * Constructs ...
     *
     */
    private OOoCombinedChartController() {}

    /**
     * this method is a singleton that creates a unique (single) instance of this class
     * @return unique instance of this (class)
     */
    public static OOoCombinedChartController getInstance() {
        return instance;
    }

    /**
     * Method description
     *
     *
     * @param m_xContext
     */
    void setXComponentContext(XComponentContext m_xContext) {
        this.m_xContext = m_xContext;
    }

    /**
     * Method description
     *
     *
     * @param m_xFrame
     */
    void setXFrame(XFrame m_xFrame) {
        this.m_xFrame = m_xFrame;
    }

    public XController getXController() {
        return xController;
    }

    public void setXController(XController xController) {
        this.xController = xController;
    }

    /**
     * <PRE>controller method</PRE>
     * <PRE>dispatchs path execution to specific methods depending on openoffice path</PRE>
     * @param path openoffice UNO path
     */
    public void execute(String path) {
        if (path.equals("OpenOfficeCombinedChart")) {
            SwingUtilities.invokeLater(new Runnable() {
                public void run() {
                    try {
                        SpreadsheetInfo sheetInfo = document.getRangeSelection();

                        // UnoDialog dlg=new UnoDialog(m_xContext);
                        // dlg.showDialog();
                        CombinedChart   chart     = new CombinedChart(m_xContext, document, sheetInfo);

                        chart.createChart();
                    } catch (RangeException ex) {
                        Logger.getLogger(OOoCombinedChartController.class.getName()).log(Level.SEVERE, null, ex);
                        new MessageBox(m_xContext, "Error", ex.getMessage()).start();

                        return;
                    } catch (com.sun.star.uno.Exception ex) {
                        Logger.getLogger(OOoCombinedChartController.class.getName()).log(Level.SEVERE, null, ex);
                        new MessageBox(m_xContext, "Error", ex.getMessage()).start();

                        return;
                    }
                }
            });
        }
    }

    /**
     * returns SpreadsheetDoc
     * @return SpreadsheetDoc instance
     */
    public SpreadsheetDoc getDocument() {
        return document;
    }

    /**
     * sets SpreadsheetDoc
     * @param document SpreadsheetDoc instance
     */
    public void setDocument(SpreadsheetDoc document) {
        this.document = document;
    }
}
