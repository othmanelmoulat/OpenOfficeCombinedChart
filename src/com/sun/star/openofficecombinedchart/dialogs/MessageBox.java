/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */



package com.sun.star.openofficecombinedchart.dialogs;

//~--- JDK imports ------------------------------------------------------------

import com.sun.star.awt.Rectangle;
import com.sun.star.awt.XMessageBox;
import com.sun.star.awt.XMessageBoxFactory;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 *
 * @author othman
 */
public class MessageBox extends Thread {
    String                         _sMessage;
    String                         _sTitle;
    XComponentContext              m_xContext;
    private XMultiComponentFactory m_xMCF;

    public MessageBox(XComponentContext m_xContext, String _sTitle, String _sMessage) {
        this.m_xContext = m_xContext;
        this._sTitle    = _sTitle;
        this._sMessage  = _sMessage;
        m_xMCF          = m_xContext.getServiceManager();
    }

    public void run() {
        XComponent xComponent = null;

        try {
            m_xMCF = m_xContext.getServiceManager();

            Object             oToolkit           = m_xMCF.createInstanceWithContext("com.sun.star.awt.Toolkit",
                                                        m_xContext);
            XMessageBoxFactory xMessageBoxFactory =
                (XMessageBoxFactory) UnoRuntime.queryInterface(XMessageBoxFactory.class, oToolkit);

            // rectangle may be empty if position is in the center of the parent peer
            Rectangle                 aRectangle = new Rectangle();
            com.sun.star.awt.XToolkit xToolkit   =
                (com.sun.star.awt.XToolkit) UnoRuntime.queryInterface(com.sun.star.awt.XToolkit.class,
                    m_xMCF.createInstanceWithContext("com.sun.star.awt.Toolkit", m_xContext));

            // Describe the properties of the container window.
            com.sun.star.awt.WindowDescriptor aDescriptor = new com.sun.star.awt.WindowDescriptor();

            aDescriptor.Type              = com.sun.star.awt.WindowClass.TOP;
            aDescriptor.WindowServiceName = "window";
            aDescriptor.ParentIndex       = -1;
            aDescriptor.Parent            = null;
            aDescriptor.Bounds            = new com.sun.star.awt.Rectangle(0, 0, 0, 0);
            aDescriptor.WindowAttributes  = com.sun.star.awt.WindowAttribute.BORDER
                                            | com.sun.star.awt.WindowAttribute.MOVEABLE
                                            | com.sun.star.awt.WindowAttribute.SIZEABLE
                                            | com.sun.star.awt.WindowAttribute.CLOSEABLE;

            com.sun.star.awt.XWindowPeer xPeer       = xToolkit.createWindow(aDescriptor);
            XMessageBox                  xMessageBox = xMessageBoxFactory.createMessageBox(xPeer, aRectangle,
                                                           "errorbox", com.sun.star.awt.MessageBoxButtons.BUTTONS_OK,
                                                           _sTitle, _sMessage);

            xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, xMessageBox);

            if (xMessageBox != null) {
                short nResult = xMessageBox.execute();
            }
        } catch (com.sun.star.uno.Exception ex) {
            ex.printStackTrace(System.out);
        } finally {

            // make sure always to dispose the component and free the memory!
            if (xComponent != null) {
                xComponent.dispose();
            }
        }
    }
}
