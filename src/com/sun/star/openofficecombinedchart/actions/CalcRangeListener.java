/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */



package com.sun.star.openofficecombinedchart.actions;

/**
 *
 * @author othman
 */
public class CalcRangeListener implements com.sun.star.sheet.XRangeSelectionListener {
    public String aResult;

    public void done(com.sun.star.sheet.RangeSelectionEvent aEvent) {
        aResult = aEvent.RangeDescriptor;

        synchronized (this) {
            notify();
        }
    }

    public void aborted(com.sun.star.sheet.RangeSelectionEvent aEvent) {
        synchronized (this) {
            notify();
        }
    }

    public void disposing(com.sun.star.lang.EventObject aObj) {}
}
