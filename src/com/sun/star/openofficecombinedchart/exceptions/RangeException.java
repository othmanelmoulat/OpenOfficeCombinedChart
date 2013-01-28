/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */



package com.sun.star.openofficecombinedchart.exceptions;

/**
 *
 * @author othman
 */
public class RangeException extends Exception {

    /**
     * Creates a new instance of <code>RangeException</code> without detail message.
     */
    public RangeException() {}

    /**
     * Constructs an instance of <code>RangeException</code> with the specified detail message.
     * @param msg the detail message.
     */
    public RangeException(String msg) {
        super(msg);
    }
}
