/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */



package com.sun.star.openofficecombinedchart.util;

//~--- JDK imports ------------------------------------------------------------

import com.sun.star.openofficecombinedchart.spreadsheet.SpreadsheetDoc;

import java.net.InetAddress;
import java.net.UnknownHostException;

import java.util.Date;
import java.util.Random;

/**
 *
 * @author othman
 */
public class Utils {
    public static final String ERROR_MSG =
        "Range Selection Error!\n\nPlease select a range that contains following series in order:\n\nDATE-OPEN-HIGH-LOW-CLOSE-(Zero or more Line data series)\n\n";
    public static final String DEFAULT_CHART_SHEET_NAME = "CombinedChart" + Integer.toString(1);
    private static String[]    byteToStr                = {
        "00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "0a", "0b", "0c", "0d", "0e", "0f", "10", "11",
        "12", "13", "14", "15", "16", "17", "18", "19", "1a", "1b", "1c", "1d", "1e", "1f", "20", "21", "22", "23",
        "24", "25", "26", "27", "28", "29", "2a", "2b", "2c", "2d", "2e", "2f", "30", "31", "32", "33", "34", "35",
        "36", "37", "38", "39", "3a", "3b", "3c", "3d", "3e", "3f", "40", "41", "42", "43", "44", "45", "46", "47",
        "48", "49", "4a", "4b", "4c", "4d", "4e", "4f", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59",
        "5a", "5b", "5c", "5d", "5e", "5f", "60", "61", "62", "63", "64", "65", "66", "67", "68", "69", "6a", "6b",
        "6c", "6d", "6e", "6f", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", "7a", "7b", "7c", "7d",
        "7e", "7f", "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "8a", "8b", "8c", "8d", "8e", "8f",
        "90", "91", "92", "93", "94", "95", "96", "97", "98", "99", "9a", "9b", "9c", "9d", "9e", "9f", "a0", "a1",
        "a2", "a3", "a4", "a5", "a6", "a7", "a8", "a9", "aa", "ab", "ac", "ad", "ae", "af", "b0", "b1", "b2", "b3",
        "b4", "b5", "b6", "b7", "b8", "b9", "ba", "bb", "bc", "bd", "be", "bf", "c0", "c1", "c2", "c3", "c4", "c5",
        "c6", "c7", "c8", "c9", "ca", "cb", "cc", "cd", "ce", "cf", "d0", "d1", "d2", "d3", "d4", "d5", "d6", "d7",
        "d8", "d9", "da", "db", "dc", "dd", "de", "df", "e0", "e1", "e2", "e3", "e4", "e5", "e6", "e7", "e8", "e9",
        "ea", "eb", "ec", "ed", "ee", "ef", "f0", "f1", "f2", "f3", "f4", "f5", "f6", "f7", "f8", "f9", "fa", "fb",
        "fc", "fd", "fe", "ff"
    };
    private static int iSequence = 1048576;

    /**
     * Generate an universal unique identifier
     *
     * @return An hexadecimal string of 32 characters, created using the machine
     *         IP address, current system date, a randon number and a sequence.
     */
    public static String generateUUID() {
        int          iRnd;
        long         lSeed = new Date().getTime();
        Random       oRnd  = new Random(lSeed);
        String       sHex;
        StringBuffer sUUID       = new StringBuffer(32);
        byte[]       localIPAddr = new byte[4];

        try {

            // 8 characters Code IP address of this machine
            localIPAddr = InetAddress.getLocalHost().getAddress();
            sUUID.append(byteToStr[((int) localIPAddr[0]) & 255]);
            sUUID.append(byteToStr[((int) localIPAddr[1]) & 255]);
            sUUID.append(byteToStr[((int) localIPAddr[2]) & 255]);
            sUUID.append(byteToStr[((int) localIPAddr[3]) & 255]);
        } catch (UnknownHostException e) {

            // Use localhost by default
            sUUID.append("7F000000");
        }

        // Append a seed value based on current system date
        sUUID.append(Long.toHexString(lSeed));

        // 6 characters - an incremental sequence
        sUUID.append(Integer.toHexString(iSequence++));

        if (iSequence > 16777000) {
            iSequence = 1048576;
        }

        do {
            iRnd = oRnd.nextInt();

            if (iRnd > 0) {
                iRnd = -iRnd;
            }

            sHex = Integer.toHexString(iRnd);
        } while (0 == iRnd);

        // Finally append a random number
        sUUID.append(sHex);

        return sUUID.substring(0, 32);
    }    // generateUUID()

    /**
     * makes a String unique by appending a numerical suffix
     * @param _xElementContainer the com.sun.star.container.XNameAccess container
     * that the new Element is going to be inserted to
     * @param _sElementName the StemName of the Element
     */
    public static String createUniqueName(SpreadsheetDoc xDoc) {
        boolean bElementexists = true;
        int     i              = 0;
        String  _sElementName  = generateUUID();

        while (bElementexists) {
            bElementexists = xDoc.hasSheetByname(_sElementName);

            if (bElementexists) {
                _sElementName = generateUUID();
            }
        }

        return _sElementName;
    }

    public static void print(String[] a) {
        if ((a == null) || (a.length <= 0)) {
            return;
        }

        StringBuffer buff = new StringBuffer();

        buff.append("[");

        for (int i = 0; i < a.length; i++) {
            buff.append(a[i] + ",");
        }

        buff.deleteCharAt(buff.length() - 1);
        buff.append("]");
        System.out.println(buff.toString());
    }
}
