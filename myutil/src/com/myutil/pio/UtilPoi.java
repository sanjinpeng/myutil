package com.myutil.pio;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;


import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

public class UtilPoi {

    public static SheetDatasHandler read(Path file) throws Exception {
        SheetDatasHandler handler = new SheetDatasHandler();
        read(file, handler);
        return handler;
    }

    public static void read(Path file, RowMapper mapper) throws Exception {
        final long size = Files.size(file);
        try (InputStream in = new BufferedInputStream(new FileInputStream(file.toFile()), size > Integer.MAX_VALUE ? 1024 * 1024 * 10 : (int) size)) {
            read(in, mapper);
        }
    }

    public static void read(InputStream in, RowMapper mapper) throws Exception {
        XSSFReader reader = new XSSFReader(OPCPackage.open(in));
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");

        mapper.setSharedStringsTable(reader.getSharedStringsTable());
        parser.setContentHandler(mapper);

        for (Iterator<InputStream> iter = reader.getSheetsData(); iter.hasNext(); ) {
            try (InputStream sheetIn = iter.next()) {
                parser.parse(new InputSource(sheetIn));
            }
        }
    }
}