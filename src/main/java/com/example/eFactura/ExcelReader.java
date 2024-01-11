package com.example.eFactura;

import org.apache.poi.ss.usermodel.*;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;
import java.io.*;

public class ExcelReader {
    public static void main(String[] args) throws FileNotFoundException {
        // Provide the path to your Excel file
        String filePath = "C:/Users/Aurika/Desktop/eFactura/Book2ex.xlsx";
        String xmlFile = "C:/Users/Aurika/Desktop/eFactura/Book2ex.xml";

        //FileOutputStream xmlFile = new FileOutputStream(new File("C:/Users/Aurika/Desktop/eFactura/Book2ex.xml"));
        //XMLStreamWriter xmlStreamWriter = XMLOutputFactory.newFactory().createXMLStreamWriter(xmlFile);
        //  xmlStreamWriter.writeStartDocument();

        try {
            FileInputStream fis = new FileInputStream(new File(filePath));
            FileOutputStream xFis = new FileOutputStream(new File(xmlFile));

            Workbook workbook = WorkbookFactory.create(fis);
            XMLStreamWriter xmlStreamWriter = XMLOutputFactory.newFactory().createXMLStreamWriter(xFis);

            xmlStreamWriter.writeStartDocument();
            // Assuming you are working with the first sheet, you can change the index if needed
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate through rows
            for (Row row : sheet) {
                // Iterate through cells in the row
                for (Cell cell : row) {
                    // Print the cell value
                    System.out.print(cell.toString() + "\t");

                    xmlStreamWriter.writeStartElement("cbc:ID"); // Element for each cell

                    xmlStreamWriter.writeCharacters(cell.toString());
                    xmlStreamWriter.writeEndElement();

                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));
                }
                System.out.println(); // Move to the next row
            }

            for (Row row : sheet) {
                // Iterate through cells in the row
                for (Cell cell : row) {
                    // Print the cell value
                    System.out.print(cell.toString() + "\t");

                    xmlStreamWriter.writeStartElement("cbc:date"); // Element for each cell

                    xmlStreamWriter.writeCharacters(cell.toString());
                    xmlStreamWriter.writeEndElement();

                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));
                }

            }
        } catch (XMLStreamException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}


