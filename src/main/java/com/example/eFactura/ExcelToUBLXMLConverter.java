package com.example.eFactura;

import java.io.FileOutputStream;
import java.util.Iterator;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToUBLXMLConverter {

    public static void main(String[] args) {
        try {
            // Read from Excel file
            Workbook workbook = new XSSFWorkbook("C:/Users/Aurika/Desktop/eFactura/Book2ex.xlsx");
            Sheet sheet = workbook.getSheetAt(0);

            // Create UBL XML file and writer
            XMLOutputFactory outputFactory = XMLOutputFactory.newInstance();
            XMLStreamWriter writer = outputFactory.createXMLStreamWriter(new FileOutputStream("C:/Users/Aurika/Desktop/eFactura/Book2ex.xml"));

            // Start writing UBL XML document
            writer.writeStartDocument();
            writer.writeStartElement("Invoice");
            writer.setPrefix("II","urn:oasis:names:specification:ubl:schema:xsd:Invoice-2");
            writer.setDefaultNamespace("urn:oasis:names:specification:ubl:schema:xsd:Invoice-2");

            // Iterate through Excel rows and columns
            for (Row row : sheet) {
                for (Cell cell : row) {
                writer.writeStartElement("cbc:ID");

              //  Cell productCell = row.getCell(0);  // Assuming product information is in the first column
              //  Cell quantityCell = row.getCell(1); // Assuming quantity information is in the second column

                // Write UBL XML elements based on your specific UBL structure
              //  writer.writeStartElement("Item");
             //   writer.writeCharacters(productCell.toString());
                writer.writeEndElement(); // Item

                Iterator<Sheet> sheetIterator = workbook.sheetIterator();


                writer.writeStartElement("InvoicedQuantity");
                writer.writeCharacters(cell.toString());
                writer.writeEndElement(); // InvoicedQuantity



                }
            }

            // End writing UBL XML document
            writer.writeEndElement(); // Invoice
            writer.writeEndDocument();

            // Close resources
            writer.close();
            workbook.close();

            System.out.println("Conversion complete.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

