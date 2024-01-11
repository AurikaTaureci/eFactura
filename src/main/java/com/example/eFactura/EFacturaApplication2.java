package com.example.eFactura;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;
import java.io.*;
import java.util.Iterator;

@SpringBootApplication
    public class EFacturaApplication2 {

    public static void main(String[] args) throws XMLStreamException {
        SpringApplication.run(com.example.eFactura.EFacturaApplication.class, args);

        try {
            // Load Excel file
            FileInputStream excelFile = new FileInputStream(new File("C:/Users/Aurika/Desktop/eFactura/Book2ex.xlsx"));
            Workbook workbook = new XSSFWorkbook(excelFile);

            Sheet sheet = workbook.getSheetAt(0);

            //Invoice invoice = createUBLObjectFromExcel(sheet);
//            Invoice invoice = new Invoice();
//            invoice.setUblVersionID("2.1");
            // Create XML file
            FileOutputStream xmlFile = new FileOutputStream(new File("C:/Users/Aurika/Desktop/eFactura/Book2ex.xml"));
            XMLStreamWriter xmlStreamWriter = XMLOutputFactory.newFactory().createXMLStreamWriter(xmlFile);
            xmlStreamWriter.writeStartDocument();

            StringWriter outputXmlStringWriter = new StringWriter();
            String outputXmlString = outputXmlStringWriter.toString();
            xmlFile.write(outputXmlString.getBytes("UTF-8"));


            xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));
            xmlStreamWriter.writeStartElement("Invoice");
            xmlStreamWriter.writeNamespace("", "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2");

            xmlStreamWriter.writeNamespace("cac", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
            xmlStreamWriter.writeNamespace("cbc", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
            // xmlStreamWriter.setDefaultNamespace("urn:oasis:names:specification:ubl:schema:xsd:Invoice-2");

            //
            xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

            xmlStreamWriter.writeStartElement("cbc:UBLVersionID");
            xmlStreamWriter.writeCharacters("2.1");
            xmlStreamWriter.writeEndElement();
            xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

            //
            xmlStreamWriter.writeStartElement("cbc:CustomizationID");
            xmlStreamWriter.writeCharacters("urn:cen.eu:en16931:2017#compliant#urn:efactura.mfinante.ro:CIUS-RO:1.0.0");
            xmlStreamWriter.writeEndElement();
            xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

            //


            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Iterate over all cells in the row
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();


                    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                        // Iterate through sheets and rows
                          //for (int j = 1; j < workbook.getNumberOfSheets(); j++) {
                        Sheet sheet1 = workbook.getSheetAt(i);
                       Iterator<Row> rowIterator1 = sheet1.iterator();

                        //while (rowIterator.hasNext()) {
                            Row row1 = rowIterator.next();
                            Cell cell1 = row1.getCell(0);
                            xmlStreamWriter.writeStartElement("cbc:ID");
                            xmlStreamWriter.writeCharacters(cell1.toString());
                            xmlStreamWriter.writeEndElement();
                            xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


                            //// ultimul

                            // Sheet sheet2 = workbook.getSheetAt(i);
                             Iterator<Row> rowIterator2 = sheet1.iterator();
                            //rowIterator.hasNext();
                            //Row row2 = rowIterator.next();
                             Row row2 = rowIterator2.next();
                            Cell cell2 = row2.getCell(1);
                            // Cell cell2 = row.getCell(1);
                            xmlStreamWriter.writeStartElement("cbc:IssueDate"); // Element for each cell

                            //xmlStreamWriter.writeCharacters(cell2.toString()); //Date din excel
                            xmlStreamWriter.writeCharacters(cell.toString()); //Date din excel
                            xmlStreamWriter.writeEndElement();

                            xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


                            xmlStreamWriter.writeEndElement(); // Close root element
                            xmlStreamWriter.writeEndDocument();
                            xmlStreamWriter.writeEndElement();
                            xmlStreamWriter.close();
                            xmlFile.close();
                            workbook.close();
                              // }
                        }
                    }
                }
            //}


        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (UnsupportedEncodingException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}