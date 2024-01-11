//package com.example.eFactura;
//
//    import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import javax.xml.bind.JAXBContext;
//import javax.xml.bind.JAXBException;
//import javax.xml.bind.Marshaller;
//import java.io.File;
//import java.util.ArrayList;
//import java.util.Iterator;
//import java.util.List;
//
//    public class ExcelToUBLConverter {
//
//        public static void main(String[] args) {
//            try {
//                Workbook workbook = new XSSFWorkbook(new File("C:/Users/Aurika/Desktop/eFactura/Book1Ex.xlsx"));
//                Sheet sheet = workbook.getSheetAt(0);
//
//                // Create UBL XML
//                Invoice invoice = createUBLObjectFromExcel(sheet);
//
//                // Write XML to file
//                marshalToXML(invoice, "C:/Users/Aurika/Desktop/eFactura/Book1Ex.xml");
//
//            } catch (Exception e) {
//                e.printStackTrace();
//            }
//        }
//
//        private static Invoice createUBLObjectFromExcel(Sheet sheet) {
//            Invoice invoice = new Invoice();
//            List<InvoiceLine> invoiceLines = new ArrayList<>();
//
//            Iterator<Row> rowIterator = sheet.iterator();
//            // Skip header row
//            if (rowIterator.hasNext()) {
//                rowIterator.next();
//            }
//
//            while (rowIterator.hasNext()) {
//                Row row = rowIterator.next();
//                InvoiceLine invoiceLine = new InvoiceLine();
//
//                invoiceLine.setId(getStringValue(row.getCell(0)));
//                invoiceLine.setItemCode(getStringValue(row.getCell(1)));
//                invoiceLine.setDescription(getStringValue(row.getCell(2)));
//                invoiceLine.setQuantity(getNumericValue(row.getCell(3)));
//                invoiceLine.setPrice(getNumericValue(row.getCell(4)));
//
//                invoiceLines.add(invoiceLine);
//            }
//
//            invoice.setInvoiceLines(invoiceLines);
//            return invoice;
//        }
//
//        private static String getStringValue(Cell cell) {
//            return cell == null ? null : cell.getStringCellValue();
//        }
//
//        private static Double getNumericValue(Cell cell) {
//            return cell == null ? null : cell.getNumericCellValue();
//        }
//
//        private static void marshalToXML(Object object, String filePath) throws JAXBException {
//            JAXBContext jaxbContext = JAXBContext.newInstance(object.getClass());
//            Marshaller marshaller = jaxbContext.createMarshaller();
//            marshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, true);
//            marshaller.marshal(object, new File(filePath));
//        }
//    }
//
//
