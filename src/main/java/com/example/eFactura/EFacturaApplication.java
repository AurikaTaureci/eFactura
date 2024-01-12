package com.example.eFactura;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;
import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

@SpringBootApplication
public class EFacturaApplication {

    public static void main(String[] args) {
        SpringApplication.run(EFacturaApplication.class, args);

        try {
            // Load Excel file
            FileInputStream excelFile = new FileInputStream(new File("C:/Users/Aurika/Desktop/eFactura/F1.xlsx"));
            Workbook workbook = new XSSFWorkbook(excelFile);

            //Sheet sheet = workbook.getSheetAt(0);

            //Invoice invoice = createUBLObjectFromExcel(sheet);
//            Invoice invoice = new Invoice();
//            invoice.setUblVersionID("2.1");
            // Create XML file
            FileOutputStream xmlFile = new FileOutputStream(new File("C:/Users/Aurika/Desktop/eFactura/F1.xml"));
            XMLStreamWriter xmlStreamWriter = XMLOutputFactory.newFactory().createXMLStreamWriter(xmlFile);
            //xmlStreamWriter.writeStartDocument();

            xmlFile.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>".getBytes());

           // StringWriter outputXmlStringWriter = new StringWriter();
            //String outputXmlString = outputXmlStringWriter.toString();
           // xmlFile.write(outputXmlString.getBytes("UTF-8"));
           // xmlFile.write("encoding=\"UTF-8\">".getBytes());

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


            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

                Sheet sheet = workbook.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();

                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();

                    Cell cell = row.getCell(1);
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters(cell.toString());
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    Row row2 = rowIterator.next();
                    Cell cell1 = row2.getCell(1);
                    xmlStreamWriter.writeStartElement("cbc:IssueDate"); // Element for each cell

                    DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
                    String requiredDate = df.format(cell1.getDateCellValue());
                    xmlStreamWriter.writeCharacters(requiredDate);

                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


                    // cbc:DueDate
                    Row row3 = rowIterator.next();
                    Cell cell3 = row3.getCell(1);
                    xmlStreamWriter.writeStartElement("cbc:DueDate");

                    String requiredDate1 = df.format(cell3.getDateCellValue());
                   // xmlStreamWriter.writeCharacters(cell3.toString()); //Date din excel
                    xmlStreamWriter.writeCharacters(requiredDate1);

                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


                    // <cac:InvoicePeriod>
                    //    <cbc:EndDate>2022-05-31</cbc:EndDate>
                    //  </cac:InvoicePeriod>

                    Row row4 = rowIterator.next();
                    Cell cell4 = row4.getCell(1);
                    xmlStreamWriter.writeStartElement("cac:InvoicePeriod"); // <cac:InvoicePeriod>
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));
                    xmlStreamWriter.writeStartElement("cbc:EndDate"); // <cbc:EndDate>2022-05-31</cbc:EndDate>
                   // xmlStreamWriter.writeCharacters(cell4.toString()); //Date din excel

                    String requiredDate2 = df.format(cell4.getDateCellValue());
                    xmlStreamWriter.writeCharacters(requiredDate2);

                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


                    //<cbc:InvoiceTypeCode>380</cbc:InvoiceTypeCode><!--BT-3-->
                    xmlStreamWriter.writeStartElement("cbc:InvoiceTypeCode");
                    xmlStreamWriter.writeCharacters("380");
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // <cbc:Note></cbc:Note>
                    int rowIndex43 = 42; // (indicele începe de la 0)
                    int columnIndex0 = 0; // (indicele începe de la 0)

                    Row row43= rowIterator.next().getSheet().getRow(rowIndex43);
                    Cell cell43 = row43.getCell(columnIndex0);

                    xmlStreamWriter.writeStartElement("cbc:Note");
                    xmlStreamWriter.writeCharacters(String.valueOf(cell43.toString()));
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


                    // <cbc:DocumentCurrencyCode>RON</cbc:DocumentCurrencyCode>
                    xmlStreamWriter.writeStartElement("cbc:DocumentCurrencyCode");
                    xmlStreamWriter.writeCharacters("RON");
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));



//                      <cac:AccountingSupplierParty> <!-- BG-4 VÂNZĂTOR -->
//                      <cac:Party>
//                      <cac:PartyName>
//                           <cbc:Name>Seller SRL</cbc:Name> --> Technology Reply S.R.L.
//                       </cac:PartyName>


                    //<cac:AccountingSupplierParty>
                    xmlStreamWriter.writeStartElement("cac:AccountingSupplierParty");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // <cac:Party>
                    xmlStreamWriter.writeStartElement("cac:Party");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                     //<cac:PartyName>
                    xmlStreamWriter.writeStartElement("cac:PartyName");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:Name>Seller SRL</cbc:Name> --> Technology Reply S.R.L.
                    int rowIndex8 = 8; // (indicele începe de la 0)
                    int columnIndex5 = 5; // (indicele începe de la 0)

                    Row row8= rowIterator.next().getSheet().getRow(rowIndex8);
                    Cell cell5 = row8.getCell(columnIndex5);
                    xmlStreamWriter.writeStartElement("cbc:Name");
                    xmlStreamWriter.writeCharacters(cell5.toString());

                    // end  <cbc:Name>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //  end <cac:PartyName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


 //<cac:PostalAddress>
//        <cbc:StreetName>line1</cbc:StreetName> --> Strada Ceasornicului, Nr.17, Corp A, Etaj 5, District 1, Bucharest
//         cbc:CityName>SECTOR1</cbc:CityName> --> Bucuresti
//        <cbc:PostalZone>013329</cbc:PostalZone>
//        <cbc:CountrySubentity>RO-B</cbc:CountrySubentity>
// <cac:Country>
//          <cbc:IdentificationCode>RO</cbc:IdentificationCode> --> RO
//  </cac:Country>
//      </cac:PostalAddress>

                    /* <cac:PostalAddress>*/
//                    int rowIndex = 13; // rândul 13 (indicele începe de la 0)
//                    int columnIndex = 1; // coloana 1 (indicele începe de la 0)
//
//                    Row row13 = rowIterator.next().getSheet().getRow(rowIndex);
//                    Cell cell14 = row13.getCell(columnIndex);

                    xmlStreamWriter.writeStartElement("cac:PostalAddress");
                   // xmlStreamWriter.writeCharacters(String.valueOf(cell14.toString()));
                   // xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:StreetName>line1</cbc:StreetName> --> Strada Ceasornicului, Nr.17, Corp A, Etaj 5, District 1, Bucharest
                    int rowStreetNameSupplier9 =9;
                    int cellNameSupplier5=5;

                    Row rowStreetNameSupplier= rowIterator.next().getSheet().getRow(rowStreetNameSupplier9);
                    Cell cellNameSupplier = rowStreetNameSupplier.getCell(cellNameSupplier5);
                    xmlStreamWriter.writeStartElement("cbc:StreetName");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellNameSupplier.toString()));

                    // end <cbc:StreetName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:CityName>
                    int rowCityNameSupplier10 =10;
                    int cellCityNameSupplier5=5;

                    Row rowCityNameSupplier= rowIterator.next().getSheet().getRow(rowCityNameSupplier10);
                    Cell cellCityNameSupplier = rowCityNameSupplier.getCell(cellCityNameSupplier5);
                    xmlStreamWriter.writeStartElement("cbc:CityName");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellCityNameSupplier.toString()));

                    // end <cbc:CityName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:PostalZone>013329</cbc:PostalZone>
                    int rowPostalZone11 =11;
                    int cellPostalZone5=5;

                    Row rowPostalZone= rowIterator.next().getSheet().getRow(rowPostalZone11);
                    Cell cellPostalZone = rowPostalZone.getCell(cellPostalZone5);
                    xmlStreamWriter.writeStartElement("cbc:PostalZone");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellPostalZone.toString()));

                    //end <cbc:PostalZone>013329</cbc:PostalZone>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // <cbc:CountrySubentity>RO-B</cbc:CountrySubentity>
                    xmlStreamWriter.writeStartElement("cbc:CountrySubentity");
                    xmlStreamWriter.writeCharacters("RO-B");
                    // end  </cbc:CountrySubentity>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

  // <cac:Country>
//          <cbc:IdentificationCode>RO</cbc:IdentificationCode> --> RO
//  </cac:Country>
                    xmlStreamWriter.writeStartElement("cac:Country");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));
                    xmlStreamWriter.writeStartElement("cbc:IdentificationCode");
                    xmlStreamWriter.writeCharacters("RO");
                    // end </cbc:IdentificationCode>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // end  </cac:Country
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


                    // end </cac:PostalAddress>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

//   <cac:PartyTaxScheme>
//        <cbc:CompanyID>RO1234567890</cbc:CompanyID>
//	 <cac:TaxScheme>
//          <cbc:ID>VAT</cbc:ID>
//        </cac:TaxScheme>
//</cac:PartyTaxScheme>

                    //<cac:PartyTaxScheme>
                    xmlStreamWriter.writeStartElement("cac:PartyTaxScheme");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:CompanyID>RO1234567890</cbc:CompanyID>
                    int rowCompanyID =14;
                    int cellCompanyID=5;

                    Row rowCompanyId= rowIterator.next().getSheet().getRow(rowCompanyID);
                    Cell cellrowCompanyId = rowCompanyId.getCell(cellCompanyID);
                    xmlStreamWriter.writeStartElement("cbc:CompanyID");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellrowCompanyId.toString()));

                    // end </cbc:CompanyID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

 // <cac:TaxScheme>
//          <cbc:ID>VAT</cbc:ID>

//  </cac:TaxScheme>

                    // <cac:TaxScheme>
                    xmlStreamWriter.writeStartElement("cac:TaxScheme");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                  //  <cbc:ID>VAT</cbc:ID>
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters("VAT");

                    // end </cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //end  <cac:TaxScheme>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // end <cac:PartyTaxScheme>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


//   <cac:PartyLegalEntity>
//        <cbc:RegistrationName>Seller SRL</cbc:RegistrationName>
//        <cbc:CompanyLegalForm>J40/12345/1998</cbc:CompanyLegalForm>
//   </cac:PartyLegalEntity>


                    //<cac:PartyLegalEntity>
                    xmlStreamWriter.writeStartElement("cac:PartyLegalEntity");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:RegistrationName>Seller SRL</cbc:RegistrationName>
                    xmlStreamWriter.writeStartElement("cbc:RegistrationName");
                    xmlStreamWriter.writeCharacters(cell5.toString());

                    // end </cbc:RegistrationName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:CompanyLegalForm>J40/12345/1998</cbc:CompanyLegalForm>
                    int rowCompanyLegalForm =15;
                    int cellCompanyLegalForm=5;

                    Row rowCompanyLegal= rowIterator.next().getSheet().getRow(rowCompanyLegalForm);
                    Cell cellCompanyLegal = rowCompanyLegal.getCell(cellCompanyLegalForm);
                    xmlStreamWriter.writeStartElement("cbc:CompanyLegalForm");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellCompanyLegal.toString()));

                    // end </cbc:CompanyLegalForm>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // end </cac:PartyLegalEntity>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


                    //end </cac:Party>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // end </cac:AccountingSupplierParty>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

        /* CUSTOMER*/
         //<cac:AccountingCustomerParty> <!-- BG-7 CUMPĂRĂTOR -->
    //<cac:Party>

                    //<cac:AccountingCustomerParty>
                    xmlStreamWriter.writeStartElement("cac:AccountingCustomerParty");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // <cac:Party>
                    xmlStreamWriter.writeStartElement("cac:Party");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

//     <cac:PartyIdentification>
//	        <cbc:ID>123456</cbc:ID>
//      </cac:PartyIdentification>

                    // <cac:PartyIdentification>
                    xmlStreamWriter.writeStartElement("cac:PartyIdentification");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:ID>123456</cbc:ID>
                    int rowIdS =10;
                    int cellIdS=1;

                    Row rowS= rowIterator.next().getSheet().getRow(rowIdS);
                    Cell cellS= rowS.getCell(cellIdS);
                    DataFormatter dataFormatter = new DataFormatter();
                    String formattedValue = dataFormatter.formatCellValue(cellS);
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters(formattedValue);

                    //end </cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // end </cac:PartyIdentification>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


//        <cac:PartyName>
//        <cbc:Name>Buyer name</cbc:Name>
//      </cac:PartyName>

                    //<cac:PartyName>
                    xmlStreamWriter.writeStartElement("cac:PartyName");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:Name>Buyer name</cbc:Name>
                    int rowCustomerPartyName =7;
                    int cellCustomerPartyName=1;

                    Row rowC= rowIterator.next().getSheet().getRow(rowCustomerPartyName);
                    Cell cellC= rowC.getCell(cellCustomerPartyName);
                    xmlStreamWriter.writeStartElement("cbc:Name");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellC.toString()));

                    // end </cbc:Name>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //end </cac:PartyName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

//   <cac:PostalAddress>
//        <cbc:StreetName>BD DECEBAL NR 1 ET1</cbc:StreetName>
//        <cbc:CityName>ARAD</cbc:CityName>
//        <cbc:PostalZone>123456</cbc:PostalZone> --> nu l-am scris
//        <cbc:CountrySubentity>RO-AR</cbc:CountrySubentity>

//        <cac:Country>
//          <cbc:IdentificationCode>RO</cbc:IdentificationCode>
//        </cac:Country>
//      </cac:PostalAddress>


                  //  <cac:PostalAddress>
                    xmlStreamWriter.writeStartElement("cac:PostalAddress");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:StreetName>BD DECEBAL NR 1 ET1</cbc:StreetName>
                    int rowCustomerStreetName=12;
                   // int cellCustomerStreetName=1;

                    Row rowStreetName= rowIterator.next().getSheet().getRow(rowCustomerStreetName);
                    Cell cellStreetName= rowStreetName.getCell(1);
                    xmlStreamWriter.writeStartElement("cbc:StreetName");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellStreetName.toString()));

                    // end </cbc:StreetName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:CityName>Bucuresti</cbc:CityName>
                    int rowCustomerCityName=13;
                    Row rowCityName= rowIterator.next().getSheet().getRow(rowCustomerCityName);
                    Cell cellCityName= rowCityName.getCell(1);
                    xmlStreamWriter.writeStartElement("cbc:CityName");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellCityName.toString()));

                    // end </cbc:CityName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // <cbc:CountrySubentity>RO-AR</cbc:CountrySubentity>
                    xmlStreamWriter.writeStartElement("cbc:CountrySubentity");
                    xmlStreamWriter.writeCharacters("RO-B");

                    // end </cbc:CountrySubentity>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

//    <cac:Country>
//          <cbc:IdentificationCode>RO</cbc:IdentificationCode> -->  se repeta tagul<!-- Codul tarii (BT-55)***-->
//   </cac:Country>

                    xmlStreamWriter.writeStartElement("cac:Country");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));
                    xmlStreamWriter.writeStartElement("cbc:IdentificationCode");
                    xmlStreamWriter.writeCharacters("RO");
                    // end </cbc:IdentificationCode>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // end  </cac:Country
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));




                    // end  </cac:PostalAddress>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


//   <cac:PartyTaxScheme>
//              <cbc:CompanyID>RO987456123</cbc:CompanyID>
//        <cac:TaxScheme>
//              <cbc:ID>VAT</cbc:ID>
//        </cac:TaxScheme>
//  </cac:PartyTaxScheme>


                    //   <cac:PartyTaxScheme>
                    xmlStreamWriter.writeStartElement("cac:PartyTaxScheme");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                     //<cbc:CompanyID>RO987456123</cbc:CompanyID>

                    int rowCompanyIdS =9;
                    int cellCompanyIdS=1;

                    Row rowCompanyS= rowIterator.next().getSheet().getRow(rowCompanyIdS);
                    Cell cellrowCompanyS = rowCompanyS.getCell(1);
                    xmlStreamWriter.writeStartElement("cbc:CompanyID");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellrowCompanyS.toString()));

                    // end </cbc:CompanyID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // <cac:TaxScheme>
                    xmlStreamWriter.writeStartElement("cac:TaxScheme");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //  <cbc:ID>VAT</cbc:ID>
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters("VAT");

                    // end </cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //end  <cac:TaxScheme>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


                    // end </cac:PartyTaxScheme>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

//<cac:PartyLegalEntity>
//        <cbc:RegistrationName>Buyer SRL</cbc:RegistrationName>
//        <cbc:CompanyID>J02/321/2010</cbc:CompanyID>
//	</cac:PartyLegalEntity>


                    //<cac:PartyLegalEntity>
                    xmlStreamWriter.writeStartElement("cac:PartyLegalEntity");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:RegistrationName>Buyer SRL</cbc:RegistrationName>
                    int rowRegistrationNameBuyer =7;
                    int cellRegistrationNameBuyer=1;

                    Row rowRegistrationName= rowIterator.next().getSheet().getRow(rowRegistrationNameBuyer);
                    Cell cellRegistrationName = rowRegistrationName.getCell(1);
                    xmlStreamWriter.writeStartElement("cbc:RegistrationName");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellRegistrationName.toString()));

                    // and </cbc:RegistrationName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:CompanyID>J02/321/2010</cbc:CompanyID>
                    int rowRegistrationNameBuyerCompanyID =11;
                    int cellRegistrationNameBuyerCompanyID=1;

                    Row rowRegistrationNameC= rowIterator.next().getSheet().getRow(rowRegistrationNameBuyerCompanyID);
                    Cell cellRegistrationNameC = rowRegistrationNameC.getCell(1);
                    xmlStreamWriter.writeStartElement("cbc:CompanyID");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellRegistrationNameC.toString()));

                    // and </cbc:CompanyID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // and </cac:PartyLegalEntity>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // end </cac:Party>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    // end </cac:AccountingCustomerParty>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));
/* END---> CUSTOMER*/


//<cac:PaymentMeans> <!-- BG-16 INSTRUCŢIUNI DE PLATĂ -->
//    <cbc:PaymentMeansCode>31</cbc:PaymentMeansCode>
//    <cac:PayeeFinancialAccount>
//      <cbc:ID>RO80RNCB0067054355123456</cbc:ID>
//    </cac:PayeeFinancialAccount>
//  </cac:PaymentMeans>

                    //<cac:PaymentMeans>
                    xmlStreamWriter.writeStartElement("cac:PaymentMeans");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //<cbc:PaymentMeansCode>42</cbc:PaymentMeansCode>
                    int rowPaymentMeansCode =43;
                    int cellPaymentMeansCode=1;

                    Row rowPaymentMeansCode1= rowIterator.next().getSheet().getRow(rowPaymentMeansCode);
                    Cell cellPaymentMeansCode1 = rowPaymentMeansCode1.getCell(1);
                    DataFormatter dataFormatterPaymentMeansCode = new DataFormatter();
                    String formattedValuePaymentMeansCode= dataFormatter.formatCellValue(cellPaymentMeansCode1);
                    xmlStreamWriter.writeStartElement("cbc:PaymentMeansCode");
                    xmlStreamWriter.writeCharacters(formattedValuePaymentMeansCode);

                    // end </cbc:PaymentMeansCode>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                // <cac:PayeeFinancialAccount>
                    xmlStreamWriter.writeStartElement("cac:PayeeFinancialAccount");
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                // <cbc:ID>RO80RNCB0067054355123456</cbc:ID> --> IBAN (BT-84)
                    int rowIBAN =16;
                    int cellIBAN=5;

                    Row rowPaymentMeansCodeIBAN= rowIterator.next().getSheet().getRow(rowIBAN);
                    Cell cellPaymentMeansCodeIBAN = rowPaymentMeansCodeIBAN.getCell(5);
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters(cellPaymentMeansCodeIBAN.toString());



                    // end </cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


                // END   </cac:PayeeFinancialAccount>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));

                    //end </cac:PaymentMeans>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty("line.separator"));


//// all

                    xmlStreamWriter.writeEndElement(); // Close root element
                    xmlStreamWriter.writeEndDocument();
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.close();
                    xmlFile.close();
                    workbook.close();
                    //   }
                }
            }

        } catch (XMLStreamException e) {
            throw new RuntimeException(e);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


//    public static void f() throws XMLStreamException {
//        DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
//        HSSFCell cell1 = null;
//        String requiredDate = df.format(cell1.getDateCellValue());
//        XMLStreamWriter xmlStreamWriter = null;
//        xmlStreamWriter.writeCharacters(requiredDate);
//    }
}





