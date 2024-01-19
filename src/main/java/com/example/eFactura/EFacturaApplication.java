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

    private static final String LINESEPARATOR = "line.separator";

    public static void main(String[] args) {
        SpringApplication.run(EFacturaApplication.class, args);

        try {
            // Load Excel file
            FileInputStream excelFile = new FileInputStream(new File("C:/Users/Aurika/Desktop/eFactura/F8.xlsx"));
            Workbook workbook = new XSSFWorkbook(excelFile);

            // Create XML file
            FileOutputStream xmlFile = new FileOutputStream(new File("C:/Users/Aurika/Desktop/eFactura/F8.xml"));
            XMLStreamWriter xmlStreamWriter = XMLOutputFactory.newFactory().createXMLStreamWriter(xmlFile);
            //xmlStreamWriter.writeStartDocument();

            xmlFile.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>".getBytes());

            xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));
            xmlStreamWriter.writeStartElement("Invoice");
            xmlStreamWriter.writeNamespace("", "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2");
            xmlStreamWriter.writeNamespace("cac", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
            xmlStreamWriter.writeNamespace("cbc", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");

            xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

            xmlStreamWriter.writeStartElement("cbc:UBLVersionID");
            xmlStreamWriter.writeCharacters("2.1");
            xmlStreamWriter.writeEndElement();
            xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

            xmlStreamWriter.writeStartElement("cbc:CustomizationID");
            xmlStreamWriter.writeCharacters("urn:cen.eu:en16931:2017#compliant#urn:efactura.mfinante.ro:CIUS-RO:1.0.1");
            xmlStreamWriter.writeEndElement();
            xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));


            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

                Sheet sheet = workbook.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();

                while (rowIterator.hasNext()) {

                    // BT-1
                    int rowInvoiceNo =4;
                    int cellInvoiceNo =5;
                    Row rowInvoiceNoBT_1= rowIterator.next().getSheet().getRow(rowInvoiceNo);
                    Cell cellInvoiceNoBT_1 = rowInvoiceNoBT_1.getCell(cellInvoiceNo);
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters(cellInvoiceNoBT_1.toString());
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));



                    // BT-2 <cbc:IssueDate>
                    Cell dateCell = sheet.getRow(5).getCell(5);
                    Date dateValue = dateCell.getDateCellValue();
                    xmlStreamWriter.writeStartElement("cbc:IssueDate");
                    SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                    String formattedDate = dateFormat.format(dateValue);
                    xmlStreamWriter.writeCharacters(formattedDate);
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));


                    // BT-9 <cbc:DueDate>
                    Cell cellDueDateBT_9 = sheet.getRow(6).getCell(5);
                    Date dateValueBT_9 = cellDueDateBT_9.getDateCellValue();
                    xmlStreamWriter.writeStartElement("cbc:DueDate");
                    String requiredDueDateBT_9 = dateFormat.format(dateValueBT_9);
                    xmlStreamWriter.writeCharacters(requiredDueDateBT_9);
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));


                    // BT-3 <cbc:InvoiceTypeCode>380</cbc:InvoiceTypeCode>
                    xmlStreamWriter.writeStartElement("cbc:InvoiceTypeCode");
                    xmlStreamWriter.writeCharacters("380");
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // BT-22 <cbc:Note></cbc:Note>
                    int rowNote = 39;
                    int cellNote = 0;
                    Row rowNoteBT_22= rowIterator.next().getSheet().getRow(rowNote);
                    Cell cellNoteBT_22 = rowNoteBT_22.getCell(cellNote);
                    xmlStreamWriter.writeStartElement("cbc:Note");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellNoteBT_22.toString()));
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));


                    // BT-5 <cbc:DocumentCurrencyCode>RON</cbc:DocumentCurrencyCode>
                    xmlStreamWriter.writeStartElement("cbc:DocumentCurrencyCode");
                    xmlStreamWriter.writeCharacters("RON");
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));



                   // <!-- BG-4 VÂNZĂTOR -->

                    //<cac:AccountingSupplierParty>
                    xmlStreamWriter.writeStartElement("cac:AccountingSupplierParty");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cac:Party>
                    xmlStreamWriter.writeStartElement("cac:Party");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                     //<cac:PartyName>
                    xmlStreamWriter.writeStartElement("cac:PartyName");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // BT-27 <cbc:Name>Technology Reply S.R.L.</cbc:Name>
                    int rowSupplierName = 2;
                    int cellSupplierName = 8;
                    Row rowSupplierNameBT_27= rowIterator.next().getSheet().getRow(rowSupplierName);
                    Cell cellSupplierNameBT_27 = rowSupplierNameBT_27.getCell(cellSupplierName);
                    xmlStreamWriter.writeStartElement("cbc:Name");
                    xmlStreamWriter.writeCharacters(cellSupplierNameBT_27.toString());
                    // end </cbc:Name>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //  end </cac:PartyName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //<cac:PostalAddress> ADRESA POSTALA (BG-5)
                    xmlStreamWriter.writeStartElement("cac:PostalAddress");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // BT-35 <cbc:StreetName>
                    int rowStreetNameSupplier =2;
                    int cellStreetNameSupplier=8;
                    Row rowStreetNameSupplierBT_35= rowIterator.next().getSheet().getRow(rowStreetNameSupplier);
                    Cell cellStreetNameSupplierBT_35 = rowStreetNameSupplierBT_35.getCell(cellStreetNameSupplier);
                    xmlStreamWriter.writeStartElement("cbc:StreetName");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellStreetNameSupplierBT_35.toString()));
                    // end <cbc:StreetName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //BT-37 <cbc:CityName>
                    int rowCityNameSupplier =8;
                    int cellCityNameSupplier=8;
                    Row rowCityNameSupplierBT_37= rowIterator.next().getSheet().getRow(rowCityNameSupplier);
                    Cell cellCityNameSupplierBT_37 = rowCityNameSupplierBT_37.getCell(cellCityNameSupplier);
                    xmlStreamWriter.writeStartElement("cbc:CityName");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellCityNameSupplierBT_37.toString()));
                    // end <cbc:CityName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //BT-38 <cbc:PostalZone>013329</cbc:PostalZone>
                    int rowPostalZone =9;
                    int cellPostalZone=8;
                    Row rowPostalZoneBT_38= rowIterator.next().getSheet().getRow(rowPostalZone);
                    Cell cellPostalZoneBT_38 = rowPostalZoneBT_38.getCell(cellPostalZone);
                    xmlStreamWriter.writeStartElement("cbc:PostalZone");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellPostalZoneBT_38.toString()));
                    //end </cbc:PostalZone>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // BT-39 <cbc:CountrySubentity>RO-B</cbc:CountrySubentity>
                    xmlStreamWriter.writeStartElement("cbc:CountrySubentity");
                    xmlStreamWriter.writeCharacters("RO-B");
                    // end </cbc:CountrySubentity>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cac:Country>
                    xmlStreamWriter.writeStartElement("cac:Country");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));
                    // BT-40 <cbc:IdentificationCode>
                    xmlStreamWriter.writeStartElement("cbc:IdentificationCode");
                    xmlStreamWriter.writeCharacters("RO");
                    // end </cbc:IdentificationCode>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));
                    // end  </cac:Country
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end </cac:PostalAddress>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //<cac:PartyTaxScheme>
                    xmlStreamWriter.writeStartElement("cac:PartyTaxScheme");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //BT-31 <cbc:CompanyID>RO1234567890</cbc:CompanyID>
                    int rowCompanyID =4;
                    int cellCompanyID=8;
                    Row rowCompanyIdBT_31= rowIterator.next().getSheet().getRow(rowCompanyID);
                    Cell cellrowCompanyIdBT_31 = rowCompanyIdBT_31.getCell(cellCompanyID);
                    xmlStreamWriter.writeStartElement("cbc:CompanyID");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellrowCompanyIdBT_31.toString()));
                    // end </cbc:CompanyID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                     // <cac:TaxScheme>
                    xmlStreamWriter.writeStartElement("cac:TaxScheme");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));
                    //  <cbc:ID>VAT</cbc:ID>
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters("VAT");
                    // end </cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));
                    //end <cac:TaxScheme>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end </cac:PartyTaxScheme>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));


//   <cac:PartyLegalEntity>
//        <cbc:RegistrationName>Seller SRL</cbc:RegistrationName>
//        <cbc:CompanyLegalForm>J40/12345/1998</cbc:CompanyLegalForm>
//   </cac:PartyLegalEntity>


                    //<cac:PartyLegalEntity>
                    xmlStreamWriter.writeStartElement("cac:PartyLegalEntity");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //<cbc:RegistrationName>Seller SRL</cbc:RegistrationName>
                    xmlStreamWriter.writeStartElement("cbc:RegistrationName");
                    xmlStreamWriter.writeCharacters(cellSupplierNameBT_27.toString());
                    // end </cbc:RegistrationName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //BT-30 <cbc:CompanyLegalForm>J40/12345/1998</cbc:CompanyLegalForm>
                    int rowCompanyLegalForm =6;
                    int cellCompanyLegalForm=8;
                    Row rowCompanyLegalBT_30= rowIterator.next().getSheet().getRow(rowCompanyLegalForm);
                    Cell cellCompanyLegalBT_30 = rowCompanyLegalBT_30.getCell(cellCompanyLegalForm);
                    xmlStreamWriter.writeStartElement("cbc:CompanyLegalForm");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellCompanyLegalBT_30.toString()));
                    // end </cbc:CompanyLegalForm>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end </cac:PartyLegalEntity>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));


                    //end </cac:Party>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end </cac:AccountingSupplierParty>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));


        /* CUSTOMER*/
         //<cac:AccountingCustomerParty> <!-- BG-7 CUMPĂRĂTOR -->

                    //<cac:AccountingCustomerParty>
                    xmlStreamWriter.writeStartElement("cac:AccountingCustomerParty");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cac:Party>
                    xmlStreamWriter.writeStartElement("cac:Party");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cac:PartyIdentification>
                    xmlStreamWriter.writeStartElement("cac:PartyIdentification");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //<cbc:ID>16986329</cbc:ID> Fiscal Code
                    int rowIdCustomer =5;
                    int cellIdCustomer=1;
                    Row rowS= rowIterator.next().getSheet().getRow(rowIdCustomer);
                    Cell cellS= rowS.getCell(cellIdCustomer);
                    DataFormatter dataFormatter = new DataFormatter();
                    String formattedValue = dataFormatter.formatCellValue(cellS);
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters(formattedValue);
                    //end </cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end </cac:PartyIdentification>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));


                    //<cac:PartyName>
                    xmlStreamWriter.writeStartElement("cac:PartyName");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //BT-44 <cbc:Name>Buyer name</cbc:Name>
                    int rowCustomerPartyName =2;
                    int cellCustomerPartyName=1;
                    Row rowCustomerPartyNameBT_44= rowIterator.next().getSheet().getRow(rowCustomerPartyName);
                    Cell cellCustomerPartyNameBT_44= rowCustomerPartyNameBT_44.getCell(cellCustomerPartyName);
                    xmlStreamWriter.writeStartElement("cbc:Name");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellCustomerPartyNameBT_44.toString()));
                    // end </cbc:Name>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //end </cac:PartyName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

//   <cac:PostalAddress>
//        <cbc:StreetName>BD DECEBAL NR 1 ET1</cbc:StreetName>
//        <cbc:CityName>ARAD</cbc:CityName>
//        <cbc:PostalZone>123456</cbc:PostalZone> --> nu l-am scris
//        <cbc:CountrySubentity>RO-AR</cbc:CountrySubentity>

//        <cac:Country>
//          <cbc:IdentificationCode>RO</cbc:IdentificationCode>
//        </cac:Country>
//      </cac:PostalAddress>


                  //  ADRESA POSTALA (BG-8) <cac:PostalAddress>
                    xmlStreamWriter.writeStartElement("cac:PostalAddress");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // BT-50 <cbc:StreetName>
                    int rowCustomerStreetName=7;
                    int cellCustomerStreetName=1;

                    Row rowStreetNameBT_50= rowIterator.next().getSheet().getRow(rowCustomerStreetName);
                    Cell cellStreetNameBT_50= rowStreetNameBT_50.getCell(cellCustomerStreetName);
                    xmlStreamWriter.writeStartElement("cbc:StreetName");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellStreetNameBT_50.toString()));
                    // end </cbc:StreetName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //BT-52 <cbc:CityName>SECTOR2</cbc:CityName>
                    int rowCustomerCityName=8;
                    int cellCustomerCityName=1;
                    Row rowCustomerCityNameBT_52= rowIterator.next().getSheet().getRow(rowCustomerCityName);
                    Cell cellCustomerCityNameBT_52= rowCustomerCityNameBT_52.getCell(cellCustomerCityName);
                    xmlStreamWriter.writeStartElement("cbc:CityName");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellCustomerCityNameBT_52.toString()));
                    // end </cbc:CityName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cbc:CountrySubentity>RO-B</cbc:CountrySubentity>
                    xmlStreamWriter.writeStartElement("cbc:CountrySubentity");
                    xmlStreamWriter.writeCharacters("RO-B");

                    // end </cbc:CountrySubentity>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //BT-55 <cac:Country>
                    xmlStreamWriter.writeStartElement("cac:Country");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));
                    xmlStreamWriter.writeStartElement("cbc:IdentificationCode");
                    xmlStreamWriter.writeCharacters("RO");
                    // end </cbc:IdentificationCode>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));
                    // end  </cac:Country
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end  </cac:PostalAddress>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cac:PartyTaxScheme>
                    xmlStreamWriter.writeStartElement("cac:PartyTaxScheme");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                     //<cbc:CompanyID>RO987456123</cbc:CompanyID>
                    int rowCompanyIdS =4;
                    int cellCompanyIdS=1;
                    Row rowCompanyS= rowIterator.next().getSheet().getRow(rowCompanyIdS);
                    Cell cellrowCompanyS = rowCompanyS.getCell(cellCompanyIdS);
                    xmlStreamWriter.writeStartElement("cbc:CompanyID");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellrowCompanyS.toString()));
                    // end </cbc:CompanyID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cac:TaxScheme>
                    xmlStreamWriter.writeStartElement("cac:TaxScheme");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //  <cbc:ID>VAT</cbc:ID>
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters("VAT");

                    // end </cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //end  <cac:TaxScheme>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end </cac:PartyTaxScheme>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

//<cac:PartyLegalEntity>
//        <cbc:RegistrationName>Buyer SRL</cbc:RegistrationName>
//        <cbc:CompanyID>J02/321/2010</cbc:CompanyID>
//	</cac:PartyLegalEntity>


                    //<cac:PartyLegalEntity>
                    xmlStreamWriter.writeStartElement("cac:PartyLegalEntity");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //<cbc:RegistrationName>Buyer SRL</cbc:RegistrationName>
                    int rowRegistrationNameBuyer =2;
                    int cellRegistrationNameBuyer=1;
                    Row rowRegistrationName= rowIterator.next().getSheet().getRow(rowRegistrationNameBuyer);
                    Cell cellRegistrationName = rowRegistrationName.getCell(cellRegistrationNameBuyer);
                    xmlStreamWriter.writeStartElement("cbc:RegistrationName");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellRegistrationName.toString()));
                    // and </cbc:RegistrationName>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //BT-47<cbc:CompanyID>J40/19595/2004</cbc:CompanyID>
                    int rowRegistrationNameBuyerCompanyID =6;
                    int cellRegistrationNameBuyerCompanyID=1;
                    Row rowRegistrationNameC= rowIterator.next().getSheet().getRow(rowRegistrationNameBuyerCompanyID);
                    Cell cellRegistrationNameC = rowRegistrationNameC.getCell(cellRegistrationNameBuyerCompanyID);
                    xmlStreamWriter.writeStartElement("cbc:CompanyID");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellRegistrationNameC.toString()));
                    // and </cbc:CompanyID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // and </cac:PartyLegalEntity>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end </cac:Party>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end </cac:AccountingCustomerParty>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));
/* END---> CUSTOMER*/


//<cac:PaymentMeans> <!-- BG-16 INSTRUCŢIUNI DE PLATĂ -->
//    <cbc:PaymentMeansCode>31</cbc:PaymentMeansCode>
//    <cac:PayeeFinancialAccount>
//      <cbc:ID>RO80RNCB0067054355123456</cbc:ID>
//    </cac:PayeeFinancialAccount>
//  </cac:PaymentMeans>

                    //<cac:PaymentMeans>
                    xmlStreamWriter.writeStartElement("cac:PaymentMeans");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //BT-81<cbc:PaymentMeansCode>42</cbc:PaymentMeansCode>
                    int rowPaymentMeansCode =41;
                    int cellPaymentMeansCode=1;
                    Row rowPaymentMeansCodeBT_81= rowIterator.next().getSheet().getRow(rowPaymentMeansCode);
                    Cell cellPaymentMeansCodeBT_81 = rowPaymentMeansCodeBT_81.getCell(cellPaymentMeansCode);
                    String formattedValuePaymentMeansCode= dataFormatter.formatCellValue(cellPaymentMeansCodeBT_81);
                    xmlStreamWriter.writeStartElement("cbc:PaymentMeansCode");
                    xmlStreamWriter.writeCharacters(formattedValuePaymentMeansCode);
                    // end </cbc:PaymentMeansCode>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                // <cac:PayeeFinancialAccount>
                    xmlStreamWriter.writeStartElement("cac:PayeeFinancialAccount");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                // BT-84 <cbc:ID>RO80RNCB0067054355123456</cbc:ID> --> IBAN
                    int rowIBAN =10;
                    int cellIBAN=8;
                    Row rowPaymentMeansCodeIBAN= rowIterator.next().getSheet().getRow(rowIBAN);
                    Cell cellPaymentMeansCodeIBAN = rowPaymentMeansCodeIBAN.getCell(cellIBAN);
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters(cellPaymentMeansCodeIBAN.toString());
                    // end </cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                // END  </cac:PayeeFinancialAccount>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //end </cac:PaymentMeans>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cac:TaxTotal>
                    xmlStreamWriter.writeStartElement("cac:TaxTotal");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // BT-110 <cbc:TaxAmount currencyID="RON">7238.06</cbc:TaxAmount> --> Valoare totala TVA (BT-110)
                    int rowTaxAmount =35;
                    int cellTaxAmount=8;
                    Row rowTaxAmountBT_110= rowIterator.next().getSheet().getRow(rowTaxAmount);
                    Cell cellTaxAmountBT_110 = rowTaxAmountBT_110.getCell(cellTaxAmount);
                    double doubleValue = cellTaxAmountBT_110.getNumericCellValue();
                    String formattedValue1 = String.format("%.2f", doubleValue);
                    xmlStreamWriter.writeStartElement("cbc:TaxAmount");
                    xmlStreamWriter.writeAttribute("currencyID","RON");
                    xmlStreamWriter.writeCharacters(formattedValue1);
                    // end </cbc:TaxAmount>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //<cac:TaxSubtotal>
                    xmlStreamWriter.writeStartElement("cac:TaxSubtotal");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                     //BT-116 <cbc:TaxableAmount currencyID="RON">38095.05</cbc:TaxableAmount> <!--BT-116-->
                    int rowTaxableAmount =18;
                    int cellTaxableAmount=5;
                    Row rowTaxableAmountBT_116= rowIterator.next().getSheet().getRow(rowTaxableAmount);
                    Cell cellTaxableAmountBT_116 = rowTaxableAmountBT_116.getCell(cellTaxableAmount);
                    double doubleValueTaxableAmountVAT = cellTaxableAmountBT_116.getNumericCellValue();
                    String formattedValueTaxableAmountVAT = String.format("%.2f", doubleValueTaxableAmountVAT);
                    xmlStreamWriter.writeStartElement("cbc:TaxableAmount");
                    xmlStreamWriter.writeAttribute("currencyID","RON");
                    xmlStreamWriter.writeCharacters(formattedValueTaxableAmountVAT);
                    // end </cbc:TaxableAmount>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //BT-117<cbc:TaxAmount currencyID="RON">7238.06</cbc:TaxAmount> --> <!--BT-117-->
                    int rowTaxAmountBT_117 =18;
                    int cellTaxAmountBT_117=8;
                    Row rowTaxAmountVAT_BT_117= rowIterator.next().getSheet().getRow(rowTaxAmountBT_117);
                    Cell cellTaxAmountVAT_BT_117 = rowTaxAmountVAT_BT_117.getCell(cellTaxAmountBT_117);
                    double doubleValueBT_117 = cellTaxAmountVAT_BT_117.getNumericCellValue();
                    String formattedValueBT_117 = String.format("%.2f", doubleValueBT_117);
                    xmlStreamWriter.writeStartElement("cbc:TaxAmount");
                    xmlStreamWriter.writeAttribute("currencyID","RON");
                    xmlStreamWriter.writeCharacters(formattedValueBT_117);

                    // end </cbc:TaxAmount>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cac:TaxCategory>
                    xmlStreamWriter.writeStartElement("cac:TaxCategory");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //<cbc:ID>S</cbc:ID> --> !--BT-118-->
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters("S");
                    //end </cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                   // <cbc:Percent>19.00</cbc:Percent>
                    xmlStreamWriter.writeStartElement("cbc:Percent");
                    xmlStreamWriter.writeCharacters("19.00");

                    // end </cbc:Percent>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cac:TaxScheme>
                    xmlStreamWriter.writeStartElement("cac:TaxScheme");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //<cbc:ID>VAT</cbc:ID>
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters("VAT");

                    // END</cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // END</cac:TaxScheme>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end  </cac:TaxCategory>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // END </cac:TaxSubtotal>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // END </cac:TaxTotal>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                   // <cac:LegalMonetaryTotal> <!-- BG-22 TOTALURI ALE DOCUMENTULUI-->
                    xmlStreamWriter.writeStartElement("cac:LegalMonetaryTotal");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //BT-106 <cbc:LineExtensionAmount currencyID="RON">
                    int rowLineExtensionAmount=35;
                    int cellLineExtensionAmount=6;
                    Row rowLineExtensionAmount_BT106= rowIterator.next().getSheet().getRow(rowLineExtensionAmount);
                    Cell cellLineExtensionAmount_BT106 = rowLineExtensionAmount_BT106.getCell(cellLineExtensionAmount);
                    double doubleLineExtensionAmount_BT106= cellLineExtensionAmount_BT106.getNumericCellValue();
                    String formattedLineExtensionAmount_BT106 = String.format("%.2f", doubleLineExtensionAmount_BT106);
                    xmlStreamWriter.writeStartElement("cbc:LineExtensionAmount");
                    xmlStreamWriter.writeAttribute("currencyID","RON");
                    xmlStreamWriter.writeCharacters(formattedLineExtensionAmount_BT106);
                    // end </cbc:LineExtensionAmount>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //BT-109 <cbc:TaxExclusiveAmount currencyID="RON">
                    int rowTaxExclusiveAmount=35;
                    int cellTaxExclusiveAmount=6;
                    Row rowTaxExclusiveAmount_BT109= rowIterator.next().getSheet().getRow(rowTaxExclusiveAmount);
                    Cell cellTaxExclusiveAmount_BT109 = rowTaxExclusiveAmount_BT109.getCell(cellTaxExclusiveAmount);
                    double doubleTaxExclusiveAmount_BT109= cellTaxExclusiveAmount_BT109.getNumericCellValue();
                    String formattedTaxExclusiveAmount_BT109 = String.format("%.2f", doubleTaxExclusiveAmount_BT109);
                    xmlStreamWriter.writeStartElement("cbc:TaxExclusiveAmount");
                    xmlStreamWriter.writeAttribute("currencyID","RON");
                    xmlStreamWriter.writeCharacters(formattedTaxExclusiveAmount_BT109);
                    // END </cbc:TaxExclusiveAmount>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));


                     //BT-112 <cbc:TaxInclusiveAmount currencyID="RON">
                    int rowTaxInclusiveAmount=37;
                    int cellTaxInclusiveAmount=8;
                    Row rowTaxInclusiveAmount_BT112= rowIterator.next().getSheet().getRow(rowTaxInclusiveAmount);
                    Cell cellTaxInclusiveAmount_BT112 = rowTaxInclusiveAmount_BT112.getCell(cellTaxInclusiveAmount);
                    double doubleTaxInclusiveAmount_BT112= cellTaxInclusiveAmount_BT112.getNumericCellValue();
                    String formattedTaxInclusiveAmount_BT112 = String.format("%.2f", doubleTaxInclusiveAmount_BT112);
                    xmlStreamWriter.writeStartElement("cbc:TaxInclusiveAmount");
                    xmlStreamWriter.writeAttribute("currencyID","RON");
                    xmlStreamWriter.writeCharacters(formattedTaxInclusiveAmount_BT112);
                     // end </cbc:TaxInclusiveAmount>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));


                    //BT-115 <cbc:PayableAmount currencyID="RON">
                    int rowPayableAmount=37;
                    int cellPayableAmount=8;
                    Row rowPayableAmount_BT115= rowIterator.next().getSheet().getRow(rowPayableAmount);
                    Cell cellPayableAmount_BT115 = rowPayableAmount_BT115.getCell(cellPayableAmount);
                    double doublePayableAmount_BT115= cellPayableAmount_BT115.getNumericCellValue();
                    String formattedPayableAmount_BT115 = String.format("%.2f", doublePayableAmount_BT115);
                    xmlStreamWriter.writeStartElement("cbc:PayableAmount");
                    xmlStreamWriter.writeAttribute("currencyID","RON");
                    xmlStreamWriter.writeCharacters(formattedPayableAmount_BT115);
                    // end </cbc:PayableAmount>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // END <cac:LegalMonetaryTotal>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

// <cac:InvoiceLine>
//      <cbc:ID>1</cbc:ID>
//      <!--  BT-115 Suma de plata  -->
//      <cbc:InvoicedQuantity unitCode="C62">1</cbc:InvoicedQuantity>
//      <cbc:LineExtensionAmount currencyID="RON">38095.05</cbc:LineExtensionAmount>

//      <cac:Item>
//         <cbc:Name>Maintenance November-December</cbc:Name>
//         <cac:SellersItemIdentification>
//            <cbc:ID>PD102238</cbc:ID>
//         </cac:SellersItemIdentification>

//         <cac:ClassifiedTaxCategory>
//            <cbc:ID>S</cbc:ID>
//            <cbc:Percent>19.0</cbc:Percent>
//            <cac:TaxScheme>
//               <cbc:ID>VAT</cbc:ID>
//            </cac:TaxScheme>
//         </cac:ClassifiedTaxCategory>

//      </cac:Item>

//      <cac:Price>
//         <cbc:PriceAmount  currencyID="RON">38095.05</cbc:PriceAmount >
//         <cbc:BaseQuantity  unitCode="C62">1</cbc:BaseQuantity >
//      </cac:Price>
//   </cac:InvoiceLine>


                    //<cac:InvoiceLine>
                    xmlStreamWriter.writeStartElement("cac:InvoiceLine");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //<cbc:ID>1</cbc:ID>
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters("1");
                    // END </cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //<cbc:InvoicedQuantity unitCode="C62">1</cbc:InvoicedQuantity>
                    xmlStreamWriter.writeStartElement("cbc:InvoicedQuantity");
                    xmlStreamWriter.writeAttribute("unitCode","C62");
                    xmlStreamWriter.writeCharacters("1");
                    // END </cbc:InvoicedQuantity>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));


                    //BT-131 <cbc:LineExtensionAmount currencyID="RON">38095.05</cbc:LineExtensionAmount>
                    int rowLineExtensionAmountInvoiceLine=18;
                    int cellLineExtensionAmountInvoiceLine=6;
                    Row rowLineExtensionAmount_BT131= rowIterator.next().getSheet().getRow(rowLineExtensionAmountInvoiceLine);
                    Cell cellLineExtensionAmount_BT131 = rowLineExtensionAmount_BT131.getCell(cellLineExtensionAmountInvoiceLine);
                    double doubleLineExtensionAmount_BT131= cellLineExtensionAmount_BT131.getNumericCellValue();
                    String formattedLineExtensionAmount_BT131 = String.format("%.2f", doubleLineExtensionAmount_BT131);
                    xmlStreamWriter.writeStartElement("cbc:LineExtensionAmount");
                    xmlStreamWriter.writeAttribute("currencyID","RON");
                    xmlStreamWriter.writeCharacters(formattedLineExtensionAmount_BT131);
                    // END </cbc:LineExtensionAmount>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //<cac:Item>
                    xmlStreamWriter.writeStartElement("cac:Item");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                   // <cbc:Name>Maintenance November-December</cbc:Name>
                    int rowInvoiceName =18;
                    int cellInvoiceName=1;
                    Row rowName= rowIterator.next().getSheet().getRow(rowInvoiceName);
                    Cell cellName = rowName.getCell(cellInvoiceName);
                    xmlStreamWriter.writeStartElement("cbc:Name");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellName.toString()));
                   // end </cbc:Name>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cac:SellersItemIdentification>
                    xmlStreamWriter.writeStartElement("cac:SellersItemIdentification");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cbc:ID>PD102238</cbc:ID>
                    int rowSellersItemIdentification =18;
                    int cellSellersItemIdentification=0;
                    Row rowSellersItem= rowIterator.next().getSheet().getRow(rowSellersItemIdentification);
                    Cell cellSellersItem= rowSellersItem.getCell(cellSellersItemIdentification);
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters(String.valueOf(cellSellersItem.toString()));
                    // End </cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // END <cac:SellersItemIdentification>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cac:ClassifiedTaxCategory>
                    xmlStreamWriter.writeStartElement("cac:ClassifiedTaxCategory");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                   // <cbc:ID>S</cbc:ID>
                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters("S");
                    // end </cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //  <cbc:Percent>19.0</cbc:Percent>
                    xmlStreamWriter.writeStartElement("cbc:Percent");
                    xmlStreamWriter.writeCharacters("19.00");

                    // END </cbc:Percent>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cac:TaxScheme>
                    xmlStreamWriter.writeStartElement("cac:TaxScheme");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    xmlStreamWriter.writeStartElement("cbc:ID");
                    xmlStreamWriter.writeCharacters("VAT");
                    // end </cbc:ID>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end </cac:TaxScheme>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end <cac:ClassifiedTaxCategory>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end </cac:Item>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cac:Price>
                    xmlStreamWriter.writeStartElement("cac:Price");
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // BT-146 <cbc:PriceAmount  currencyID="RON">38095.05
                    int rowPriceAmount=18;
                    int cellPriceAmount=5;
                    Row rowPriceAmount_BT146= rowIterator.next().getSheet().getRow(rowPriceAmount);
                    Cell cellPriceAmount_BT146 = rowPriceAmount_BT146.getCell(cellPriceAmount);
                    double doublePriceAmount_BT146= cellPriceAmount_BT146.getNumericCellValue();
                    String formattedPriceAmount_BT146 = String.format("%.2f", doublePriceAmount_BT146);
                    xmlStreamWriter.writeStartElement("cbc:PriceAmount");
                    xmlStreamWriter.writeAttribute("currencyID","RON");
                    xmlStreamWriter.writeCharacters(formattedPriceAmount_BT146);
                    // END </cbc:PriceAmount >
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // <cbc:BaseQuantity  unitCode="C62">1
                    xmlStreamWriter.writeStartElement("cbc:BaseQuantity");
                    xmlStreamWriter.writeAttribute("unitCode","C62");
                    xmlStreamWriter.writeCharacters("1");

                    // end </cbc:BaseQuantity >
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    // end </cac:Price>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    //End <cac:InvoiceLine>
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.writeCharacters(System.getProperty(LINESEPARATOR));

                    xmlStreamWriter.writeEndElement(); // Close root element
                    xmlStreamWriter.writeEndDocument();
                    xmlStreamWriter.writeEndElement();
                    xmlStreamWriter.close();
                    xmlFile.close();
                    workbook.close();
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

}





