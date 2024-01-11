package com.example.eFactura;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Marshaller;
import java.io.File;

public class UBLService {

    public static void createUBLXML(Invoice invoice, String filePath) {
        try {
            // Create a JAXBContext
            JAXBContext jaxbContext = JAXBContext.newInstance(Invoice.class);

            // Create a Marshaller
            Marshaller marshaller = jaxbContext.createMarshaller();
            marshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, true);

            // Marshal the object to XML and write to file
            marshaller.marshal(invoice, new File(filePath));
        } catch (JAXBException e) {
            e.printStackTrace();
            // Handle JAXBException (e.g., log, throw, etc.)
        }
    }
}

