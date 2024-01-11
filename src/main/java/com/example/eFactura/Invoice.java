
package com.example.eFactura;

import jakarta.xml.bind.annotation.*;

import java.util.List;

@XmlRootElement(name = "Invoice")
public class Invoice {

    private String ublVersionID;
    private List<InvoiceLine> invoiceLines;

   @XmlElement(name = "UBLVersionID" , namespace = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")
    public String getUblVersionID() {
        return ublVersionID;
    }

    public void setUblVersionID(String ublVersionID) {
        this.ublVersionID = ublVersionID;
    }

    @XmlElement(name = "InvoiceLine")
    public List<InvoiceLine> getInvoiceLines() {
        return invoiceLines;
    }

    public void setInvoiceLines(List<InvoiceLine> invoiceLines) {
        this.invoiceLines = invoiceLines;
    }

    // ... Other properties and methods
}
