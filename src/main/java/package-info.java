@XmlSchema(
        namespace = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
        elementFormDefault = XmlNsForm.QUALIFIED,
        attributeFormDefault = XmlNsForm.UNQUALIFIED,
        xmlns = {
                @XmlNs(prefix = "cbc", namespaceURI = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
                @XmlNs(prefix = "cac",namespaceURI = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")

        }
)
package com.example;

import jakarta.xml.bind.annotation.XmlNs;
import jakarta.xml.bind.annotation.XmlNsForm;
import jakarta.xml.bind.annotation.XmlSchema;




