/**
 * ============================================================================
 * EU VAT NUMBER VALIDATOR FOR GOOGLE SHEETS
 * ============================================================================
 *
 * This script validates EU VAT numbers using the official European Commission
 * VIES (VAT Information Exchange System) service.
 *
 * @author    Jeroen de Jong <jeroen@telartis.nl>
 * @copyright 2015,2026 Telartis BV
 *
 * @wsdl https://ec.europa.eu/taxation_customs/vies/services/checkVatService.wsdl
 * @url  https://ec.europa.eu/taxation_customs/vies/faq.html
 *
 * ============================================================================
 * INSTALLATION GUIDE
 * ============================================================================
 *
 * Step 1: Open Your Google Sheet
 *   - Open the Google Sheet where you want to validate VAT numbers
 *
 * Step 2: Access Apps Script Editor
 *   - Click on "Extensions" in the menu bar
 *   - Select "Apps Script"
 *   - This opens the Apps Script editor in a new tab
 *
 * Step 3: Create New Script File
 *   - In the Apps Script editor, you'll see a default "Code.gs" file
 *   - Delete any existing code in Code.gs (or create a new file)
 *   - To create a new file: click the "+" next to "Files" and select "Script"
 *   - Name it something like "CheckVAT" (optional)
 *
 * Step 4: Copy This Script
 *   - Copy this entire JavaScript file
 *   - Paste it into the Apps Script editor
 *
 * Step 5: Save the Project
 *   - Click the disk icon or press Ctrl+S (Cmd+S on Mac)
 *   - Give your project a name like "VAT Validator"
 *
 * Step 6: Authorization (First Run Only)
 *   - The first time you use the function, Google will ask for permissions
 *   - Click "Review permissions" when prompted
 *   - Select your Google account
 *   - Click "Advanced" if you see a warning
 *   - Click "Go to [Project Name] (unsafe)" - don't worry, it's your own script
 *   - Click "Allow" to grant the necessary permissions
 *
 * ============================================================================
 * USAGE IN GOOGLE SHEETS
 * ============================================================================
 *
 * Basic Usage:
 *   In any cell, type the formula:
 *   =CHECK_VAT_SERVICE("NL123456789B01")
 *
 * Input Format:
 *   - Include the 2-letter country code (e.g., NL, BE, DE, FR)
 *   - Followed by the VAT number
 *   - Spaces and special characters are automatically removed
 *   - Examples: "NL123456789B01", "BE0123456789", "DE123456789"
 *
 * Output Format:
 *   The function returns an array with 5 columns:
 *   [Request Date, Valid, Company Name, Address, Error Message]
 *
 *   Column A: Request Date (e.g., "2015-03-09+01:00")
 *   Column B: Valid (true/false)
 *   Column C: Company Name (e.g., "Acme Corp B.V.")
 *   Column D: Address (e.g., "KERKSTRAAT 00123\n1234AB AMSTERDAM")
 *   Column E: Error Message (empty if successful)
 *
 * Example Spreadsheet Layout:
 *
 *   |    A         |       B               |   C     |      D      |     E     |      F      |
 *   |--------------|-----------------------|---------|-------------|-----------|-------------|
 *   | VAT Number   | Formula               | Date    | Valid       | Name      | Address     |
 *   | NL1234567... | =CHECK_VAT_SERVICE(A2)| 2015... | true        | ACME COR..| KERKSTRA... |
 *   | BE0123456... | =CHECK_VAT_SERVICE(A3)| 2015... | false       |           | INVALID...  |
 *
 * Advanced Usage - Multiple VAT Numbers:
 *   1. Put VAT numbers in column A (starting row 2)
 *   2. In cell B2, enter: =CHECK_VAT_SERVICE(A2)
 *   3. The result will automatically spread across 5 columns (B through F)
 *   4. Copy the formula down to validate multiple VAT numbers
 *
 * Handling Errors:
 *   - Network errors: "Network error: [message]"
 *   - Invalid format: "Invalid VAT number: too short"
 *   - Service errors: "HTTP error 500: Service unavailable"
 *   - Invalid VAT: Check column E for "INVALID_INPUT" or similar
 *
 * Rate Limiting:
 *   - The script includes a 1-second delay between requests
 *   - This prevents overwhelming the EU VIES service
 *   - For large batches, expect ~60 validations per minute
 *
 * Tips:
 *   - The VIES service may be temporarily unavailable (maintenance)
 *   - Not all EU VAT numbers are in the VIES system immediately
 *   - Some countries may have limited data (name/address might be empty)
 *   - Save your results to avoid re-querying the same VAT numbers
 *
 * ============================================================================
 */

/**
 * Check validity of VAT-numbers of companies registered in EU.
 *
 * @param {vat_number} input The vat number including country code
 * @return requestDate, valid, name, address, error
 * @customfunction
 */
function CHECK_VAT_SERVICE(vat_number) {
    try {
        // Validate input
        if (!vat_number || typeof vat_number !== 'string') {
            return [['', false, '', '', 'Invalid input: VAT number is required']];
        }

        vat_number = vat_number.toUpperCase().replace(/[^A-Z0-9]/g, '');

        // Check minimum length (country code + at least one digit)
        if (vat_number.length < 3) {
            return [['', false, '', '', 'Invalid VAT number: too short']];
        }

        var countryCode = vat_number.substring(0, 2),
            vatNumber   = vat_number.substring(2);

        var url = 'https://ec.europa.eu/taxation_customs/vies/services/checkVatService';

        var payload = '<?xml version="1.0" encoding="UTF-8"?>\
<SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns1="urn:ec.europa.eu:taxud:vies:services:checkVat:types">\
<SOAP-ENV:Body>\
<ns1:checkVat>\
<ns1:countryCode>' + countryCode + '</ns1:countryCode>\
<ns1:vatNumber>' + vatNumber + '</ns1:vatNumber>\
</ns1:checkVat>\
</SOAP-ENV:Body>\
</SOAP-ENV:Envelope>';

        var options = {
            contentType : 'text/xml; charset=utf-8',
            method : 'POST',
            payload: payload,
            muteHttpExceptions: true
        };

        var response, code, xml;

        try {
            response = UrlFetchApp.fetch(url, options);
            code = response.getResponseCode();
            xml = response.getContentText();
        } catch (fetchError) {
            return [['', false, '', '', 'Network error: ' + fetchError.message]];
        }

        // Check HTTP response code
        if (code !== 200) {
            return [['', false, '', '', 'HTTP error ' + code + ': Service unavailable']];
        }

        // Check if we got a response
        if (!xml || xml.length === 0) {
            return [['', false, '', '', 'Empty response from service']];
        }

        var result = [], requestDate = '', valid = false, name = '', address = '', error = '';

        try {
            var doc = XmlService.parse(xml);
            var root = doc.getRootElement();

            if (!root) {
                return [['', false, '', '', 'Invalid XML response: no root element']];
            }

            // Get the Body element (works with both soap: and SOAP-ENV: namespaces)
            var bodyElements = root.getChildren();
            var body = null;
            for (var i = 0; i < bodyElements.length; i++) {
                var name_lower = bodyElements[i].getName().toLowerCase();
                if (name_lower === 'body') {
                    body = bodyElements[i];
                    break;
                }
            }

            if (!body) {
                return [['', false, '', '', 'Invalid XML response: no body element found']];
            }

            // Get the response or fault element inside Body
            var responseElements = body.getChildren();
            if (!responseElements || responseElements.length === 0) {
                return [['', false, '', '', 'Invalid XML response: empty body element']];
            }

            var responseElement = responseElements[0];
            var responseType = responseElement.getName().toLowerCase();

            // Check if it's a SOAP Fault
            if (responseType === 'fault') {
                var faultElements = responseElement.getChildren();
                for (var i = 0; i < faultElements.length; i++) {
                    var tag = faultElements[i].getName().toLowerCase();
                    var value = faultElements[i].getText();
                    if (tag === 'faultstring') {
                        error = value;
                        break;
                    }
                }
                if (!error) {
                    error = 'SOAP Fault received';
                }
            } else {
                // It's a checkVatResponse - parse the data
                var dataElements = responseElement.getChildren();
                for (var i = 0; i < dataElements.length; i++) {
                    var tag = dataElements[i].getName();
                    var value = dataElements[i].getText();
                    if (tag === 'requestDate') requestDate = value;
                    if (tag === 'valid')       valid = value;
                    if (tag === 'name')        name = value;
                    if (tag === 'address')     address = value;
                }
            }
        } catch (parseError) {
            return [['', false, '', '', 'XML parsing error: ' + parseError.message]];
        }

        Utilities.sleep(1000);

        result.push([requestDate, valid, name, address, error]);

        return result;

    } catch (e) {
        // Catch-all for any unexpected errors
        return [['', false, '', '', 'Unexpected error: ' + e.message]];
    }
}


/**
 * ============================================================================
 * DOCUMENTATION & EXAMPLES
 * ============================================================================
 */

/**
 * EXAMPLE REQUEST/RESPONSE
 * Example VAT validation for Acme Corp B.V. (NL123456789B01)
 *
 * Request Headers:
 * POST /taxation_customs/vies/services/checkVatService HTTP/1.1
 * Host: ec.europa.eu
 * Connection: Keep-Alive
 * User-Agent: PHP-SOAP/5.5.14
 * Content-Type: text/xml; charset=utf-8
 * SOAPAction: ""
 * Content-Length: 342
 *
 * Request Body:
 * <?xml version="1.0" encoding="UTF-8"?>
 * <SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/"
 *                    xmlns:ns1="urn:ec.europa.eu:taxud:vies:services:checkVat:types">
 *   <SOAP-ENV:Body>
 *     <ns1:checkVat>
 *       <ns1:countryCode>NL</ns1:countryCode>
 *       <ns1:vatNumber>123456789B01</ns1:vatNumber>
 *     </ns1:checkVat>
 *   </SOAP-ENV:Body>
 * </SOAP-ENV:Envelope>
 *
 * Response Headers:
 * HTTP/1.1 200 OK
 * Date: Mon, 09 Mar 2015 11:45:56 GMT
 * Transfer-Encoding: chunked
 * Content-Type: text/xml; charset=UTF-8
 * Server: Europa
 * Connection: Keep-Alive
 *
 * Response Body (Success):
 * <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 *   <soap:Body>
 *     <checkVatResponse xmlns="urn:ec.europa.eu:taxud:vies:services:checkVat:types">
 *       <countryCode>NL</countryCode>
 *       <vatNumber>123456789B01</vatNumber>
 *       <requestDate>2015-03-09+01:00</requestDate>
 *       <valid>true</valid>
 *       <name>Acme Corp B.V.</name>
 *       <address>
 *         KERKSTRAAT 001234
 *         1234AB AMSTERDAM
 *       </address>
 *     </checkVatResponse>
 *   </soap:Body>
 * </soap:Envelope>
 *
 * Parsed Response Array:
 * [
 *   "countryCode"  => "NL",
 *   "vatNumber"    => "123456789B01",
 *   "requestDate"  => "2015-03-09+01:00",
 *   "valid"        => true,
 *   "name"         => "Acme Corp B.V.",
 *   "address"      => "KERKSTRAAT 001234\n1234AB AMSTERDAM"
 * ]
 */

/**
 * ERROR RESPONSE EXAMPLE
 * When invalid input is provided or validation fails
 *
 * <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 *   <soap:Body>
 *     <soap:Fault>
 *       <faultcode>soap:Server</faultcode>
 *       <faultstring>INVALID_INPUT</faultstring>
 *     </soap:Fault>
 *   </soap:Body>
 * </soap:Envelope>
 */

/**
 * ALTERNATIVE IMPLEMENTATION APPROACHES
 * ============================================================================
 */

/**
 * Approach 1: JSON to XML conversion (deprecated - Xml.parseJS no longer available)
 *
 * var json = {
 *     'SOAP-ENV:Envelope': {
 *         'SOAP-ENV:Body': {
 *             'ns1:checkVat': {
 *                 'ns1:countryCode': countryCode,
 *                 'ns1:vatNumber': vatNumber
 *             }
 *         }
 *     }
 * };
 * var payload = Xml.parseJS(json).toXmlString();
 */

/**
 * Approach 2: XmlService element creation
 * Reference: https://developers.google.com/apps-script/reference/xml-service/
 *
 * var envelope = XmlService.createElement('SOAP-ENV:Envelope')
 *     .setAttribute('xmlns:SOAP-ENV', 'http://schemas.xmlsoap.org/soap/envelope/')
 *     .setAttribute('xmlns:ns1', 'urn:ec.europa.eu:taxud:vies:services:checkVat:types');
 * var body = XmlService.createElement('SOAP-ENV:Body');
 * var checkVat = XmlService.createElement('ns1:checkVat');
 * var child1 = XmlService.createElement('ns1:countryCode').setText(countryCode);
 * var child2 = XmlService.createElement('ns1:vatNumber').setText(vatNumber);
 * checkVat.addContent(child1);
 * checkVat.addContent(child2);
 * body.addContent(checkVat);
 * envelope.addContent(body);
 * var document = XmlService.createDocument(envelope);
 * var xml = XmlService.getPrettyFormat().format(document);
 * Logger.log(xml);
 */


