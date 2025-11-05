Office.onReady(() => {
  document.getElementById('saveBtn').onclick = saveAsEml;
});

function saveAsEml(event) {
  const item = Office.context.mailbox.item;
  if (!item) {
    alert("Cannot get the item.");
    return;
  }

  // Get the EWS request SOAP envelope with IncludeMimeContent = true
  const itemId = item.itemId;
  const ewsUrl = Office.context.mailbox.ewsUrl;

  const soapRequest =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
                   'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
                   'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
                   'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
      '<soap:Header>' +
        '<t:RequestServerVersion Version="Exchange2016"/>' +  // or appropriate version
      '</soap:Header>' +
      '<soap:Body>' +
        '<m:GetItem>' +
          '<m:ItemShape>' +
            '<t:IncludeMimeContent>true</t:IncludeMimeContent>' +
            '<t:BodyType>Best</t:BodyType>' +
            '<t:AdditionalProperties>' +
              '<t:FieldURI FieldURI="item:MimeContent"/>' +
            '</t:AdditionalProperties>' +
          '</m:ItemShape>' +
          '<m:ItemIds>' +
            '<t:ItemId Id="' + itemId + '"/>' +
          '</m:ItemIds>' +
        '</m:GetItem>' +
      '</soap:Body>' +
    '</soap:Envelope>';

  Office.context.mailbox.makeEwsRequestAsync(soapRequest, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const response = asyncResult.value;
      // Should parse out the MIME content (Base64) from response
      const mimeBase64 = parseMimeFromResponse(response);
      const mimeBlob = b64toBlob(mimeBase64, 'message/rfc822');
      
      let fileName = item.subject || "message";
      fileName = fileName.replace(/[\\\/:*?"<>|]+/g, "_") + ".eml";
      saveAs(mimeBlob, fileName);
    } else {
      console.error("EWS request failed:", asyncResult.error);
      alert("Could not get MIME content: " + asyncResult.error.message);
    }
  });
}

// Utility: convert base64 to Blob
function b64toBlob(b64Data, contentType) {
  const byteCharacters = atob(b64Data);
  const byteNumbers = new Array(byteCharacters.length);
  for (let i = 0; i < byteCharacters.length; i++) {
    byteNumbers[i] = byteCharacters.charCodeAt(i);
  }
  const byteArray = new Uint8Array(byteNumbers);
  return new Blob([byteArray], { type: contentType });
}

// Utility: parse the MIME content from the SOAP response
function parseMimeFromResponse(responseXml) {
  // Simple parser – assuming the <MimeContent> element contains Base64
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(responseXml, "text/xml");
  const mimeNode = xmlDoc.getElementsByTagName("t:MimeContent")[0];
  if (!mimeNode) {
    throw new Error("MimeContent element not found");
  }
  const base64 = mimeNode.textContent;
  return base64;
}
