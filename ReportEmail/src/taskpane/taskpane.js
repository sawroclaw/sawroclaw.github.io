/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
/* Minimalny taskpane dla zgłaszania phishing — zbiera checkboxy, dołącza oryginalną wiadomość i wysyła EWS CreateItem */
Office.onReady((info) => {
  try {
    if (info.host === Office.HostType.Outlook) {
      // Pokaż UI (jeśli istnieje)
      const appBody = document.getElementById("app-body");
      if (appBody) appBody.style.display = "block";

      // Wstaw temat aktywnej wiadomości (jeśli element istnieje)
      const item = Office.context.mailbox.item;
      const subjEl = document.getElementById("item-subject");
      if (subjEl) subjEl.textContent = (item && item.subject) ? item.subject : "(brak tematu)";

      // Podłącz przycisk zgłoszenia (jeśli element istnieje). Jeśli element nie istnieje teraz,
      // podłączenie zostanie spróbowane po DOMContentLoaded poniżej.
      const tryAttach = () => {
        const btn = document.getElementById("report-phish");
        if (btn) btn.onclick = sendPhishingReport;
      };
      tryAttach();
      // jeśli skrypt załadował się przed DOM, upewnij się, że dołączymy handler po pełnym załadowaniu
      if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", tryAttach);
      }
    }
  } catch (err) {
    // Zawsze loguj błędy do konsoli — ułatwia debugowanie "Script error"
    // (Outlook/Edge czasem maskuje szczegóły więc zapisujemy tu co możemy)
    console.error("Office.onReady handler error:", err);
  }
});

function sendPhishingReport() {
  const REPORT_ADDRESS = "rafal.sulkowski@wroclaw.sa.gov.pl"; // <- ustaw właściwy adres
  const item = Office.context.mailbox.item;
  const statusEl = document.getElementById("phish-status");

  if (!item || !item.itemId) {
    setStatus("Brak aktywnej wiadomości.", true);
    return;
  }

  const opts = [
    { id: "opt-answered", text: "Odpowiedziano na e-mail." },
    { id: "opt-downloaded", text: "Pobrano plik." },
    { id: "opt-opened", text: "Otworzono załącznik." },
    { id: "opt-visited", text: "Odwiedzono link." },
    { id: "opt-password", text: "Wprowadzono swoje hasło." },
    { id: "opt-forwarded", text: "Przesłano dalej ten e-mail." },
    { id: "opt-none", text: "Nie zrobiono nic z e-mailem." }
  ];

  const selected = opts.filter(o => document.getElementById(o.id)?.checked).map(o => o.text);
  const selectionHtml = selected.length ? `<ul>${selected.map(s => `<li>${escapeXml(s)}</li>`).join("")}</ul>` : "<p>Brak zaznaczonych opcji.</p>";
  const reportSubject = `Zgłoszenie phishing — ${item.subject || ""}`;
  const reportBody = `<p>Użytkownik zgłosił podejrzaną wiadomość. Wybrane opcje:</p>${selectionHtml}<hr/><p>Oryginalny temat: ${escapeXml(item.subject || "(brak)")}</p>`;

  // Spróbuj użyć itemId bez konwersji; jeśli w środowisku okaże się konieczny ConvertId, dorzucę konwersję SOAP.
  let ewsId = item.itemId;
  try {
    if (Office.context.mailbox.convertToEwsId) {
      ewsId = Office.context.mailbox.convertToEwsId(item.itemId);
    }
  } catch (e) {
    // fallback - użyj item.itemId
  }

  const req = `<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                   xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                   xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
      <soap:Header>
        <t:RequestServerVersion Version="Exchange2016" />
      </soap:Header>
      <soap:Body>
        <m:CreateItem MessageDisposition="SendAndSaveCopy">
          <m:SavedItemFolderId>
            <t:DistinguishedFolderId Id="sentitems" />
          </m:SavedItemFolderId>
          <m:Items>
            <t:Message>
              <t:Subject>${escapeXml(reportSubject)}</t:Subject>
              <t:Body BodyType="HTML">${escapeXml(reportBody)}</t:Body>
              <t:ToRecipients>
                <t:Mailbox>
                  <t:EmailAddress>${REPORT_ADDRESS}</t:EmailAddress>
                </t:Mailbox>
              </t:ToRecipients>
              <t:Attachments>
                <t:ItemAttachment>
                  <t:Name>SuspectedMessage.msg</t:Name>
                  <t:Item>
                    <t:Message>
                      <t:ItemId Id="${escapeXml(ewsId)}" />
                    </t:Message>
                  </t:Item>
                </t:ItemAttachment>
              </t:Attachments>
            </t:Message>
          </m:Items>
        </m:CreateItem>
      </soap:Body>
    </soap:Envelope>`;

  setStatus("Wysyłanie zgłoszenia...", false);
  Office.context.mailbox.makeEwsRequestAsync(req, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      setStatus("Zgłoszenie phishing wysłane.", false);
    } else {
      // pokaż więcej informacji, jeśli są dostępne
      const err = asyncResult.error || {};
      console.error("EWS error:", err);
      const msg = err.message || JSON.stringify(err);
      setStatus("Błąd wysyłania zgłoszenia: " + msg, true);
    }
  });

  function setStatus(msg, isError) {
    if (!statusEl) return;
    statusEl.textContent = msg;
    statusEl.style.color = isError ? "#b00020" : "#2b7a0b";
  }
}

function escapeXml(s) {
  return (!s) ? "" : s.replace(/[<>&'"]/g, c => ({ "<":"&lt;", ">":"&gt;", "&":"&amp;", "'":"&apos;", '"':"&quot;" }[c]));
}
