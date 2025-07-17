 Office.onReady(() => {
      const rs = Office.context.roamingSettings;
      const apiKey = rs.get("apiKey") || "";
      const agreementId = rs.get("agreementId") || ""; 
      
    // Check if config is missing
    if (!apiKey || !agreementId) {
      const loader = document.getElementById("loader-overlay");
      const content = document.getElementById("content");
      const buttons = document.getElementById("buttons");

      if (loader) loader.style.display = "none";
      if (buttons) buttons.style.display = "none";
      if (content) {
        content.style.display = "block";
        content.innerHTML = `
          <div style="text-align: center; padding: 40px;">
            <h3 style="color: #d32f2f;">⚠️ Please set up configuration first.</h3>
            <p style="color: #666;">API Key or Agreement ID is missing.</p>
          </div>
        `;
      }

      return; // Stop execution
    }
    const item = Office.context.mailbox.item;
    const evidence = [];
    let payload = null;
    const pushIfValid = (key, value) => {
      if (value && value.trim() !== "") {
        evidence.push({ key, value });
      }
    };

     const parseEmail = (input, fallback) => {
      if (!input) return fallback || "";
        const match = /<([^>]+)>/.exec(input);
        if (match && match[1]) {
          return match[1].trim();
        }
        // Fallback if input is already an email (no brackets)
        return input.includes("@") ? input.trim() : (fallback || "");
      };

      const getRecipientEmails = (list) => {
        if (!Array.isArray(list)) return "";
        return list
          .map(r => parseEmail(r.displayName || r.emailAddress?.address || ""))
          .filter(email => !!email)
          .join(", ");
      };

      // From
      if (item.from) {
        console.log(item.from);
        const email = parseEmail(item.from.emailAddress || item.from.displayName|| "");
        pushIfValid("From", email);
      }

      // To, CC, BCC
      pushIfValid("To", getRecipientEmails(item.to));
      pushIfValid("CC", getRecipientEmails(item.cc));
      pushIfValid("BCC", getRecipientEmails(item.bcc));

      pushIfValid("Subject", item.subject || "");
      pushIfValid("DateTime", item.dateTimeCreated?.toISOString?.() || new Date().toISOString());

      item.body.getAsync("text", result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          pushIfValid("Body", result.value || "");
        }

        payload = {
          evidence,
          tags: ["Outlook", "Evidence"],
          header: {
            sourceSystemDispatchReference: (item.subject && item.subject.trim()) ? item.subject : "OutlookAddin",
            serviceAgreementIdentifier: agreementId,
            where: "Outlook",
            when: new Date().toISOString()
          }
        };
        document.getElementById("loader-overlay").style.display = "none";
        document.getElementById("content").style.display = "block";

        // Render evidence in styled blocks (not JSON)
        const contentDiv = document.getElementById("content");
        contentDiv.innerHTML = "";
        evidence.forEach(({ key, value }) => {
          const section = document.createElement("div");
          section.className = "section";

          const title = document.createElement("h4");
          title.textContent = key;

          const content = document.createElement("p");
          content.textContent = value;

          section.appendChild(title);
          section.appendChild(content);
          contentDiv.appendChild(section);
        });

        document.getElementById("buttons").style.display = "flex";
      }); 
  document.getElementById("downloadBtn").addEventListener("click", () => {
  if (!evidence || evidence.length === 0) return;

  let textContent = "\n";
  evidence.forEach(({ key, value }) => {
    textContent += `${key}:\n${value}\n\n`;
  });

  const blob = new Blob([textContent], { type: "text/plain" });
  const url = URL.createObjectURL(blob);
  const now = new Date();
  const timestamp = now.toISOString().replace(/[:.]/g, "-").slice(0, 16);
  const filename = `evidence_${timestamp}.txt`;

  const isSafariMac = /^((?!chrome|android).)*safari/i.test(navigator.userAgent) && navigator.platform.toUpperCase().indexOf('MAC') >= 0;

  if (isSafariMac) {
    // Use window.open for Safari on macOS (requires pop-ups enabled)
    const newTab = window.open(url);
    if (!newTab) {
      alert("Please allow pop-ups in Safari to enable the file download.");
    }
  } else {
    // Default download method for Windows/Chrome/Edge etc.
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  URL.revokeObjectURL(url);
});


 
      function dispatchEvidence(payload, apiKey) {
        console.log(payload);
      const dispatchUrl = "https://evidentproof-api.azurewebsites.net/api/v2/Evidence/Dispatch";

      return fetch(dispatchUrl, {
        method: "PUT",
        headers: {
          "Accept": "text/plain",
          "Authorization": apiKey,
          "Content-Type": "application/json-patch+json"
        },
        body: JSON.stringify(payload)
      })
        .then(response => {
          if (!response.ok) {
            throw new Error(`Dispatch failed with status ${response.status}`);
          }
          return response.text();
        });
      }
      
      const dispatchBtn = document.getElementById("dispatchBtn");
      const statusBox = document.getElementById("status");

      dispatchBtn.addEventListener("click", () => {
        if (!payload) return;

        // Disable the button to prevent double clicks
        dispatchBtn.disabled = true;
        dispatchBtn.textContent = "Dispatching...";

        dispatchEvidence(payload, apiKey)
          .then(responseText => {
            try {
              const response = JSON.parse(responseText);
              const receiptId = response?.id;

              if (receiptId) {
                statusBox.textContent = `✅ Evidence dispatched successfully. Receipt ID: ${receiptId}`;
                statusBox.className = "status success";
                statusBox.style.display = "block";
                dispatchBtn.style.display = "none"; // Hide button
              } else {
                throw new Error("Invalid response from server.");
              }
            } catch (e) {
              statusBox.textContent = "❌ Something went wrong. Please try again later.";
              statusBox.className = "status error"; 
              statusBox.style.display = "block";
                    dispatchBtn.disabled = false; 


				// Reset button inner HTML (restore icon and text)
				dispatchBtn.innerHTML = `
					<span style="display: inline-block; margin-right: 8px; transform: rotate(-90deg)">
					⤵
					</span>  
					Dispatch Evidence
				`;
            }
          })
          .catch(error => {
            console.error("Dispatch error:", error);
            statusBox.textContent = "❌ Something went wrong. Please try again later.";
            statusBox.className = "status error";
            
            statusBox.style.display = "block";
           dispatchBtn.disabled = false; 
			// Reset button inner HTML (restore icon and text)
			dispatchBtn.innerHTML = `
				<span style="display: inline-block; margin-right: 8px; transform: rotate(-90deg)">
				⤵
				</span>  
				Dispatch Evidence
			`;
          })
          .finally(() => { 
            // Keep the button disabled after click no matter the result
            dispatchBtn.disabled = false; 

          });
      });

    }); 
