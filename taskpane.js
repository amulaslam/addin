Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    loadDropdownData();
  }
});

async function loadDropdownData() {
  try {
    const accessToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
    const siteId = "alaaliinternational.sharepoint.com,9ffde265-4f33-4f30-84b7-ea0fd5de1655,db811aa3-e2b7-4004-b2c4-8ae99904dca9";
    const listId = "86a3f828-4843-42cd-bc00-20f98b531d66";

    const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields`, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    if (!response.ok) {
      throw new Error("Graph API request failed");
    }

    const data = await response.json();
    const items = data.value;

    // Clear dropdowns
    clearDropdown("emailTopic");
    clearDropdown("emailSubtopic");
    clearDropdown("emailClassification");

    items.forEach(item => {
      const fields = item.fields;
      if (fields.EmailTopic) {
        addOption("emailTopic", fields.EmailTopic);
      }
      if (fields.EmailSubtopic) {
        addOption("emailSubtopic", fields.EmailSubtopic);
      }
      if (fields.EmailClassification) {
        addOption("emailClassification", fields.EmailClassification);
      }
    });
  } catch (error) {
    console.error("Error fetching SharePoint list data:", error);
    showMessage("Failed to load dropdown data. Please check permissions or list setup.");
  }
}

function clearDropdown(id) {
  const dropdown = document.getElementById(id);
  dropdown.innerHTML = "<option value=''>--Select--</option>";
}

function addOption(id, text) {
  const dropdown = document.getElementById(id);
  const option = document.createElement("option");
  option.value = text;
  option.text = text;
  dropdown.appendChild(option);
}

function showMessage(message) {
  const msg = document.getElementById("statusMessage");
  if (msg) {
    msg.textContent = message;
    msg.style.display = "block";
  } else {
    alert(message);
  }
}
