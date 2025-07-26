(function () {
    "use strict";

    const siteId = "alaaliinternational.sharepoint.com,9ffde265-4f33-4f30-84b7-ea0fd5de1655,db811aa3-e2b7-4004-b2c4-8ae99904dca9";
    const listId = "86a3f828-4843-42cd-bc00-20f98b531d66";
    let accessToken;
    let currentItem;

    Office.onReady(() => {
        Office.context.mailbox.addHandler(Office.EventType.ItemSend, onMessageSend);
        loadDropdowns();
        updateSubjectField();
    });

    async function acquireToken() {
        try {
            accessToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
        } catch (error) {
            console.error("Error acquiring token:", error);
        }
    }

    async function loadDropdowns() {
        await acquireToken();

        try {
            const res = await axios.get(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields`, {
                headers: { Authorization: `Bearer ${accessToken}` }
            });

            const items = res.data.value;

            const projectDropdown = document.getElementById("projectDropdown");
            const categoryDropdown = document.getElementById("categoryDropdown");
            const subcategoryDropdown = document.getElementById("subcategoryDropdown");

            items.forEach(item => {
                const fields = item.fields;

                // Populate Project
                if (fields.ProjectName) {
                    const option = document.createElement("option");
                    option.value = fields.ProjectName;
                    option.text = fields.ProjectName;
                    projectDropdown.appendChild(option);
                }

                // Populate Category
                if (fields.Category && !Array.from(categoryDropdown.options).some(o => o.value === fields.Category)) {
                    const option = document.createElement("option");
                    option.value = fields.Category;
                    option.text = fields.Category;
                    categoryDropdown.appendChild(option);
                }

                // Populate Subcategory
                if (fields.Subcategory && !Array.from(subcategoryDropdown.options).some(o => o.value === fields.Subcategory)) {
                    const option = document.createElement("option");
                    option.value = fields.Subcategory;
                    option.text = fields.Subcategory;
                    subcategoryDropdown.appendChild(option);
                }
            });

            projectDropdown.addEventListener("change", updateSubjectField);
            categoryDropdown.addEventListener("change", updateSubjectField);
            subcategoryDropdown.addEventListener("change", updateSubjectField);
        } catch (error) {
            console.error("Error loading SharePoint data:", error);
        }
    }

    function updateSubjectField() {
        const project = document.getElementById("projectDropdown").value;
        const category = document.getElementById("categoryDropdown").value;
        const subcategory = document.getElementById("subcategoryDropdown").value;
        const subjectField = document.getElementById("subjectField");

        Office.context.mailbox.item.subject.getAsync(result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                let prefix = "";
                if (project) prefix += `[${project}]`;
                if (category) prefix += `[${category}]`;
                if (subcategory) prefix += `[${subcategory}]`;
                subjectField.value = prefix + (prefix ? " " : "") + result.value;
            }
        });
    }

    function onMessageSend(event) {
        currentItem = Office.context.mailbox.item;
        Office.context.ui.displayDialogAsync(
            window.location.origin + "/taskpane.html",
            { height: 50, width: 30, displayInIframe: true },
            result => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const dialog = result.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, args => {
                        const message = JSON.parse(args.message);
                        handleDialogMessage(message, event, dialog);
                    });
                } else {
                    event.completed({ allowEvent: true });
                }
            }
        );
        event.completed({ allowEvent: false });
    }

    async function handleDialogMessage(message, event, dialog) {
        if (message.action === "send") {
            if (!document.getElementById("projectDropdown").value) {
                document.getElementById("errorMessage").style.display = "block";
                return;
            }

            const subjectField = document.getElementById("subjectField").value;
            currentItem.subject.setAsync(subjectField, async result => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    await archiveEmail();
                    event.completed({ allowEvent: true });
                } else {
                    event.completed({ allowEvent: false });
                }
                dialog.close();
            });
        } else if (message.action === "reset") {
            resetForm();
        } else if (message.action === "ignore") {
            event.completed({ allowEvent: true });
            dialog.close();
        } else if (message.action === "dontSend") {
            event.completed({ allowEvent: false });
            dialog.close();
        }
    }

    function resetForm() {
        document.getElementById("projectDropdown").value = "";
        document.getElementById("categoryDropdown").value = "";
        document.getElementById("subcategoryDropdown").value = "";
        document.querySelectorAll('input[name="audience"]').forEach(radio => radio.checked = false);
        document.getElementById("errorMessage").style.display = "none";
        updateSubjectField();
    }

    async function archiveEmail() {
        try {
            const project = document.getElementById("projectDropdown").value;
            const subjectField = document.getElementById("subjectField").value;
            const emlContent = await getEmailAsEml();
            const folderPath = `sites/Projects/${project}/Emails`;

            await axios.put(
                `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${folderPath}/${new Date().toISOString().split("T")[0]}_${subjectField}.eml:/content`,
                emlContent,
                {
                    headers: {
                        Authorization: `Bearer ${accessToken}`,
                        "Content-Type": "message/rfc822"
                    }
                }
            );
        } catch (error) {
            console.error("Error archiving email:", error);
        }
    }

    async function getEmailAsEml() {
        // Placeholder for actual email export logic
        return "MIME-Version: 1.0\nContent-Type: text/plain\nSubject: Sample\n\nThis is a test email.";
    }

    // UI button event listeners
    document.getElementById("sendButton").addEventListener("click", () => {
        Office.context.ui.messageParent(JSON.stringify({ action: "send" }));
    });

    document.getElementById("resetButton").addEventListener("click", () => {
        resetForm();
        Office.context.ui.messageParent(JSON.stringify({ action: "reset" }));
    });

    document.getElementById("ignoreButton").addEventListener("click", () => {
        Office.context.ui.messageParent(JSON.stringify({ action: "ignore" }));
    });

    document.getElementById("dontSendButton").addEventListener("click", () => {
        Office.context.ui.messageParent(JSON.stringify({ action: "dontSend" }));
    });

})();
