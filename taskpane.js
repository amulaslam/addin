(function () {
    "use strict";

    Office.onReady(() => {
        Office.context.mailbox.addHandler(Office.EventType.ItemSend, onMessageSend);
        loadDropdowns();
        updateSubjectField();
    });

    let currentItem;
    let accessToken;

    async function loadDropdowns() {
        try {
            await acquireToken();
            const projectDropdown = document.getElementById("projectDropdown");
            const categoryDropdown = document.getElementById("categoryDropdown");
            const subcategoryDropdown = document.getElementById("subcategoryDropdown");

            // Fetch SharePoint list data (replace with your site and list IDs)
            const response = await axios.get("https://graph.microsoft.com/v1.0/sites/your-site-id/lists/your-list-id/items", {
                headers: { Authorization: `Bearer ${accessToken}` }
            });

            response.data.value.forEach(item => {
                const option = document.createElement("option");
                option.value = item.fields.ProjectID;
                option.text = item.fields.ProjectName;
                projectDropdown.appendChild(option);
            });

            projectDropdown.addEventListener("change", updateSubjectField);
            categoryDropdown.addEventListener("change", updateSubjectField);
            subcategoryDropdown.addEventListener("change", updateSubjectField);
        } catch (error) {
            console.error("Error loading dropdowns:", error);
        }
    }

    async function acquireToken() {
        try {
            accessToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
        } catch (error) {
            console.error("Error acquiring token:", error);
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
                    event.completed({ allowEvent: true });
                    await archiveEmail();
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
            const emlContent = await getEmailAsEml(); // Placeholder
            const folderPath = `sites/Projects/${project}/Emails`;
            await axios.post(
                `https://graph.microsoft.com/v1.0/sites/your-site-id/drive/root:/${folderPath}/${new Date().toISOString().split("T")[0]}_${subjectField}.eml:/content`,
                emlContent,
                { headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "message/rfc822" } }
            );
        } catch (error) {
            console.error("Error archiving email:", error);
        }
    }

    async function getEmailAsEml() {
        // Placeholder: Replace with actual Graph API call to get MIME content
        return "MIME-Version: 1.0\nContent-Type: text/plain\nSubject: Example\n\nSample email content";
    }

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
