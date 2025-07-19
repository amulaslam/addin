(function () {
    "use strict";

    Office.onReady(() => {
        loadDropdowns();
        updateSubjectField();
    });

    let currentItem;
    let accessToken;

    async function loadDropdowns() {
        try {
            // Placeholder: Fetch data from SharePoint (needs customization)
            const projectDropdown = document.getElementById("projectDropdown");
            // Add sample data for testing
            const sampleProjects = [
                { id: "P-123", name: "Project Alpha" },
                { id: "P-456", name: "Project Beta" }
            ];
            sampleProjects.forEach(project => {
                const option = document.createElement("option");
                option.value = project.id;
                option.text = project.name;
                projectDropdown.appendChild(option);
            });

            // Update subject when dropdowns change
            document.getElementById("projectDropdown").addEventListener("change", updateSubjectField);
            document.getElementById("categoryDropdown").addEventListener("change", updateSubjectField);
            document.getElementById("subcategoryDropdown").addEventListener("change", updateSubjectField);
        } catch (error) {
            console.error("Error loading dropdowns:", error);
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
                    // Placeholder: Archive email (needs customization)
                    console.log("Email would be archived to SharePoint");
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

    // Button event listeners
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

    Office.context.mailbox.addHandler(Office.EventType.ItemSend, onMessageSend);
})();