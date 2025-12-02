(function () {
    const originalFetch = window.fetch;
    let cachedFileList = [];
    let observer = null;

    window.fetch = async function (...args) {
        const requestUrl = args[0];

        if (requestUrl.includes("/files")) {
            //console.log(`🚀 Hitchhiking on request: ${requestUrl}`);

            const response = await originalFetch.apply(this, args);
            const clonedResponse = response.clone();

            const contentType = clonedResponse.headers.get("Content-Type");
            if (contentType && contentType.includes("application/json")) {
                const jsonData = await clonedResponse.json();

                cachedFileList = jsonData;
                startObserver(); // Start observer when files are fetched
            }

            return response;
        }

        return originalFetch.apply(this, args);
    };

    function startObserver() {
        if (observer) return; // Prevent duplicate observers

        observer = new MutationObserver((mutations) => {
            const dropdown = document.querySelector("td.parameters-col_description select");

            if (dropdown) {
                if (!dropdown.dataset.updated || dropdown.children.length !== cachedFileList.length) {
                    console.log("🔄 Restoring cached file list...");
                    updateDropdown(dropdown, cachedFileList);
                    updateAvailableValues(cachedFileList);
                    dropdown.dataset.updated = "true"; // ✅ Mark dropdown as updated
                }
            }
        });

        observer.observe(document.body, { childList: true, subtree: true });
    }

    function updateDropdown(dropdown, fileList) {
        const options = fileList.map(file => {
            const option = document.createElement("option");
            option.value = file;
            option.textContent = file;
            return option;
        });

        dropdown.replaceChildren(...options); // ✅ More efficient replacement method

        console.log("✅ Dropdown successfully updated!");
    }
})();

function updateAvailableValues(fileList) {
    const availableValuesParagraph = document.querySelector(".parameter__enum p");

    if (!availableValuesParagraph) {
        console.error("⚠ Available values `<p>` not found.");
        return;
    }
    availableValuesParagraph.innerHTML = `<i>Available values</i>: ${fileList.join(", ")}`;
}
