Office.onReady(() => {
  const apiKeyInput = document.getElementById("apiKey");
  const agreementInput = document.getElementById("agreement");
  const actionBtn = document.getElementById("actionBtn");
  const form = document.getElementById("configForm");
  const statusBox = document.getElementById("status");

  const rs = Office.context.roamingSettings;
  let isEditMode = true;

  function showStatus(msg, type) {
    statusBox.textContent = msg;
    statusBox.className = "status " + type;
    statusBox.style.display = "block";
  }

  function updateButtonUI() { 
     if (isEditMode) {
    actionBtn.innerHTML = '<i class="fa fa-save"></i>  Save';
    actionBtn.classList.add("save-mode");
  } else {
    actionBtn.innerHTML = '<i class="fa fa-edit"></i>  Edit';
    actionBtn.classList.remove("save-mode");
  }
  }

  function loadSettings() {
    const key = rs.get("apiKey") || "";
    const agreement = rs.get("agreementId") || "";

    if (key || agreement) {
      apiKeyInput.value = key;
      agreementInput.value = agreement;
      apiKeyInput.readOnly = true;
      agreementInput.readOnly = true;
      isEditMode = false;
    } else {
      apiKeyInput.readOnly = false;
      agreementInput.readOnly = false;
      isEditMode = true;
    }

    updateButtonUI();
  }

  form.addEventListener("submit", (e) => {
    e.preventDefault();

    if (!isEditMode) {
      // Enter edit mode
      apiKeyInput.readOnly = false;
      agreementInput.readOnly = false;
      isEditMode = true;
      updateButtonUI();
      return;
    }

    const apiKey = apiKeyInput.value.trim();
    const agreement = agreementInput.value.trim();

    if (!apiKey || !agreement) {
      showStatus("❌ Please fill in all fields.", "error");
      return;
    }

    rs.set("apiKey", apiKey);
    rs.set("agreementId", agreement);
    rs.set("savedOn", new Date().toISOString());

    rs.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        apiKeyInput.readOnly = true;
        agreementInput.readOnly = true;
        isEditMode = false;
        updateButtonUI();
        showStatus("✅ Configuration saved successfully.", "success");
      } else {
        showStatus("❌ Failed to save settings.", "error");
      }
    });
  });

  loadSettings();
});
