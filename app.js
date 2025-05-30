const msalConfig = {
  auth: {
    clientId: "ce16da31-dc22-483f-a139-6f7b966049c9",
    authority: "https://login.microsoftonline.com/62345b7a-94ed-4671-b8f2-624e28c8253a",
    redirectUri: window.location.origin
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const loginBtn = document.getElementById("btnLogin");
const statusDiv = document.getElementById("status");
const fileListDiv = document.getElementById("fileList");

loginBtn.addEventListener("click", async () => {
  statusDiv.textContent = "";
  fileListDiv.innerHTML = "";
  try {
    await msalInstance.loginPopup({
      scopes: ["User.Read", "Sites.Read.All"],
      prompt: "select_account"
    });

    const account = msalInstance.getActiveAccount();
    if (!account) throw new Error("Nenhuma conta ativa encontrada.");

    statusDiv.textContent = `Logado como ${account.username}`;
    loginBtn.style.display = "none";

    listarContraCheques(account.username);
  } catch (err) {
    console.error(err);
    statusDiv.textContent = "Erro no login: " + err.message;
  }
});

async function listarContraCheques(email) {
  statusDiv.textContent = "Carregando contra-cheques...";
  try {
    const tokenResponse = await msalInstance.acquireTokenSilent({
      scopes: ["Sites.Read.All"],
      account: msalInstance.getActiveAccount(),
    });

    const accessToken = tokenResponse.accessToken;

    // Pega siteId do site Adm
    const siteResponse = await fetch(
      "https://graph.microsoft.com/v1.0/sites/gsilvainfo.sharepoint.com:/sites/Adm",
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!siteResponse.ok) throw new Error("Erro ao buscar site");

    const siteData = await siteResponse.json();
    const siteId = siteData.id;

    // Pega o driveId (Documentos)
    const driveResponse = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!driveResponse.ok) throw new Error("Erro ao buscar drives");

    const driveData = await driveResponse.json();
    const driveId = driveData.value[0].id;

    // Pega arquivos na pasta pessoal (ContraCheque/{email})
    const folderPath = `ContraCheque/${email}`;
    const filesResponse = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodeURIComponent(folderPath)}:/children`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!filesResponse.ok) {
      if (filesResponse.status === 404) {
        fileListDiv.textContent = "Nenhum arquivo encontrado na sua pasta.";
        statusDiv.textContent = "";
        return;
      }
      throw new Error("Erro ao buscar arquivos");
    }
    const filesData = await filesResponse.json();

    if (filesData.value.length === 0) {
      fileListDiv.textContent = "Nenhum arquivo encontrado na sua pasta.";
      statusDiv.textContent = "";
      return;
    }

    // Monta lista de links para download
    fileListDiv.innerHTML = "";
    filesData.value.forEach(file => {
      if (!file.file) return; // ignora pastas

      const a = document.createElement("a");
      a.href = file["@microsoft.graph.downloadUrl"];
      a.textContent = file.name;
      a.target = "_blank";
      a.rel = "noopener noreferrer";
      fileListDiv.appendChild(a);
    });

    statusDiv.textContent = "Clique nos arquivos para baixar.";
  } catch (err) {
    console.error(err);
    statusDiv.textContent = "Erro ao carregar arquivos: " + err.message;
  }
}
