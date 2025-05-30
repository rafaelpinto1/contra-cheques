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

// ... seu código MSAL e login permanece igual ...

loginBtn.addEventListener("click", async () => {
  statusDiv.textContent = "";
  fileListDiv.innerHTML = "";
  try {
    const loginResponse = await msalInstance.loginPopup({
      scopes: ["User.Read", "Sites.Read.All"],
      prompt: "select_account"
    });

    msalInstance.setActiveAccount(loginResponse.account);

    const account = msalInstance.getActiveAccount();
    if (!account) throw new Error("Nenhuma conta ativa encontrada.");

    statusDiv.textContent = `Logado como ${account.username}`;
    loginBtn.style.display = "none";

    // Primeiro listar pastas dentro de ContraCheque
    await listarPastasContraCheque();

    // Depois listar arquivos do usuário
    // await listarContraCheques(account.username);
  } catch (err) {
    console.error(err);
    statusDiv.textContent = "Erro no login: " + err.message;
  }
});

async function listarPastasContraCheque() {
  statusDiv.textContent = "Listando pastas dentro de ContraCheque...";
  fileListDiv.innerHTML = "";

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

    // Pega o driveId da biblioteca Documentos
    const driveResponse = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!driveResponse.ok) throw new Error("Erro ao buscar drives");

    const driveData = await driveResponse.json();
    const driveId = driveData.value[0].id;

    // Listar pastas dentro de ContraCheque (não arquivo, só pastas)
    const folderPath = "ContraCheque";

    const foldersResponse = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodeURIComponent(folderPath)}:/children`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    if (!foldersResponse.ok) throw new Error("Erro ao buscar pastas");

    const foldersData = await foldersResponse.json();

    // Filtra só pastas
    const folders = foldersData.value.filter(item => item.folder);

    if (folders.length === 0) {
      fileListDiv.textContent = "Nenhuma pasta encontrada dentro de ContraCheque.";
      statusDiv.textContent = "";
      return;
    }

    folders.forEach(folder => {
      const div = document.createElement("div");
      div.textContent = folder.name;
      div.style.margin = "5px 0";
      fileListDiv.appendChild(div);
    });

    statusDiv.textContent = "Pastas listadas acima. Confira o nome exato para usar na busca dos arquivos.";
  } catch (err) {
    console.error(err);
    statusDiv.textContent = "Erro ao listar pastas: " + err.message;
  }
}
