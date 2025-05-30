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
    const loginResponse = await msalInstance.loginPopup({
      scopes: ["User.Read", "Sites.Read.All"],
      prompt: "select_account"
    });

    msalInstance.setActiveAccount(loginResponse.account);

    const account = msalInstance.getActiveAccount();
    if (!account) throw new Error("Nenhuma conta ativa encontrada.");

    statusDiv.textContent = `Logado como ${account.username}`;
    loginBtn.style.display = "none";

    await listarContraCheques(account.username);
  } catch (err) {
    console.error(err);
    statusDiv.textContent = "Erro no login: " + err.message;
  }
});

async function listarContraCheques(email) {
  statusDiv.textContent = "Carregando contra-cheques...";
  fileListDiv.innerHTML = "";

  try {
    const tokenResponse = await msalInstance.acquireTokenSilent({
      scopes: ["Sites.Read.All"],
      account: msalInstance.getActiveAccount(),
    });
    const accessToken = tokenResponse.accessToken;

    // 1. Pega o site Adm
    const siteResponse = await fetch(
      "https://graph.microsoft.com/v1.0/sites/gsilvainfo.sharepoint.com:/sites/Adm",
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!siteResponse.ok) throw new Error("Erro ao buscar site");
    const siteData = await siteResponse.json();
    const siteId = siteData.id;

    // 2. Pega drives do site
    const driveResponse = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!driveResponse.ok) throw new Error("Erro ao buscar drives");
    const driveData = await driveResponse.json();

    // 3. Busca o drive chamado 'Documentos'
    const documentDrive = driveData.value.find(d => d.name === "Documentos");
    if (!documentDrive) throw new Error("Drive 'Documentos' não encontrado");
    const driveId = documentDrive.id;

    // 4. Lista o conteúdo raiz do drive Documentos para diagnosticar
    const rootResponse = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!rootResponse.ok) throw new Error("Erro ao listar raiz do drive");
    const rootData = await rootResponse.json();
    console.log("Conteúdo raiz do drive Documentos:", rootData);

    // 5. Verifica se existe a pasta ContraCheque na raiz
    const contraChequeFolder = rootData.value.find(
      item => item.folder && item.name.toLowerCase() === "contraCheque".toLowerCase()
    );
    if (!contraChequeFolder) {
      fileListDiv.textContent = "Pasta 'ContraCheque' não encontrada no drive Documentos.";
      statusDiv.textContent = "";
      return;
    }

    // 6. Lista pastas dentro da pasta ContraCheque
    const contraChequeFolderResponse = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${contraChequeFolder.id}/children`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!contraChequeFolderResponse.ok) throw new Error("Erro ao buscar pastas dentro de ContraCheque");
    const contraChequeFolders = await contraChequeFolderResponse.json();

    // 7. Procura a pasta que corresponde ao email do usuário (case insensitive)
    const userFolder = contraChequeFolders.value.find(
      item => item.folder && item.name.toLowerCase() === email.toLowerCase()
    );

    if (!userFolder) {
      fileListDiv.textContent = "Nenhuma pasta encontrada para seu usuário.";
      statusDiv.textContent = "";
      return;
    }

    // 8. Lista arquivos dentro da pasta do usuário
    const userFilesResponse = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${userFolder.id}/children`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!userFilesResponse.ok) throw new Error("Erro ao buscar arquivos na pasta do usuário");
    const userFiles = await userFilesResponse.json();

    if (userFiles.value.length === 0) {
      fileListDiv.textContent = "Nenhum arquivo encontrado na sua pasta.";
      statusDiv.textContent = "";
      return;
    }

    // 9. Mostra links para download dos arquivos
    userFiles.value.forEach(file => {
      if (!file.file) return; // ignora pastas

      const a = document.createElement("a");
      a.href = file["@microsoft.graph.downloadUrl"];
      a.textContent = file.name;
      a.target = "_blank";
      a.rel = "noopener noreferrer";
      a.style.display = "block";
      a.style.margin = "8px 0";
      fileListDiv.appendChild(a);
    });

    statusDiv.textContent = "Clique nos arquivos para baixar.";
  } catch (err) {
    console.error(err);
    statusDiv.textContent = "Erro ao carregar arquivos: " + err.message;
  }
}
