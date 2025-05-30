const msalConfig = {
  auth: {
    clientId: "SEU_CLIENT_ID_AQUI",
    authority: "https://login.microsoftonline.com/SEU_TENANT_ID_AQUI",
    redirectUri: window.location.origin,
  },
};
const loginScopes = ["User.Read"];
const graphScopes = ["Sites.Read.All", "Sites.ReadWrite.All"]; // ajustar conforme necessidade

const msalInstance = new msal.PublicClientApplication(msalConfig);

const loginBtn = document.getElementById("btnLogin");
const userInfoDiv = document.getElementById("userInfo");
const arquivosContainer = document.getElementById("arquivosContainer");
const statusDiv = document.getElementById("status");

let currentAccount = null;

loginBtn.addEventListener("click", async () => {
  try {
    const loginResponse = await msalInstance.loginPopup({ scopes: loginScopes });
    currentAccount = loginResponse.account;
    loginBtn.style.display = "none";
    userInfoDiv.textContent = `Olá, ${currentAccount.name}`;
    carregarArquivos();
  } catch (err) {
    alert("Erro no login: " + err.message);
  }
});

async function carregarArquivos() {
  try {
    statusDiv.textContent = "Carregando contracheques...";
    const tokenResponse = await msalInstance.acquireTokenSilent({ scopes: graphScopes, account: currentAccount })
      .catch(() => msalInstance.acquireTokenPopup({ scopes: graphScopes }));
    const accessToken = tokenResponse.accessToken;

    const siteId = await getSiteId(accessToken);
    const driveId = await getDriveId(siteId, accessToken);

    const arquivos = await getUserContraCheques(accessToken, siteId, driveId, currentAccount.username);
    statusDiv.textContent = "";

    if (arquivos.length === 0) {
      arquivosContainer.innerHTML = `<p>Nenhum contracheque encontrado.</p>`;
      return;
    }

    arquivosContainer.innerHTML = "";
    arquivos.forEach((arquivo) => {
      const item = document.createElement("div");
      item.className = "list-group-item d-flex justify-content-between align-items-center";

      const link = document.createElement("a");
      link.href = arquivo.webUrl;
      link.target = "_blank";
      link.textContent = arquivo.name;

      const btnGroup = document.createElement("div");

      const btnConcorda = document.createElement("button");
      btnConcorda.className = "btn btn-success btn-sm me-2";
      btnConcorda.textContent = "Concorda";
      btnConcorda.onclick = () => confirmarRecebimento(arquivo.name, "concorda");

      const btnDiscorda = document.createElement("button");
      btnDiscorda.className = "btn btn-danger btn-sm";
      btnDiscorda.textContent = "Discorda";
      btnDiscorda.onclick = () => confirmarRecebimento(arquivo.name, "discorda");

      btnGroup.appendChild(btnConcorda);
      btnGroup.appendChild(btnDiscorda);

      item.appendChild(link);
      item.appendChild(btnGroup);

      arquivosContainer.appendChild(item);
    });
  } catch (err) {
    statusDiv.textContent = "Erro: " + err.message;
  }
}

async function getSiteId(accessToken) {
  const url = "https://graph.microsoft.com/v1.0/sites/gsilvainfo.sharepoint.com:/sites/Adm";
  const resp = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
  if (!resp.ok) throw new Error("Erro ao buscar site: " + resp.statusText);
  const data = await resp.json();
  return data.id;
}

async function getDriveId(siteId, accessToken) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive`;
  const resp = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
  if (!resp.ok) throw new Error("Erro ao buscar drive: " + resp.statusText);
  const data = await resp.json();
  return data.id;
}

async function getUserContraCheques(accessToken, siteId, driveId, userEmail) {
  const folderPath = `ContraCheque/${encodeURIComponent(userEmail)}`;
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${folderPath}:/children`;
  const resp = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
  if (!resp.ok) throw new Error("Erro ao listar arquivos: " + resp.statusText);
  const data = await resp.json();
  return data.value;
}

async function confirmarRecebimento(nomeArquivo, status) {
  statusDiv.textContent = "Enviando confirmação...";
  try {
    const payload = {
      funcionario: currentAccount.username,
      nomeArquivo,
      status,
      data: new Date().toISOString(),
    };

    const response = await fetch("URL_DA_SUA_AZURE_FUNCTION", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (response.ok) {
      statusDiv.textContent = "Confirmação enviada com sucesso!";
    } else {
      statusDiv.textContent = "Erro ao enviar confirmação.";
    }
  } catch (err) {
    statusDiv.textContent = "Erro: " + err.message;
  }
}
