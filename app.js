async function listarPastaContraCheque() {
  try {
    const tokenResponse = await msalInstance.acquireTokenSilent({
      scopes: ["Sites.Read.All"],
      account: msalInstance.getActiveAccount(),
    });

    const accessToken = tokenResponse.accessToken;

    // 1) Pega siteId do site Adm
    const siteResponse = await fetch(
      "https://graph.microsoft.com/v1.0/sites/gsilvainfo.sharepoint.com:/sites/Adm",
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!siteResponse.ok) throw new Error("Erro ao buscar site");
    const siteData = await siteResponse.json();
    const siteId = siteData.id;

    // 2) Pega o driveId da biblioteca Documentos
    const driveResponse = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!driveResponse.ok) throw new Error("Erro ao buscar drives");
    const driveData = await driveResponse.json();

    // Procura pela drive Documentos, que normalmente tem o nome "Documents" ou "Documentos"
    const documentDrive = driveData.value.find(d => d.name.toLowerCase() === "documentos" || d.name.toLowerCase() === "documents");
    if (!documentDrive) throw new Error("Drive Documentos não encontrada");
    const driveId = documentDrive.id;

    // 3) Lista arquivos e pastas dentro de ContraCheque
    const folderPath = "ContraCheque";
    const filesResponse = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${folderPath}:/children`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    if (!filesResponse.ok) {
      if (filesResponse.status === 404) {
        fileListDiv.textContent = "Pasta ContraCheque não encontrada.";
        statusDiv.textContent = "";
        return;
      }
      throw new Error("Erro ao buscar arquivos na pasta ContraCheque");
    }

    const filesData = await filesResponse.json();
    if (filesData.value.length === 0) {
      fileListDiv.textContent = "Nenhum arquivo ou pasta dentro de ContraCheque.";
      statusDiv.textContent = "";
      return;
    }

    fileListDiv.innerHTML = "";
    filesData.value.forEach(item => {
      const el = document.createElement("div");
      el.textContent = item.name + (item.folder ? " (Pasta)" : " (Arquivo)");
      el.style.margin = "4px 0";
      fileListDiv.appendChild(el);
    });

    statusDiv.textContent = "Listagem da pasta ContraCheque concluída.";

  } catch (err) {
    console.error(err);
    statusDiv.textContent = "Erro ao listar pasta ContraCheque: " + err.message;
  }
}
