function log(msg) {
  document.getElementById("log").textContent = msg;
}

function apiBase() {
  return document.getElementById("api").value.trim();
}

async function carregarMenu() {
  try {
    log("Carregando menu...");
    const r = await fetch(`${apiBase()}/menu`);
    const j = await r.json();

    const div = document.getElementById("botoes");
    div.innerHTML = "";

    j.opcoes.forEach((nome, i) => {
      const b = document.createElement("button");
      b.textContent = `▶ ${nome}`;
      b.onclick = () => executar(i);
      div.appendChild(b);
    });

    log("Menu carregado. Clique em uma opção.");
  } catch (e) {
    log("Erro carregando menu: " + e);
  }
}

async function executar(indice) {
  try {
    log(`Executando ação ${indice}...`);
    const r = await fetch(`${apiBase()}/menu/acao/${indice}`, {
      method: "POST"
    });
    const j = await r.json();
    log(JSON.stringify(j, null, 2));
  } catch (e) {
    log("Erro executando: " + e);
  }
}

carregarMenu();
