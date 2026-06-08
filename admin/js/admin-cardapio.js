// admin-cardapio.js — Cardápio do Admin Homy
// Fluxo correto:
// 1. Subir PDF Vascon
// 2. Extrair dados para rascunho
// 3. Mostrar conferência editável
// 4. Salvar no SharePoint somente após confirmação

const AdminCardapio = window.AdminCardapio = {

  _editando: null,
  _rascunhoPDF: [],
  _modoRascunho: false,

  OPCOES: ["principal", "light", "carne", "massa", "lanche"],

  OPCAO_LABEL: {
    principal: "Principal",
    light: "Light",
    carne: "Carne",
    massa: "Massa",
    lanche: "Lanche"
  },

  async load(semanaId) {
    this._bindBotoes(semanaId);
    await this._renderAtual(semanaId);
  },

  async _renderAtual(semanaId) {
    const wrap = document.getElementById("cardapioAtual");
    if (!wrap) return;

    wrap.innerHTML = `<div class="alert alert-info">Carregando cardápio...</div>`;

    try {
      await SP.init();
      const items = await SP.getCardapio(semanaId);

      if (!items.length) {
        wrap.innerHTML = `<div class="alert alert-warning">Nenhum cardápio cadastrado para esta semana.</div>`;
        return;
      }

      this._modoRascunho = false;
      wrap.innerHTML = this._htmlCardapioAgrupado(items, false);

    } catch (e) {
      wrap.innerHTML = `<div class="alert alert-warning">Erro: ${AdminUtils.esc(e.message)}</div>`;
    }
  },

  _htmlCardapioAgrupado(items, rascunho = false) {
    const norm = v => AdminUtils.norm(v);

    return `
      <div style="display:flex;flex-direction:column;gap:.65rem">
        ${AdminUtils.DIAS.map(dia => {
          const diaItems = items.filter(i => norm(SP.pick(i, "Dia", "dia")) === norm(dia));
          if (!diaItems.length) return "";

          return `
            <div class="dashboard-panel">
              <div class="dashboard-panel-title">${AdminUtils.DIA_LABEL[dia]}</div>
              <div style="display:flex;gap:.5rem;flex-wrap:wrap">
                ${diaItems.map(op => this._htmlBadgeItem(op, dia, rascunho)).join("")}
              </div>
            </div>
          `;
        }).join("")}
      </div>
    `;
  },

  _htmlBadgeItem(op, dia, rascunho = false) {
    const id = AdminUtils.esc(op.id || op.tempId || "");
    const opcaoOriginal = SP.pick(op, "Opcao", "opcao") || "";
    const opcao = AdminUtils.esc(opcaoOriginal);
    const label = this.OPCAO_LABEL[AdminUtils.norm(opcaoOriginal)] || opcaoOriginal;
    const nome = AdminUtils.esc(SP.pick(op, "Nome_Prato", "nome") || "");
    const detalhes = AdminUtils.esc(SP.pick(op, "Detalhes", "detalhes") || "");

    const editAction = rascunho
      ? `AdminCardapio.abrirEdicaoRascunho('${id}')`
      : `AdminCardapio.abrirEdicao('${id}','${AdminUtils.esc(dia)}','${opcao}')`;

    const deleteAction = rascunho
      ? `AdminCardapio.excluirRascunho('${id}')`
      : `AdminCardapio.excluirItem('${id}','${AdminUtils.esc(AdminState.getSemanaId())}')`;

    return `
      <span class="badge badge-blue"
        title="${detalhes}"
        style="display:inline-flex;align-items:center;gap:6px;padding:6px 10px">
        ${AdminUtils.esc(label)} — ${nome}
        <button class="btn-icon" style="width:22px;height:22px;font-size:.72rem" onclick="${editAction}">✏️</button>
        <button class="btn-icon danger" style="width:22px;height:22px;font-size:.72rem" onclick="${deleteAction}">🗑️</button>
      </span>
    `;
  },

  abrirNovo() {
    this._editando = null;
    this._preencherModal(
      { dia: "segunda", opcao: "principal", nome: "", detalhes: "" },
      "Cadastrar item",
      "As alterações serão salvas no SharePoint ao clicar em Salvar."
    );
    AdminUtils.openModal("modalCardapioManualV20");
  },

  async abrirEdicao(itemId, dia, opcao) {
    try {
      await SP.init();

      const semanaId = AdminState.getSemanaId();
      const items = await SP.getCardapio(semanaId);
      const item = items.find(i => String(i.id) === String(itemId));

      if (!item) {
        AdminUtils.toast("Item não encontrado.", "error");
        return;
      }

      this._editando = {
        tipo: "sharepoint",
        itemId,
        dia: SP.pick(item, "Dia") || dia,
        opcao: SP.pick(item, "Opcao") || opcao,
        nome: SP.pick(item, "Nome_Prato") || "",
        detalhes: SP.pick(item, "Detalhes") || ""
      };

      this._preencherModal(
        this._editando,
        "Editar item",
        "As alterações serão salvas no SharePoint ao clicar em Salvar."
      );

      AdminUtils.openModal("modalCardapioManualV20");

    } catch (e) {
      AdminUtils.toast("Erro ao abrir edição: " + e.message, "error");
    }
  },

  abrirEdicaoRascunho(tempId) {
    const item = this._rascunhoPDF.find(i => String(i.tempId) === String(tempId));

    if (!item) {
      AdminUtils.toast("Item do rascunho não encontrado.", "error");
      return;
    }

    this._editando = {
      tipo: "rascunho",
      itemId: tempId,
      dia: item.Dia,
      opcao: item.Opcao,
      nome: item.Nome_Prato,
      detalhes: item.Detalhes || ""
    };

    this._preencherModal(
      this._editando,
      "Editar item do rascunho",
      "Este item ainda não está no SharePoint. Ele será salvo somente ao confirmar o cardápio."
    );

    AdminUtils.openModal("modalCardapioManualV20");
  },

  _preencherModal({ dia, opcao, nome, detalhes }, titulo, aviso) {
    const t = document.querySelector("#modalCardapioManualV20 .modal-title");
    if (t) t.textContent = titulo;

    AdminUtils.setVal("manualCardapioDiaV20", dia);
    AdminUtils.setVal("manualCardapioOpcaoV20", opcao);
    AdminUtils.setVal("manualCardapioNomeV20", nome);
    AdminUtils.setVal("manualCardapioDetalhesV20", detalhes);

    this._atualizarAvisoModal(aviso);

    setTimeout(() => document.getElementById("manualCardapioNomeV20")?.focus(), 80);
  },

  _atualizarAvisoModal(texto) {
    const modal = document.getElementById("modalCardapioManualV20");
    if (!modal) return;

    let aviso = modal.querySelector(".cardapio-modal-aviso");

    if (!aviso) {
      const body = modal.querySelector(".modal-body") || modal.querySelector(".modal-content") || modal;
      aviso = document.createElement("div");
      aviso.className = "alert alert-info cardapio-modal-aviso";
      aviso.style.marginBottom = "1rem";
      body.prepend(aviso);
    }

    aviso.textContent = texto || "As alterações serão salvas ao clicar em Salvar.";
  },

  async salvar() {
    const semanaId = AdminState.getSemanaId();
    const dia = AdminUtils.getVal("manualCardapioDiaV20");
    const opcao = AdminUtils.getVal("manualCardapioOpcaoV20");
    const nome = AdminUtils.getVal("manualCardapioNomeV20");
    const detalhes = AdminUtils.getVal("manualCardapioDetalhesV20");

    if (!nome) {
      AdminUtils.toast("Informe o nome do prato.", "error");
      return;
    }

    try {
      if (this._editando?.tipo === "rascunho") {
        const item = this._rascunhoPDF.find(i => String(i.tempId) === String(this._editando.itemId));

        if (!item) {
          AdminUtils.toast("Item do rascunho não encontrado.", "error");
          return;
        }

        item.Dia = dia;
        item.Opcao = opcao;
        item.Nome_Prato = nome;
        item.Detalhes = detalhes;

        AdminUtils.closeModal("modalCardapioManualV20");
        AdminUtils.toast("Rascunho atualizado.", "success");

        this._editando = null;
        this._renderConferenciaPDF();
        return;
      }

      await SP.init();
      await SP.saveCardapio(semanaId, dia, opcao, nome, detalhes);

      AdminUtils.closeModal("modalCardapioManualV20");
      AdminUtils.toast("Cardápio salvo no SharePoint.", "success");

      this._editando = null;
      await this._renderAtual(semanaId);

    } catch (e) {
      AdminUtils.toast("Erro ao salvar: " + e.message, "error");
    }
  },

  async excluirItem(itemId, semanaId) {
    if (!confirm("Excluir este item do cardápio?")) return;

    try {
      await SP.init();
      await SP.deleteItem("Cardapio", itemId);

      AdminUtils.toast("Item excluído.", "success");
      await this._renderAtual(semanaId || AdminState.getSemanaId());

    } catch (e) {
      AdminUtils.toast("Erro ao excluir: " + e.message, "error");
    }
  },

  excluirRascunho(tempId) {
    if (!confirm("Remover este item do rascunho?")) return;

    this._rascunhoPDF = this._rascunhoPDF.filter(i => String(i.tempId) !== String(tempId));
    this._renderConferenciaPDF();
    AdminUtils.toast("Item removido do rascunho.", "success");
  },

  async processarPDF(file) {
    if (!file) return;

    AdminUtils.toast("Lendo PDF da Vascon...", "info");

    try {
      const texto = await this._extrairTextoPDF(file);
      const parsed = this._parseVascon(texto);
      const rascunho = this._parsedParaLista(parsed);

      if (!rascunho.length) {
        AdminUtils.toast("Não consegui extrair o PDF. Use Editar manual.", "error");
        return;
      }

      this._rascunhoPDF = rascunho;
      this._modoRascunho = true;

      this._renderConferenciaPDF();

      AdminUtils.toast(`PDF lido. Confira ${rascunho.length} itens antes de salvar.`, "success");

    } catch (e) {
      AdminUtils.toast("Erro ao processar PDF: " + e.message, "error");
    }
  },

  _parsedParaLista(parsed) {
    const lista = [];

    for (const dia of AdminUtils.DIAS) {
      for (const opcao of this.OPCOES) {
        const valor = (parsed[dia]?.[opcao] || "").trim();
        if (!valor) continue;

        const normalizado = this._normalizarItemExtraido(valor, opcao);

        lista.push({
          tempId: `tmp-${dia}-${opcao}-${Date.now()}-${Math.random().toString(16).slice(2)}`,
          Dia: dia,
          Opcao: opcao,
          Nome_Prato: normalizado.nome,
          Detalhes: normalizado.detalhes
        });
      }
    }

    return lista;
  },

  _normalizarItemExtraido(valor, opcao) {
    const partes = String(valor || "")
      .split(";")
      .map(p => p.trim())
      .filter(Boolean);

    if (!partes.length) {
      return { nome: "", detalhes: "" };
    }

    if (opcao === "principal") {
      const detalhes = partes.join("\n");
      return {
        nome: "Prato Principal",
        detalhes
      };
    }

    const nome = partes[0];
    const detalhes = partes.slice(1).join("\n");

    return {
      nome,
      detalhes
    };
  },

  _renderConferenciaPDF() {
    const wrap = document.getElementById("cardapioAtual");
    if (!wrap) return;

    if (!this._rascunhoPDF.length) {
      wrap.innerHTML = `<div class="alert alert-warning">Nenhum item no rascunho do PDF.</div>`;
      return;
    }

    wrap.innerHTML = `
      <div class="dashboard-panel" style="border-color:rgba(255,200,80,.35);background:rgba(255,180,0,.05)">
        <div class="dashboard-panel-title">Conferência antes de salvar no SharePoint</div>
        <div style="font-size:.86rem;color:rgba(220,235,255,.82);line-height:1.45;margin-bottom:1rem">
          Confira os itens extraídos do PDF Vascon. Você pode editar ou remover qualquer item.
          Nada foi salvo no SharePoint ainda.
        </div>

        <div style="display:flex;gap:.6rem;flex-wrap:wrap;margin-bottom:1rem">
          <button class="btn btn-primary" onclick="AdminCardapio.confirmarSalvarPDF()">Salvar no SharePoint</button>
          <button class="btn" onclick="AdminCardapio.cancelarRascunhoPDF()">Cancelar rascunho</button>
        </div>

        ${this._htmlCardapioAgrupado(this._rascunhoPDF, true)}
      </div>
    `;
  },

  cancelarRascunhoPDF() {
    if (!confirm("Cancelar este rascunho do PDF? Nada será salvo no SharePoint.")) return;

    this._rascunhoPDF = [];
    this._modoRascunho = false;

    AdminUtils.toast("Rascunho cancelado.", "info");
    this._renderAtual(AdminState.getSemanaId());
  },

  async confirmarSalvarPDF() {
    if (!this._rascunhoPDF.length) {
      AdminUtils.toast("Não há itens para salvar.", "error");
      return;
    }

    const ok = confirm(
      "Confirmar gravação do cardápio no SharePoint?\n\n" +
      "A semana atual será substituída pelos itens conferidos."
    );

    if (!ok) return;

    const semanaId = AdminState.getSemanaId();

    try {
      AdminUtils.toast("Salvando cardápio no SharePoint...", "info");

      await SP.init();

      const existentes = await SP.getCardapio(semanaId);

      for (const item of existentes) {
        await SP.deleteItem("Cardapio", item.id);
      }

      for (const item of this._rascunhoPDF) {
        await SP.saveCardapio(
          semanaId,
          item.Dia,
          item.Opcao,
          item.Nome_Prato,
          item.Detalhes || ""
        );
      }

      this._rascunhoPDF = [];
      this._modoRascunho = false;

      AdminUtils.toast("Cardápio salvo no SharePoint com sucesso.", "success");
      await this._renderAtual(semanaId);

    } catch (e) {
      AdminUtils.toast("Erro ao salvar PDF no SharePoint: " + e.message, "error");
    }
  },

  async _extrairTextoPDF(file) {
    if (!window.pdfjsLib) {
      await new Promise((res, rej) => {
        const s = document.createElement("script");
        s.src = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js";
        s.onload = res;
        s.onerror = rej;
        document.head.appendChild(s);
      });

      pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
    }

    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

    let texto = "";

    for (let p = 1; p <= Math.min(pdf.numPages, 2); p++) {
      const page = await pdf.getPage(p);
      const content = await page.getTextContent();
      texto += "\n" + content.items.map(i => i.str).join(" ");
    }

    return texto;
  },

  _parseVascon(texto) {
    const raw = String(texto || "")
      .toUpperCase()
      .replace(/\r/g, "\n");

    const result = {
      segunda: {},
      terca: {},
      quarta: {},
      quinta: {},
      sexta: {}
    };

    const dias = [
      ["SEGUNDA", "segunda"],
      ["TERÇA", "terca"],
      ["TERCA", "terca"],
      ["QUARTA", "quarta"],
      ["QUINTA", "quinta"],
      ["SEXTA", "sexta"]
    ];

    for (let i = 0; i < dias.length; i++) {
      const [pat, key] = dias[i];
      const nextPats = dias.slice(i + 1).map(d => d[0]).join("|") || "OBS";
      const re = new RegExp(`${pat}[\\s\\S]*?(?=${nextPats}|OBS|$)`, "i");
      const m = raw.match(re);

      if (!m) continue;

      const lines = m[0]
        .split(/\n|\s{2,}/)
        .map(x => x.replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim())
        .filter(x =>
          x &&
          x.length > 2 &&
          !/VASCON|DATA|OPCAO|OPÇÃO|SERV|CARD|OBS|SEGUNDA|TERÇA|TERCA|QUARTA|QUINTA|SEXTA|SOBREMESA/.test(x)
        );

      const has = rx => lines.find(x => rx.test(x));

      const carne = has(/LINGUI|STEAK|FILE DE PEIXE|FILÉ DE PEIXE|COXA DE FRANGO|BIFE|TOSCANA/);
      const massa = has(/ENROLADINHO|GRAVATINHA|RAVIOLI|MACARR|MASSA/);
      const lanche = has(/X[- ]?SALADA|HAMBURGUER|HAMBÚRGUER|LANCHE/);
      const light = lines.filter(x =>
        /OMELETE|MANDIOCA|CHICORIA|CHICÓRIA|SALADA DA PISTA|ALMONDEGA|ALMÔNDEGA|COUVE|ABOBRINHA|LIGHT/.test(x)
      );

      const used = new Set([carne, massa, lanche, ...light].filter(Boolean));
      const principal = lines.filter(x => !used.has(x)).slice(0, 10).join("; ");

      if (principal) result[key].principal = principal;
      if (light.length) result[key].light = light.join("; ");
      if (carne) result[key].carne = carne;
      if (massa) result[key].massa = massa;
      if (lanche) result[key].lanche = lanche;
    }

    return result;
  },

  _bindBotoes(semanaId) {
    const btnManual = document.getElementById("btnCardapioManual");

    if (btnManual && !btnManual.dataset.bound) {
      btnManual.dataset.bound = "1";
      btnManual.addEventListener("click", () => this.abrirNovo());
    }

    const btnPDF = document.getElementById("btnUploadPDF");
    const inputPDF = document.getElementById("pdfInput");

    if (btnPDF && !btnPDF.dataset.bound) {
      btnPDF.dataset.bound = "1";
      btnPDF.addEventListener("click", () => {
        if (inputPDF) {
          inputPDF.value = "";
          inputPDF.click();
        }
      });
    }

    if (inputPDF && !inputPDF.dataset.bound) {
      inputPDF.dataset.bound = "1";
      inputPDF.addEventListener("change", () => this.processarPDF(inputPDF.files[0]));
    }

    const btnSalvar = document.getElementById("btnSalvarCardapioModal");

    if (btnSalvar && !btnSalvar.dataset.bound) {
      btnSalvar.dataset.bound = "1";
      btnSalvar.addEventListener("click", () => this.salvar());
    }

    const btnCancelar = document.getElementById("btnCancelarCardapioModal");

    if (btnCancelar && !btnCancelar.dataset.bound) {
      btnCancelar.dataset.bound = "1";
      btnCancelar.addEventListener("click", () => AdminUtils.closeModal("modalCardapioManualV20"));
    }
  }
};
