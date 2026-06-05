// admin-cardapio.js — Cardápio do Admin Homy

const AdminCardapio = window.AdminCardapio = {

  _editando: null, // { itemId, dia, opcao, nome, detalhes }

  async load(semanaId) {
    await this._renderAtual(semanaId);
    this._bindBotoes(semanaId);
  },

  // ── Render da tabela atual ───────────────────────────────────
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

      const OPCAO_LABEL = { principal: "Principal", light: "Light", carne: "Carne", massa: "Massa", lanche: "Lanche" };
      const norm = v => AdminUtils.norm(v);

      wrap.innerHTML = `<div style="display:flex;flex-direction:column;gap:.65rem">` +
        AdminUtils.DIAS.map(dia => {
          const diaItems = items.filter(i => norm(SP.pick(i, "Dia")) === norm(dia));
          if (!diaItems.length) return "";
          return `<div class="dashboard-panel">
            <div class="dashboard-panel-title">${AdminUtils.DIA_LABEL[dia]}</div>
            <div style="display:flex;gap:.5rem;flex-wrap:wrap">
              ${diaItems.map(op => {
                const id    = AdminUtils.esc(op.id || "");
                const opcao = AdminUtils.esc(SP.pick(op, "Opcao") || "");
                const prato = AdminUtils.esc(SP.pick(op, "Nome_Prato") || "");
                const label = OPCAO_LABEL[AdminUtils.norm(opcao)] || opcao;
                return `<span class="badge badge-blue" style="display:inline-flex;align-items:center;gap:6px;padding:6px 10px">
                  ${label} — ${prato}
                  <button class="btn-icon" style="width:22px;height:22px;font-size:.72rem"
                    onclick="AdminCardapio.abrirEdicao('${id}','${AdminUtils.esc(dia)}','${opcao}')">✏️</button>
                  <button class="btn-icon danger" style="width:22px;height:22px;font-size:.72rem"
                    onclick="AdminCardapio.excluirItem('${id}','${AdminUtils.esc(semanaId)}')">🗑️</button>
                </span>`;
              }).join("")}
            </div>
          </div>`;
        }).join("") + `</div>`;
    } catch (e) {
      wrap.innerHTML = `<div class="alert alert-warning">Erro: ${AdminUtils.esc(e.message)}</div>`;
    }
  },

  // ── Modal de cadastro/edição ────────────────────────────────
  abrirNovo() {
    this._editando = null;
    this._preencherModal({ dia: "segunda", opcao: "principal", nome: "", detalhes: "" }, "Cadastrar item");
    AdminUtils.openModal("modalCardapioManualV20");
  },

  async abrirEdicao(itemId, dia, opcao) {
    try {
      await SP.init();
      const semanaId = AdminState.getSemanaId();
      const items    = await SP.getCardapio(semanaId);
      const item     = items.find(i => String(i.id) === String(itemId));
      if (!item) { AdminUtils.toast("Item não encontrado.", "error"); return; }

      this._editando = {
        itemId,
        dia:     SP.pick(item, "Dia")       || dia,
        opcao:   SP.pick(item, "Opcao")     || opcao,
        nome:    SP.pick(item, "Nome_Prato")|| "",
        detalhes: SP.pick(item, "Detalhes") || ""
      };
      this._preencherModal(this._editando, "Editar item");
      AdminUtils.openModal("modalCardapioManualV20");
    } catch (e) {
      AdminUtils.toast("Erro ao abrir edição: " + e.message, "error");
    }
  },

  _preencherModal({ dia, opcao, nome, detalhes }, titulo) {
    const t = document.querySelector("#modalCardapioManualV20 .modal-title");
    if (t) t.textContent = titulo;
    AdminUtils.setVal("manualCardapioDiaV20",     dia);
    AdminUtils.setVal("manualCardapioOpcaoV20",   opcao);
    AdminUtils.setVal("manualCardapioNomeV20",    nome);
    AdminUtils.setVal("manualCardapioDetalhesV20", detalhes);
    setTimeout(() => document.getElementById("manualCardapioNomeV20")?.focus(), 80);
  },

  async salvar() {
    const semanaId = AdminState.getSemanaId();
    const dia      = AdminUtils.getVal("manualCardapioDiaV20");
    const opcao    = AdminUtils.getVal("manualCardapioOpcaoV20");
    const nome     = AdminUtils.getVal("manualCardapioNomeV20");
    const detalhes = AdminUtils.getVal("manualCardapioDetalhesV20");

    if (!nome) { AdminUtils.toast("Informe o nome do prato.", "error"); return; }

    try {
      await SP.init();
      // saveCardapio já verifica se existe e atualiza em vez de criar
      await SP.saveCardapio(semanaId, dia, opcao, nome, detalhes);
      AdminUtils.closeModal("modalCardapioManualV20");
      AdminUtils.toast("Cardápio salvo.", "success");
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

  // ── PDF Vascon ───────────────────────────────────────────────
  async processarPDF(file) {
    if (!file) return;
    AdminUtils.toast("Lendo PDF da Vascon...", "info");
    try {
      const texto = await this._extrairTextoPDF(file);
      const parsed = this._parseVascon(texto);
      const total  = Object.values(parsed).reduce((s, d) => s + Object.keys(d).length, 0);
      if (!total) { AdminUtils.toast("Não consegui extrair o PDF. Use Editar manual.", "error"); return; }

      const semanaId = AdminState.getSemanaId();
      await SP.init();
      for (const dia of AdminUtils.DIAS) {
        for (const opcao of ["principal", "light", "carne", "massa", "lanche"]) {
          const nome = (parsed[dia]?.[opcao] || "").trim();
          if (nome) await SP.saveCardapio(semanaId, dia, opcao, nome, "Extraído do PDF Vascon");
        }
      }
      AdminUtils.toast(`Cardápio salvo (${total} itens).`, "success");
      await this._renderAtual(semanaId);
    } catch (e) {
      AdminUtils.toast("Erro ao processar PDF: " + e.message, "error");
    }
  },

  async _extrairTextoPDF(file) {
    if (!window.pdfjsLib) {
      await new Promise((res, rej) => {
        const s = document.createElement("script");
        s.src = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js";
        s.onload = res; s.onerror = rej;
        document.head.appendChild(s);
      });
      pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
    }
    const buf  = await file.arrayBuffer();
    const pdf  = await pdfjsLib.getDocument({ data: buf }).promise;
    let texto  = "";
    for (let p = 1; p <= Math.min(pdf.numPages, 2); p++) {
      const page    = await pdf.getPage(p);
      const content = await page.getTextContent();
      texto += "\n" + content.items.map(i => i.str).join(" ");
    }
    return texto;
  },

  _parseVascon(texto) {
    const raw = String(texto || "").toUpperCase().replace(/\r/g, "\n");
    const result = { segunda: {}, terca: {}, quarta: {}, quinta: {}, sexta: {} };
    const aliases = { SEGUNDA: "segunda", "TERÇA": "terca", TERCA: "terca", QUARTA: "quarta", QUINTA: "quinta", SEXTA: "sexta" };
    const dias = [["SEGUNDA","segunda"],["TERÇA","terca"],["TERCA","terca"],["QUARTA","quarta"],["QUINTA","quinta"],["SEXTA","sexta"]];

    for (let i = 0; i < dias.length; i++) {
      const [pat, key] = dias[i];
      const nextPats   = dias.slice(i + 1).map(d => d[0]).join("|") || "OBS";
      const re  = new RegExp(`${pat}[\\s\\S]*?(?=${nextPats}|OBS|$)`, "i");
      const m   = raw.match(re);
      if (!m) continue;

      const lines = m[0].split(/\n|\s{2,}/)
        .map(x => x.replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim())
        .filter(x => x && x.length > 2 && !/VASCON|DATA|OPCAO|SERV|CARD|OBS|SEGUNDA|TERÇA|TERCA|QUARTA|QUINTA|SEXTA|SOBREMESA/.test(x));

      const has = rx => lines.find(x => rx.test(x));
      const carne  = has(/LINGUI|STEAK|FILE DE PEIXE|COXA DE FRANGO|BIFE|TOSCANA/);
      const massa  = has(/ENROLADINHO|GRAVATINHA|RAVIOLI|MACARR/);
      const lanche = has(/X[- ]?SALADA|HAMBURGUER|LANCHE/);
      const light  = lines.filter(x => /OMELETE|MANDIOCA|CHICORIA|SALADA DA PISTA|ALMONDEGA|COUVE|ABOBRINHA|LIGHT/.test(x));
      const used   = new Set([carne, massa, lanche, ...light].filter(Boolean));
      const principal = lines.filter(x => !used.has(x)).slice(0, 6).join("; ");

      if (principal) result[key].principal = principal;
      if (light.length) result[key].light  = light.join("; ");
      if (carne)  result[key].carne  = carne;
      if (massa)  result[key].massa  = massa;
      if (lanche) result[key].lanche = lanche;
    }
    return result;
  },

  // ── Bindings ─────────────────────────────────────────────────
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
      btnPDF.addEventListener("click", () => { inputPDF.value = ""; inputPDF.click(); });
    }
    if (inputPDF && !inputPDF.dataset.bound) {
      inputPDF.dataset.bound = "1";
      inputPDF.addEventListener("change", () => this.processarPDF(inputPDF.files[0]));
    }

    // Botão Salvar do modal
    const btnSalvar = document.getElementById("btnSalvarCardapioModal");
    if (btnSalvar && !btnSalvar.dataset.bound) {
      btnSalvar.dataset.bound = "1";
      btnSalvar.addEventListener("click", () => this.salvar());
    }

    // Cancelar — só fecha, não limpa, não salva
    const btnCancelar = document.getElementById("btnCancelarCardapioModal");
    if (btnCancelar && !btnCancelar.dataset.bound) {
      btnCancelar.dataset.bound = "1";
      btnCancelar.addEventListener("click", () => AdminUtils.closeModal("modalCardapioManualV20"));
    }
  }
};
