// admin-cardapio.js — Cardápio do Admin Homy
// Fluxo preservado:
// 1. Subir PDF Vascon
// 2. Extrair dados para rascunho
// 3. Mostrar conferência editável
// 4. Salvar no SharePoint somente após confirmação
//
// Correção desta versão:
// - Remove parser fixo da semana anterior
// - Lê o PDF atual por posição de tabela: dia + colunas Principal/Light/Carne/Massa/Lanche
// - Não cria Lanche em dias onde a coluna está vazia
// - Estrutura semanal dinâmica: Principal, Light, Carne, Massa e Lanche quando preenchidos
// - Se o PDF trouxer FERIADO no bloco do dia, salva o dia como FERIADO

const AdminCardapio = window.AdminCardapio = {
  _editando: null,
  _rascunhoPDF: [],
  _modoRascunho: false,

  OPCOES: ["feriado", "principal", "light", "carne", "massa", "lanche"],
  OPCAO_LABEL: {
    feriado: "Feriado",
    principal: "Principal",
    light: "Light",
    carne: "Carne",
    massa: "Massa",
    lanche: "Lanche"
  },
  OBS_VASCON: [
    "O cardápio acima está sujeito à alterações.",
    "As opções Carne e Massa substituem a carne do prato principal.",
    "As opções Light e Lanche substituem o prato principal completo."
  ].join("\n"),

  async load(semanaId) {
    this._bindBotoes();
    await this._renderAtual(semanaId);
  },

  async _renderAtual(semanaId) {
    const wrap = document.getElementById("cardapioAtual");
    if (!wrap) return;
    wrap.innerHTML = `<div class="alert alert-info">Carregando cardápio...</div>`;

    try {
      await SP.init();
      const items = await SP.getCardapio(semanaId);
      const limpos = this._deduplicarItens(items);

      if (!limpos.length) {
        wrap.innerHTML = `<div class="alert alert-warning">Nenhum cardápio cadastrado para esta semana.</div>`;
        return;
      }

      this._modoRascunho = false;
      wrap.innerHTML = this._htmlCardapioAgrupado(limpos, false);
    } catch (e) {
      console.error("[AdminCardapio] _renderAtual:", e);
      wrap.innerHTML = `<div class="alert alert-warning">Erro: ${AdminUtils.esc(e.message || e)}</div>`;
    }
  },

  _deduplicarItens(items) {
    const mapa = new Map();

    for (const item of items || []) {
      const dia = AdminUtils.norm(SP.pick(item, "Dia", "dia"));
      const opcao = AdminUtils.norm(SP.pick(item, "Opcao", "opcao"));
      if (!dia || !opcao) continue;

      const chave = `${dia}||${opcao}`;
      if (!mapa.has(chave)) {
        mapa.set(chave, item);
        continue;
      }

      const atual = mapa.get(chave);
      const atualDetalhes = String(SP.pick(atual, "Detalhes", "detalhes") || "");
      const novoDetalhes = String(SP.pick(item, "Detalhes", "detalhes") || "");
      if (novoDetalhes.length > atualDetalhes.length) mapa.set(chave, item);
    }

    return Array.from(mapa.values());
  },

  _htmlCardapioAgrupado(items, rascunho = false) {
    const norm = v => AdminUtils.norm(v);
    return `
      <div style="display:flex;flex-direction:column;gap:.65rem">
        ${AdminUtils.DIAS.map(dia => {
          const diaItems = (items || []).filter(i => norm(SP.pick(i, "Dia", "dia")) === norm(dia));
          if (!diaItems.length) return "";

          const feriado = diaItems.find(i =>
            norm(SP.pick(i, "Opcao", "opcao")) === "feriado" ||
            norm(SP.pick(i, "Nome_Prato", "nome", "Title")).includes("feriado")
          );

          if (feriado) {
            return `
              <div class="dashboard-panel" style="border-color:rgba(255,80,80,.35);background:rgba(192,40,28,.08)">
                <div class="dashboard-panel-title">${AdminUtils.esc(AdminUtils.DIA_LABEL[dia] || dia)}</div>
                <div style="border:1px solid rgba(255,80,80,.28);background:rgba(192,40,28,.12);border-radius:14px;padding:1.3rem;text-align:center">
                  <div style="font-family:'Barlow Condensed',sans-serif;font-size:2rem;font-weight:800;letter-spacing:.12em;color:#fff;text-transform:uppercase">FERIADO</div>
                  <div style="font-size:.78rem;color:rgba(255,200,200,.82);margin-top:.25rem">Dia sem marcação de refeição.</div>
                  <div style="margin-top:.7rem">${this._htmlBadgeItem(feriado, dia, rascunho)}</div>
                </div>
              </div>`;
          }

          return `
            <div class="dashboard-panel">
              <div class="dashboard-panel-title">${AdminUtils.esc(AdminUtils.DIA_LABEL[dia] || dia)}</div>
              <div style="display:flex;gap:.5rem;flex-wrap:wrap">
                ${this.OPCOES.filter(opcao => opcao !== "feriado").map(opcao => {
                  const item = diaItems.find(i => norm(SP.pick(i, "Opcao", "opcao")) === norm(opcao));
                  return item ? this._htmlBadgeItem(item, dia, rascunho) : "";
                }).join("")}
              </div>
            </div>`;
        }).join("")}
      </div>`;
  },

  _htmlBadgeItem(op, dia, rascunho = false) {
    const id = AdminUtils.esc(op.id || op.tempId || "");
    const opcaoOriginal = SP.pick(op, "Opcao", "opcao") || "";
    const opcaoNorm = AdminUtils.norm(opcaoOriginal);
    const opcao = AdminUtils.esc(opcaoOriginal);
    const label = this.OPCAO_LABEL[opcaoNorm] || opcaoOriginal;
    const nome = AdminUtils.esc(SP.pick(op, "Nome_Prato", "nome") || "");
    const editAction = rascunho
      ? `AdminCardapio.abrirEdicaoRascunho('${id}')`
      : `AdminCardapio.abrirEdicao('${id}','${AdminUtils.esc(dia)}','${opcao}')`;
    const deleteAction = rascunho
      ? `AdminCardapio.excluirRascunho('${id}')`
      : `AdminCardapio.excluirItem('${id}','${AdminUtils.esc(AdminState.getSemanaId())}')`;
    const badgeClass = opcaoNorm === "feriado" ? "badge-red" : "badge-blue";

    return `<span class="badge ${badgeClass}" style="display:inline-flex;align-items:center;gap:6px;padding:6px 10px">
      ${AdminUtils.esc(label)} — ${nome}
      <button class="btn-icon" style="width:22px;height:22px;font-size:.72rem" onclick="${editAction}">✏️</button>
      <button class="btn-icon danger" style="width:22px;height:22px;font-size:.72rem" onclick="${deleteAction}">🗑️</button>
    </span>`;
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
      console.error("[AdminCardapio] abrirEdicao:", e);
      AdminUtils.toast("Erro ao abrir edição: " + (e.message || e), "error");
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
    this._garantirOpcaoFeriadoModal();
    const t = document.querySelector("#modalCardapioManualV20 .modal-title");
    if (t) t.textContent = titulo;

    AdminUtils.setVal("manualCardapioDiaV20", dia);
    AdminUtils.setVal("manualCardapioOpcaoV20", opcao);
    AdminUtils.setVal("manualCardapioNomeV20", nome);
    AdminUtils.setVal("manualCardapioDetalhesV20", detalhes);
    this._atualizarAvisoModal(aviso);

    setTimeout(() => document.getElementById("manualCardapioNomeV20")?.focus(), 80);
  },


  _garantirOpcaoFeriadoModal() {
    const sel = document.getElementById("manualCardapioOpcaoV20");
    if (!sel || sel.querySelector('option[value="feriado"]')) return;
    const opt = document.createElement("option");
    opt.value = "feriado";
    opt.textContent = "Feriado";
    sel.appendChild(opt);
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

    if (!dia) { AdminUtils.toast("Informe o dia.", "error"); return; }
    if (!opcao) { AdminUtils.toast("Informe a opção.", "error"); return; }
    if (!nome) { AdminUtils.toast("Informe o nome do prato.", "error"); return; }

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
        this._rascunhoPDF = this._deduplicarItens(this._rascunhoPDF);

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
      console.error("[AdminCardapio] salvar:", e);
      AdminUtils.toast("Erro ao salvar: " + (e.message || e), "error");
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
      console.error("[AdminCardapio] excluirItem:", e);
      AdminUtils.toast("Erro ao excluir: " + (e.message || e), "error");
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
      const dadosPDF = await this._extrairDadosPDF(file);
      const parsed = this._parseVascon(dadosPDF);
      const rascunho = this._deduplicarItens(this._parsedParaLista(parsed));

      if (!rascunho.length) {
        AdminUtils.toast("Não consegui extrair o PDF. Use Editar manual.", "error");
        return;
      }

      this._rascunhoPDF = rascunho;
      this._modoRascunho = true;
      this._renderConferenciaPDF();
      AdminUtils.toast(`PDF lido. Confira ${rascunho.length} itens antes de salvar.`, "success");
    } catch (e) {
      console.error("[AdminCardapio] processarPDF:", e);
      AdminUtils.toast("Erro ao processar PDF: " + (e.message || e), "error");
    }
  },

  _parsedParaLista(parsed) {
    const lista = [];

    for (const dia of AdminUtils.DIAS) {
      const dadosDia = parsed[dia] || {};

      if (dadosDia.feriado) {
        lista.push({
          tempId: this._tempId(dia, "feriado"),
          Dia: dia,
          Opcao: "feriado",
          Nome_Prato: "FERIADO",
          Detalhes: "Não haverá marcação de refeição para este dia."
        });
        continue;
      }

      const principal = Array.isArray(dadosDia.principal) ? dadosDia.principal : [];

      if (principal.length) {
        lista.push({
          tempId: this._tempId(dia, "principal"),
          Dia: dia,
          Opcao: "principal",
          Nome_Prato: "Prato Principal",
          Detalhes: principal.join("\n")
        });
      }

      for (const opcao of ["light", "carne", "massa", "lanche"]) {
        const item = dadosDia[opcao];
        if (!item?.nome) continue;

        const detalhesBase = (item.detalhes || []).filter(Boolean);
        lista.push({
          tempId: this._tempId(dia, opcao),
          Dia: dia,
          Opcao: opcao,
          Nome_Prato: item.nome,
          Detalhes: [...detalhesBase, "", this.OBS_VASCON].join("\n").trim()
        });
      }
    }

    return lista;
  },

  _tempId(dia, opcao) {
    return `tmp-${dia}-${opcao}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  },

  _renderConferenciaPDF() {
    const wrap = document.getElementById("cardapioAtual");
    if (!wrap) return;

    if (!this._rascunhoPDF.length) {
      wrap.innerHTML = `<div class="alert alert-warning">Nenhum item no rascunho do PDF.</div>`;
      return;
    }

    wrap.innerHTML = `
      <div class="dashboard-panel" style="border-color:rgba(255,180,0,.35);background:rgba(255,180,0,.045)">
        <div class="dashboard-panel-title">Conferência antes de salvar no SharePoint</div>
        <div style="font-size:.88rem;color:rgba(200,220,255,.88);line-height:1.5;margin-bottom:1rem">
          Confira os itens extraídos do PDF Vascon. Você pode editar ou remover qualquer item.
          Nada foi salvo no SharePoint ainda.
        </div>
        <div style="display:flex;gap:.6rem;flex-wrap:wrap;margin-bottom:1rem">
          <button class="btn-primary" onclick="AdminCardapio.confirmarSalvarPDF()">Salvar no SharePoint</button>
          <button class="btn-secondary" onclick="AdminCardapio.cancelarRascunhoPDF()">Cancelar rascunho</button>
        </div>
        ${this._htmlCardapioAgrupado(this._rascunhoPDF, true)}
      </div>`;
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

      const todos = await SP.getItems("Cardapio");
      const existentes = (todos || []).filter(item =>
        String(SP.pick(item, "Semana_id", "SemanaId", "SemanaID", "Semanaid")) === String(semanaId)
      );

      for (const item of existentes) {
        await SP.deleteItem("Cardapio", item.id);
      }

      const rascunhoFinal = this._deduplicarItens(this._rascunhoPDF);
      for (const item of rascunhoFinal) {
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
      console.error("[AdminCardapio] confirmarSalvarPDF:", e);
      AdminUtils.toast("Erro ao salvar PDF no SharePoint: " + (e.message || e), "error");
    }
  },

  async _carregarPDFJS() {
    if (window.pdfjsLib) return;

    await new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js";
      s.onload = res;
      s.onerror = rej;
      document.head.appendChild(s);
    });

    pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
  },

  async _extrairDadosPDF(file) {
    await this._carregarPDFJS();

    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    const paginas = [];
    let texto = "";

    for (let p = 1; p <= Math.min(pdf.numPages, 2); p++) {
      const page = await pdf.getPage(p);
      const viewport = page.getViewport({ scale: 1 });
      const content = await page.getTextContent();
      const items = (content.items || [])
        .map(i => ({
          text: String(i.str || "").trim(),
          x: Number(i.transform?.[4] || 0),
          y: Number(i.transform?.[5] || 0),
          w: Number(i.width || 0),
          h: Number(i.height || 0)
        }))
        .filter(i => i.text);

      texto += "\n" + items.map(i => i.text).join(" ");
      paginas.push({ width: viewport.width, height: viewport.height, items });
    }

    return { texto, paginas };
  },

  _parseVascon(dadosPDF) {
    const result = this._parseVasconPorCoordenadas(dadosPDF);
    const total = Object.values(result).reduce((s, d) => s + Object.keys(d || {}).length, 0);
    if (total) return result;

    console.warn("[AdminCardapio] Parser por coordenadas não retornou itens. Tentando parser textual.");
    return this._parseVasconTexto(dadosPDF?.texto || dadosPDF || "");
  },

  _parseVasconPorCoordenadas(dadosPDF) {
    const result = { segunda: {}, terca: {}, quarta: {}, quinta: {}, sexta: {} };
    const pagina = (dadosPDF?.paginas || [])[0];
    if (!pagina?.items?.length) return result;

    const linhas = this._agruparLinhasPDF(pagina.items);
    const linhasComTexto = linhas.map(l => ({
      ...l,
      text: l.items.map(i => i.text).join(" ").toUpperCase()
    }));

    const diaInfo = [
      ["segunda", /\bSEGUNDA\b/],
      ["terca", /\bTER[ÇC]A\b/],
      ["quarta", /\bQUARTA\b/],
      ["quinta", /\bQUINTA\b/],
      ["sexta", /\bSEXTA\b/]
    ].map(([dia, rx]) => {
      const linha = linhasComTexto.find(l => rx.test(l.text));
      return linha ? { dia, y: linha.y } : null;
    }).filter(Boolean);

    if (diaInfo.length < 2) return result;

    const cresceParaBaixo = diaInfo[1].y > diaInfo[0].y;
    const obsLinha = linhasComTexto.find(l => /\bOBS\b/.test(l.text));
    const obsY = obsLinha?.y;
    const primeiraLinhaDia = linhasComTexto.find(l => /\bSEGUNDA\b/.test(l.text));
    const linhasCabecalho = primeiraLinhaDia
      ? linhasComTexto.filter(l => cresceParaBaixo ? l.y < primeiraLinhaDia.y - 4 : l.y > primeiraLinhaDia.y + 4)
      : linhasComTexto;
    const columns = this._detectarColunasVascon(linhasCabecalho, pagina.width);

    for (let i = 0; i < diaInfo.length; i++) {
      const atual = diaInfo[i];
      const prox = diaInfo[i + 1];
      const bloco = linhasComTexto.filter(l => {
        if (cresceParaBaixo) {
          const fim = prox ? prox.y : (Number.isFinite(obsY) ? obsY : Infinity);
          return l.y >= atual.y - 3 && l.y < fim - 3;
        }
        const fim = prox ? prox.y : (Number.isFinite(obsY) ? obsY : -Infinity);
        return l.y <= atual.y + 3 && l.y > fim + 3;
      });

      const parsedDia = this._parseBlocoDiaVascon(bloco, columns);
      result[atual.dia] = parsedDia;
    }

    return result;
  },

  _agruparLinhasPDF(items) {
    const sorted = [...items].sort((a, b) => {
      if (Math.abs(a.y - b.y) > 3) return a.y - b.y;
      return a.x - b.x;
    });
    const linhas = [];

    for (const item of sorted) {
      let linha = linhas.find(l => Math.abs(l.y - item.y) <= 3.2);
      if (!linha) {
        linha = { y: item.y, items: [] };
        linhas.push(linha);
      }
      linha.items.push(item);
      linha.y = (linha.y * (linha.items.length - 1) + item.y) / linha.items.length;
    }

    for (const linha of linhas) {
      linha.items.sort((a, b) => a.x - b.x);
    }

    return linhas.sort((a, b) => a.y - b.y);
  },

  _detectarColunasVascon(linhas, pageWidth) {
    const all = [];
    for (const linha of linhas) {
      for (const item of linha.items) {
        all.push({ text: item.text.toUpperCase(), x: item.x });
      }
    }

    const media = arr => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null;
    const xsPrincipal = all.filter(i => i.text.includes("PRATO") || i.text.includes("PRINCIPAL")).map(i => i.x);
    const xLight = media(all.filter(i => i.text.includes("LIGHT")).map(i => i.x));
    const xCarne = media(all.filter(i => i.text.includes("CARNE")).map(i => i.x));
    const xMassa = media(all.filter(i => i.text.includes("MASSA")).map(i => i.x));
    const xLanche = media(all.filter(i => i.text.includes("LANCHE")).map(i => i.x));
    const xPrincipal = media(xsPrincipal.filter(x => x > 120 && x < 330)) || 235;

    if ([xLight, xCarne, xMassa, xLanche].every(Number.isFinite)) {
      return {
        dayEnd: Math.max(120, Math.min(150, xPrincipal - 90)),
        principalEnd: (xPrincipal + xLight) / 2,
        lightEnd: (xLight + xCarne) / 2,
        carneEnd: (xCarne + xMassa) / 2,
        massaEnd: (xMassa + xLanche) / 2,
        lancheEnd: pageWidth || 9999
      };
    }

    return {
      dayEnd: 135,
      principalEnd: 335,
      lightEnd: 500,
      carneEnd: 605,
      massaEnd: 700,
      lancheEnd: pageWidth || 9999
    };
  },

  _colunaPorX(x, cols) {
    if (x < cols.dayEnd) return "dia";
    if (x < cols.principalEnd) return "principal";
    if (x < cols.lightEnd) return "light";
    if (x < cols.carneEnd) return "carne";
    if (x < cols.massaEnd) return "massa";
    return "lanche";
  },

  _parseBlocoDiaVascon(bloco, cols) {
    const linhasPorColuna = {
      principal: [], light: [], carne: [], massa: [], lanche: []
    };

    for (const linha of bloco) {
      const grupos = { principal: [], light: [], carne: [], massa: [], lanche: [] };

      for (const item of linha.items) {
        const col = this._colunaPorX(item.x, cols);
        if (!grupos[col]) continue;
        grupos[col].push(item.text);
      }

      for (const col of Object.keys(grupos)) {
        const texto = this._limparTextoCardapio(grupos[col].join(" "));
        if (!texto) continue;
        if (this._ignorarLinhaVascon(texto)) continue;
        linhasPorColuna[col].push(texto);
      }
    }

    const colunasComFeriado = Object.keys(linhasPorColuna).filter(col =>
      (linhasPorColuna[col] || []).some(t => this._ehFeriadoOuVazio(t))
    );

    // Regra Vascon/Homy: se o bloco do dia trouxer FERIADO em qualquer coluna de opção,
    // o dia inteiro é tratado como feriado. Não cadastramos opções vazias nem refeição normal.
    if (colunasComFeriado.length) {
      return {
        feriado: {
          nome: "FERIADO",
          detalhes: ["Não haverá marcação de refeição para este dia."]
        }
      };
    }

    const principal = this._limparListaCardapio(linhasPorColuna.principal);
    const light = this._montarOpcaoComDetalhes(linhasPorColuna.light, true);
    const carne = this._montarOpcaoComDetalhes(linhasPorColuna.carne, false);
    const massa = this._montarOpcaoComDetalhes(linhasPorColuna.massa, false);
    const lanche = this._montarOpcaoComDetalhes(linhasPorColuna.lanche, false);

    const out = {};
    if (principal.length) out.principal = principal;
    if (light?.nome) out.light = light;
    if (carne?.nome) out.carne = carne;
    if (massa?.nome) out.massa = massa;
    if (lanche?.nome) out.lanche = lanche;
    return out;
  },

  _montarOpcaoComDetalhes(linhas, primeiraLinhaComoNome = true) {
    const limpas = this._limparListaCardapio(linhas)
      .filter(x => !this._ehFeriadoOuVazio(x));
    if (!limpas.length) return null;

    if (primeiraLinhaComoNome) {
      return { nome: this._titleCaseCardapio(limpas[0]), detalhes: limpas.slice(1).map(x => this._titleCaseCardapio(x)) };
    }

    return { nome: this._titleCaseCardapio(limpas.join(" ")), detalhes: [] };
  },

  _limparListaCardapio(lista) {
    const out = [];
    for (const item of lista || []) {
      const texto = this._limparTextoCardapio(item);
      if (!texto || this._ignorarLinhaVascon(texto) || this._ehFeriadoOuVazio(texto)) continue;
      if (out[out.length - 1] !== texto) out.push(this._titleCaseCardapio(texto));
    }
    return out;
  },

  _limparTextoCardapio(v) {
    return String(v || "")
      .replace(/\([^)]*\)/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  },

  _ignorarLinhaVascon(texto) {
    const n = AdminUtils.norm(texto).toUpperCase();
    return /^(SEGUNDA|TERCA|TERÇA|QUARTA|QUINTA|SEXTA)$/.test(n) ||
      /^\d{2}\/\d{2}\/\d{2}$/.test(n) ||
      /^(PRATO|PRINCIPAL|OPCAO|OPÇÃO|LIGHT|CARNE|MASSA|LANCHE|DATA|VASCON|SERVICOS|SERVIÇOS|ALIMENTACAO|ALIMENTAÇÃO|OBS)$/.test(n);
  },

  _ehFeriadoOuVazio(texto) {
    const n = AdminUtils.norm(texto);
    return !n || n === "feriado" || n === "-" || n === "nao se aplica";
  },

  _titleCaseCardapio(texto) {
    const manterMaiusculo = new Set(["E", "À", "A", "AO", "DE", "DA", "DO", "DAS", "DOS", "EM", "COM"]);
    return String(texto || "")
      .toLowerCase()
      .replace(/(^|[\s/\-])([a-záéíóúâêôãõçà])/g, (m, sep, c) => sep + c.toUpperCase())
      .split(" ")
      .map((p, i) => i > 0 && manterMaiusculo.has(p.toUpperCase()) ? p.toLowerCase() : p)
      .join(" ")
      .replace(/\bFeijao\b/g, "Feijão")
      .replace(/\bOpcao\b/g, "Opção");
  },

  _parseVasconTexto(texto) {
    const raw = String(texto || "")
      .normalize("NFC")
      .replace(/\r/g, "\n")
      .replace(/\s+/g, " ")
      .toUpperCase()
      .trim();

    const result = { segunda: {}, terca: {}, quarta: {}, quinta: {}, sexta: {} };
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
      if (result[key] && Object.keys(result[key]).length) continue;
      const next = dias.slice(i + 1).map(d => d[0]).join("|") || "OBS";
      const re = new RegExp(`${pat}[\\s\\S]*?(?=${next}|OBS|$)`, "i");
      const m = raw.match(re);
      if (!m) continue;

      const blocoOriginal = m[0]
        .replace(new RegExp(pat, "i"), " ")
        .replace(/\([^)]*\)/g, " ")
        .replace(/\s+/g, " ")
        .trim();

      if (/FERIADO/i.test(blocoOriginal)) {
        result[key].feriado = { nome: "FERIADO", detalhes: ["Não haverá marcação de refeição para este dia."] };
        continue;
      }

      const bloco = blocoOriginal.replace(/\s+/g, " ").trim();

      // Fallback conservador: salva somente o texto completo como prato principal.
      // Se o PDF não trouxer posição, não inventa opções alternativas.
      if (bloco) result[key].principal = [this._titleCaseCardapio(bloco)];
    }

    return result;
  },

  _bindBotoes() {
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
