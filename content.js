(() => {
  try {
    const BTN_ID = 'similImportarCustosXlsxBtn';
    const INPUT_ID = 'similImportarCustosXlsxInput';
    const TOAST_ID = 'similImportarCustosXlsxToast';

    const ETAPAS_LIMITE = 20;

    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    const normalize = (s) =>
      String(s ?? '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();

    const isVisible = (el) => {
      if (!el) return false;
      const style = getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return (
        style.display !== 'none' &&
        style.visibility !== 'hidden' &&
        rect.width > 0 &&
        rect.height > 0
      );
    };

    const isEditableInput = (el) =>
      el &&
      el.tagName === 'INPUT' &&
      !el.disabled &&
      !el.readOnly &&
      isVisible(el);

    const toast = (msg) => {
      let t = document.getElementById(TOAST_ID);
      if (!t) {
        t = document.createElement('div');
        t.id = TOAST_ID;
        t.style.cssText = `
          position: fixed;
          right: 16px;
          bottom: 160px;
          z-index: 2147483647;
          background: rgba(0,0,0,.82);
          color: #fff;
          padding: 10px 12px;
          border-radius: 10px;
          font: 13px Arial, sans-serif;
          max-width: 460px;
          white-space: pre-line;
          box-shadow: 0 8px 24px rgba(0,0,0,.18);
        `;
        (document.body || document.documentElement).appendChild(t);
      }

      t.textContent = msg;
      t.style.display = 'block';

      clearTimeout(window.__similImportarCustosToastTimer);
      window.__similImportarCustosToastTimer = setTimeout(() => {
        t.style.display = 'none';
      }, 4500);
    };

    function setNativeValue(input, value) {
      const descriptor = Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value'
      );
      const setter = descriptor && descriptor.set;
      if (setter) {
        setter.call(input, value);
      } else {
        input.value = value;
      }
    }

    function preencherCampo(input, valor) {
      if (!input) return false;

      input.focus();
      setNativeValue(input, valor);
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      input.dispatchEvent(new Event('blur', { bubbles: true }));
      input.blur();

      return true;
    }

    function formatarMoedaBR(valor) {
      const numero = Number(valor || 0);
      return numero.toLocaleString('pt-BR', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      });
    }

    function parseNumero(valor) {
      if (valor == null || valor === '') return 0;

      if (typeof valor === 'number') {
        return Number.isFinite(valor) ? valor : 0;
      }

      let s = String(valor).trim();
      if (!s) return 0;

      s = s.replace(/\s/g, '');

      if (s.includes('.') && s.includes(',')) {
        s = s.replace(/\./g, '').replace(',', '.');
      } else if (s.includes(',')) {
        s = s.replace(',', '.');
      }

      const n = Number(s);
      return Number.isFinite(n) ? n : 0;
    }

    function obterValorCelula(sheet, row, col) {
      const addr = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = sheet[addr];
      return cell ? cell.v : null;
    }

    function encontrarBlocoCustosNaAba(sheet) {
      if (!sheet || !sheet['!ref']) return null;

      const range = XLSX.utils.decode_range(sheet['!ref']);
      const maxRow = Math.min(range.e.r, 500);
      const maxCol = range.e.c;

      const candidatos = [];

      for (let r = range.s.r; r <= maxRow; r++) {
        for (let c = range.s.c; c <= maxCol; c++) {
          const txt = normalize(obterValorCelula(sheet, r, c));
          if (txt !== 'servicos') continue;

          for (let rr = r; rr <= Math.min(r + 2, maxRow); rr++) {
            for (let cc = c + 1; cc <= maxCol; cc++) {
              const txt2 = normalize(obterValorCelula(sheet, rr, cc));
              if (
                txt2 === 'custos [r$]' ||
                txt2 === 'custo [r$]' ||
                txt2 === 'custos' ||
                txt2 === 'custo'
              ) {
                candidatos.push({
                  serviceCol: c,
                  costCol: cc,
                  headerRow: Math.min(r, rr),
                  startRow: Math.max(r, rr) + 1
                });
              }
            }
          }
        }
      }

      if (!candidatos.length) return null;

      let melhor = null;

      for (const cand of candidatos) {
        const custos = coletarCustosDoBloco(sheet, cand.serviceCol, cand.costCol, cand.startRow);
        const temOutros = custos.some(x => normalize(x.nome).startsWith('outros'));
        const score = (custos.length >= ETAPAS_LIMITE ? 1000 : 0) + (temOutros ? 100 : 0) + custos.length;

        if (!melhor || score > melhor.score) {
          melhor = { ...cand, custos, score };
        }
      }

      return melhor && melhor.custos.length ? melhor : null;
    }

    function coletarCustosDoBloco(sheet, serviceCol, costCol, startRow) {
      const custos = [];
      let linhasVaziasSeguidas = 0;

      for (let r = startRow; r <= startRow + 80; r++) {
        const nomeRaw = obterValorCelula(sheet, r, serviceCol);
        const custoRaw = obterValorCelula(sheet, r, costCol);

        const nome = String(nomeRaw ?? '').trim();
        const nomeNorm = normalize(nome);

        if (!nome) {
          linhasVaziasSeguidas++;
          if (custos.length && linhasVaziasSeguidas >= 3) break;
          continue;
        }

        linhasVaziasSeguidas = 0;

        if (nomeNorm === 'servicos') continue;
        if (nomeNorm.startsWith('custo total')) break;
        if (nomeNorm === 'bdi') break;

        custos.push({
          nome,
          valor: parseNumero(custoRaw)
        });

        if (nomeNorm.startsWith('outros')) break;
        if (custos.length >= ETAPAS_LIMITE) break;
      }

      return custos;
    }

    function extrairCustosWorkbook(workbook) {
      if (!workbook || !workbook.SheetNames?.length) return null;

      let melhor = null;

      for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const bloco = encontrarBlocoCustosNaAba(sheet);
        if (!bloco) continue;

        if (!melhor || bloco.score > melhor.score) {
          melhor = {
            sheetName,
            custos: bloco.custos
          };
        }
      }

      if (!melhor || !melhor.custos?.length) return null;

      const custos = melhor.custos
        .slice(0, ETAPAS_LIMITE)
        .map(item => ({
          nome: item.nome,
          valor: Number.isFinite(item.valor) ? item.valor : 0
        }));

      return {
        sheetName: melhor.sheetName,
        custos
      };
    }

    function encontrarSecaoValoresCustos() {
      const nodes = [...document.querySelectorAll('h1,h2,h3,h4,h5,h6,legend,span,div,strong,label')];

      for (const node of nodes) {
        const txt = normalize(node.textContent || '');
        if (txt !== 'valores/custos' && txt !== 'valores custos') continue;

        let atual = node;
        for (let i = 0; i < 6 && atual; i++) {
          const inputs = [...atual.querySelectorAll('input[type="text"], input:not([type])')]
            .filter(isEditableInput);

          if (inputs.length >= 10) return atual;
          atual = atual.parentElement;
        }
      }

      return null;
    }

    function encontrarInputsCustosPorRotulo() {
      const labels = [...document.querySelectorAll('label, span, div, td, strong')]
        .filter(el => normalize(el.textContent || '') === 'custo dos servicos' && isVisible(el));

      const encontrados = [];

      for (const label of labels) {
        const scopes = [
          label,
          label.parentElement,
          label.parentElement?.parentElement,
          label.parentElement?.parentElement?.parentElement
        ];

        let inputEncontrado = null;

        for (const scope of scopes) {
          if (!scope) continue;

          const inputs = [...scope.querySelectorAll('input[type="text"], input:not([type])')]
            .filter(isEditableInput);

          if (inputs.length === 1) {
            inputEncontrado = inputs[0];
            break;
          }

          if (inputs.length > 1) {
            inputEncontrado = inputs[inputs.length - 1];
            break;
          }
        }

        if (inputEncontrado && !encontrados.includes(inputEncontrado)) {
          encontrados.push(inputEncontrado);
        }
      }

      if (encontrados.length >= ETAPAS_LIMITE) {
        return encontrados.slice(0, ETAPAS_LIMITE);
      }

      const secao = encontrarSecaoValoresCustos();
      if (!secao) return encontrados;

      const inputsSecao = [...secao.querySelectorAll('input[type="text"], input:not([type])')]
        .filter(isEditableInput);

      for (const input of inputsSecao) {
        if (!encontrados.includes(input)) {
          encontrados.push(input);
        }
      }

      return encontrados.slice(0, ETAPAS_LIMITE);
    }

    async function preencherCustosNoSimil(custos) {
      const inputs = encontrarInputsCustosPorRotulo();

      if (inputs.length < ETAPAS_LIMITE) {
        throw new Error(
          `Encontrei apenas ${inputs.length} campos de custo no SIMIL.`
        );
      }

      const resultados = [];

      for (let i = 0; i < ETAPAS_LIMITE; i++) {
        const item = custos[i] || { nome: `Etapa ${i + 1}`, valor: 0 };
        const valorFormatado = formatarMoedaBR(item.valor);
        const ok = preencherCampo(inputs[i], valorFormatado);

        resultados.push({
          ordem: i + 1,
          nome: item.nome,
          valor: valorFormatado,
          ok
        });

        await sleep(40);
      }

      return resultados;
    }

    async function processarPlanilha(file) {
      if (typeof XLSX === 'undefined') {
        toast('Biblioteca XLSX não encontrada. Coloque o arquivo xlsx.full.min.js na pasta da extensão.');
        return;
      }

      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });

      const extraido = extrairCustosWorkbook(workbook);

      if (!extraido || !extraido.custos?.length) {
        toast('Não consegui localizar o bloco de Custos na planilha.');
        return;
      }

      if (extraido.custos.length < ETAPAS_LIMITE) {
        toast(
          `Encontrei apenas ${extraido.custos.length} etapas de custos na aba "${extraido.sheetName}".`
        );
        return;
      }

      const resultado = await preencherCustosNoSimil(extraido.custos);
      const preenchidos = resultado.filter(x => x.ok).length;

      console.log('[SIMIL-CUSTOS-XLSX] Aba encontrada:', extraido.sheetName);
      console.log('[SIMIL-CUSTOS-XLSX] Custos extraídos:', extraido.custos);
      console.log('[SIMIL-CUSTOS-XLSX] Resultado preenchimento:', resultado);

      toast(
        `Planilha lida com sucesso.\n` +
        `Aba: ${extraido.sheetName}\n` +
        `Preenchidos ${preenchidos}/${resultado.length} campos de custos.`
      );
    }

    function criarInputArquivo() {
      let input = document.getElementById(INPUT_ID);
      if (input) return input;

      input = document.createElement('input');
      input.id = INPUT_ID;
      input.type = 'file';
      input.accept = '.xlsx,.xls,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel';
      input.style.display = 'none';

      input.addEventListener('change', async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;

        try {
          await processarPlanilha(file);
        } catch (err) {
          console.error('[SIMIL-CUSTOS-XLSX] Erro:', err);
          toast(`Erro ao processar a planilha: ${err.message || err}`);
        } finally {
          input.value = '';
        }
      });

      (document.body || document.documentElement).appendChild(input);
      return input;
    }

    function paginaTemSecaoCustos() {
      return (
        !!encontrarSecaoValoresCustos() ||
        encontrarInputsCustosPorRotulo().length >= 5
      );
    }

    function ensureButton() {
      if (!paginaTemSecaoCustos()) return;
      if (document.getElementById(BTN_ID)) return;

      const btn = document.createElement('button');
      btn.id = BTN_ID;
      btn.type = 'button';
      btn.textContent = 'Importar Custos XLSX';
      btn.title = 'Importar custos da planilha para o SIMIL';

      btn.style.cssText = `
        position: fixed;
        right: 160px;
        bottom: 16px;
        z-index: 2147483647;
        padding: 10px 12px;
        border: 0;
        border-radius: 10px;
        background: #f39201;
        color: #fff;
        font: 600 13px Arial, sans-serif;
        cursor: pointer;
        box-shadow: 0 8px 24px rgba(0,0,0,.18);
      `;

      btn.addEventListener('click', () => {
        const input = criarInputArquivo();
        input.click();
      });

      (document.body || document.documentElement).appendChild(btn);
    }

    ensureButton();

    const obs = new MutationObserver(() => ensureButton());
    obs.observe(document.documentElement, { childList: true, subtree: true });

  } catch (err) {
    console.error('[SIMIL-CUSTOS-XLSX] erro:', err);
  }
})();