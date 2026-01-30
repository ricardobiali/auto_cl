// frontend/script.js
document.addEventListener('DOMContentLoaded', function () {

    // =====================================================
    // 1) WELCOME (robusto: tenta várias funções no backend)
    // =====================================================
    async function getWelcomeProfileSafe() {
        // tenta get_welcome_user -> get_user_profile -> get_welcome_name
        try {
            if (eel.get_welcome_user) return await eel.get_welcome_user()(); // {name, gender}
        } catch (e) { /* ignore */ }

        try {
            if (eel.get_user_profile) return await eel.get_user_profile()(); // {name, gender}
        } catch (e) { /* ignore */ }

        try {
            if (eel.get_welcome_name) {
                const nome = await eel.get_welcome_name()();
                return { name: nome, gender: "m" };
            }
        } catch (e) { /* ignore */ }

        return { name: "", gender: "m" };
    }

    async function atualizarNomeUsuario() {
        try {
            const profile = await getWelcomeProfileSafe();
            const nome = (profile?.name || "").trim();
            const gender = (profile?.gender || "m").toLowerCase();

            const welcomeEl = document.querySelector('p.mb-0');
            if (welcomeEl) {
                const saudacao = (gender === "f") ? "Seja bem-vinda, " : "Seja bem-vindo, ";
                welcomeEl.textContent = `${saudacao}${nome}!`;
            }
        } catch (err) {
            console.error("Erro ao obter dados do usuário:", err);
        }
    }

    atualizarNomeUsuario();

    // =====================================================
    // TABELA (seu código original)
    // =====================================================
    const tbody = document.getElementById('rows-body');

    // =====================================================
    // BOTÃO: Importar planilha (visível só com switch1)
    // =====================================================
    const tableContainer = document.querySelector('.table-container');

    const importWrapper = document.createElement('div');
    importWrapper.className = 'import-wrapper';
    importWrapper.style.display = 'none'; // começa oculto

    const importBtn = document.createElement('button');
    importBtn.textContent = 'Importar planilha…';
    importBtn.className = 'import-planilha-btn';

    const importStatus = document.createElement('span');
    importStatus.className = 'import-status';
    importStatus.textContent = '';

    function setImportStatus(kind, text) {
        importStatus.classList.remove('ok', 'err', 'info');
        if (kind) importStatus.classList.add(kind);
        importStatus.textContent = text || '';
    }

    importBtn.addEventListener('click', async () => {
        try {
            setImportStatus('info', 'Importando...');

            const switches = {
                report_SAP: document.getElementById('switch1')?.checked || false,
                completa: document.getElementById('switch2')?.checked || false,
                reduzida: document.getElementById('switch3')?.checked || false,
                diretos: document.getElementById('switch4')?.checked || false,
                indiretos: document.getElementById('switch5')?.checked || false,
                estoques: document.getElementById('switch6')?.checked || false
            };

            if (!switches.report_SAP) {
                setImportStatus('err', 'Ative o Switch 1 para importar.');
                return;
            }

            const paths = {
                path1: document.querySelector("input[name='path1']")?.value || "",
                path2: document.querySelector("input[name='path2']")?.value || "",
                path3: document.querySelector("input[name='path3']")?.value || "",
                path4: document.querySelector("input[name='path4']")?.value || "",
                path5: document.querySelector("input[name='path5']")?.value || "",
                path6: document.querySelector("input[name='path6']")?.value || ""
            };

            // chama o backend
            const res = await eel.import_planilha(switches, paths)();

            if (!res) {
                setImportStatus('err', 'Falha ao importar.');
                return;
            }

            if (res.status === 'cancelled') {
                setImportStatus('', '');
                return;
            }

            if (res.status === 'imported') {
                const n = res.imported_rows ?? '';
                setImportStatus('ok', `Importação realizada com sucesso${n ? ` (${n} linhas)` : ''}.`);
                return;
            }

            // outros status
            setImportStatus('err', res.error || `Erro: ${res.status}`);
        } catch (err) {
            console.error(err);
            setImportStatus('err', 'Erro ao importar planilha.');
        }
    });

    importWrapper.appendChild(importBtn);
    importWrapper.appendChild(importStatus);

    // insere acima da tabela
    if (tableContainer) {
        tableContainer.parentNode.insertBefore(importWrapper, tableContainer);
    }

    const rows = 10;

    for (let i = 1; i <= rows; i++) {
        const tr = document.createElement('tr');
        const tdIndex = document.createElement('td');
        tdIndex.textContent = i;
        tr.appendChild(tdIndex);

        function tdInput(type, name) {
            const td = document.createElement('td');
            const input = document.createElement('input');
            input.className = 'form-control';
            input.name = name;
            input.type = type || 'text';
            td.appendChild(input);
            return td;
        }

        function tdSelect(name, options) {
            const td = document.createElement('td');
            const select = document.createElement('select');
            select.className = 'form-select';
            select.name = name;
            const emptyOption = document.createElement('option');
            emptyOption.value = "";
            emptyOption.textContent = "Selecione...";
            select.appendChild(emptyOption);
            options.forEach(opt => {
                const option = document.createElement('option');
                option.value = opt;
                option.textContent = opt;
                select.appendChild(option);
            });
            td.appendChild(select);
            return td;
        }

        function tdSwitch(name) {
            const td = document.createElement('td');
            const div = document.createElement('div');
            div.className = 'form-check form-switch d-flex justify-content-center';
            const input = document.createElement('input');
            input.type = 'checkbox';
            input.className = 'form-check-input';
            input.name = name;
            input.role = 'switch';
            div.appendChild(input);
            td.appendChild(div);
            return td;
        }

        tr.appendChild(tdInput('text', `empresa_${i}`));
        tr.appendChild(tdInput('text', `exercicio_${i}`));
        tr.appendChild(tdSelect(`trimestre_${i}`, [1, 2, 3, 4]));
        tr.appendChild(tdInput('text', `campo_${i}`));
        tr.appendChild(tdSelect(`fase_${i}`, ['E', 'D', 'P']));
        tr.appendChild(tdSelect(`status_${i}`, [1, 2, 3, 4, 5, 6]));
        tr.appendChild(tdSelect(`versao_${i}`, [0, 1, 2, 3]));
        tr.appendChild(tdSelect(`secao_${i}`, ['ANP_0901', 'CL_PADRAO']));
        tr.appendChild(tdInput('text', `defprojeto_${i}`));

        const tdDate = document.createElement('td');
        const dateInput = document.createElement('input');
        dateInput.className = 'form-control';
        dateInput.name = `datainicio_${i}`;
        dateInput.type = 'date';
        tdDate.appendChild(dateInput);
        tr.appendChild(tdDate);

        tr.appendChild(tdInput('text', `bidround_${i}`));
        tr.appendChild(tdSwitch(`rit_${i}`));
        tbody.appendChild(tr);
    }

    // =====================================================
    // Captura de seções (seu código original)
    // =====================================================
    const optionsSection = document.querySelectorAll('.form-section')[0];
    const tableSection = document.querySelector('.table-container');
    const directoriesSection = document.querySelectorAll('.form-section')[1];
    const runSection = document.getElementById('runSection');
    const cancelSection = document.getElementById('cancelSection');
    const cancelBtn = document.getElementById('cancelBtn');

    const pathInputs = {
        path1: document.querySelector('[name="path1"]').closest('.mb-3'),
        path2: document.querySelector('[name="path2"]').closest('.mb-3'),
        path3: document.querySelector('[name="path3"]').closest('.mb-3'),
        path4: document.querySelector('[name="path4"]').closest('.mb-3'),
        path5: document.querySelector('[name="path5"]').closest('.mb-3'),
        path6: document.querySelector('[name="path6"]').closest('.mb-3')
    };

    [tableSection, directoriesSection, runSection, cancelSection, ...Object.values(pathInputs)].forEach(el => {
        if (!el) return;
        el.classList.add('fade-toggle');
    });

    function toggleFade(element, show) {
        if (!element) return;
        const duration = 400;
        if (show) {
            element.style.display = '';
            void element.offsetWidth;
            element.classList.add('show');
        } else {
            element.classList.remove('show');
            setTimeout(() => {
                if (!element.classList.contains('show')) {
                    element.style.display = 'none';
                }
            }, duration + 20);
        }
    }

    [tableSection, directoriesSection, runSection, cancelSection, ...Object.values(pathInputs)].forEach(el => {
        if (!el) return;
        if (!el.classList.contains('show')) {
            el.style.display = 'none';
        }
    });

    function updateVisibility() {
        const s1 = document.getElementById('switch1').checked;
        const s2 = document.getElementById('switch2').checked;
        const s3 = document.getElementById('switch3').checked;
        const s4 = document.getElementById('switch4').checked;
        const s5 = document.getElementById('switch5').checked;
        const s6 = document.getElementById('switch6').checked;

        const anySwitchOn = s1 || s2 || s3 || s4 || s5 || s6;

        toggleFade(tableSection, s1);
        importWrapper.style.display = s1 ? 'flex' : 'none';
        if (!s1) setImportStatus('', '');
        toggleFade(pathInputs.path1, s1);
        toggleFade(pathInputs.path2, s2);
        toggleFade(pathInputs.path3, s3);
        toggleFade(pathInputs.path4, s4);
        toggleFade(pathInputs.path5, s5);
        toggleFade(pathInputs.path6, s6);
        toggleFade(directoriesSection, anySwitchOn);
        toggleFade(runSection, anySwitchOn);
    }

    for (let i = 1; i <= 6; i++) {
        const sw = document.getElementById(`switch${i}`);
        if (sw) sw.addEventListener('change', updateVisibility);
    }

    const runBtn = document.getElementById('runBtn');
    const avisosContainer = document.getElementById('avisosContainer');
    const avisosContent = document.getElementById('avisosContent');

    // =====================================================
    // 2) AVISOS + LOGS em tempo real
    // =====================================================
    function exibirAvisos(mensagens) {
        avisosContent.innerHTML = mensagens.map(msg => `
            <li>
                <div class="spinner-border" role="status" aria-hidden="true"></div>
                <span>${msg}</span>
            </li>
        `).join('');
        avisosContainer.style.display = 'block';
        avisosContainer.classList.add('fade-in');
    }

    function atualizarAvisoFinal(iconHtml, texto) {
        avisosContent.innerHTML = `
            <li style="display:flex; align-items:center; gap:6px;">
                ${iconHtml}
                <span>${texto}</span>
            </li>
        `;
    }

    // ✅ NOVO: renderiza logs do backend (SAP) enquanto roda
    let _lastLogsLen = 0;

    function renderLogsAoVivo(logs) {
        if (!Array.isArray(logs) || logs.length === 0) return;

        // se resetou (novo job), reinicia
        if (logs.length < _lastLogsLen) _lastLogsLen = 0;

        // adiciona somente os novos
        const novos = logs.slice(_lastLogsLen);
        _lastLogsLen = logs.length;
        if (novos.length === 0) return;

        // garante que o container está visível
        avisosContainer.style.display = 'block';
        avisosContainer.classList.add('fade-in');

        // adiciona no final do <ul>
        const html = novos.map(l => `
            <li style="display:flex; align-items:center; gap:6px;">
                <span style="font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono', 'Courier New', monospace;">
                    ${escapeHtml(String(l))}
                </span>
            </li>
        `).join('');

        avisosContent.insertAdjacentHTML('beforeend', html);

        // rola pra baixo
        avisosContainer.scrollTop = avisosContainer.scrollHeight;
    }

    function escapeHtml(str) {
        return str
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll('"', "&quot;")
            .replaceAll("'", "&#039;");
    }

    function hideSectionsOnRun() {
        const sectionsToHide = [
            optionsSection,
            tableSection,
            directoriesSection,
            runSection,
            ...Object.values(pathInputs)
        ];
        sectionsToHide.forEach(el => toggleFade(el, false));
    }

    function atualizarAvisoEmExecucao(texto) {
        if (!texto) return;

        // tenta achar o primeiro aviso "principal" (o que tem spinner)
        const principal = avisosContent.querySelector("li .spinner-border + span");
        if (principal) {
            principal.textContent = texto;
        }
    }


    // =====================================================
    // Polling status (✅ agora também puxa logs)
    // =====================================================
    async function checarStatus() {
        try {
            const status = await eel.get_job_status()();

            // ✅ NOVO: mostra logs enquanto roda
            if (status?.logs) {
                renderLogsAoVivo(status.logs);
            }

            // ✅ NOVO: mensagem viva de progresso
            if (status.running && status.message) {
                atualizarAvisoEmExecucao(status.message);
            }

            if (!status.running) {
                atualizarAvisoFinal(
                    status.success
                        ? '<i class="bi bi-check-circle-fill" style="color: green; font-size:1.1rem;"></i>'
                        : '<i class="bi bi-x" style="color: red; font-size:1.1rem;"></i>',
                    status.message
                );

                const newBtn = cancelBtn.cloneNode(true);
                newBtn.textContent = "Reiniciar";
                newBtn.onclick = () => window.location.reload();
                cancelBtn.parentNode.replaceChild(newBtn, cancelBtn);

            } else {
                setTimeout(checarStatus, 1000);
            }

        } catch (err) {
            console.error(err);
            setTimeout(checarStatus, 1500);
        }
    }

    // =====================================================
    // RUN
    // =====================================================
    if (runBtn) {
        runBtn.addEventListener('click', async function () {
            const data = [];
            for (let i = 1; i <= rows; i++) {
                const rawDate = document.querySelector(`[name="datainicio_${i}"]`).value;
                let formattedDate = "";
                if (rawDate) {
                    const [year, month, day] = rawDate.split("-");
                    formattedDate = `${day}${month}${year}`;
                }

                const rowObj = {
                    empresa: document.querySelector(`[name="empresa_${i}"]`).value.trim(),
                    exercicio: document.querySelector(`[name="exercicio_${i}"]`).value.trim(),
                    trimestre: document.querySelector(`[name="trimestre_${i}"]`).value.trim(),
                    campo: document.querySelector(`[name="campo_${i}"]`).value.trim(),
                    fase: document.querySelector(`[name="fase_${i}"]`).value.trim(),
                    status: document.querySelector(`[name="status_${i}"]`).value.trim(),
                    versao: document.querySelector(`[name="versao_${i}"]`).value.trim(),
                    secao: document.querySelector(`[name="secao_${i}"]`).value.trim(),
                    defprojeto: document.querySelector(`[name="defprojeto_${i}"]`).value.trim(),
                    datainicio: formattedDate,
                    bidround: document.querySelector(`[name="bidround_${i}"]`).value.trim(),
                    rit: document.querySelector(`[name="rit_${i}"]`).checked
                };

                const hasValue = Object.entries(rowObj)
                    .filter(([key]) => key !== "rit")
                    .some(([, value]) => value !== "");
                if (hasValue) data.push(rowObj);
            }

            const paths = [{
                path1: document.querySelector("input[name='path1']").value || "",
                path2: document.querySelector("input[name='path2']")?.value || "",
                path3: document.querySelector("input[name='path3']")?.value || "",
                path4: document.querySelector("input[name='path4']")?.value || "",
                path5: document.querySelector("input[name='path5']")?.value || "",
                path6: document.querySelector("input[name='path6']")?.value || ""
            }];

            const switches = {
                report_SAP: document.getElementById('switch1').checked,
                completa: document.getElementById('switch2').checked,
                reduzida: document.getElementById('switch3').checked,
                diretos: document.getElementById('switch4').checked,
                indiretos: document.getElementById('switch5').checked,
                estoques: document.getElementById('switch6').checked
            };

            const res = await eel.import_planilha(switches, paths)();
            console.log(res);

            const payload = { paths: paths, requests: data, switches: switches };

            // ✅ reset do log incremental (novo job)
            _lastLogsLen = 0;

            const mensagens = [];
            if (document.getElementById('switch1')?.checked) mensagens.push("Aguardando requisição da base do SAP");
            if (document.getElementById('switch2')?.checked) mensagens.push("Aguardando relatório completo");
            if (document.getElementById('switch3')?.checked) mensagens.push("Aguardando relatório reduzido");
            if (document.getElementById('switch4')?.checked) mensagens.push("Aguardando relatório de Gastos Diretos");
            if (document.getElementById('switch5')?.checked) mensagens.push("Aguardando relatório de Gastos Indiretos");
            if (document.getElementById('switch6')?.checked) mensagens.push("Aguardando relatório de Estoques");
            if (mensagens.length > 0) exibirAvisos(mensagens);

            hideSectionsOnRun();
            toggleFade(cancelSection, true);

            await eel.save_requests(payload)();

            const result = await eel.start_job(switches, paths[0])();
            if (result.status === "started") {
                checarStatus();
            } else if (result.status === "already_running") {
                atualizarAvisoFinal(
                    '<i class="bi bi-exclamation-triangle-fill" style="color: orange; font-size:1.1rem;"></i>',
                    "Um job já está em execução. Aguarde ele terminar antes de iniciar outro."
                );
            }

            console.log("Linhas:", data);
            console.log("Paths:", paths);
        });
    }

    // =====================================================
    // Cancel
    // =====================================================
    if (cancelBtn) {
        cancelBtn.addEventListener('click', async () => {
            try {
                atualizarAvisoFinal(
                    '<i class="bi bi-exclamation-octagon-fill" style="color: orange; font-size:1.1rem;"></i>',
                    'Cancelando automação...'
                );

                await eel.cancel_job()();

                setTimeout(() => {
                    window.location.reload();
                }, 1000);

            } catch (err) {
                console.error("Erro ao cancelar:", err);
            }
        });
    }

    // =====================================================
    // Dialog de diretório
    // =====================================================
    async function selecionarDiretorio() {
        try {
            const caminho = await eel.selecionar_diretorio()();
            return caminho;
        } catch (error) {
            console.error("Erro ao selecionar diretório:", error);
            return "";
        }
    }

    document.querySelectorAll(".input-group-text").forEach(icon => {
        icon.addEventListener("click", async (e) => {
            const input = e.target.closest(".input-group").querySelector("input");
            if (input) {
                const caminho = await selecionarDiretorio();
                if (caminho) input.value = caminho;
            }
        });
    });

    updateVisibility();
});
