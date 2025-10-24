document.addEventListener('DOMContentLoaded', function () {
    const tbody = document.getElementById('rows-body');
    const rows = 20;

    // Cria a tabela dinamicamente
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

        tr.appendChild(tdInput('text', `empresa_${i}`));
        tr.appendChild(tdInput('text', `exercicio_${i}`));
        tr.appendChild(tdInput('text', `trimestre_${i}`));
        tr.appendChild(tdInput('text', `campo_${i}`));
        tr.appendChild(tdInput('text', `fase_${i}`));
        tr.appendChild(tdInput('text', `status_${i}`));
        tr.appendChild(tdInput('text', `versao_${i}`));
        tr.appendChild(tdInput('text', `secao_${i}`));
        tr.appendChild(tdInput('text', `defprojeto_${i}`));

        const tdDate = document.createElement('td');
        const dateInput = document.createElement('input');
        dateInput.className = 'form-control';
        dateInput.name = `datainicio_${i}`;
        dateInput.type = 'date';
        tdDate.appendChild(dateInput);
        tr.appendChild(tdDate);

        tr.appendChild(tdInput('text', `bidround_${i}`));
        tbody.appendChild(tr);
    }

    // 📁 Captura das seções
    const tableSection = document.querySelector('.table-container');
    const directoriesSection = document.querySelectorAll('.form-section')[1]; // segunda form-section (diretórios)
    const runSection = document.getElementById('runSection');

    const pathInputs = {
        path1: document.querySelector('[name="path1"]').closest('.mb-3'),
        path2: document.querySelector('[name="path2"]').closest('.mb-3'),
        path3: document.querySelector('[name="path3"]').closest('.mb-3'),
        path4: document.querySelector('[name="path4"]').closest('.mb-3'),
        path5: document.querySelector('[name="path5"]').closest('.mb-3'),
        path6: document.querySelector('[name="path6"]').closest('.mb-3')
    };

    [tableSection, directoriesSection, runSection, ...Object.values(pathInputs)].forEach(el => {
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

    [tableSection, directoriesSection, runSection, ...Object.values(pathInputs)].forEach(el => {
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
                    bidround: document.querySelector(`[name="bidround_${i}"]`).value.trim()
                };

                const hasValue = Object.values(rowObj).some(v => v !== "");
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

            const payload = { paths: paths, requests: data };

            // 🔹 Exibe avisos com spinner
            const mensagens = [];
            if (document.getElementById('switch1')?.checked) mensagens.push("Aguardando requisição da base do SAP");
            if (document.getElementById('switch2')?.checked) mensagens.push("Aguardando relatório completo");
            if (document.getElementById('switch3')?.checked) mensagens.push("Aguardando relatório reduzido");
            if (document.getElementById('switch4')?.checked) mensagens.push("Aguardando relatório de Gastos Diretos");
            if (document.getElementById('switch5')?.checked) mensagens.push("Aguardando relatório de Gastos Indiretos");
            if (document.getElementById('switch6')?.checked) mensagens.push("Aguardando relatório de Estoques");

            if (mensagens.length > 0) exibirAvisos(mensagens);

            // 🚀 Chamada ao backend Python
            try {
                const response = await fetch("http://127.0.0.1:8000/save_requests", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify(payload)
                });

                const result = await response.json();

                if (result.status === "success") {
                atualizarAvisoFinal(
                    '<i class="bi bi-check-circle-fill" style="color: green; font-size:1.1rem;"></i>',
                    result.message || "Arquivo criado com sucesso!"
                );
                } else {
                throw new Error(result.message || "Erro desconhecido");
                }


            } catch (err) {
                //  erro
                atualizarAvisoFinal(
                    '<i class="bi bi-x" style="color: red; font-size:1.1rem;"></i>',
                    "Ocorreu um erro na operação, revise os dados e tente novamente. Caso o erro persista, entre em contato com o administrador do programa"
                );
            }

            console.log("🔹 Linhas:", data);
            console.log("🔹 Paths:", paths);
        });
    }

    updateVisibility();
});
