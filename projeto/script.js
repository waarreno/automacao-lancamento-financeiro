const { chromium } = require('playwright');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// ===== CONFIGURAÇÕES GLOBAIS =====
const CONFIG = {
    TIMEOUT_NAVEGACAO: 30000,
    MAX_TENTATIVAS_ERRO: 3,
    TIMEOUT_MENSAGEM_SUCESSO: 15000,
    TIMEOUT_LIMPEZA_FORMULARIO: 3000,
    TIMEOUT_CARREGAMENTO_CAMPO: 4000
};

// ===== FUNÇÃO UTILITÁRIA PARA LOG COM TIMESTAMP =====
function logComTimestamp(mensagem) {
    const timestamp = new Date().toLocaleString('pt-BR');
    console.log(`[${timestamp}] ${mensagem}`);
}

// ===== FUNÇÃO PARA VERIFICAR SUCESSO DO LANÇAMENTO =====
async function verificarSucessoLancamento(page) {
    logComTimestamp('🔍 Verificando se apareceu mensagem de sucesso...');
    
    try {
        const frameContent = page.locator('#iframeasp').contentFrame()
            .locator('iframe[name="principal2"]').contentFrame();
        
        const mensagemSucesso = frameContent.locator('#txt_msg');
        
        // Aguarda a mensagem aparecer
        await mensagemSucesso.waitFor({ 
            state: 'visible', 
            timeout: CONFIG.TIMEOUT_MENSAGEM_SUCESSO 
        });
        
        const textoMensagem = await mensagemSucesso.textContent();
        
        if (textoMensagem?.includes('Operação realizada com sucesso')) {
            const match = textoMensagem.match(/Item\(s\) Pagamento gerado\(s\) : (\d+)/);
            const numeroPagamento = match ? match[1] : 'N/A';
            
            logComTimestamp(`✅ SUCESSO confirmado! Pagamento gerado: ${numeroPagamento}`);
            return {
                sucesso: true,
                mensagem: textoMensagem.trim(),
                numeroPagamento: numeroPagamento
            };
        } else {
            logComTimestamp(`⚠️ Mensagem encontrada mas sem confirmação de sucesso: "${textoMensagem}"`);
            return {
                sucesso: false,
                mensagem: textoMensagem ? textoMensagem.trim() : 'Mensagem vazia',
                numeroPagamento: null,
                motivo: 'Mensagem não indica sucesso'
            };
        }
        
    } catch (timeoutError) {
        logComTimestamp('❌ FALHA: Nenhuma mensagem de sucesso apareceu no tempo esperado');
        return {
            sucesso: false,
            mensagem: null,
            numeroPagamento: null,
            motivo: 'Timeout aguardando mensagem de sucesso'
        };
    }
}

// ===== VALIDAÇÃO DA PLANILHA =====
async function validarPlanilha(dados) {
    logComTimestamp('🔍 Validando dados da planilha...');
    
    const errosValidacao = [];
    
    dados.forEach((linha, index) => {
        const numeroLinha = index + 1;
        
        // Validações obrigatórias
        if (!linha.cod_tipo_rubrica) {
            errosValidacao.push(`Linha ${numeroLinha}: Código Tipo Rubrica está vazio`);
        }
        
        if (!linha.cod_prestador) {
            errosValidacao.push(`Linha ${numeroLinha}: Código Prestador está vazio`);
        }
        
        if (!linha.val_bruto) {
            errosValidacao.push(`Linha ${numeroLinha}: Valor Bruto está vazio`);
        }
        
        // Validação de formato de data (se preenchida)
        if (linha.dt_pgto_prevista && linha.dt_pgto_prevista.length > 0) {
            const formatoData = /^\d{2}\/\d{2}\/\d{4}$/;
            if (!formatoData.test(linha.dt_pgto_prevista)) {
                errosValidacao.push(`Linha ${numeroLinha}: Data de pagamento deve estar no formato DD/MM/AAAA`);
            }
        }
        
        // Validação de formato de mês/ano (se preenchida)
        if (linha.mes_ano_ref && linha.mes_ano_ref.length > 0) {
            const formatoMesAno = /^\d{2}\/\d{4}$/;
            if (!formatoMesAno.test(linha.mes_ano_ref)) {
                errosValidacao.push(`Linha ${numeroLinha}: Mês/Ano deve estar no formato MM/AAAA`);
            }
        }
    });
    
    if (errosValidacao.length > 0) {
        console.error('❌ ERROS DE VALIDAÇÃO ENCONTRADOS:');
        errosValidacao.forEach(erro => console.error(`   - ${erro}`));
        throw new Error(`${errosValidacao.length} erro(s) de validação encontrado(s)`);
    }
    
    logComTimestamp(`✅ Validação concluída: ${dados.length} linha(s) válida(s)`);
}

async function lerPlanilhaExcel() {
    try {
        logComTimestamp('📊 Lendo planilha Excel...');
        
        if (!fs.existsSync('dados_lancamento.xlsx')) {
            throw new Error('Arquivo dados_lancamento.xlsx não encontrado');
        }
        
        const workbook = XLSX.readFile('dados_lancamento.xlsx');
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const dados = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        // Remove linhas completamente vazias
        const dadosLimpos = dados.filter(linha => 
            linha && linha.some(celula => celula !== undefined && celula !== null && celula.toString().trim() !== '')
        );
        
        logComTimestamp(`✅ Planilha lida com sucesso! ${dadosLimpos.length} linha(s) encontrada(s)`);
        
        const dadosFormatados = dadosLimpos.map((linha, index) => {
            return {
                linha: index + 1,
                cod_tipo_rubrica: linha[0]?.toString().trim() || '',
                cod_prestador: linha[1]?.toString().trim() || '',
                mes_ano_ref: linha[2]?.toString().trim() || '',
                val_bruto: linha[3]?.toString().trim() || '',
                dt_pgto_prevista: linha[4]?.toString().trim() || '',
                observacoes: linha[5]?.toString().trim() || ''
            };
        });
        
        await validarPlanilha(dadosFormatados);
        return dadosFormatados;
        
    } catch (error) {
        console.error('❌ Erro ao ler planilha Excel:', error.message);
        throw error;
    }
}

async function aguardarLimpezaFormulario(frameContent) {
    logComTimestamp('🧹 Aguardando formulário ser limpo/preparado...');
    
    try {
        // Aguarda o campo cod_tipo_rubrica estar limpo ou pronto para preenchimento
        await frameContent.locator('#cod_tipo_rubrica').waitFor({ state: 'attached' });
        
        // Aguarda um pouco mais para garantir que o formulário foi completamente limpo
        await new Promise(resolve => setTimeout(resolve, CONFIG.TIMEOUT_LIMPEZA_FORMULARIO));
        
        logComTimestamp('✅ Formulário preparado para preenchimento');
    } catch (error) {
        logComTimestamp('⚠️ Erro ao aguardar limpeza do formulário, continuando...');
    }
}

async function aguardarCarregamentoCampo(frameContent, campo, valorPreenchido) {
    logComTimestamp(`⏳ Aguardando carregamento de informações do campo ${campo}...`);
    
    try {
        // Para cod_tipo_rubrica e cod_prestador, aguarda possíveis mudanças de estado
        // ou carregamento de informações relacionadas
        await new Promise(resolve => setTimeout(resolve, CONFIG.TIMEOUT_CARREGAMENTO_CAMPO));
        
        logComTimestamp(`✅ Carregamento do campo ${campo} concluído`);
    } catch (error) {
        logComTimestamp(`⚠️ Erro ao aguardar carregamento do campo ${campo}`);
    }
}

async function preencherFormulario(page, dados) {
    logComTimestamp(`📋 Preenchendo formulário com dados da linha ${dados.linha}...`);
    
    const frameContent = page.locator('#iframeasp').contentFrame()
        .locator('iframe[name="principal2"]').contentFrame();
    
    try {
        // Função auxiliar para preencher campo com aguardo específico
        async function preencherCampo(campo, valor, usarTab = false, aguardarCarregamento = false) {
            if (!valor) return; // Pula campos vazios
            
            logComTimestamp(`✏️ Preenchendo ${campo}: ${valor}`);
            
            const campoElement = frameContent.locator(`#${campo}`);
            await campoElement.fill(valor);
            
            if (usarTab) {
                await campoElement.press('Tab');
            }
            
            // Aguarda carregamento específico para campos que carregam informações
            if (aguardarCarregamento) {
                await aguardarCarregamentoCampo(frameContent, campo, valor);
            }
        }
        
        // Preenche campos críticos com aguardo de carregamento
        await preencherCampo('cod_tipo_rubrica', dados.cod_tipo_rubrica, true, true);
        await preencherCampo('cod_prestador', dados.cod_prestador, true, true);
        
        // Preenche demais campos normalmente
        await preencherCampo('mes_ano_ref', dados.mes_ano_ref);
        await preencherCampo('val_bruto', dados.val_bruto);
        await preencherCampo('dt_pgto_prevista', dados.dt_pgto_prevista);
        await preencherCampo('txt_obs_lm', dados.observacoes);
        
        logComTimestamp(`✅ Formulário da linha ${dados.linha} preenchido com sucesso!`);
        
    } catch (error) {
        console.error(`❌ Erro ao preencher formulário da linha ${dados.linha}:`, error.message);
        throw error;
    }
}

async function criarRelatorioExecucao(dadosProcessados, sucessos, erros, tempoExecucao) {
    const agora = new Date();
    const timestamp = agora.toISOString().replace(/[:.]/g, '-').slice(0, 19);
    
    const relatorio = {
        data_execucao: agora.toLocaleString('pt-BR'),
        tempo_execucao_minutos: Math.round(tempoExecucao / 60000 * 100) / 100,
        total_linhas: dadosProcessados.length,
        sucessos: sucessos,
        erros: erros,
        taxa_sucesso: `${((sucessos / dadosProcessados.length) * 100).toFixed(1)}%`,
        configuracoes_utilizadas: CONFIG,
        detalhes: dadosProcessados
    };
    
    const nomeArquivo = `relatorio_execucao_${timestamp}.json`;
    
    try {
        fs.writeFileSync(nomeArquivo, JSON.stringify(relatorio, null, 2), 'utf8');
        logComTimestamp(`📄 Relatório de execução salvo: ${nomeArquivo}`);
        return nomeArquivo;
    } catch (error) {
        console.error('❌ Erro ao salvar relatório:', error.message);
        return null;
    }
}

async function aguardarLogin(page) {
    logComTimestamp('⏳ Aguardando login manual...');
    
    // Aguarda mudança na URL ou elementos específicos da página logada
    try {
        await page.waitForFunction(() => {
            return window.location.href.includes('/Home') || 
                   window.location.href.includes('/Dashboard') ||
                   !window.location.href.includes('/Account/Login');
        }, { timeout: 120000 }); // 2 minutos
        
        logComTimestamp('✅ Login detectado com sucesso!');
        return true;
    } catch (error) {
        throw new Error('Timeout aguardando login manual. Verifique se fez login corretamente.');
    }
}

async function automatizarUnimed() {
    const inicioExecucao = Date.now();
    
    console.log('🚀 Iniciando Automação');
    console.log('⚠️ ATENÇÃO: Esta execução VAI SALVAR os lançamentos no sistema!');
    console.log('════════════════════════════════════════════════════════════');
    
    const todosOsDados = await lerPlanilhaExcel();
    logComTimestamp(`📊 Total de lançamentos a processar: ${todosOsDados.length}`);
    
    let sucessos = 0;
    let erros = 0;
    const resultadosDetalhados = [];
    
    const browser = await chromium.launch({ 
        headless: false,
        timeout: 60000
    });
    
    const page = await browser.newPage();
    page.setDefaultTimeout(CONFIG.TIMEOUT_NAVEGACAO);
    
    try {
        // ===== LOGIN MANUAL =====
        logComTimestamp('📝 Acessando página de login...');
        await page.goto('https://unimedcerrado.topsaude.com.br/TSNMVC/Account/Login');
        
        console.log('\n╔══════════════════════════════════════════════════════════════╗');
        console.log('║                    🔐 LOGIN MANUAL                           ║');
        console.log('╚══════════════════════════════════════════════════════════════╝');
        console.log('');
        console.log('🌐 A página de login foi aberta no navegador');
        console.log('📝 Por favor, faça seu login manualmente');
        console.log('⚠️  A automação continuará automaticamente após o login');
        console.log('');
        
        await aguardarLogin(page);
        
        logComTimestamp('🧭 Navegando para Lançamento Manual...');
        await page.getByRole('link', { name: ' Pagamento' }).click();
        await page.getByRole('link', { name: 'Lançamento Manual' }).click();
        
        // ===== PROCESSAMENTO DE CADA LINHA =====
        logComTimestamp('\n🔄 === INICIANDO PROCESSAMENTO DOS LANÇAMENTOS ===');
        
        const toolbarFrame = page.locator('#iframeasp').contentFrame()
            .locator('iframe[name="toolbarMvcToAsp"]').contentFrame();
        const incluirButton = toolbarFrame.getByRole('img', { name: 'Incluir' });
        
        for (let i = 0; i < todosOsDados.length; i++) {
            const dadosLinha = todosOsDados[i];
            const numeroLancamento = i + 1;
            
            logComTimestamp(`\n🔄 === PROCESSANDO LANÇAMENTO ${numeroLancamento} de ${todosOsDados.length} ===`);
            logComTimestamp(`📋 Dados: ${dadosLinha.cod_tipo_rubrica} | ${dadosLinha.cod_prestador} | ${dadosLinha.val_bruto}`);
            
            let tentativasRestantes = CONFIG.MAX_TENTATIVAS_ERRO;
            let sucessoLancamento = false;
            let resultadoValidacao = null;
            
            while (tentativasRestantes > 0 && !sucessoLancamento) {
                try {
                    // PRIMEIRO CLIQUE: Incluir para abrir/limpar formulário
                    logComTimestamp('➕ [1/4] Clicando no botão Incluir para abrir formulário...');
                    await incluirButton.click();
                    
                    // CRÍTICO: Aguarda o formulário ser limpo/preparado
                    const frameContent = page.locator('#iframeasp').contentFrame()
                        .locator('iframe[name="principal2"]').contentFrame();
                    await aguardarLimpezaFormulario(frameContent);
                    
                    // Preenche o formulário
                    logComTimestamp('📝 [2/4] Preenchendo formulário...');
                    await preencherFormulario(page, dadosLinha);
                    
                    // SEGUNDO CLIQUE: Incluir para salvar
                    logComTimestamp(`💾 [3/4] Clicando no botão Incluir para salvar lançamento ${numeroLancamento}...`);
                    await incluirButton.click();
                    
                    // VALIDAÇÃO: Verifica se apareceu mensagem de sucesso
                    logComTimestamp('🔍 [4/4] Validando se lançamento foi salvo com sucesso...');
                    resultadoValidacao = await verificarSucessoLancamento(page);
                    
                    if (resultadoValidacao.sucesso) {
                        sucessos++;
                        sucessoLancamento = true;
                        resultadosDetalhados.push({
                            linha: numeroLancamento,
                            status: 'SUCESSO',
                            dados: dadosLinha,
                            mensagem_sucesso: resultadoValidacao.mensagem,
                            numero_pagamento: resultadoValidacao.numeroPagamento,
                            tentativas_utilizadas: CONFIG.MAX_TENTATIVAS_ERRO - tentativasRestantes + 1,
                            timestamp: new Date().toLocaleString('pt-BR')
                        });
                        
                        logComTimestamp(`🎉 Lançamento ${numeroLancamento} CONFIRMADO com sucesso! Pagamento: ${resultadoValidacao.numeroPagamento}`);
                        
                    } else {
                        throw new Error(`Lançamento não foi salvo: ${resultadoValidacao.motivo}`);
                    }
                    
                } catch (error) {
                    tentativasRestantes--;
                    const tentativaAtual = CONFIG.MAX_TENTATIVAS_ERRO - tentativasRestantes;
                    
                    console.error(`❌ ERRO na tentativa ${tentativaAtual} do lançamento ${numeroLancamento}:`, error.message);
                    
                    if (tentativasRestantes > 0) {
                        logComTimestamp(`🔄 Tentando novamente... (${tentativasRestantes} tentativa(s) restante(s))`);
                    } else {
                        erros++;
                        resultadosDetalhados.push({
                            linha: numeroLancamento,
                            status: 'ERRO',
                            dados: dadosLinha,
                            erro: error.message,
                            resultado_validacao: resultadoValidacao,
                            tentativas_utilizadas: CONFIG.MAX_TENTATIVAS_ERRO,
                            timestamp: new Date().toLocaleString('pt-BR')
                        });
                        
                        const nomeScreenshot = `erro_lancamento_${numeroLancamento}_${Date.now()}.png`;
                        await page.screenshot({ 
                            path: nomeScreenshot, 
                            fullPage: true 
                        });
                        
                        logComTimestamp(`📸 Screenshot do erro salvo como: ${nomeScreenshot}`);
                        logComTimestamp('⏭️ Continuando para o próximo lançamento...');
                    }
                }
            }
        }
        
        // ===== RELATÓRIO FINAL =====
        const fimExecucao = Date.now();
        const tempoExecucao = fimExecucao - inicioExecucao;
        
        console.log(`\n🎉 === PROCESSAMENTO CONCLUÍDO ===`);
        logComTimestamp(`📊 RESULTADOS:`);
        logComTimestamp(`   ✅ Sucessos: ${sucessos}`);
        logComTimestamp(`   ❌ Erros: ${erros}`);
        logComTimestamp(`   📋 Total: ${todosOsDados.length}`);
        logComTimestamp(`   📈 Taxa de sucesso: ${((sucessos / todosOsDados.length) * 100).toFixed(1)}%`);
        logComTimestamp(`   ⏱️ Tempo total: ${Math.round(tempoExecucao / 60000 * 100) / 100} minutos`);
        
        const arquivoRelatorio = await criarRelatorioExecucao(resultadosDetalhados, sucessos, erros, tempoExecucao);
        
        if (arquivoRelatorio) {
            logComTimestamp(`📄 Relatório detalhado disponível em: ${arquivoRelatorio}`);
        }
        
        console.log('\n════════════════════════════════════════════════════════════');
        if (sucessos === todosOsDados.length) {
            console.log('🎉 PARABÉNS! Todos os lançamentos foram processados com sucesso!');
        } else if (sucessos > 0) {
            console.log('⚠️ Processamento concluído com alguns erros. Verifique o relatório.');
        } else {
            console.log('❌ Nenhum lançamento foi processado com sucesso. Verifique os erros.');
        }
        console.log('════════════════════════════════════════════════════════════');
        
        // ===== LOGOFF ANTES DE FECHAR =====
        logComTimestamp('🚪 Fazendo logoff do sistema...');
        try {
            await page.goto('https://unimedcerrado.topsaude.com.br/TSNMVC/TSNMVC/Account/Login?ssoUserLogon=S');
            logComTimestamp('✅ Logoff realizado com sucesso!');
        } catch (logoffError) {
            logComTimestamp(`⚠️ Erro durante logoff: ${logoffError.message}`);
            logComTimestamp('🔄 Continuando para fechar navegador...');
        }
        
    } catch (error) {
        console.error('❌ Erro crítico durante automação:', error.message);
        console.error('Stack trace:', error.stack);
        
        const timestampErro = Date.now();
        const nomeScreenshotCritico = `erro_critico_${timestampErro}.png`;
        
        try {
            await page.screenshot({ path: nomeScreenshotCritico, fullPage: true });
            logComTimestamp(`📸 Screenshot do erro crítico salvo como: ${nomeScreenshotCritico}`);
        } catch (screenshotError) {
            logComTimestamp('❌ Não foi possível salvar screenshot do erro crítico');
        }
        
        const logErro = {
            timestamp: new Date().toLocaleString('pt-BR'),
            erro: error.message,
            stack: error.stack,
            dados_processados_ate_erro: resultadosDetalhados
        };
        
        try {
            fs.writeFileSync(`log_erro_critico_${timestampErro}.json`, JSON.stringify(logErro, null, 2));
            logComTimestamp(`📄 Log do erro crítico salvo como: log_erro_critico_${timestampErro}.json`);
        } catch (logError) {
            logComTimestamp('❌ Não foi possível salvar log do erro crítico');
        }
        
        throw error;
        
    } finally {
        try {
            await browser.close();
            logComTimestamp('🏁 Navegador fechado');
        } catch (closeError) {
            logComTimestamp('⚠️ Erro ao fechar navegador:', closeError.message);
        }
    }
}

// ===== VERIFICAÇÃO DE SEGURANÇA ANTES DA EXECUÇÃO =====
async function verificacaoSeguranca() {
    console.log('🔒 === VERIFICAÇÃO DE SEGURANÇA ===');
    console.log('⚠️  Esta automação VAI SALVAR dados no sistema!');
    console.log('⚠️  Certifique-se de que todos os dados estão corretos!');
    console.log('🔐  O login será feito manualmente no navegador');
    console.log('════════════════════════════════════════════════════════════');
    
    if (!fs.existsSync('dados_lancamento.xlsx')) {
        throw new Error('❌ Arquivo dados_lancamento.xlsx não encontrado!');
    }
    
    try {
        const workbook = XLSX.readFile('dados_lancamento.xlsx');
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const dados = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        const dadosLimpos = dados.filter(linha => 
            linha && linha.some(celula => celula !== undefined && celula !== null && celula.toString().trim() !== '')
        );
        
        console.log(`📊 Preview da planilha (${dadosLimpos.length} linha(s)):`);
        console.log('┌─────────────────────────────────────────────────────────────┐');
        
        dadosLimpos.slice(0, 3).forEach((linha, index) => {
            const resumo = `${linha[0] || ''} | ${linha[1] || ''} | ${linha[3] || ''}`;
            console.log(`│ Linha ${index + 1}: ${resumo.padEnd(50)} │`);
        });
        
        if (dadosLimpos.length > 3) {
            console.log(`│ ... e mais ${dadosLimpos.length - 3} linha(s)                    │`);
        }
        console.log('└─────────────────────────────────────────────────────────────┘');
        
        console.log('\n🔐 LEMBRE-SE:');
        console.log('• Tenha suas credenciais do TopSaude prontas');
        console.log('• O navegador abrirá para login manual');
        console.log('• NÃO feche o navegador durante o processo');
        
    } catch (error) {
        console.error('❌ Erro ao ler preview da planilha:', error.message);
        throw error;
    }
}

// ===== EXECUÇÃO PRINCIPAL =====
async function main() {
    try {
        await verificacaoSeguranca();
        
        console.log('\n⏰ Iniciando em 5 segundos...');
        console.log('⏰ Pressione Ctrl+C para cancelar!');
        
        // Countdown de segurança
        for (let i = 5; i > 0; i--) {
            process.stdout.write(`⏰ ${i}... `);
            await new Promise(resolve => setTimeout(resolve, 1000));
        }
        
        console.log('\n🚀 INICIANDO AUTOMAÇÃO!\n');
        
        await automatizarUnimed();
        
        console.log('\n✅ Automação concluída com sucesso!');
        process.exit(0);
        
    } catch (error) {
        console.error('\n❌ Falha na automação:', error.message);
        console.error('📝 Verifique os arquivos de log e screenshots para mais detalhes');
        process.exit(1);
    }
}

// Captura sinais de interrupção para limpeza
process.on('SIGINT', () => {
    console.log('\n🛑 Automação interrompida pelo usuário');
    console.log('🧹 Fazendo limpeza...');
    process.exit(0);
});

process.on('SIGTERM', () => {
    console.log('\n🛑 Automação terminada externamente');
    process.exit(0);
});

process.on('unhandledRejection', (reason, promise) => {
    console.error('❌ Erro não tratado:', reason);
    process.exit(1);
});

// Executa o programa principal
main();