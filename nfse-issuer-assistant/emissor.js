require('dotenv').config();
const { chromium } = require('playwright');
const xlsx = require('xlsx');
const readline = require('readline');
const fs = require('fs');
const path = require('path');

// =========================================================================
// CONFIGURAÇÕES (login vem do arquivo .env, CNPJs vêm do tomadores.json)
// Nenhum dado sensível fica escrito dentro do código.
// =========================================================================
const LOGIN_CREDENCIAIS = {
    usuario: process.env.NFSE_USER,
    senha: process.env.NFSE_PASSWORD
};

if (!LOGIN_CREDENCIAIS.usuario || !LOGIN_CREDENCIAIS.senha) {
    console.error('❌ [Erro]: Defina NFSE_USER e NFSE_PASSWORD no arquivo .env (veja .env.example).');
    process.exit(1);
}

let BANCO_CNPJ_TOMADORES = {};
try {
    BANCO_CNPJ_TOMADORES = JSON.parse(
        fs.readFileSync(path.join(__dirname, 'tomadores.json'), 'utf8')
    );
} catch (e) {
    console.error('❌ [Erro]: Arquivo tomadores.json não encontrado ou inválido (veja tomadores.example.json).');
    process.exit(1);
}

function fazerPerguntaTerminal(mensagem) {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    return new Promise(resolve => rl.question(mensagem, resposta => {
        rl.close();
        resolve(resposta);
    }));
}

function buscarEnderecoCelula(sheet, textoChave) {
    if (!sheet['!ref']) return null;
    const range = xlsx.utils.decode_range(sheet['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = xlsx.utils.encode_cell({ r: R, c: C });
            const cell = sheet[cellAddress];
            if (cell && cell.v && String(cell.v).toUpperCase().includes(textoChave.toUpperCase())) {
                return { r: R, c: C };
            }
        }
    }
    return null;
}

function buscarValorAoLado(sheet, textoChave) {
    const pos = buscarEnderecoCelula(sheet, textoChave);
    if (pos) {
        const proximaCelula = sheet[xlsx.utils.encode_cell({ r: pos.r, c: pos.c + 1 })];
        return proximaCelula ? proximaCelula.v : '';
    }
    return '';
}

async function rodarEmissorTotal() {
    const pastaExecutar = path.join(__dirname, 'executar');

    // Cria a pasta automaticamente se ela não existir
    if (!fs.existsSync(pastaExecutar)) {
        fs.mkdirSync(pastaExecutar);
    }

    // Lê os arquivos dentro da pasta 'executar'
    const arquivos = fs.readdirSync(pastaExecutar).filter(f => f.endsWith('.xlsx'));

    if (arquivos.length === 0) {
        console.error('❌ [Erro]: Nenhuma planilha .xlsx foi encontrada dentro da pasta "executar"!');
        console.log('👉 Coloque o rascunho da regional que você quer rodar lá dentro e tente novamente.');
        return;
    }

    // Pega o primeiro arquivo encontrado na pasta
    const nomeArquivoReal = arquivos[0];
    const caminhoCompletoPlanilha = path.join(pastaExecutar, nomeArquivoReal);

    console.log(`🤖 [Robô]: Encontrei o arquivo "${nomeArquivoReal}" na pasta "executar". Iniciando leitura...`);

    let workbook, sheetName, sheet;
    try {
        workbook = xlsx.readFile(caminhoCompletoPlanilha);
        sheetName = workbook.SheetNames[0];
        sheet = workbook.Sheets[sheetName];
    } catch (e) {
        console.error(`❌ [Erro]: Falha ao tentar ler o arquivo ${nomeArquivoReal}`);
        return;
    }

    // Procura a localidade de forma dinâmica na planilha lida
    let localidadeLida = String(buscarValorAoLado(sheet, 'Municipio:') || buscarValorAoLado(sheet, 'Estado:') || '').toUpperCase().trim();

    if (!localidadeLida) {
        if (buscarEnderecoCelula(sheet, 'GURUPI')) localidadeLida = 'GURUPI';
        else if (buscarEnderecoCelula(sheet, 'PORTO NACIONAL')) localidadeLida = 'PORTO NACIONAL';
        else if (buscarEnderecoCelula(sheet, 'GOIÂNIA')) localidadeLida = 'GOIÂNIA';
        else if (buscarEnderecoCelula(sheet, 'APARECIDA DE GOIÂNIA')) localidadeLida = 'APARECIDA DE GOIÂNIA';
        else if (buscarEnderecoCelula(sheet, 'TOCANTINS') || buscarEnderecoCelula(sheet, 'TO')) localidadeLida = 'TO_PADRAO';
        else if (buscarEnderecoCelula(sheet, 'GOIAS') || buscarEnderecoCelula(sheet, 'GOIÁS')) localidadeLida = 'GO';
        else if (buscarEnderecoCelula(sheet, 'MATO GROSSO DO SUL') || buscarEnderecoCelula(sheet, 'MS')) localidadeLida = 'MS';
        else if (buscarEnderecoCelula(sheet, 'MATO GROSSO') || buscarEnderecoCelula(sheet, 'MT')) localidadeLida = 'MT';
        else if (buscarEnderecoCelula(sheet, 'RIO GRANDE DO SUL') || buscarEnderecoCelula(sheet, 'RS')) localidadeLida = 'RS';
    }

    let cnpjTomadorFinal = BANCO_CNPJ_TOMADORES[localidadeLida] || '';

    if (!cnpjTomadorFinal) {
        if (localidadeLida.includes('GO')) cnpjTomadorFinal = BANCO_CNPJ_TOMADORES['GO'];
        if (localidadeLida.includes('TO')) cnpjTomadorFinal = BANCO_CNPJ_TOMADORES['TO_PADRAO'];
        if (localidadeLida.includes('MS')) cnpjTomadorFinal = BANCO_CNPJ_TOMADORES['MS'];
        if (localidadeLida.includes('MT')) cnpjTomadorFinal = BANCO_CNPJ_TOMADORES['MT'];
        if (localidadeLida.includes('RS')) cnpjTomadorFinal = BANCO_CNPJ_TOMADORES['RS'];
    }

    console.log(`✅ [Planilha Processada]:\n ➔ Arquivo Lido: ${nomeArquivoReal}\n ➔ Localidade Identificada: ${localidadeLida}`);

    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext();
    const page = await context.newPage();

    try {
        console.log('🚀 [Tela 1/Login]: Acessando o portal nacional...');
        await page.goto('https://nfse.gov.br', { waitUntil: 'load' });
        await page.fill('#CnpjCpf', String(LOGIN_CREDENCIAIS.usuario).replace(/\D/g, ''));
        await page.fill('#Senha', LOGIN_CREDENCIAIS.senha);
        await page.click('button:has-text("Entrar")');
        await page.waitForURL('**/PortalContribuinte/Home**', { timeout: 30000 });
        console.log('➔ Login efetuado com sucesso!');

        console.log('🚀 [Tela 2 e 3/Menu]: Navegando até a emissão completa...');
        await page.click('.fa-plus-circle, [title*="Emissão"]');
        await page.waitForSelector('text=Emissão completa');
        await page.click('text=Emissão completa');
        await page.waitForLoadState('networkidle');

        console.log('🚀 [Aba Pessoas/Tomador]: Aplicando regras do cliente...');
        const cnpjLimpo = String(cnpjTomadorFinal).replace(/\D/g, '');
        if (cnpjLimpo) {
            await page.fill('#VisualizacaoTomador_CnpjCpf', cnpjLimpo);
            await page.click('button:has-text("Avançar")');
        }
        await page.waitForLoadState('networkidle');

        console.log('🚀 [Aba Serviço]: Preenchendo localidade...');
        await page.click('button:has-text("Avançar")');
        await page.waitForLoadState('networkidle');

        console.log('🚀 [Aba Valores]: Formulário carregado.');
        console.log('\n=============================================================');
        console.log('🟩 [ROBÔ PAUSADO PARA INTERAÇÃO HUMANA]');
        console.log('➔ Olhe para o campo "Alíquota" na tela do navegador do governo.');
        console.log('=============================================================\n');

        const aliquotaDigitada = await fazerPerguntaTerminal('Digite apenas o número da Alíquota que apareceu no sistema (ex: 5 ou 3): ');

        const posISSQN = buscarEnderecoCelula(sheet, 'ISSQN');
        if (posISSQN) {
            const celulaPorcentagem = xlsx.utils.encode_cell({ r: posISSQN.r, c: posISSQN.c + 1 });
            sheet[celulaPorcentagem] = { t: 'n', v: parseFloat(aliquotaDigitada) / 100, z: '0%' };
            xlsx.writeFile(workbook, caminhoCompletoPlanilha);
            console.log(`💾 [Excel Atualizado]: Alíquota de ${aliquotaDigitada}% salva na planilha com sucesso!`);
        }

        await fazerPerguntaTerminal('Confira os dados finais. Pressione ENTER para finalizar e fechar...');

    } catch (error) {
        console.error('❌ [Erro durante a execução]:', error);
    } finally {
        await browser.close();
    }
}

rodarEmissorTotal();
