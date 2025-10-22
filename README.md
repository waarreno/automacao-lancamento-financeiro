# 🤖 Automação de Lançamento Financeiro Manual

Sistema automatizado para processamento em lote de lançamentos financeiros no TopSaude (Unimed Cerrado).

![Versão](https://img.shields.io/badge/versão-1.4-blue)
![Node](https://img.shields.io/badge/node-%3E%3D18.0.0-green)
![Status](https://img.shields.io/badge/status-ativo-success)

---

## 📋 Índice

- [Visão Geral](#-visão-geral)
- [Funcionalidades](#-funcionalidades)
- [Requisitos](#-requisitos)
- [Instalação](#-instalação)
- [Uso](#-uso)
- [Estrutura da Planilha](#-estrutura-da-planilha)
- [Estrutura do Projeto](#-estrutura-do-projeto)
- [Configurações](#️-configurações)
- [Logs e Relatórios](#-logs-e-relatórios)
- [Tratamento de Erros](#-tratamento-de-erros)
- [Segurança](#-segurança)
- [Troubleshooting](#-troubleshooting)
- [Autor](#-autor)

---

## 🎯 Visão Geral

Esta automação foi desenvolvida para otimizar o processo de lançamento financeiro manual no sistema TopSaude, permitindo o processamento em lote de múltiplos lançamentos através de uma planilha Excel.

O sistema:
- ✅ Lê dados de uma planilha Excel
- ✅ Valida as informações antes do processamento
- ✅ Preenche automaticamente os formulários do sistema
- ✅ Verifica o sucesso de cada lançamento
- ✅ Gera relatórios detalhados
- ✅ Captura screenshots em caso de erro
- ✅ Implementa retry automático

---

## ⚡ Funcionalidades

### Principais Recursos

- **Processamento em Lote**: Processa múltiplos lançamentos de uma só vez
- **Validação Inteligente**: Valida dados antes do envio
- **Login Manual Seguro**: Login realizado manualmente pelo usuário
- **Retry Automático**: Tenta novamente em caso de erro (até 3 tentativas)
- **Verificação de Sucesso**: Confirma cada lançamento antes de prosseguir
- **Relatórios Detalhados**: Gera relatório completo da execução
- **Screenshots de Erro**: Captura evidências visuais quando há falhas
- **Node.js Portable**: Funciona sem instalação global do Node.js
- **Logoff Automático**: Encerra sessão ao finalizar

---

## 📦 Requisitos

### Sistema Operacional
- Windows 10 ou superior
- Conexão com internet (apenas para instalação inicial)

### Acesso
- Credenciais válidas do sistema TopSaude (Unimed Cerrado)
- Permissões para realizar lançamentos financeiros

### Hardware Recomendado
- 4GB de RAM ou mais
- 500MB de espaço em disco livre

---

## 🚀 Instalação

### Passo 1: Download
Baixe todos os arquivos do projeto e extraia para uma pasta de sua preferência.

### Passo 2: Instalação das Dependências
Execute o arquivo `Instalar_Dependencias.bat`:

```batch
Instalar_Dependencias.bat
```

Este script irá:
1. Baixar o Node.js portable (versão mais recente)
2. Instalar as bibliotecas necessárias (Playwright e XLSX)
3. Baixar o navegador Chromium
4. Configurar o ambiente completo

**⚠️ Importante**: Este processo requer conexão com internet e pode levar alguns minutos.

### Passo 3: Preparar a Planilha
Coloque sua planilha `dados_lancamento.xlsx` na pasta `projeto/`.

---

## 📖 Uso

### Execução Básica

1. **Prepare sua planilha** com os dados dos lançamentos
2. **Execute** o arquivo `Executar.bat`
3. **Aguarde** o navegador abrir automaticamente
4. **Faça login** manualmente no sistema TopSaude
5. **Aguarde** o processamento automático dos lançamentos

### Comandos Disponíveis

```batch
# Executar a automação
Executar.bat

# Reinstalar dependências (se necessário)
Instalar_Dependencias.bat
```

### Fluxo de Execução

```
Início
  ↓
Verificação de Segurança
  ↓
Preview da Planilha (primeiras 3 linhas)
  ↓
Countdown de 5 segundos
  ↓
Abertura do Navegador
  ↓
Aguarda Login Manual do Usuário
  ↓
Processamento Automático dos Lançamentos
  ↓
Geração de Relatório
  ↓
Logoff Automático
  ↓
Fim
```

---

## 📊 Estrutura da Planilha

### Formato do Arquivo
- **Nome**: `dados_lancamento.xlsx`
- **Localização**: pasta `projeto/`
- **Formato**: Excel (.xlsx)

### Colunas Obrigatórias

| Coluna | Nome | Tipo | Formato | Obrigatório | Exemplo |
|--------|------|------|---------|-------------|---------|
| A | Código Tipo Rubrica | Número | - | ✅ Sim | 12 |
| B | Código Prestador | Número | - | ✅ Sim | 0880000912 |
| C | Mês/Ano Referência | Texto | MM/AAAA | ✅ Sim | 10/2025 |
| D | Valor Bruto | Número | Decimal | ✅ Sim | 1500.50 |
| E | Data Pagamento Prevista | Texto | DD/MM/AAAA | ✅ Sim | 15/11/2025 |
| F | Observações | Texto | - | ✅ Sim | Texto livre |

### Exemplo de Planilha

```
| Cód Rubrica | Cód Prestador | Mês/Ano | Valor  | Data Pgto  | Observações    |
|-------------|---------------|---------|--------|------------|----------------|
| 12          | 0880000912    | 10/2025 | -1500  | 15/11/2025 | Pagamento ref  |
| 12          | 0880000912    | 10/2025 | -2300.5| 15/11/2025 | Ajuste mensal  |
| 12          | 0880000912    | 10/2025 | -800   | 15/11/2025 | Sem referência |
```

### Validações Automáticas

O sistema valida automaticamente:
- ✅ Campos obrigatórios preenchidos
- ✅ Formato correto de datas (DD/MM/AAAA)
- ✅ Formato correto de mês/ano (MM/AAAA)
- ✅ Valores numéricos válidos (com sinal negativo em caso de débito - rubrica=12)
- ✅ Linhas vazias são ignoradas

---

## 📁 Estrutura do Projeto

```
projeto/
│
├── 📄 Executar.bat                         # Script para executar a automação (Windows-1252)
├── 📄 Instalar_Dependencias.bat            # Script para instalação inicial (Windows-1252)
├── 📄 .gitignore                           # Arquivos ignorados pelo Git
├── 📄 package.json                         # Configurações do Node.js
├── 📄 README.md                            # Este arquivo
│
├── 📁 projeto/
│   ├── 📄 script.js                        # Script principal da automação
│   ├── 📊 dados_lancamento.xlsx            # Planilha com os dados (criar)
│   ├── 📊 ConverterJSONParaExcel.xlsm      # Planilha para visualização de relatório .json
├── 📁 node_modules/                        # Dependências (gerado)
│   └── 📄 package.json                     # Configurações locais
│
├── 📁 node-portable/                       # Node.js portable (gerado)
│   ├── node.exe
│   ├── npm.cmd
│   └── ...
│
└── 📁 Outputs (gerados durante execução)/
    ├── 📄 relatorio_execucao_*.json        # Relatórios
    ├── 🖼️ erro_lancamento_*.png            # Screenshots de erros
    └── 📄 log_erro_critico_*.json          # Logs de erros críticos
```

---

## ⚙️ Configurações

### Configurações Globais (script.js)

Você pode ajustar os timeouts editando o objeto `CONFIG` no início do `script.js`:

```javascript
const CONFIG = {
    TIMEOUT_NAVEGACAO: 30000,           // Timeout de navegação (30s)
    MAX_TENTATIVAS_ERRO: 3,             // Tentativas em caso de erro
    TIMEOUT_MENSAGEM_SUCESSO: 15000,    // Aguarda mensagem de sucesso (15s)
    TIMEOUT_LIMPEZA_FORMULARIO: 3000,   // Aguarda limpeza do form (3s)
    TIMEOUT_CARREGAMENTO_CAMPO: 4000    // Aguarda carregamento campo (4s)
};
```

### Personalização

Para modificar comportamentos específicos, edite as funções em `script.js`:
- `preencherFormulario()` - Lógica de preenchimento
- `verificarSucessoLancamento()` - Validação de sucesso
- `validarPlanilha()` - Regras de validação

---

## 📝 Logs e Relatórios

### Relatório de Execução

Gerado automaticamente ao final: `relatorio_execucao_YYYYMMDD_HHMMSS.json`

Contém:
- ✅ Resumo geral (sucessos/erros/taxa de sucesso)
- ✅ Tempo total de execução
- ✅ Detalhes de cada lançamento processado
- ✅ Números de pagamento gerados
- ✅ Erros encontrados com descrições

### Logs no Console

Durante a execução, o sistema exibe logs detalhados:
```
[20/10/2025 14:30:15] 📊 Lendo planilha Excel...
[20/10/2025 14:30:16] ✅ Planilha lida com sucesso! 10 linha(s) encontrada(s)
[20/10/2025 14:30:17] 🔍 Validando dados da planilha...
[20/10/2025 14:30:17] ✅ Validação concluída: 10 linha(s) válida(s)
```

### Screenshots de Erro

Capturados automaticamente quando há falhas:
- `erro_lancamento_X_TIMESTAMP.png` - Erros em lançamentos específicos
- `erro_critico_TIMESTAMP.png` - Erros críticos do sistema

---

## ⚠️ Tratamento de Erros

### Sistema de Retry

- **Tentativas automáticas**: Até 3 tentativas por lançamento
- **Intervalo**: Aguarda limpeza do formulário entre tentativas
- **Continuidade**: Erros não param todo o processamento

### Tipos de Erro

| Tipo | Ação |
|------|------|
| Validação | Para antes de iniciar |
| Timeout | Retry automático |
| Campo não encontrado | Screenshot + log + retry |
| Erro crítico | Screenshot + log JSON + encerramento |

### Recuperação

Em caso de erro crítico:
1. Verifique o arquivo `log_erro_critico_*.json`
2. Analise o screenshot `erro_critico_*.png`
3. Corrija os dados na planilha
4. Remove linhas já processadas (se necessário)
5. Execute novamente

---

## 🔒 Segurança

### Proteções Implementadas

✅ **Login Manual**: Credenciais não são armazenadas no código
✅ **Verificação de Segurança**: Preview antes de processar
✅ **Countdown**: 5 segundos para cancelar (Ctrl+C)
✅ **Validação**: Dados verificados antes do envio
✅ **Logoff Automático**: Encerra sessão ao finalizar
✅ **Sem armazenamento**: Credenciais nunca são salvas

### Boas Práticas

- ⚠️ Nunca compartilhe suas credenciais
- ⚠️ Revise a planilha antes de executar
- ⚠️ Mantenha backups dos dados
- ⚠️ Execute em ambiente seguro
- ⚠️ Não compartilhe screenshots de erro (podem conter dados sensíveis)

### .gitignore

O arquivo `.gitignore` está configurado para proteger:
- Planilhas com dados
- Screenshots
- Logs e relatórios
- Credenciais
- Node.js portable

---

## 🔧 Troubleshooting

### Problema: "Node.js portable não encontrado"
**Solução**: Execute `Instalar_Dependencias.bat` novamente

### Problema: "Planilha dados_lancamento.xlsx não encontrada"
**Solução**: Certifique-se de que a planilha está na pasta `projeto/`

### Problema: "Erro ao baixar Node.js"
**Solução**: 
- Verifique sua conexão com internet
- Execute como administrador
- Verifique firewall/antivírus

### Problema: Timeout ao aguardar mensagem de sucesso
**Solução**: 
- Aumente `TIMEOUT_MENSAGEM_SUCESSO` no script.js
- Verifique se o sistema está respondendo normalmente

### Problema: Campos não estão sendo preenchidos
**Solução**:
- Aumente `TIMEOUT_CARREGAMENTO_CAMPO`
- Verifique se os seletores no código ainda estão corretos
- Analise o screenshot de erro

### Problema: Muitos erros de validação
**Solução**:
- Revise o formato das datas (DD/MM/AAAA)
- Verifique o formato do mês/ano (MM/AAAA)
- Confirme que campos obrigatórios estão preenchidos

### Problema: Caracteres estranhos ou erros com acentos nos arquivos .bat
**Solução**: 
- Os arquivos `.bat` devem estar em encoding **Windows-1252 (ANSI)**
- O comando `chcp 65001` no início do arquivo ajuda com a exibição
- Não use UTF-8 com BOM em arquivos .bat - isso causa erros
- Se editou os arquivos, salve-os com encoding Windows-1252/ANSI

### Problema: Navegador não abre
**Solução**:
- Execute `npx playwright install chromium` manualmente
- Verifique se há espaço em disco suficiente
- Reinstale as dependências

---

## 👨‍💻 Autor

**Wárreno Hendrick Costa Lima Guimarães**

Coordenador de Contas Médicas

- Versão: 1.4
- Ano: 2025

---

## 📄 Licença

Este projeto é de uso interno e privado. Todos os direitos reservados.

**UNLICENSED** - Não licenciado para distribuição ou uso externo.

---

## 🔄 Histórico de Versões

### v1.4 (Atual)
- ✨ Node.js portable implementado
- ✨ Sistema de retry aprimorado
- ✨ Validação de planilha melhorada
- ✨ Relatórios mais detalhados

### v1.3
- ✨ Verificação de sucesso implementada
- ✨ Screenshots de erro automáticos

### v1.2
- ✨ Sistema de logs com timestamp
- ✨ Tratamento de erros aprimorado

### v1.1
- ✨ Suporte para leitura de planilha Excel
- ✨ Validação de dados

### v1.0
- 🎉 Versão inicial

---

## 📞 Suporte

Em caso de dúvidas ou problemas:

1. Consulte a seção [Troubleshooting](#-troubleshooting)
2. Verifique os logs e relatórios gerados
3. Analise os screenshots de erro
4. Revise a estrutura da planilha

---

## ⚡ Dicas de Performance

- **Processamento em lote**: Processe até 50 lançamentos por vez
- **Horários**: Evite horários de pico do sistema
- **Internet**: Use conexão estável
- **Hardware**: Feche aplicações desnecessárias durante a execução

---

## 🎯 Roadmap Futuro

Possíveis melhorias futuras:
- [ ] Interface gráfica
- [ ] Modo de teste (dry-run)
- [ ] Agendamento de execução
- [ ] Dashboard de estatísticas
- [ ] Exportação de relatórios em Excel
- [ ] Notificações por email
- [ ] Modo headless (sem interface gráfica)

---

**🤖 Automação desenvolvida com foco em segurança, confiabilidade e facilidade de uso.**

*Última atualização: Outubro/2025*

Feito com ❤️ para a área de Contas Médicas da Unimed Cerrado