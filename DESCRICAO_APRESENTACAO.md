# Yuta – Central de Processos  
## Descrição para Apresentação

---

### O que é o Yuta?

**Yuta** é um sistema desktop de **automação de processos** voltado à área de faturamento e gestão portuária. Ele centraliza em uma única tela as principais rotinas do dia a dia: faturamentos (vários tipos), registro de ponto, criação de pastas, geração de relatórios e envio de e-mails, reduzindo retrabalho e erros manuais.

---

### Para quem é?

- Equipes de **faturamento** que trabalham com planilhas Excel, PDFs (OGMO) e documentos Word  
- Empresas que precisam padronizar processos de **vigia**, **de acordo** e **São Sebastião**  
- Quem usa **OneDrive** ou pastas em rede e quer um fluxo único e configurável  

---

### Principais funcionalidades

| Recurso | Descrição |
|--------|-----------|
| **Faturamento (Normal)** | Preenche planilhas a partir do arquivo do navio e do PDF OGMO: FRONT VIGIA, REPORT VIGIA, NF, geração de PDF e atualização da planilha de controle. |
| **Faturamento (Atípico)** | Fluxo para períodos atípicos: lê datas, períodos e valores do RESUMO do navio e replica no REPORT VIGIA com ordem correta. |
| **Faturamento São Sebastião** | Processa PDFs OGMO (incluindo múltiplos arquivos, ex.: Sea Side): extrai dados, identifica cliente/porto (Wilson Sons, Sea Side PSS, Aquarius etc.) e preenche o modelo correto. |
| **De Acordo** | Gera faturamento “De Acordo” com preenchimento da FRONT, regras por cliente (Unimar, Delta, North Star) e atualização da planilha de controle. |
| **Fazer Ponto** | Copia período (06h, 12h, 18h, 00h) de uma data de referência para a data escolhida, respeitando domingos e feriados. |
| **Desfazer Ponto** | Remove um período já lançado e recalcula totais do dia e total geral. |
| **Criar Pasta** | Cria a estrutura de pasta do cliente/navio/DN com sugestão de próximo DN e lista de clientes. |
| **Relatório** | Módulo preparado para geração de relatórios. |
| **Configurações** | Define o caminho base de faturamentos (manual ou auto-detecção OneDrive/SANPORT). |
| **Preview e rascunho de e-mail** | Pré-visualização em PDF antes de gerar e criação de rascunho no Outlook com anexos e texto padrão. |

---

### Benefícios

- **Menos trabalho manual**: menos copiar/colar entre planilhas e menos digitação repetida.  
- **Padronização**: mesmo fluxo para todos os tipos de faturamento e clientes.  
- **Rastreabilidade**: log em tempo real na tela e atualização da planilha de controle.  
- **Integração**: Excel (xlwings), Word (recibos), PDF (leitura e geração), Outlook (rascunho).  
- **Flexibilidade**: funciona com arquivos na rede e OneDrive; configuração por máquina.  

---

### Tecnologias

- **Interface**: aplicação desktop em Python (Tkinter).  
- **Planilhas**: Excel via xlwings e openpyxl.  
- **Documentos**: Word (COM) e PDF (pdfplumber, pdf2image, OCR quando necessário).  
- **Configuração**: JSON (caminho base, auto-detecção).  
- **Opcional**: API FastAPI para uso via navegador.  

---

### Em uma frase

**Yuta é a central de processos que automatiza faturamentos, ponto e rotinas de escritório da área portuária, em uma única tela, com preview, controle e integração com Excel, PDF e Outlook.**

---

*Documento gerado para apoio a apresentações do projeto Yuta.*
