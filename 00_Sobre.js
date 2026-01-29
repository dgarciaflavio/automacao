/**
 * ============================================================================
 * 📘 MANUAL TÉCNICO ANALÍTICO - ORQUESTRADOR GERAL (SISTEMA INCA)
 * ============================================================================
 * @project  Orquestrador de Estoque, Empenhos, Distribuição e Inteligência
 * @author   Flavio Garcia Diniz
 * @version  15.0 (Edição Definitiva "Caixa Preta Aberta")
 * @date     2026-01-29
 * * ESTE DOCUMENTO DESCREVE EXAUSTIVAMENTE A LÓGICA, CÁLCULOS E FLUXOS DO SISTEMA.
 * NÃO HÁ RESUMOS. CADA MÓDULO É EXPLICADO EM SEU NÍVEL DE ENGENHARIA.
 * ============================================================================
 * * 🏗️ ARQUITETURA DE DADOS E CONEXÕES
 * ============================================================================
 * O sistema opera sob uma arquitetura de "Hub & Spoke" (Centro e Raios).
 * O script não processa dados isolados; ele cruza informações de 6 bancos de dados distintos
 * em tempo real para gerar uma "Única Fonte de Verdade".
 * * 1. FONTE DE DADOS GERAL (EMS - ID: 1s44YD...):
 * - A "Verdade Absoluta" financeira e logística extraída do ERP.
 * - Filtro de Otimização: O script aplica um filtro rígido (Data >= 2023) na memória RAM.
 * Dados anteriores a 2023 são descartados na leitura para evitar estouro de tempo (Timeout),
 * garantindo que o processamento foque apenas na gestão atual.
 * * 2. MATERIAIS (Remoto - ID: 1jXd...): Entrada manual da equipe de Almoxarifado.
 * 3. MEDICAMENTOS (Remoto - ID: 16_jA...): Entrada manual da equipe de Farmácia.
 * 4. CORREÇÃO EXTERNA (ID: 1r8l...): Base de auditoria para divergências.
 * 5. COMPILADOS (Local - ID: 1ZLe...): O Hub onde este script reside.
 * 6. PAINEL EQUIPE (ID: 1Rc5...): Saída de dados para os planejadores (Bianca, Katia, etc.).
 * * ============================================================================
 * 🧠 MOTOR DE INTELIGÊNCIA E LÓGICA (O CÉREBRO)
 * ============================================================================
 * * ----------------------------------------------------------------------------
 * A. O ALGORITMO DE STATUS UNIFICADO (_calcularStatusUnificado em 03_Helpers.js)
 * ----------------------------------------------------------------------------
 * Esta é a função mais crítica do sistema. Ela decide o estado de um item baseada
 * em 4 variáveis: Qtd Empenhada (E), Saída Oficial (SO), Saldo Físico (SF) e Flag Provisório (P).
 * * A hierarquia de decisão (IF/ELSE) é estrita e segue esta ordem:
 * * 1. RECEBIDO A MAIOR (Erro Grave):
 * - Lógica: SE (E > 0) E (SO > E).
 * - Significado: O sistema registra mais entregas do que o comprado. Bloqueia pagamento.
 * * 2. CONCLUÍDO (Sucesso):
 * - Lógica: SE (E > 0) E (SO == E).
 * - Significado: A entrega oficial bateu exatamente com o empenho. Processo encerrado.
 * * 3. ERRO DE CADASTRO (Recebido s/ Associação):
 * - Lógica: SE (E == 0) E (SO > 0).
 * - Significado: O item entrou no almoxarifado, mas alguém esqueceu de lançar o empenho na planilha.
 * * 4. ITEM FANTASMA (Solicitar Associação):
 * - Lógica: SE (E == 0) E (SO == 0).
 * - Significado: Item listado mas sem nenhuma movimentação ou registro válido.
 * * 5. LÓGICA HÍBRIDA (Oficial vs Físico/Provisório):
 * - O sistema prioriza a Nota Fiscal (Oficial). Porém, a mercadoria chega antes da nota.
 * - SE (Flag Provisório Existe) E (SO == 0):
 * - O sistema entra em "Modo Físico". Ele ignora que o oficial é zero.
 * - Sub-regra: SE (Saldo Físico <= 10% do Empenho): Status = "Resíduo 10%".
 * - Sub-regra: SE (Saldo Físico > 0): Status = "Recebimento Provisório".
 * - IMPORTANTE: Assim que (SO > 0), o sistema AUTOMATICAMENTE sai do modo provisório
 * e assume o status oficial, prevenindo duplicidade de contagem.
 * * 6. PENDÊNCIA E RESÍDUO TÉCNICO:
 * - SE (SO == 0) E (SF == E): Status = "Pendente" (Nada chegou).
 * - SE (SF > 10% de E): Status = "Pendente com Resíduo" (Chegou parte, falta muito).
 * - SE (SF > 0 e SF <= 10% de E): Status = "Resíduo 10%" (Considerado entregue contabilmente).
 * * ============================================================================
 * 📂 DETALHAMENTO ANALÍTICO POR MENU (FUNCIONALIDADES)
 * ============================================================================
 * * ----------------------------------------------------------------------------
 * MENU 1: CICLO COMPLETO (`04_Ciclo_Completo.js`)
 * ----------------------------------------------------------------------------
 * - Função: Orquestração síncrona de atualização.
 * - Segurança: Exige senha administrativa (armazenada em ScriptProperties) para evitar execução acidental.
 * - Fluxo de Dados:
 * 1. Leitura Global: Carrega ~50.000 linhas do EMS na RAM (filtradas por ano >= 2023).
 * 2. Injeção Remota: Envia os dados processados para as planilhas de Materiais e Medicamentos.
 * - Isso garante que as planilhas satélites vejam o status real antes da compilação.
 * 3. Compilação Reversa: Puxa os dados atualizados das satélites de volta para a Local.
 * 4. Sincronização de Estoque: Recalcula coberturas e sugestões de compra.
 * - Por que essa ordem? Para garantir integridade referencial. O local só é atualizado
 * depois que o remoto confirmou o recebimento dos dados globais.
 * * ----------------------------------------------------------------------------
 * MENU 2 & 3: PROCESSAMENTO REMOTO (`05_Materiais` e `06_Medicamentos`)
 * ----------------------------------------------------------------------------
 * - Diferença Crucial: Materiais lida com locais 'ALM', 'MAI', '5x5'. Medicamentos lida com 'FAR' e códigos numéricos.
 * - Cálculo de Atraso:
 * - Data Limite = Data Envio do Empenho + 10 dias corridos.
 * - SE (Hoje > Data Limite) E (Status != Concluído/Resíduo):
 * - O script calcula os dias de atraso e escreve "X dias e Y meses" na célula.
 * - Preservação de Dados:
 * - O script lê as anotações manuais (Colunas J a N) antes de limpar a aba.
 * - Ao reescrever os dados atualizados, ele "devolve" as anotações para as linhas corretas
 * usando uma Chave Única composta por (NúmeroEmpenho + CódigoItem).
 * * ----------------------------------------------------------------------------
 * MENU 6: DISTRIBUIÇÃO DE EQUIPE (`16_Distribuicao_Equipe.js`)
 * ----------------------------------------------------------------------------
 * - Lógica de Atribuição Dinâmica:
 * - Não existem "nomes fixos" no código (hardcoded).
 * - O script lê a aba "Config_Equipe". Se você mudar a família "Saneantes" de "Bianca" para "Katia" lá,
 * o script redireciona os itens na próxima execução automaticamente.
 * - Cálculo de CMA Histórico (Consumo Médio Ajustado):
 * - O sistema ignora a média simples. Ele analisa o histórico de 15 meses (anos 2022, 2023, 2025).
 * - Fórmula: (Soma das Saídas dos últimos 3 anos / 3) / 12 * 15.
 * - Objetivo: Suavizar a sazonalidade e projetar um consumo para 15 meses de segurança.
 * - Detecção de Conflitos:
 * - Se um item pertence à família X (Katia) mas o código específico está mapeado para Y (Rafaelle),
 * o script duplica o item nas duas abas, pinta de VERMELHO e adiciona nota: "⚠️ COMPARTILHADO".
 * * ----------------------------------------------------------------------------
 * MENU INTELIGÊNCIA: ANÁLISE DE TENDÊNCIA (`23_Analise_Tendencia.js`) **[NOVO]**
 * ----------------------------------------------------------------------------
 * - Objetivo: Detectar "Aceleração" ou "Frenagem" de consumo antes que o estoque acabe.
 * - Metodologia Matemática:
 * 1. Calcula o Consumo Real dos últimos 30 dias (baseado na data de movimentação global).
 * 2. Compara com o CMM (Média Histórica).
 * 3. Fórmula de Desvio: (Consumo30d - CMM) / CMM.
 * - Gatilhos de Alerta:
 * - SE Desvio > +30%: Diagnóstico "🔥 Aceleração Alta". (Risco de ruptura iminente).
 * - SE Desvio < -30%: Diagnóstico "❄️ Desaceleração". (Estoque parado/excesso).
 * - Caso contrário: "Estável".
 * * ----------------------------------------------------------------------------
 * MENU ESTOQUE: GESTÃO E SUGESTÃO (`08_Gestao_Estoque.js`)
 * ----------------------------------------------------------------------------
 * - Cálculo de Cobertura (Dias):
 * - Fórmula: Estoque Atual / (CMM / 30).
 * - Se CMM for 0: Retorna "Sem Consumo" ou "Zerado" (infinito técnico).
 * - Cálculo de Sugestão de Compra (Meta 6 Meses) **[ATUALIZADO]**:
 * - Meta de Estoque = CMM * 6.
 * - Sugestão = Meta de Estoque - Estoque Atual.
 * - Se o resultado for negativo (temos excesso), a sugestão é 0.
 * - Previsão de Esgotamento Projetada:
 * - Calcula a data futura onde o estoque chegará a zero SE a sugestão de compra for atendida.
 * - Fórmula: Hoje + ((EstoqueAtual + Sugestão) / ConsumoDiario).
 * * ----------------------------------------------------------------------------
 * MENU OPERACIONAL: OPERAÇÃO CONTINGÊNCIA (`22_Operacao_Contingencia.js`) **[ATUALIZADO]**
 * ----------------------------------------------------------------------------
 * - Objetivo: Relatório de crise para itens com risco imediato de falta.
 * - Critério de Seleção (Filtro Rígido):
 * - Saldo em Dias (Coluna I da aba dados) <= 91 dias.
 * - Lógica de Agrupamento:
 * - O script agrupa os itens pelo "Processo SEI".
 * - Tratativa de Exceção: Se o processo for vazio, "-", ou "0", o script renomeia para "Item sem processo"
 * e força o responsável para "Não mapeado".
 * - Enriquecimento Financeiro:
 * - O script busca na Base Global a "Última Entrada" (maior data) daquele item.
 * - Captura o Valor Unitário dessa entrada e insere no relatório para cálculo de custo de reposição.
 * - Status Visual:
 * - <= 30 dias: Status "CRÍTICO" (Vermelho).
 * - 31 a 91 dias: Status "ALERTA" (Amarelo).
 * * ----------------------------------------------------------------------------
 * MENU OPERACIONAL: STATUS REPORT / MUTIRÃO (`15_Mutirao.js`) **[ATUALIZADO]**
 * ----------------------------------------------------------------------------
 * - Funcionalidade: Ferramenta de trabalho para preenchimento de solicitações.
 * - Automação Financeira (OnEdit Simulado):
 * - O script injeta uma fórmula na Coluna E: `=IF(ISNUMBER(D2); C2*D2; 0)`.
 * - Isso permite que o usuário digite a Qtd Solicitada (Col D) e o Valor Total (Col E)
 * seja calculado instantaneamente pelo Sheets, sem precisar rodar o script novamente.
 * - Busca de Preço Inteligente:
 * - Varre a Base Global.
 * - Para cada código, encontra a entrada com a data mais recente.
 * - Preenche a Coluna C com esse "Último Preço Praticado".
 * - Geração de Documentos:
 * - Gera PDF (visualização limpa) e Excel (editável) contendo todas as colunas financeiras.
 * * ----------------------------------------------------------------------------
 * MENU BI: RELATÓRIOS FINANCEIROS E PERFORMANCE (`14_Relatorios_BI.js`)
 * ----------------------------------------------------------------------------
 * - Panorama Financeiro (Executivo):
 * - Solicita Ano Inicial e Final.
 * - Soma Empenhos (Passivo Total).
 * - Subtrai Entregas (Passivo Baixado).
 * - Calcula "Restos a Pagar" real (Passivo Líquido).
 * - Separa o que é Pendência Real do que é Resíduo Técnico (<10%).
 * - Performance de Fornecedores:
 * - Métrica: (Total de Itens Pendentes / Total de Itens Empenhados).
 * - Gera um ranking dos fornecedores com maior taxa de falha na entrega.
 * * ============================================================================
 * 🛡️ SEGURANÇA E AUDITORIA (`20_Logger_Hub.js` e `01_Config.js`)
 * ============================================================================
 * - Logger de Acesso:
 * - Monitora silenciosamente quem abre a planilha.
 * - Registra E-mail, Data e Hora na aba oculta "Log_Acesso_Hub".
 * - Possui autolimpeza (mantém apenas os últimos 2000 acessos).
 * - Validação de Conexões:
 * - Ao iniciar, o script tenta "tocar" em todas as 6 planilhas conectadas.
 * - Se algum ID estiver errado ou sem permissão, ele bloqueia a execução e alerta
 * exatamente qual planilha falhou, prevenindo erros em cascata.
 * * ============================================================================
 */
