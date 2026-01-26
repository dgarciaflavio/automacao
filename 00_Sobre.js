/**
 * ============================================================================
 * 📘 DOCUMENTAÇÃO TÉCNICA - ORQUESTRADOR GERAL (SISTEMA DE GESTÃO INCA)
 * ============================================================================
 * * @project  Orquestrador de Estoque, Empenhos e Distribuição
 * @author   Flavio Garcia Diniz
 * @version  12.0 (Documentação Definitiva "Sem Economia de Palavras")
 * @date     2026-01-16
 * * ============================================================================
 * 🎯 VISÃO GERAL DO SISTEMA
 * ============================================================================
 * Este ecossistema de automação (Orquestrador) atua como o "Cérebro Central" 
 * da logística. Ele não apenas processa dados, mas conecta 6 bases de dados 
 * distintas (Planilhas Google) para garantir que a informação de Materiais, 
 * Medicamentos, Estoque e Financeiro esteja sincronizada em tempo real.
 * * O sistema foi construído em arquitetura modular (Blocos 01 a 20) para 
 * permitir manutenção isolada sem quebrar o todo.
 * * ============================================================================
 * 📂 DETALHAMENTO PROFUNDO DOS MÓDULOS (ARQUIVOS)
 * ============================================================================
 * * ----------------------------------------------------------------------------
 * 1. 01_Config.js (O Mapa do Tesouro)
 * ----------------------------------------------------------------------------
 * - **Função:** É o arquivo mais crítico do sistema. Ele armazena as chaves de acesso (IDs)
 * e configurações globais. Se uma planilha mudar, é aqui que corrigimos.
 * - **Fontes de Dados Conectadas (IDs Reais):**
 * 1. **Materiais:** `1jXd4uEnyGZvLv4ozDfFi5ZMlumw0TvleGdMKlgGcPtU`
 * - Onde a equipe de Materiais lança os empenhos manuais.
 * 2. **Medicamentos:** `16_jA8i4zOKqgXDUdOyelrE0zMTGR27RaYQjZdlSInOE`
 * - Onde a equipe de Farmácia lança seus controles.
 * 3. **Fonte de Dados Geral (EMS):** `1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA`
 * - A "Verdade Absoluta" extraída do sistema ERP (EMS). Contém todas as entradas oficiais.
 * 4. **Correção Externa:** `1r8lYhlCecGTKlx7hSM3Fj6KFQqyt1JYKTfQ7HGGj6Zc`
 * - Usada para auditoria de divergências entre o sistema e o controle manual.
 * 5. **Compilados (Local):** `1ZLebBqhR1bMZgrnr_dfXikyIY22oi0B2pqXDz1UdRZM`
 * - A planilha onde este script roda. É o "Hub" que recebe tudo.
 * 6. **Painel de Equipe:** `1Rc5fUr-nP3g8SU9083Y-N83QZtg2wMCMlmlwcXJYl9k`
 * - Planilha externa onde Bianca, Katia, Leonardo, etc., recebem suas tarefas.
 * - **Paleta de Cores:** Define hexadecimalmente as cores de status (ex: Pendente = #f4cccc).
 * * ----------------------------------------------------------------------------
 * 2. 02_Menu.js (A Interface)
 * ----------------------------------------------------------------------------
 * - **Função:** Cria o menu visual "🚀 Orquestrador Geral" na barra superior.
 * - **Estrutura:** Organiza as 15+ funções do sistema em categorias lógicas para o usuário:
 * - Execução Master (Ciclo Completo).
 * - Processamentos Individuais (Remoto vs Local).
 * - Inteligência (BI, Snapshots).
 * - Operacional (Relatórios de Atas, PDF).
 * * ----------------------------------------------------------------------------
 * 3. 03_Helpers.js (O Motor Lógico & Filtro)
 * ----------------------------------------------------------------------------
 * - **Função:** Contém a "Inteligência" matemática do sistema.
 * - **Destaque: `obterDadosEntradasGlobal()`**
 * - Conecta na planilha Fonte de Dados (`1s44YD...`).
 * - **Otimização:** Aplica um filtro de data (`>= 2023`). Dados anteriores a este ano
 * são descartados da memória RAM instantaneamente, garantindo performance e evitando
 * estouro de tempo limite.
 * - **Destaque: `_calcularStatusUnificado()` (As 8 Regras de Ouro)**
 * Esta função decide o destino de cada empenho:
 * 1. **Recebido a Maior:** Se (Qtd Entregue > Qtd Empenho). Erro grave.
 * 2. **Concluído:** Se (Qtd Entregue == Qtd Empenho). Sucesso.
 * 3. **Falta Associar EMS:** Se (Qtd Empenho == 0) mas (Qtd Entregue > 0). Erro de cadastro.
 * 4. **Solicitar Associação:** Se não tem empenho nem entrega. Item fantasma.
 * 5. **Recebimento Provisório:** Se consta na aba manual de provisórios E a entrega oficial é ZERO.
 * 6. **Pendente:** Se nada foi entregue.
 * 7. **Pendente com Resíduo:** Falta entregar, e o saldo é relevante (> 10%).
 * 8. **Resíduo 10%:** Falta entregar, mas é "mixaria" (Saldo <= 10%). Considera-se entregue.
 * * ----------------------------------------------------------------------------
 * 4. 04_Ciclo_Completo.js (O Maestro Seguro)
 * ----------------------------------------------------------------------------
 * - **Função:** Executa tudo em ordem cronológica correta.
 * - **Segurança (Senha):** Antes de iniciar, exige a senha (`inca2026`). Isso impede
 * execuções acidentais por usuários não autorizados.
 * - **Sequência de Eventos:**
 * 1. **Leitura Otimizada:** Carrega a Fonte Global (Filtrada 2023+).
 * 2. **Escrita Remota 1:** Atualiza a planilha de Materiais (`1jXd...`).
 * 3. **Escrita Remota 2:** Atualiza a planilha de Medicamentos (`16_jA...`).
 * 4. **Compilação:** Puxa os dados das duas remotas de volta para a Local.
 * 5. **Estoque:** Sincroniza e recálcula a aba "Cont.Estoque".
 * * ----------------------------------------------------------------------------
 * 5. 05_Materiais.js (Processamento Remoto)
 * ----------------------------------------------------------------------------
 * - **Alvo:** Planilha de Materiais (`1jXd...`).
 * - **Mecanismo:**
 * - Lê os empenhos manuais da aba "Empenhos Enviados".
 * - Cruza com a memória do EMS (Entradas Globais).
 * - Verifica a aba "Rec.Provisorio" local daquela planilha.
 * - Preserva anotações (colunas de Obs) feitas pela equipe.
 * - Filtra Locais: Só processa itens de 'ALM', 'MAI' ou '5x5'.
 * - **Matemática:** Usa `Math.round()` agressivamente para evitar que 14.00000001
 * seja diferente de 14.
 * * ----------------------------------------------------------------------------
 * 6. 06_Medicamentos.js (Processamento Remoto)
 * ----------------------------------------------------------------------------
 * - **Alvo:** Planilha de Medicamentos (`16_jA...`).
 * - **Diferença:** Foca em itens com local "FAR" ou códigos puramente numéricos.
 * - **Visual:** Aplica formatação condicional (cores) diretamente na planilha de destino
 * para que o farmacêutico veja os atrasos em vermelho instantaneamente.
 * * ----------------------------------------------------------------------------
 * 7. 07_Compilacao_Local.js (O Funil)
 * ----------------------------------------------------------------------------
 * - **Função:** Traz a "Verdade" de volta para casa.
 * - **Fluxo:** Vai até as planilhas remotas (Mat e Med), copia o que foi processado
 * e cola na aba "Compilados" desta planilha.
 * - **Validação Dupla:** Re-executa a lógica de status localmente. Isso garante que,
 * mesmo se alguém mexer manualmente na planilha remota, o Painel Central (Compilados)
 * sempre mostrará o status calculado matematicamente correto.
 * * ----------------------------------------------------------------------------
 * 8. 08_Gestao_Estoque.js (Cérebro de Suprimentos)
 * ----------------------------------------------------------------------------
 * - **Fontes:** Lê as planilhas remotas para saber quais empenhos estão "vivos" e a
 * aba local "dados" para pegar Estoque Atual e CMM (Consumo Médio Mensal).
 * - **Cálculos Avançados:**
 * 1. **Cobertura (Dias):** Estoque Atual / (CMM / 30).
 * 2. **Previsão de Esgotamento:** Data de Hoje + Dias de Cobertura.
 * 3. **Sugestão de Compra (Regra 6 Meses):**
 * - Meta = CMM * 6.
 * - Sugestão = Meta - Estoque Atual. (Se negativo, é zero).
 * 4. **Semáforo:**
 * - Crítico: < 2 meses de estoque.
 * - Atenção: 2 a 5 meses.
 * - Ok: > 5 meses ou sem consumo.
 * * ----------------------------------------------------------------------------
 * 9. 09_Relatorios_Locais.js (Ferramentas do Dia a Dia)
 * ----------------------------------------------------------------------------
 * - **Função:** Gera relatórios operacionais sob demanda.
 * 1. **Validade de Atas:** Lê datas em L1/M1 e busca itens cujas atas vencem no período.
 * 2. **Resíduo 10%:** Lista itens que sobraram "migalhas" para limpeza da base.
 * 3. **Atrasos > 10 Dias:** Varre a base, ignora itens "Concluídos" ou "Resíduo" e
 * lista quem está devendo há mais de 10 dias.
 * 4. **Lista:** Preenche automaticamente as colunas C a N da aba "Lista" baseado apenas
 * nos códigos digitados na coluna A.
 * * ----------------------------------------------------------------------------
 * 10. 10_Helpers_Relatorios.js (Apoio)
 * ----------------------------------------------------------------------------
 * - **Função:** Funções utilitárias para os relatórios.
 * - **Destaque:** `parseAtrasoParaDias()` - Converte texto humano como "1 mês e 5 dias"
 * para o número "35", permitindo cálculos matemáticos de atraso.
 * * ----------------------------------------------------------------------------
 * 11. 11_Sincronizacao_Externa.js (Auditoria)
 * ----------------------------------------------------------------------------
 * - **Alvo:** Planilha de Correção (`1r8l...`).
 * - **Função:** Compara o que temos no controle manual com o que existe no EMS (`1s44YD...`).
 * - **Diagnóstico:** Aponta "Itens Faltantes" (estão no sistema mas esquecemos de lançar)
 * e "Códigos Errados" (digitamos errado no manual). Permite envio automático da correção.
 * * ----------------------------------------------------------------------------
 * 12. 12_Dashboard.js (Visualização)
 * ----------------------------------------------------------------------------
 * - **Função:** Gera a aba "Dashboard".
 * - **Mecanismo:** Conta via script a frequência de cada status na aba "Compilados"
 * e desenha um Gráfico de Pizza 3D nativo do Google Sheets.
 * * ----------------------------------------------------------------------------
 * 13. 13_Relatorio_Email.js (Reporte Automático)
 * ----------------------------------------------------------------------------
 * - **Função:** Envia e-mail para chefia (`CONFIG.emails`).
 * - **Formato:** Gera um HTML limpo com tabela de resumo (Pendentes vs Concluídos)
 * e envia via GmailApp.
 * * ----------------------------------------------------------------------------
 * 14. 14_Relatorios_BI.js (Business Intelligence)
 * ----------------------------------------------------------------------------
 * - **Função:** Análise Estratégica.
 * - **BI Fornecedores:** Cria um Ranking de "Inadimplência". Calcula % de itens entregues
 * vs itens atrasados por fornecedor.
 * - **BI Financeiro:** Solicita Ano Inicial/Final. Soma todo o valor empenhado (R$) e subtrai
 * o entregue para mostrar o "Passivo Financeiro" (quanto falta pagar).
 * * ----------------------------------------------------------------------------
 * 16. 16_Distribuicao_Equipe.js (Gestão Dinâmica)
 * ----------------------------------------------------------------------------
 * - **Alvo:** Painel de Equipe (`1Rc5fUr...`).
 * - **Inovação Dinâmica:** Não usa mais nomes "chumbados" no código. Lê a aba "Config_Equipe"
 * para saber quem cuida de qual Família.
 * - **CMA Híbrido:** Calcula a média de consumo somando dados históricos.
 * - **Coluna S (Híbrida):** Exibe o saldo em dias E a classificação (Crítico/Ok) na mesma célula.
 * - **Trava:** Impede duplicidade na configuração manual de CMM.
 * * ----------------------------------------------------------------------------
 * 17. 17_Extras_Snapshot_Cobranca.js (Automação de Cobrança)
 * ----------------------------------------------------------------------------
 * - **Snapshot:** Salva uma linha nova todo dia na aba "Historico_BI" com os totais do dia.
 * Permite criar gráficos de evolução temporal.
 * - **Cobrança:**
 * - Varre itens pendentes.
 * - Verifica se o atraso > 10 dias.
 * - Verifica a coluna "Última Cobrança" (Col V). Se já cobrou há menos de 15 dias, ignora.
 * - Se elegível, cria um Rascunho no Gmail com texto padrão cobrando o fornecedor.
 * * ----------------------------------------------------------------------------
 * 18. 18_Controle_Mobile.js (Uso no Celular)
 * ----------------------------------------------------------------------------
 * - **Problema:** O App do Sheets no celular não mostra menus de script.
 * - **Solução:** Monitora checkboxes na aba "Painel_Mobile". Se o usuário marcar "TRUE",
 * o script detecta a edição (OnEdit) e dispara a função correspondente.
 * * ----------------------------------------------------------------------------
 * 19. 19_Gerador_PDF.js (Documentação)
 * ----------------------------------------------------------------------------
 * - **Função:** Gera PDFs profissionais para impressão.
 * - **Recursos:**
 * - Permite selecionar múltiplos status (ex: "Pendente" + "Recebido Parcial").
 * - Remove logos para layout limpo.
 * - Centraliza cabeçalhos.
 * - Salva o PDF no Drive e gera link para download imediato.
 * * ----------------------------------------------------------------------------
 * 20. 20_Logger_Hub.js (Segurança)
 * ----------------------------------------------------------------------------
 * - **Função:** Auditoria de acesso.
 * - **Mecanismo:** Toda vez que a planilha é aberta, registra: Data, Hora e E-mail do Usuário
 * na aba oculta "Log_Acesso_Hub".
 * - **Manutenção:** Mantém apenas os últimos 2000 registros para não pesar o arquivo.
 * * ============================================================================
 */