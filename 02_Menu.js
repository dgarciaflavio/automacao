// =================================================================
// --- BLOCO 2: MENU PRINCIPAL (COMPLETO E ATUALIZADO) ---
// =================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Orquestrador Geral')
    
    // --- BLOCO PRINCIPAL ---
    .addItem('🔄 1. Executar CICLO COMPLETO (Tudo)', 'executarCicloCompleto')
    .addSeparator()
    .addItem('📦 2. Processar Materiais (Remoto)', 'processarMateriaisRemoto')
    .addItem('💊 3. Processar Medicamentos (Remoto)', 'processarMedicamentosRemoto')
    .addItem('📊 4. Apenas Compilar Dados (Local)', 'compilarDadosLocal') 
    .addSeparator()
    .addItem('📧 5. Enviar Relatório de Status (E-mail)', 'enviarRelatorioGerencial')
    
    // --- SUBMENU: DISTRIBUIÇÃO ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('👥 6. Distribuir para Equipe')
        .addItem('✅ Atualizar TODOS (Completo)', 'atualizarTodos')
        .addSeparator()
        .addItem('👤 Bianca', 'atualizarBianca')
        .addItem('👤 Katia', 'atualizarKatia')
        .addItem('👤 Leonardo', 'atualizarLeonardo')
        .addItem('👤 Moises', 'atualizarMoises')
        .addItem('👤 Rafaelle', 'atualizarRafaelle')
        .addItem('👤 Luciana', 'atualizarLuciana'))
    
    .addItem('🔍 7. Localizar Item (Qual Planejador?)', 'localizarItemNoPainelEquipe')
    .addSeparator()

    // --- SUBMENU: INTELIGÊNCIA ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🧠 Inteligência & Automação')
        .addItem('📸 Salvar Snapshot (Histórico BI)', 'salvarSnapshotHistorico')
        .addItem('📨 Gerar Rascunhos de Cobrança (Gmail)', 'gerarRascunhosCobranca')
        .addSeparator()
        // NOVA FUNÇÃO DE TENDÊNCIA AQUI:
        .addItem('📈 Analisar Tendência de Consumo (Aceleração)', 'gerarRelatorioTendencia')) 
    .addSeparator()

    // --- SUBMENU: GERENCIAL ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('💼 Relatórios Gerenciais (Chefia)')
        .addItem('🏆 Ranking de Fornecedores (Performance)', 'gerarRelatorioPerformanceFornecedores')
        .addItem('💰 Panorama Financeiro (Executivo)', 'gerarRelatorioFinanceiroExecutivo'))
    .addSeparator()

    // --- SUBMENU: ESTOQUE ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📈 Gestão de Estoque')
        .addItem('🔄 Sincronizar Cont.Estoque (Remoto)', 'sincronizarControleEstoque')
        .addItem('📊 Dashboard de Status', 'gerarDashboardStatus'))
    .addSeparator()

    // --- SUBMENU: OPERACIONAL ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📑 Relatórios Operacionais')
        .addItem('🚨 Executar Operação Contingência', 'executarOperacaoContingencia') // Contingência
        .addItem('📝 Atualizar Status Report', 'processarMutirao') // Status Report (Mutirão)
        .addSeparator()
        .addItem('Rel. Validade de Atas (Filtrar L1/M1)', 'gerarRelatorioValidadeAtas')
        .addItem('Rel. Valor Resíduo 10%', 'gerarRelatorioResiduo10')
        .addItem('Rel. Itens em atraso >10', 'gerarRelatorioAtrasos')
        .addItem('Relatórios Financeiros (Resumo)', 'atualizarResumo')
        .addItem('Processar Restos a Pagar', 'processarRestosAPagar')
        .addItem('Buscar Dados para Guia LISTA', 'buscarDadosLista'))
    .addSeparator()

    // --- SUBMENU: EXTERNO ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔗 Sincronização Externa')
        .addItem('1. Analisar Divergências', 'buscarEmpenhosCodigosErrados')
        .addItem('2. Enviar Itens Faltantes', 'sincronizarEmpenhosNaExterna')
        .addItem('3. Reparar Dados Vazios', 'repararDadosFaltantesNaExterna'))
    .addSeparator()
    
    // --- UTILITÁRIOS ---
    .addItem('📄 Gerar PDF por Status', 'abrirMenuGerarPDF') 
    .addItem('🧹 Limpar Visual (Ocultar Abas Técnicas)', 'ocultarAbasTecnicas')
    .addToUi();
}
