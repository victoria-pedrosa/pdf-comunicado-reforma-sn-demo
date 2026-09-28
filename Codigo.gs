// =============================================================================
// SCRIPT CONSOLIDADO - SIMULACAO IBS/CBS x SIMPLES NACIONAL (Reforma Tributaria)
// Planilha: Planejamento Tributario Consolidado (copia de producao)
// Gerado a partir do script real de producao (onOpen, gerarPDF_AbaAtual,
// extrairDadosExtrato_AbaAtual, TABELAS_SN_IBSCBS etc.) + novo modulo de
// consolidacao (buildConsolidado) e conformidade CNAE x NBS.
//
// ATUALIZACAO (DAS HIBRIDO 2027):
//   DAS Hibrido = (DAS atual - PIS - COFINS) + CBS(8,8% c/ reducao) + IBS(0,1% c/ reducao)
//   - IBS "fora do DAS" passa a usar a aliquota de teste de 0,1% (2027-2028,
//     art. 344 LC 214/2025), com a MESMA reducao setorial da CBS (art. 344, p.u.),
//     em vez do equivalente ISS/ICMS do Simples (que ficou so como referencia 2033).
//   - O total do "Simples Hibrido" no comunicado agora soma o IBS (antes so DAS
//     reduzido + CBS).
//   OBS.: o art. 347 reduz a CBS em 0,1 p.p. em 2027-2028 (ficaria 8,7%). Mantido
//   8,8% por opcao; para usar 8,7% basta trocar CBS_ALIQUOTA_CHEIA.
//
// FUNCOES REMOVIDAS NESTA VERSAO (utilitarios de migracao de uma unica vez,
// amarrados a estrutura antiga de 4 abas / listas de CNPJ hardcoded):
//   TEMP_diagnosticarNaoPreenchidos, TEMP_investigarCasos, TEMP_testarLinha,
//   TEMP_marcarBPO, TEMP_removeNonSN
// Se precisar delas novamente, estao preservadas no script original da
// planilha de producao (ID ID_EXEMPLO).
// =============================================================================

var COMUNICADO_DOC_ID = "ID_EXEMPLO";
var ID_PASTA_DESTINO_PDF = "ID_EXEMPLO";
var ABAS_ORIGEM_CONSOLIDACAO = ["Comercio", "Saude", "ProfLiberal", "Servico"];

var SN_HEADERS = [
  "CNPJ", "Razão Social da Empresa", "Anexo do Simples Nacional", "RBT12 (Receita Bruta Acumulada)",
  "Fator R", "Alíquota PIS/COFINS até 2026 (%)", "Valor no DAS (R$)", "% Faturamento para Clientes PJ",
  "% Estimativa Crédito CBS (Insumos)", "CBS Devido - DENTRO do DAS (R$)", "Crédito CBS - DENTRO do DAS (R$)",
  "CBS Devido - FORA do DAS (R$)", "Crédito CBS - FORA do DAS (R$)", "Recomendação", "Status / Data de Geração",
  "☑️ SELECIONAR", "IBS Devido - DENTRO do DAS (R$)", "Crédito IBS - DENTRO do DAS (R$)",
  "IBS Devido - FORA do DAS (R$)", "Crédito IBS - FORA do DAS (R$)", "Setor",
  "Redução Aplicada (%)", "Origem da Redução", "CBS+IBS Pleno 2033 Estimado (R$/ano)",
  "Custo Extra Estimado 2033 (R$/ano)", "Observação / Qualidade dos Dados",
  "Valor Híbrido Informado (R$)"
];

var CNAENBS_HEADERS = [
  "CNPJ", "Razão Social da Empresa", "Setor (origem)", "CNAE Cadastrado", "NBS de Fato Prestado",
  "Descrição NBS", "cClassTrib", "Redução Aplicável (%)", "Status Conformidade",
  "CNAE Sugerido (se divergente)", "Diretriz / Ação Recomendada"
];

var HIST_HEADERS = [
  "CNPJ", "Razão Social", "Setor", "Competência (AAAA-MM)", "Anexo", "Fator R",
  "Receita Bruta do PA - RPA (R$)", "RBT12 (R$)", "IRPJ (R$)", "CSLL (R$)",
  "COFINS (R$)", "PIS/Pasep (R$)", "INSS/CPP (R$)", "ISS (R$)", "Total DAS (R$)",
  "Alíquota Efetiva (%)", "Data de Extração"
];

// [nbs, descricao_nbs, item_lc116, cclasstrib] - 675 linhas - fonte: Anexo VIII (LCP 214/2025) x CNAE x Item da Lista (LC 116)
var REF_NBS_DATA = [["1.1502.10.00","Serviços de projeto, desenvolvimento e instalação de aplicativos e programas não personalizados (não customizados)","01.01","000001"],["1.1502.20.00","Serviços de projeto e desenvolvimento, adaptação e instalação de aplicativos e programas personalizados (customizados)","01.01","000001"],["1.1502.40.00","Serviços de projeto e desenvolvimento de estruturas e conteúdo de bancos de dados","01.01","000001"],["1.1502.50.00","Serviços de integração de sistemas em tecnologia da informação (TI)","01.01","000001"],["1.1502.90.00","Serviços de projeto e desenvolvimento de aplicativos e programas em tecnologia da informação (TI) não classificados em subposições anteriores","01.01","000001"],["1.1503.00.00","Serviços de projeto e desenvolvimento de redes em tecnologia da informação (TI)","01.01","000001"],["1.1504.00.00","Serviços de projeto e desenvolvimento de topografias de circuitos integrados","01.01","000001"],["1.1505.00.00","Serviços de projeto de circuitos integrados","01.01","000001"],["1.1507.10.00","Serviços de gerenciamento de redes em tecnologia da informação (TI)","01.01","000001"],["1.1507.20.00","Serviços de gerenciamento de sistemas computacionais","01.01","000001"],["1.1507.90.00","Serviços de gerenciamento de infraestrutura em tecnologia da informação (TI) não classificados em subposições anteriores","01.01","000001"],["1.1506.10.00","Serviços de hospedagem de sítios eletrônicos na rede mundial de computadores","01.03","000001"],["1.1506.21.00","Serviços de hospedagem de aplicativos e programas software como serviço (SaaS)","01.03","000001"],["1.1506.22.00","Serviços de fornecimento de infraestrutura como serviço (IaaS)","01.03","000001"],["1.1506.23.00","Serviços de fornecimento de plataformas como serviço (PaaS)","01.03","000001"],["1.1506.29.00","Serviços de hospedagem de aplicativos e programas não classificados em subposições anteriores","01.03","000001"],["1.1506.90.00","Serviços de hospedagem e de disponibilização de infraestrutura em tecnologia da informação (TI) não classificados em subposições anteriores","01.03","000001"],["1.1509.00.00","Serviços de processamento de dados","01.03","000001"],["1.1103.21.00","Licenciamento de direitos de produção, distribuição ou comercialização de programas de computador (software)","01.05","000001"],["1.1103.22.00","Licenciamento de direitos de uso de programas de computador (software)","01.05","000001"],["1.1103.23.00","Licenciamento de direitos sobre bancos de dados","01.05","000001"],["1.1103.29.00","Licenciamento de direitos sobre programas de computador (software) e bancos de dados não classificado em subposições anteriores","01.05","000001"],["1.1106.20.00","Cessão temporária de direitos sobre programas de computador (software)","01.05","000001"],["1.1107.20.00","Cessão definitiva de direitos sobre programas de computador (software)","01.05","000001"],["1.1501.10.00","Serviços de consultoria em tecnologia da informação (TI)","01.06","000001"],["1.1501.20.00","Serviços de segurança em tecnologia da informação (TI)","01.06","000001"],["1.1510.00.00","Serviços de tecnologia da informação (TI) não classificados em subposições anteriores","01.06","000001"],["1.1501.30.00","Serviços de suporte em tecnologia da informação (TI)","01.07","000001"],["1.1508.00.00","Serviços de manutenção de aplicativos e programas","01.07","000001"],["1.1502.30.00","Serviços de projeto e desenvolvimento de estruturas e conteúdo de páginas eletrônicas","01.08","000001"],["1.1703.10.00","Serviços de oferta de livros, jornais, periódicos, diretórios e malas diretas de acesso imediato (on-line)","01.09","000001"],["1.1703.21.00","Serviços de oferta de áudio para download","01.09","000001"],["1.1703.22.00","Serviços de oferta de áudio de conteúdo contínuo (streaming)","01.09","000001"],["1.1703.31.00","Serviços de oferta de arquivos contendo filmes e vídeos para download","01.09","000001"],["1.1703.32.00","Serviços de oferta de filmes e vídeos de conteúdo contínuo (streaming)","01.09","000001"],["1.1703.91.00","Serviços de oferta de jogos de acesso imediato (on-line)","01.09","000001"],["1.1703.92.00","Serviços de oferta de conteúdo de portais de busca na rede mundial de computadores","01.09","000001"],["1.1703.99.00","Serviços de oferta de outros conteúdos de acesso imediato (on-line) não classificados em subposições anteriores","01.09","000001"],["1.1201.11.00","Serviços de pesquisa e desenvolvimento em ciências físicas","02.01","000001"],["1.1201.12.00","Serviços de pesquisa e desenvolvimento em química e biologia","02.01","000001"],["1.1201.19.00","Serviços de pesquisa e desenvolvimento em ciências não classificadas em subposições anteriores","02.01","000001"],["1.1201.20.00","Serviços de pesquisa e desenvolvimento em biotecnologia","02.01","000001"],["1.1201.31.00","Serviços de pesquisa e desenvolvimento em Tecnologia da Informação e Comunicação (TIC)","02.01","000001"],["1.1201.32.00","Serviços de pesquisa e desenvolvimento em nanotecnologia","02.01","000001"],["1.1201.33.00","Serviços de pesquisa e desenvolvimento em engenharia e tecnologia nucleares","02.01","000001"],["1.1201.34.00","Serviços de pesquisa e desenvolvimento em engenharia e tecnologia em micro-ondas de potência","02.01","000001"],["1.1201.39.00","Serviços de pesquisa e desenvolvimento em engenharia e tecnologia não classificados em subposições anteriores","02.01","000001"],["1.1201.40.00","Serviços de pesquisa e desenvolvimento em ciências médicas, odontológicas e farmacêuticas","02.01","000001"],["1.1201.50.00","Serviços de pesquisa e desenvolvimento em ciências agrárias","02.01","000001"],["1.1201.90.00","Serviços de pesquisa e desenvolvimento em ciências, engenharia e tecnologia não classificados em subposições anteriores","02.01","000001"],["1.1202.10.00","Serviços de pesquisa e desenvolvimento em psicologia","02.01","000001"],["1.1202.20.00","Serviços de pesquisa e desenvolvimento em ciências econômicas","02.01","000001"],["1.1202.30.00","Serviços de pesquisa e desenvolvimento em direito","02.01","000001"],["1.1202.40.00","Serviços de pesquisa e desenvolvimento em línguas e literatura","02.01","000001"],["1.1202.90.00","Serviços de pesquisa e desenvolvimento em ciências sociais e humanidades não classificadas em subposições anteriores","02.01","000001"],["1.1203.00.00","Serviços de pesquisa e desenvolvimento interdisciplinar","02.01","000001"],["1.1103.33.00","Licenciamento de direitos de autor de obras publicitárias","03.02","000001"],["1.1104.20.00","Licenciamento de direitos sobre marcas","03.02","000001"],["1.1106.33.00","Cessão temporária de direitos de autor de obras publicitárias","03.02","000001"],["1.1107.33.00","Cessão definitiva de direitos de obras publicitárias","03.02","000001"],["1.1108.20.00","Cessão definitiva de direitos sobre marcas","03.02","000001"],["1.1805.31.00","Serviços de reservas para centros de convenções, auditórios e salas de exposições","03.03","200027"],["1.1806.61.00","Serviços de assistência e organização de convenções","03.03","200027"],["1.1806.62.00","Serviços de assistência e organização de feiras de negócios","03.03","200027"],["1.1806.63.00","Serviços de assistência e organização de exposições e outros eventos","03.03","200027"],["1.2508.00.00","Serviços recreativos, culturais e desportivos não classificados em posições anteriores","03.03","200027"],["1.1001.12.10","Serviços de administração e locação, sublocação, arrendamento, direito de passagem ou permissão de uso, compartilhado ou não, de ferrovia, rodovia, postes, cabos, dutos e condutos de qualquer natureza","03.04","200027"],["1.0105.70.00","Serviços de andaimes","03.05","000001"],["1.0105.50.00","Serviços de montagem de estruturas de aço","03.05","000001"],["1.2301.22.00","Serviços médicos especializados","04.01","200029"],["1.2301.93.00","Serviços laboratoriais","04.02","200029"],["1.2301.94.00","Serviços de diagnóstico por imagem","04.02","200029"],["1.2301.11.00","Serviços cirúrgicos","04.03","200029"],["1.2301.12.00","Serviços ginecológicos e obstétricos","04.03","200029"],["1.2301.13.00","Serviços psiquiátricos","04.03","200029"],["1.2301.14.00","Serviços prestados em Unidades de Terapia Intensiva","04.03","200029"],["1.2301.15.00","Serviços de atendimento de urgência","04.03","200029"],["1.2301.19.00","Serviços hospitalares não classificados em subposições anteriores","04.03","200029"],["1.2301.21.00","Serviços de clínica médica","04.03","200029"],["1.2301.99.00","Outros serviços de saúde humana não classificados em subposições anteriores","04.05","200029"],["1.2301.91.00","Serviços de enfermagem","04.06","200029"],["1.2301.97.00","Serviços de assistência ao parto e pós-parto","04.06","200029"],["1.2301.92.00","Serviços de fisioterapia","04.08","200029"],["1.2301.23.00","Serviços odontológicos","04.12","200029"],["1.2301.98.00","Serviços de psicologia","04.16","200029"],["1.2201.11.00","Serviços de creche ou entidade equivalente","04.17","200029"],["1.2302.10.00","Serviços de cuidado em saúde em unidades de acolhimento","04.17","200029"],["1.2302.21.00","Serviços de assistência a idosos em unidades de acolhimento","04.17","200029"],["1.2302.22.00","Serviços de assistência a crianças e adolescentes com deficiência em unidades de acolhimento","04.17","200029"],["1.2302.23.00","Serviços de assistência a adultos com deficiência em unidades de acolhimento","04.17","200029"],["1.2303.00.00","Serviços de assistência social com acomodação","04.17","200029"],["1.2301.95.00","Serviços de bancos de material biológico humano","04.19","200029"],["1.2301.96.00","Serviços de ambulância","04.21","200029"],["1.0910.10.00","Serviços de planos privados de assistência à saúde","04.22","011002"],["1.0910.90.00","Serviços de planos privados de assistência à saúde e serviços relacionados não classificados em subposições anteriores","04.22","011002"],["1.1405.12.00","Serviços de atendimento, assistência ou tratamento para animais domésticos","05.01","200052"],["1.1405.22.00","Serviços de atendimento, assistência ou tratamento para animais de corte","05.01","200038"],["1.1405.90.00","Serviços veterinários não classificados em subposições anteriores","05.01","200052"],["1.1405.11.00","Serviços hospitalares, com ou sem internação, para animais domésticos","05.02","000001"],["1.1405.21.00","Serviços hospitalares, com ou sem internação, para animais de corte","05.02","200038"],["1.1405.40.00","Serviços de bancos de órgãos, sangue, sêmen, tecidos, óvulos e outros materiais biológicos","05.05","000001"],["1.1405.60.00","Serviços de guarda, adestramento, embelezamento e alojamento","05.08","000001"],["1.1405.50.00","Planos de atendimento e assistência médico-veterinária","05.09","011005"],["1.2602.10.00","Serviços de cabeleireiros e barbeiros","06.01","000001"],["1.2602.20.00","Serviços de manicure, pedicure e tratamento cosmético","06.01","000001"],["1.2602.30.00","Serviços de bem-estar físico","06.02","000001"],["1.2602.90.00","Serviços de tratamento de beleza e bem-estar físico não classificados em subposições anteriores","06.02","000001"],["1.2205.12.00","Serviços de educação desportiva e recreacional","06.04","200041"],["1.2505.90.00","Serviços desportivos e recreacionais desportivos não classificados em subposições anteriores","06.04","000001"],["1.1402.11.00","Serviços de consultoria em arquitetura","07.01","200052"],["1.1402.12.00","Serviços de arquitetura para projetos de construções residenciais","07.01","200052"],["1.1402.13.00","Serviços de arquitetura para projetos de construções não residenciais","07.01","200052"],["1.1402.14.00","Serviços de arquitetura para restauração de prédios históricos","07.01","200052"],["1.1402.21.00","Serviços de planejamento urbano","07.01","200052"],["1.1402.22.00","Serviços de planejamento de áreas rurais","07.01","200052"],["1.1402.31.00","Serviços de consultoria em paisagismo","07.01","000001"],["1.1402.32.00","Serviços de paisagismo, exceto consultoria","07.01","000001"],["1.1402.90.00","Serviços de arquitetura, de planejamento urbano e de áreas rurais e de paisagismo não classificados em subposições anteriores","07.01","200052"],["1.1403.10.00","Serviços de consultoria em engenharia","07.01","200038"],["1.1403.21.10","Serviços de engenharia para projetos de construção residencial","07.01","200052"],["1.1403.21.20","Serviços de engenharia para projetos de construção não residencial","07.01","200052"],["1.1403.22.11","Serviços de engenharia para projetos de exploração de minerais","07.01","200052"],["1.1403.22.12","Serviços de engenharia para projetos de exploração de petróleo e gás","07.01","200052"],["1.1403.22.13","Serviços de engenharia para projetos de refino de petróleo e petroquímica","07.01","200052"],["1.1403.22.14","Serviços de engenharia para projetos de unidades de produção de biocombustíveis","07.01","200052"],["1.1403.22.21","Serviços de engenharia para projetos de veículos terrestres","07.01","200052"],["1.1403.22.22","Serviços de engenharia para projetos de embarcações","07.01","200052"],["1.1403.22.23","Serviços de engenharia para projetos de veículos aéreos e aeroespaciais","07.01","200052"],["1.1403.22.90","Serviços de engenharia para outros projetos industriais e de fabricação, exceto para projetos de energia","07.01","200052"],["1.1403.23.00","Serviços de engenharia para projetos de infraestrutura de transportes","07.01","200052"],["1.1403.24.00","Serviços de engenharia para projetos de energia","07.01","200052"],["1.1403.25.00","Serviços de engenharia para projetos de telecomunicações, radiodifusão e televisão","07.01","200052"],["1.1403.26.00","Serviços de engenharia para projetos de gerenciamento de resíduos (perigosos e não perigosos)","07.01","200052"],["1.1403.27.00","Serviços de engenharia para projetos de distribuição de água e rede de esgoto","07.01","200052"],["1.1403.29.00","Serviços de engenharia para outros projetos","07.01","200038"],["1.1403.30.00","Serviços de gerenciamento de projetos de construção","07.01","200052"],["1.1403.90.00","Serviços de engenharia não classificados em subposições anteriores","07.01","200052"],["1.1404.11.00","Serviços de consultoria geológica e geofísica","07.01","000001"],["1.1404.12.00","Serviços geofísicos","07.01","000001"],["1.1404.13.00","Serviços geoquímicos","07.01","000001"],["1.1404.14.00","Serviços de informações para avaliação e exploração de recursos naturais","07.01","000001"],["1.1404.19.00","Serviços geológicos, geofísicos e outros de prospecção não classificados em subposições anteriores","07.01","000001"],["1.0101.11.00","Serviços de construção de edificações residenciais de um e dois pavimentos","07.02","200046"],["1.0101.12.00","Serviços de construção de edificações residenciais com mais de dois pavimentos","07.02","200046"],["1.0101.21.00","Serviços de construção de edificações industriais","07.02","200046"],["1.0101.22.00","Serviços de construção de edificações comerciais","07.02","200046"],["1.0101.29.00","Serviços de construção de edificações não residenciais não classificados em subposições anteriores","07.02","200046"],["1.0101.30.00","Serviços de construção de edificações de uso misto (residencial e não residencial)","07.02","200046"],["1.0102.11.00","Serviços de construção de autoestradas (exceto autoestradas elevadas), ruas e estradas","07.02","200046"],["1.0102.12.00","Serviços de construção de ferrovias","07.02","200046"],["1.0102.13.00","Serviços de construção de pistas de pouso e decolagem em aeroportos e de infraestrutura aeroportuária","07.02","200046"],["1.0102.20.00","Serviços de construção de pontes, autoestradas elevadas e túneis","07.02","200046"],["1.0102.31.00","Serviços de construção de infraestrutura de proteção e acesso aquaviário","07.02","200046"],["1.0102.32.00","Serviços de construção de infraestrutura de acostagem aquaviária","07.02","200046"],["1.0102.33.00","Serviços de construção de infraestrutura terrestre e de obras de engenharia afins nos portos","07.02","200046"],["1.0102.34.00","Serviços de construção de barragens","07.02","200046"],["1.0102.35.10","Serviços de construção de adutoras em conduto livre","07.02","200046"],["1.0102.35.20","Serviços de construção de sistemas de irrigação","07.02","200046"],["1.0102.35.30","Serviços de construção relacionada ao controle dos cursos de água, inclusive canalização","07.02","200046"],["1.0102.41.10","Serviços de construção de dutos de longo curso para o transporte de petróleo, seus derivados, e gás","07.02","200046"],["1.0102.41.20","Serviços de construção de dutos de longo curso para o transporte e escoamento de águas","07.02","200046"],["1.0102.41.90","Serviços de construção de dutos de longo curso não classificados em subposições anteriores","07.02","200046"],["1.0102.42.10","Serviços de construção de linhas de comunicação de longo curso","07.02","200046"],["1.0102.42.20","Serviços de construção de linhas de transmissão de alta tensão","07.02","200046"],["1.0102.51.00","Serviços de construção de dutos locais","07.02","200046"],["1.0102.52.10","Serviços de construção de linhas locais de comunicação","07.02","200046"],["1.0102.52.20","Serviços de construção de linhas locais de transmissão de baixa e média tensão","07.02","200046"],["1.0102.53.10","Serviços de construção de sistemas de esgotos","07.02","200046"],["1.0102.53.20","Serviços de construção de sistemas de estações para elevação, tratamento e purificação de água","07.02","200046"],["1.0102.61.00","Serviços de construção de usinas de geração de energia","07.02","200046"],["1.0102.69.00","Serviços de construção de instalações industriais não classificados em subposições anteriores","07.02","200046"],["1.0102.70.00","Serviços de construção de minas e suas unidades industriais","07.02","200046"],["1.0102.80.00","Serviços de construção de instalações para recreação e atividades desportivas ao ar livre","07.02","200046"],["1.0102.90.00","Serviços de construção de obras de engenharia civil não classificados em subposições anteriores","07.02","200046"],["1.0103.20.00","Serviços de preparação de terrenos e de canteiros de obras","07.02","200046"],["1.0103.30.00","Serviços de escavação e remoção de terra","07.02","200046"],["1.0103.41.00","Serviços de perfuração de poços de água","07.02","200046"],["1.0103.42.00","Serviços de instalação de sistemas sépticos","07.02","200046"],["1.0104.00.00","Serviços de montagem e de edificação de construções pré-fabricadas","07.02","200046"],["1.0105.11.00","Serviços de estaqueamento","07.02","200046"],["1.0105.12.00","Serviços de fundação","07.02","200046"],["1.0105.21.00","Serviços de construção de estruturas de edificações","07.02","200046"],["1.0105.22.00","Serviços de construção de estruturas de telhados e coberturas","07.02","200046"],["1.0105.30.00","Serviços de construção de telhados e coberturas e serviços de impermeabilização","07.02","200046"],["1.0105.40.00","Serviços de concretagem","07.02","200046"],["1.0105.60.00","Serviços de alvenaria","07.02","200046"],["1.0105.90.00","Serviços especializados de construção não classificados em subposições anteriores","07.02","200046"],["1.0106.11.00","Serviços de instalação de fiação elétrica e componentes","07.02","200046"],["1.0106.19.00","Serviços de instalação elétrica não classificados em subposições anteriores","07.02","200046"],["1.0106.21.00","Serviços de instalação de tubulação para fornecimento de água","07.02","200046"],["1.0106.22.00","Serviços de instalação de tubulação para escoamento de água","07.02","200046"],["1.0106.60.00","Serviços de instalação de elevadores, esteiras e escadas rolantes","07.02","200046"],["1.0107.30.00","Serviços de pintura","07.02","200046"],["1.0107.40.00","Serviços de revestimento de pisos e de paredes","07.02","200046"],["1.0106.32.00","Serviços de instalação de equipamentos de ventilação e de ar condicionado","07.02","200046"],["1.0106.40.00","Serviços de instalação de gás","07.02","200046"],["1.0103.10.00","Serviços de demolição","07.04","200046"],["1.0106.50.00","Serviços de instalação de isolamentos","07.06","200046"],["1.0107.10.00","Serviços de vidraçaria","07.06","200046"],["1.0107.20.00","Serviços de gesso e de estuque","07.06","200046"],["1.0107.90.00","Serviços de acabamento não classificados em subposições anteriores","07.08","200046"],["1.2406.10.00","Serviços de varrição de vias e áreas públicas","07.09","000001"],["1.2402.20.00","Serviços de esvaziamento e limpeza de fossas sépticas","07.09","000001"],["1.2403.11.00","Serviços de coleta de resíduos de serviços de saúde e outros resíduos biológicos","07.09","000001"],["1.2403.12.00","Serviços de coleta de resíduos perigosos industriais, exceto resíduos de serviços de saúde e outros resíduos biológicos","07.09","000001"],["1.2403.19.00","Serviços de coleta de outros resíduos perigosos não classificados em subposições anteriores","07.09","000001"],["1.2403.21.00","Serviços de coleta de resíduos recicláveis, não perigosos, de origem doméstica","07.09","000001"],["1.2403.22.00","Serviços de coleta de resíduos recicláveis, não perigosos, exceto de origem doméstica","07.09","000001"],["1.2403.31.00","Serviços de coleta de resíduos gerais de origem doméstica","07.09","000001"],["1.2403.32.00","Serviços de coleta de resíduos gerais, exceto de origem doméstica","07.09","000001"],["1.2404.11.00","Serviços de triagem, preparação, consolidação e estocagem de resíduos perigosos","07.09","000001"],["1.2404.12.00","Serviços de demolição e desmantelamento de embarcações, veículos e outros bens","07.09","000001"],["1.2404.13.00","Serviços de triagem, preparação, consolidação e estocagem de resíduos recicláveis não perigosos","07.09","000001"],["1.2404.19.00","Serviços de triagem, preparação, consolidação e estocagem de resíduos não perigosos, exceto os recicláveis","07.09","000001"],["1.2404.33.00","Serviços de incineração de resíduos não perigosos","07.09","000001"],["1.1803.10.00","Serviços gerais de limpeza","07.10","000001"],["1.1803.22.00","Serviços de limpeza de janelas","07.10","000001"],["1.1803.29.00","Serviços especializados de limpeza não classificados em subposições anteriores","07.10","000001"],["1.2405.14.00","Serviços de remediação em edificações","07.10","000001"],["1.2406.90.00","Serviços de limpeza urbana e similares não classificados em subposições anteriores","07.10","000001"],["1.2407.00.00","Serviços de proteção ambiental não classificados em posições anteriores","07.10","000001"],["1.1409.11.00","Serviços de design de interiores para espaços comerciais e públicos","07.11","000001"],["1.1409.12.00","Serviços de design de interiores para espaços residenciais","07.11","000001"],["1.1806.70.00","Serviços de jardinagem","07.11","000001"],["1.2404.21.00","Serviços de tratamento de resíduos perigosos","07.12","000001"],["1.2404.22.00","Serviços de eliminação de resíduos perigosos","07.12","000001"],["1.2404.31.00","Serviços de tratamento e eliminação de resíduos não perigosos em aterros sanitários","07.12","000001"],["1.2404.32.00","Serviços de tratamento e eliminação de resíduos não perigosos em aterros, exceto os sanitários","07.12","000001"],["1.2404.39.00","Serviços de tratamento e eliminação de resíduos não perigosos não classificados em subposições anteriores","07.12","000001"],["1.2405.11.00","Serviços de remediação e limpeza do ar","07.12","000001"],["1.2405.12.00","Serviços de remediação e limpeza de águas de superfície","07.12","000001"],["1.2405.13.00","Serviços de remediação e limpeza do solo e de águas subterrâneas","07.12","000001"],["1.2405.20.00","Serviços de contenção, controle e monitoramento de áreas contaminadas","07.12","000001"],["1.2405.90.00","Serviços de remediação não classificados em subposições anteriores","07.12","000001"],["1.1803.21.00","Serviços de desinfecção e extermínio de pragas","07.13","000001"],["1.1105.41.00","Exploração de recursos vegetais, inclusive florestais","07.16","000001"],["1.1901.10.00","Serviços de apoio à agricultura","07.16","200038"],["1.1901.30.00","Serviços de apoio à produção florestal (silvicultura)","07.16","000001"],["1.0602.31.00","Serviços de armazenagem de granéis sólidos","07.16","000001"],["1.1402.15.00","Serviços de arquitetura relativos ao acompanhamento e fiscalização da execução de projetos arquitetônicos e urbanísticos","07.19","200052"],["1.1404.21.00","Serviços topográficos","07.20","000001"],["1.1404.22.00","Serviços cartográficos","07.20","000001"],["1.1902.10.00","Serviços de apoio à extração de petróleo e gás","07.21","000001"],["1.1902.90.00","Serviços de apoio à mineração não classificados em subposições anteriores","07.21","000001"],["1.2201.12.00","Serviços de pré-escola","08.01","200028"],["1.2201.19.00","Serviços de educação infantil não classificados em subposições anteriores","08.01","200028"],["1.2201.20.00","Serviços de ensino fundamental","08.01","200028"],["1.2201.30.00","Serviços de ensino médio","08.01","200028"],["1.2202.00.00","Serviços de educação técnica de nível médio","08.01","200028"],["1.2203.10.00","Serviços de ensino fundamental de jovens e adultos","08.01","200028"],["1.2203.20.00","Serviços de ensino médio de jovens e adultos","08.01","200028"],["1.2204.10.00","Serviços educacionais de graduação","08.01","200028"],["1.2204.20.00","Serviços educacionais de pós-graduação","08.01","200028"],["1.2204.30.00","Serviços educacionais de extensão","08.01","200028"],["1.2204.40.00","Serviços educacionais de cursos sequenciais","08.01","200028"],["1.2205.11.00","Serviços de educação com enfoque cultural","08.02","000001"],["1.2205.13.00","Serviços de educação em línguas estrangeiras e de sinais","08.02","200028"],["1.2205.19.00","Serviços de educação, inclusive treinamento não classificados em subposições anteriores","08.02","000001"],["1.2205.20.00","Serviços de apoio aos serviços educacionais","08.02","000001"],["1.0303.11.00","Serviços de hospedagem em quartos ou unidades de hospedagem para visitantes, com serviços diários de faxina","09.01","200048"],["1.0303.12.00","Serviços de hospedagem em quartos ou unidades de hospedagem para visitantes, sem serviços diários de faxina","09.01","200048"],["1.0303.13.00","Serviços de hospedagem em quartos ou unidades de hospedagem para visitantes, em propriedades partilhadas","09.01","200048"],["1.0303.14.00","Serviços de hospedagem em quartos ou unidades de hospedagem para visitantes, em quartos de múltipla ocupação","09.01","200048"],["1.0303.20.00","Serviços de acampamentos turísticos (camping)","09.01","200048"],["1.0303.90.00","Serviços de hospedagem para visitantes não classificados em subposições anteriores","09.01","200048"],["1.0304.10.00","Serviços de hospedagem em quartos ou unidades de hospedagem para estudantes em residências estudantis","09.01","200048"],["1.0304.20.00","Serviços de hospedagem em quartos ou unidades de hospedagem para trabalhadores em hotéis ou campos","09.01","200048"],["1.0304.90.00","Serviços de hospedagem, exceto para visitantes não classificados em subposições anteriores","09.01","200048"],["1.0401.17.10","Serviços de transporte rodoviário para passeios turísticos (sightseeing)","09.02","000001"],["1.0401.17.20","Serviços de transporte ferroviário para passeios turísticos (sightseeing)","09.02","000001"],["1.0401.17.90","Serviços de transporte terrestre para passeios turísticos (sightseeing) não classificados em itens anteriores","09.02","000001"],["1.0401.23.00","Serviços de transporte aquaviário para passeios turísticos (sightseeing)","09.02","000001"],["1.0401.43.00","Serviços de transporte aéreo para passeios turísticos (sightseeing)","09.02","000001"],["1.0402.13.20","Serviços de fretamento eventual ou turístico, nacional, exceto local","09.02","000001"],["1.1805.40.00","Serviços de operadoras de turismo","09.02","200051"],["1.1805.11.00","Serviços de reservas para transporte aéreo de passageiros","09.02","200051"],["1.1805.12.00","Serviços de reservas para transporte ferroviário de passageiros","09.02","200051"],["1.1805.13.00","Serviços de reservas para transporte rodoviário de passageiros","09.02","200051"],["1.1805.14.00","Serviços de reservas de carros de aluguel","09.02","200051"],["1.1805.19.00","Serviços de reservas para transporte de passageiros não classificados em subposições anteriores","09.02","200051"],["1.1805.21.00","Serviços de reservas de hospedagem, exceto em unidades compartilhadas","09.02","200051"],["1.1805.22.00","Serviços de reservas e intercâmbio de unidades compartilhadas (time-share)","09.02","200051"],["1.1805.23.00","Serviços de reservas em cruzeiros","09.02","200051"],["1.1805.24.00","Serviços de reservas de pacotes turísticos","09.02","200051"],["1.1805.32.00","Serviços de reservas de ingressos para eventos de entretenimento e recreativos","09.02","000001"],["1.1805.39.00","Serviços de reservas não classificados em subposições anteriores","09.02","000001"],["1.1805.61.00","Serviços de promoção turística","09.02","200051"],["1.1805.62.00","Serviços de informação a visitantes","09.02","000001"],["1.1805.50.00","Serviços de guias turísticos","09.03","200051"],["1.0906.11.00","Serviços de agenciamento e corretagem de seguros, resseguros e previdência complementar, exceto de seguros saúde","10.01","000001"],["1.0906.12.00","Serviços de corretagem de seguros saúde","10.01","011003"],["1.0910.20.00","Serviços de corretagem de planos privados de assistência à saúde","10.01","011003"],["1.0905.90.00","Serviços auxiliares aos serviços financeiros não classificados em subposições anteriores","10.01","000001"],["1.0607.00.00","Serviços de agenciamento de transporte de cargas","10.02","000001"],["1.0905.11.00","Serviços de corretagem de títulos","10.02","000001"],["1.0905.12.00","Serviços de corretagem de derivativos e commodities","10.02","000001"],["1.2501.40.00","Serviços de agenciamento para a comercialização de obras audiovisuais","10.03","000001"],["1.1001.21.00","Serviços de intermediação na compra e venda de imóveis residenciais","10.05","200046"],["1.1001.22.00","Serviços de intermediação na compra e venda de imóveis não residenciais","10.05","200046"],["1.0201.00.00","Serviços de intermediação na distribuição de mercadorias","10.05","000001"],["1.0205.00.00","Serviços de intermediação na comercialização de energia elétrica","10.05","000001"],["1.0502.29.00","Serviços de transportes aquaviário costeiro de cargas não classificados em subposições anteriores","10.06","000001"],["1.1704.10.00","Serviços de agências de notícias para jornais e periódicos","10.07","000001"],["1.1704.20.00","Serviços de agências de notícias para mídia audiovisual","10.07","000001"],["1.1406.20.00","Aquisição ou venda de espaço ou tempo para propaganda, sob comissão","10.08","000001"],["1.0604.30.00","Serviços de estacionamento","11.01","000001"],["1.0605.90.00","Serviços de apoio ao transporte aquaviário não classificados em subposições anteriores","11.01","000001"],["1.0606.19.00","Serviços de apoio ao transporte aéreo não classificados em subposições anteriores","11.01","000001"],["1.1802.50.00","Serviços de guarda e escolta armada","11.02","000001"],["1.1802.20.00","Serviços de consultoria em segurança","11.02","000001"],["1.1802.30.00","Serviços de sistemas de segurança","11.02","000001"],["1.1802.90.00","Serviços de segurança não classificados em subposições anteriores","11.02","000001"],["1.0601.10.00","Serviços de manuseio de contêineres","11.04","000001"],["1.0601.90.00","Serviços de manuseio de cargas não classificados em subposições anteriores","11.04","000001"],["1.0602.10.00","Serviços de armazenagem frigorificada","11.04","000001"],["1.0602.21.00","Serviços de armazenagem de petróleo e seus derivados","11.04","000001"],["1.0602.22.00","Serviços de armazenagem de combustíveis, lubrificantes e GLP, inclusive apresentado em botijões metálicos","11.04","000001"],["1.0602.23.00","Serviços de armazenagem de produtos químicos perigosos","11.04","000001"],["1.0602.29.00","Serviços de armazenagem de produtos perigosos não classificados em subposições anteriores","11.04","000001"],["1.0602.32.00","Serviços de armazenagem de granéis líquidos ou liquefeitos","11.04","000001"],["1.0602.33.00","Serviços de armazenagem de granéis gasosos","11.04","000001"],["1.0602.90.00","Serviços de armazenagem não classificados em subposições anteriores","11.04","000001"],["1.0608.20.00","Serviços de unitização ou desunitização de cargas no transporte multimodal","11.04","000001"],["1.0608.30.00","Serviços de movimentação de cargas no transporte multimodal","11.04","000001"],["1.2502.20.00","Serviços de produção e apresentação de atuações artísticas ao vivo","12.01","200039"],["1.2501.50.00","Serviços de projeção de filmes","12.02","200039"],["1.2502.90.00","Serviços de apresentação e promoção de atuações artísticas e outros serviços de entretenimento ao vivo não classificados em subposições anteriores","12.04","200039"],["1.2504.21.00","Serviços de jardins botânico e zoológico","12.05","200048"],["1.2504.22.00","Serviços de reserva natural, incluindo preservação de vida selvagem","12.05","200048"],["1.2507.10.00","Serviços de parques temáticos de diversão","12.05","200048"],["1.2507.90.00","Serviços de parques de diversão e atrações similares não classificados em subposições anteriores","12.05","200048"],["1.2505.10.00","Serviços de organização e promoção de eventos desportivos e recreacionais desportivos","12.10","000001"],["1.2505.20.00","Serviços de clubes desportivos","12.11","200042"],["1.2503.10.00","Serviços de atuação artística","12.12","200039"],["1.2501.21.00","Serviços de produção de programas de televisão, videoteipes e filmes","12.13","200039"],["1.2501.22.00","Serviços de produção de programas de rádio","12.13","200039"],["1.2502.10.00","Serviços de organização e promoção de atuações artísticas ao vivo","12.13","200039"],["1.2502.30.00","Serviços de apoio para atuações artísticas ao vivo","12.13","200039"],["1.2501.11.00","Serviços de gravação de som em estúdio","13.02","200039"],["1.2501.12.00","Serviços de gravação de som ao vivo","13.02","200039"],["1.2501.36.00","Serviços de legendas, títulos e dublagem em obras audiovisuais","13.02","200039"],["1.2501.37.00","Serviços de projeto e edição de som em obras audiovisuais","13.02","200039"],["1.2501.39.00","Serviços de pós-produção de obras audiovisuais não classificados em subposições anteriores","13.02","000001"],["1.1408.11.00","Serviços fotográficos de retratos","13.03","000001"],["1.1408.12.00","Serviços fotográficos e videográficos para propaganda","13.03","000001"],["1.1408.13.00","Serviços fotográficos e videográficos de eventos","13.03","000001"],["1.1408.14.00","Serviços fotográficos especiais","13.03","000001"],["1.1408.15.00","Serviços de restauração e retoque de fotografias","13.03","000001"],["1.1408.19.00","Serviços fotográficos e videográficos não classificados em subposições anteriores","13.03","000001"],["1.1408.20.00","Serviços de processamento de fotografias","13.03","000001"],["1.2101.23.00","Serviços de reprodução de mídia gravada","13.03","000001"],["1.2501.31.00","Serviços de edição de obras audiovisuais","13.03","200039"],["1.2501.32.00","Serviços de duplicação e transferência de obras audiovisuais","13.03","200039"],["1.2501.33.00","Serviços de correção de cor e restauração digital de obras audiovisuais","13.03","200039"],["1.2501.34.00","Serviços de efeitos visuais em obras audiovisuais","13.03","200039"],["1.2501.35.00","Serviços de animação","13.03","200039"],["1.2501.90.00","Serviços de produção audiovisual, de apoio e relacionados não classificados em subposições anteriores","13.03","200039"],["1.1806.51.00","Serviços de fotocópias e outros serviços de reprodução de documentos","13.04","000001"],["1.2101.10.00","Serviços de editoração","13.05","000001"],["1.2101.21.00","Serviços de impressão","13.05","000001"],["1.2101.22.00","Serviços relacionados à impressão","13.05","000001"],["1.2001.10.00","Serviços de manutenção e reparação de produtos metálicos, exceto maquinários e equipamentos","14.01","000001"],["1.2001.20.00","Serviços de manutenção e reparação de computadores e seus periféricos e de maquinário para escritório","14.01","000001"],["1.2001.31.10","Serviços de manutenção e reparação de veículos rodoviários motorizados","14.01","000001"],["1.2001.31.20","Serviços de manutenção e reparação de veículos rodoviários não motorizados","14.01","000001"],["1.2001.32.00","Serviços de manutenção e reparação de veículos sobre trilhos","14.01","000001"],["1.2001.33.00","Serviços de manutenção e reparação de veículos aquaviários","14.01","000001"],["1.2001.34.10","Serviços de manutenção e reparação de aeronaves, exceto de motores, turborreatores e turbopropulsores aeronáuticos","14.01","000001"],["1.2001.34.20","Serviços de manutenção e reparação de motores, turborreatores e turbopropulsores aeronáuticos","14.01","000001"],["1.2001.34.30","Serviços de manutenção e reparação de foguetes e equipamentos aeroespaciais","14.01","000001"],["1.2001.35.00","Serviços de manutenção e reparação de veículos militares","14.01","200044"],["1.2001.39.00","Serviços de manutenção e reparação de maquinários e equipamentos de transporte não classificados em subposições anteriores","14.01","000001"],["1.2001.40.00","Serviços de manutenção e reparação de plataformas, inclusive navios-plataforma, para extração de petróleo e gás","14.01","000001"],["1.2001.50.00","Serviços de manutenção e reparação de maquinários e equipamentos de uso industrial","14.01","000001"],["1.2001.60.00","Serviços de manutenção e reparação de maquinários e equipamentos de uso comercial","14.01","000001"],["1.2001.70.00","Serviços de manutenção e reparação de equipamentos e aparelhos de telecomunicações","14.01","000001"],["1.2001.81.00","Serviços de manutenção e reparação de aparelhos eletroeletrônicos domésticos","14.01","000001"],["1.2001.82.00","Serviços de manutenção e reparação de instrumentos e equipamentos médico-hospitalares, odontológicos, óticos e de precisão","14.01","000001"],["1.2001.83.00","Serviços de manutenção e reparação de equipamentos militares","14.01","200044"],["1.2001.89.00","Serviços de manutenção e reparação de outros maquinários e equipamentos não classificados em subposições anteriores","14.01","000001"],["1.2002.10.00","Serviços de manutenção e reparação de produtos de couro, calçados, malas e bolsas","14.01","000001"],["1.2002.20.00","Serviços de manutenção e reparação de relógios e joias","14.01","000001"],["1.2002.30.00","Serviços de manutenção e reparação de móveis","14.01","000001"],["1.2002.40.00","Serviços de manutenção de roupas e outros produtos têxteis","14.01","000001"],["1.2002.90.00","Serviços de manutenção e reparação de outros bens de consumo não classificados em subposições anteriores","14.01","000001"],["1.1804.00.00","Serviços de acondicionamento e empacotamento","14.05","000001"],["1.0106.12.00","Serviços de instalação de alarmes contra incêndio","14.06","000001"],["1.0106.13.00","Serviços de instalação de sistemas de alarmes antifurto","14.06","000001"],["1.0106.14.00","Serviços de instalação de antenas residenciais","14.06","000001"],["1.0106.31.00","Serviços de instalação de equipamentos de aquecimento","14.06","000001"],["1.0106.90.00","Serviços de instalação não classificados em subposições anteriores","14.06","000001"],["1.0107.60.00","Serviços de instalação de cercas e grades","14.06","000001"],["1.2003.10.00","Serviços de instalação de produtos metálicos, exceto maquinário e equipamentos","14.06","000001"],["1.2003.21.10","Serviços de montagem sob encomenda de turbinas industriais","14.06","000001"],["1.2003.21.90","Serviços de instalação de maquinários, aparelhos e equipamentos industriais não classificados em itens anteriores","14.06","000001"],["1.2003.22.00","Serviços de instalação de computadores e seus periféricos e maquinário de escritório","14.06","000001"],["1.2003.23.00","Serviços de instalação de equipamentos e aparelhos de comunicação, incluindo de rádio e de televisão","14.06","000001"],["1.2003.24.00","Serviços de instalação de maquinários, equipamentos, instrumentos e aparelhos médico-hospitalares, óticos e de precisão","14.06","000001"],["1.2003.25.10","Serviços de instalação de sensores e sistemas de armas","14.06","000001"],["1.2003.25.20","Serviços de instalação de maquinários, aparelhos e equipamentos de emprego militar","14.06","000001"],["1.2003.26.10","Serviços de montagem sob encomenda de motores, turborreatores e turbopropulsores aeronáuticos","14.06","000001"],["1.2003.26.90","Serviços de instalação de maquinários e equipamentos de transporte não classificados em itens anteriores","14.06","000001"],["1.2003.29.00","Serviços de instalação de maquinários, aparelhos e equipamentos não classificados em subposições anteriores","14.06","000001"],["1.2606.00.00","Serviços pessoais não classificados em posições anteriores","14.07","000001"],["1.2604.00.00","Serviços de confecção de roupas e outros artigos têxteis","14.09","000001"],["1.2601.10.00","Serviços de limpeza de têxteis, exceto quando realizados a seco","14.10","000001"],["1.2601.20.00","Serviços de limpeza a seco","14.10","000001"],["1.2601.30.00","Serviços de tinturaria","14.10","000001"],["1.2601.40.00","Serviços de passadoria de roupas e outros artigos têxteis","14.10","000001"],["1.2601.90.00","Serviços de lavanderia não classificados em subposições anteriores","14.10","000001"],["1.0107.50.00","Serviços de carpintaria e de serralharia","14.13","000001"],["1.0604.40.00","Serviços de reboque para veículos particulares e comerciais","14.14","000001"],["1.0901.40.00","Serviços de cartão de crédito","15.01","010002"],["1.0905.21.00","Serviços de gestão e administração de carteiras de ativos, exceto fundos de pensão","15.01","010002"],["1.0905.22.00","Serviços de administração fundos de investimento e fundos de pensão","15.01","010002"],["1.0905.23.00","Serviços de gestão e administração de trust","15.01","010002"],["1.0905.40.00","Serviços relacionados à administração de mercados financeiros","15.01","010002"],["1.0906.40.00","Serviços de gestão de fundos de previdência complementar","15.01","010002"],["1.0901.21.00","Serviços de depósito para pessoas jurídicas","15.02","000001"],["1.0901.22.00","Serviços de depósito para pessoas físicas","15.02","000001"],["1.0901.29.00","Serviços de depósito para outros depositantes","15.02","000001"],["1.1101.90.00","Arrendamento mercantil operacional ou locação de máquinas e equipamentos, sem operador, não classificado em subposições anteriores","15.03","000001"],["1.1301.30.00","Serviços de documentação e certificação, exceto os serviços notariais e de registro","15.04","000001"],["1.1806.10.00","Serviços de informação cadastral e análise de crédito","15.04","000001"],["1.0702.00.00","Serviços de coleta, transporte, remessa ou entrega de documentos ou encomendas, exceto remessas expressas","15.06","000001"],["1.0901.90.00","Serviços financeiros, exceto bancos de investimento, serviços de seguros e previdência complementar não classificados em subposições anteriores","15.07","000001"],["1.1806.31.00","Serviços de call center","15.07","000001"],["1.0901.33.00","Serviços de empréstimos e financiamentos pessoais","15.08","000001"],["1.0901.34.00","Serviços de empréstimos e financiamentos comerciais","15.08","000001"],["1.0901.35.00","Serviços de empréstimos e financiamentos industriais","15.08","000001"],["1.0901.36.00","Serviços de empréstimos e financiamentos agropecuários","15.08","000001"],["1.0901.39.00","Serviços de concessão de crédito não classificados em subposições anteriores","15.08","000001"],["1.0905.50.00","Serviços de consultoria financeira","15.08","000001"],["1.0901.51.11","Arrendamento mercantil financeiro de veículos rodoviários automotores para o transporte de passageiros","15.09","010002"],["1.0901.51.12","Arrendamento mercantil financeiro de veículos rodoviários automotores para o transporte de mercadorias","15.09","010002"],["1.0901.51.13","Arrendamento mercantil financeiro de veículos e equipamentos ferroviários","15.09","010002"],["1.0901.51.14","Arrendamento mercantil financeiro de outros equipamentos de transporte terrestre, inclusive de veículos de uso misto","15.09","010002"],["1.0901.51.15","Arrendamento mercantil financeiro de navios e outras embarcações","15.09","010002"],["1.0901.51.16","Arrendamento mercantil financeiro de aeronaves","15.09","010002"],["1.0901.51.17","Arrendamento mercantil financeiro de contêineres","15.09","010002"],["1.0901.51.21","Arrendamento mercantil financeiro de máquinas e equipamentos agrícolas","15.09","010002"],["1.0901.51.22","Arrendamento mercantil financeiro de máquinas e equipamentos de construção","15.09","010002"],["1.0901.51.23","Arrendamento mercantil financeiro de máquinas e equipamentos para escritórios, exceto computadores","15.09","010002"],["1.0901.51.24","Arrendamento mercantil financeiro de computadores","15.09","010002"],["1.0901.51.25","Arrendamento mercantil financeiro de equipamentos de telecomunicação","15.09","010002"],["1.0901.51.29","Arrendamento mercantil financeiro de outras máquinas e equipamentos não classificado em subposições anteriores","15.09","010002"],["1.0901.52.10","Arrendamento mercantil financeiro de televisões e outros eletroeletrônicos domésticos, bem como seus acessórios","15.09","010002"],["1.0901.52.20","Arrendamento mercantil financeiro de mídias gravadas","15.09","010002"],["1.0901.52.30","Arrendamento mercantil financeiro de móveis e eletrodomésticos","15.09","010002"],["1.0901.52.40","Arrendamento mercantil financeiro de equipamentos para diversão e lazer","15.09","010002"],["1.0901.52.50","Arrendamento mercantil financeiro de artigos de cama, mesa e banho","15.09","010002"],["1.0901.52.90","Arrendamento mercantil financeiro de outras mercadorias não classificado em subposições anteriores","15.09","010002"],["1.1101.11.00","Arrendamento mercantil operacional ou locação de veículos rodoviários automotores para o transporte de até oito passageiros, sem operador","15.09","010002"],["1.1101.12.00","Arrendamento mercantil operacional ou locação de veículos rodoviários automotores para o transporte de cargas, sem operador","15.09","010002"],["1.1101.13.00","Arrendamento mercantil operacional ou locação de veículos e equipamentos de transporte ferroviário, sem operador","15.09","010002"],["1.1101.14.00","Arrendamento mercantil operacional ou locação de outros equipamentos de transporte terrestre, inclusive de veículos de uso misto, sem operador","15.09","010002"],["1.1101.15.00","Arrendamento mercantil operacional ou locação de navios e outras embarcações, sem tripulação","15.09","010002"],["1.1101.16.00","Arrendamento mercantil operacional ou locação de aeronaves, sem tripulação","15.09","010002"],["1.1101.17.00","Arrendamento mercantil operacional ou locação de contêineres","15.09","010002"],["1.1101.20.00","Arrendamento mercantil operacional ou locação de máquinas e equipamentos agrícolas, sem operador","15.09","010002"],["1.1101.30.00","Arrendamento mercantil operacional ou locação de máquinas e equipamentos de construção, sem operador","15.09","010002"],["1.1101.40.00","Arrendamento mercantil operacional ou locação de máquinas e equipamentos para escritórios, exceto computadores, sem operador","15.09","010002"],["1.1101.50.00","Arrendamento mercantil operacional ou locação de computadores, sem operador","15.09","010002"],["1.1101.60.00","Arrendamento mercantil operacional ou locação de equipamentos de telecomunicação, sem operador","15.09","010002"],["1.1102.10.00","Arrendamento mercantil operacional ou locação de televisão e outros eletroeletrônicos, bem como seus acessórios","15.09","010002"],["1.1102.20.00","Arrendamento mercantil operacional ou locação de mídias gravadas","15.09","010002"],["1.1102.30.00","Arrendamento mercantil operacional ou locação de móveis e eletrodomésticos","15.09","010002"],["1.1102.40.00","Arrendamento mercantil operacional ou locação de equipamentos para diversão e lazer","15.09","010002"],["1.1102.50.00","Arrendamento mercantil operacional ou locação de artigos de cama, mesa e banho","15.09","010002"],["1.1102.60.00","Arrendamento mercantil operacional ou locação de roupas e calçados","15.09","010002"],["1.1102.90.00","Arrendamento mercantil operacional ou locação de outras mercadorias não classificado em subposições anteriores","15.09","010002"],["1.1806.20.00","Serviços de cobrança","15.10","000001"],["1.0905.30.00","Serviços de guarda e custódia","15.12","000001"],["1.0905.60.00","Serviços de câmbio","15.13","000001"],["1.0905.13.00","Serviços de compensação de transações financeiras, inclusive com ativos financeiros (clearinghouse)","15.15","000001"],["1.0901.31.00","Serviços de financiamentos imobiliários residenciais","15.18","000001"],["1.0901.32.00","Serviços de financiamentos imobiliários não residenciais","15.18","000001"],["1.1001.30.00","Serviços de avaliação de imóveis","15.18","000001"],["1.0401.11.19","Serviços de transporte rodoviário local regular de passageiros, exceto em áreas metropolitanas não classificados em subitens anteriores","16.01","400001"],["1.0401.16.10","Serviços de transporte ferroviário local de passageiros","16.01","200021"],["1.0401.16.20","Serviços de transporte metroviário (metrô) local de passageiros","16.01","400001"],["1.0401.16.90","Serviços de transporte local de passageiros por veículos sobre trilhos não classificados em itens anteriores","16.01","000001"],["1.0401.21.10","Serviços de transporte aquaviário local de passageiros, por navegação interior, em embarcações para travessia","16.01","200021"],["1.0401.21.90","Serviços de transporte aquaviário local de passageiros, por navegação interior não classificados em itens anteriores","16.01","000001"],["1.0401.30.00","Serviços de transporte integrado local de passageiros","16.01","400001"],["1.0401.11.11","Serviços de transporte rodoviário local prestados exclusivamente por meio de ônibus","16.02","000001"],["1.0401.12.10","Serviços de transporte escolar","16.02","000001"],["1.0401.12.20","Serviços de transporte para aeroportos (shuttle)","16.02","000001"],["1.0401.12.90","Serviços especiais de transporte rodoviário local regular de passageiros não classificados em itens anteriores","16.02","000001"],["1.0401.13.00","Serviços de táxi","16.02","000001"],["1.0401.14.00","Serviços de carro com motorista, exceto táxi","16.02","000001"],["1.0401.15.10","Serviços de fretamento contínuo, local","16.02","000001"],["1.0401.15.20","Serviços de fretamento eventual ou turístico, local","16.02","000001"],["1.0401.19.00","Serviços de transporte terrestre local de passageiros não classificados em subposições anteriores","16.02","000001"],["1.0401.21.20","Serviços de transporte aquaviário local de passageiros, por navegação interior, em embarcações para cruzeiros","16.02","000001"],["1.0401.22.00","Serviços de transporte aquaviário local de passageiros por fretamento","16.02","000001"],["1.0401.29.00","Serviços de transporte aquaviário local de passageiros não classificados em subposições anteriores","16.02","000001"],["1.0401.41.00","Serviços de táxi aéreo local","16.02","000001"],["1.0401.42.00","Serviços de transporte aéreo local de passageiros por fretamento","16.02","000001"],["1.0401.49.00","Serviços de transporte aéreo local de passageiros não classificados em subposições anteriores","16.02","000001"],["1.0401.90.00","Serviços de transporte local de passageiros não classificados não classificados em subposições anteriores","16.02","000001"],["1.0404.10.00","Locação de veículos rodoviários de passageiros com motorista","16.02","000001"],["1.0404.20.00","Locação de embarcações de passageiros com tripulação","16.02","000001"],["1.0404.30.00","Locação de aeronaves de passageiros com tripulação","16.02","000001"],["1.0405.00.00","Afretamento de embarcações de passageiros por tempo","16.02","000001"],["1.0501.11.10","Serviços de transporte rodoviário de cargas sólidas a granel","16.02","000001"],["1.0501.11.20","Serviços de transporte rodoviário de cargas líquidas, ou liquefeitas, a granel","16.02","000001"],["1.0501.11.30","Serviços de transporte rodoviário de cargas gasosas a granel","16.02","000001"],["1.0501.12.10","Serviços de transporte rodoviário de carga solta, não unitizada","16.02","000001"],["1.0501.12.20","Serviços de transporte rodoviário de carga unitizada","16.02","000001"],["1.0501.12.30","Serviços de transporte rodoviário de carga frigorificada ou climatizada","16.02","000001"],["1.0501.13.10","Serviços de transporte rodoviário de cargas em contêineres frigorificados ou climatizados","16.02","000001"],["1.0501.13.20","Serviços de transporte rodoviário de cargas em contêineres não frigorificados ou climatizados","16.02","000001"],["1.0501.14.10","Serviços de transporte rodoviário de cargas vivas","16.02","000001"],["1.0501.14.20","Serviços de transporte rodoviário de mudanças domésticas e de mobília e outros objetos de escritório","16.02","000001"],["1.0501.14.30","Serviços de transporte rodoviário de cargas de grande porte","16.02","000001"],["1.0501.14.40","Serviços de transporte rodoviário de veículos","16.02","000001"],["1.0501.14.51","Serviços de transporte rodoviário de combustíveis, lubrificantes e GLP, inclusive apresentados em botijões metálicos","16.02","000001"],["1.0501.14.52","Serviços de transporte rodoviário de produtos químicos perigosos, exceto lubrificantes e GLP","16.02","000001"],["1.0501.14.59","Serviços de transporte rodoviário de produtos perigosos não classificados em subitens anteriores","16.02","000001"],["1.0501.15.00","Serviços de transporte rodoviário de cargas postais e malotes","16.02","000001"],["1.0501.19.00","Serviços de transporte rodoviário de cargas não classificados em subposições anteriores","16.02","000001"],["1.0502.11.10","Serviços de transporte aquaviário por navegação interior de cargas sólidas, a granel","16.02","000001"],["1.0502.11.20","Serviços de transporte aquaviário por navegação interior de cargas líquidas, ou liquefeitas, a granel","16.02","000001"],["1.0502.11.30","Serviços de transporte aquaviário por navegação interior de cargas gasosas a granel","16.02","000001"],["1.0502.12.10","Serviços de transporte aquaviário por navegação interior de carga solta, não unitizada","16.02","000001"],["1.0502.12.20","Serviços de transporte aquaviário por navegação interior de carga unitizada","16.02","000001"],["1.0502.12.30","Serviços de transporte aquaviário por navegação interior de carga frigorificada ou climatizada","16.02","000001"],["1.0502.13.10","Serviços de transporte aquaviário por navegação interior de cargas em contêineres frigorificados ou climatizados","16.02","000001"],["1.0502.13.20","Serviços de transporte aquaviário por navegação interior de cargas em contêineres não frigorificados ou climatizados","16.02","000001"],["1.0502.14.10","Serviços de transporte aquaviário por navegação interior de cargas vivas","16.02","000001"],["1.0502.14.20","Serviços de transporte aquaviário por navegação interior de mudanças domésticas e de mobília e outros objetos de escritório","16.02","000001"],["1.0502.14.30","Serviços de transporte aquaviário por navegação interior de cargas de grande porte","16.02","000001"],["1.0502.14.40","Serviços de transporte aquaviário por navegação interior de veículos","16.02","000001"],["1.0502.14.51","Serviços de transporte aquaviário por navegação interior de combustíveis, lubrificantes e GLP, inclusive apresentado em botijões metálicos","16.02","000001"],["1.0502.14.52","Serviços de transporte aquaviário por navegação interior de produtos químicos perigosos, exceto lubrificantes e GLP","16.02","000001"],["1.0502.14.59","Serviços de transporte aquaviário por navegação interior de produtos perigosos não classificados em subposições anteriores","16.02","000001"],["1.0502.14.90","Serviços de transporte aquaviário por navegação interior de cargas especiais não classificados em subposições anteriores","16.02","000001"],["1.0502.19.00","Serviços de transporte aquaviário por navegação interior de cargas não classificados em subposições anteriores","16.02","000001"],["1.0502.21.10","Serviços de transporte aquaviário costeiro de cargas sólidas a granel","16.02","000001"],["1.0502.21.20","Serviços de transporte aquaviário costeiro de cargas líquidas ou liquefeitas, a granel","16.02","000001"],["1.0502.21.30","Serviços de transporte aquaviário costeiro de cargas gasosas a granel","16.02","000001"],["1.0502.22.10","Serviços de transporte aquaviário costeiro de carga solta, não unitizada","16.02","000001"],["1.0502.22.20","Serviços de transporte aquaviário costeiro de carga unitizada","16.02","000001"],["1.0502.22.30","Serviços de transporte aquaviário costeiro de carga frigorificada ou climatizada","16.02","000001"],["1.0502.23.10","Serviços de transporte aquaviário costeiro de cargas em contêineres frigorificados ou climatizados","16.02","000001"],["1.0502.23.20","Serviços de transporte aquaviário costeiro de cargas em contêineres não classificados em subposições anteriores","16.02","000001"],["1.0502.24.10","Serviços de transporte aquaviário costeiro de cargas vivas","16.02","000001"],["1.0502.24.20","Serviços de transporte aquaviário costeiro de mudanças domésticas e de mobília e outros objetos de escritório","16.02","000001"],["1.0502.24.30","Serviços de transporte aquaviário costeiro de cargas de grande porte","16.02","000001"],["1.0502.24.40","Serviços de transporte aquaviário costeiro de veículos","16.02","000001"],["1.0502.24.51","Serviços de transporte aquaviário costeiro de combustíveis, lubrificantes e GLP, inclusive apresentado em botijões metálicos","16.02","000001"],["1.0502.24.52","Serviços de transporte aquaviário costeiro de produtos químicos perigosos, exceto lubrificantes e GLP","16.02","000001"],["1.0502.24.59","Serviços de transporte aquaviário costeiro de produtos perigosos não classificados em subposições anteriores","16.02","000001"],["1.0505.10.00","Locação de veículos rodoviários de carga com operador","16.02","000001"],["1.0505.20.00","Locação de embarcações de carga com tripulação","16.02","000001"],["1.0505.30.00","Locação de aeronaves de carga com tripulação","16.02","000001"],["1.0608.40.00","Serviços de consolidação ou desconsolidação documental de cargas no transporte multimodal","17.01","000001"],["1.1001.40.00","Serviços de consultoria imobiliária","17.01","000001"],["1.1001.50.00","Serviços de assessoria de gestão condominial","17.01","000001"],["1.1001.90.00","Serviços imobiliários não classificados em subposições anteriores","17.01","000001"],["1.1303.10.00","Serviços de consultoria tributária para pessoas jurídicas","17.01","000001"],["1.1303.20.00","Serviços de consultoria tributária para pessoas físicas","17.01","000001"],["1.1401.11.00","Serviços de consultoria em gestão estratégica","17.01","000001"],["1.1401.13.00","Serviços de consultoria em gestão de recursos humanos","17.01","000001"],["1.1401.14.00","Serviços de consultoria em gestão de marketing","17.01","000001"],["1.1401.15.00","Serviços de consultoria em gestão operacional","17.01","000001"],["1.1401.16.00","Serviços de consultoria em gestão energética","17.01","000001"],["1.1401.17.00","Serviços de consultoria em gestão de cadeia logística","17.01","000001"],["1.1401.18.00","Serviços de consultoria em gestão hospitalar","17.01","000001"],["1.1401.19.00","Serviços de consultoria em gestão empresarial não classificados em subposições anteriores","17.01","000001"],["1.1401.39.00","Serviços de assessoria empresarial não classificados em subposições anteriores","17.01","000001"],["1.1410.10.00","Serviços de consultoria ambiental","17.01","000001"],["1.1410.90.00","Serviços de consultoria técnica e científica não classificados em subposições anteriores","17.01","000001"],["1.1412.00.00","Serviços para registros de marcas comerciais e de franquias empresariais, exceto as licenças de uso de direito","17.01","000001"],["1.1413.00.00","Serviços de prospecção de clientes","17.01","000001"],["1.1414.00.00","Serviços de prospecção de fornecedores","17.01","000001"],["1.1407.00.00","Pesquisas de mercado e serviços de pesquisa de opinião pública","17.01","000001"],["1.1411.00.00","Serviços de tradução e de intérpretes","17.02","000001"],["1.1806.39.00","Serviços de apoio às atividades empresariais por meio de telefone não classificados em subposições anteriores","17.02","000001"],["1.1806.40.00","Serviços combinados de escritório e apoio administrativo","17.02","000001"],["1.1806.52.00","Serviços de execução e envio de mala direta e de elaboração de listas de endereços","17.02","000001"],["1.1806.53.00","Serviços de preparação de documentos","17.02","000001"],["1.1806.59.00","Serviços especializados de apoio a escritório não classificados em subposições anteriores","17.02","000001"],["1.1401.29.00","Serviços de gestão não classificados em subposições anteriores","17.03","000001"],["1.1801.11.00","Serviço de recrutamento e seleção de profissionais executivos","17.04","000001"],["1.1801.12.00","Serviço de recrutamento e seleção de profissionais, exceto executivos","17.04","000001"],["1.1801.21.00","Serviços de fornecimento de mão de obra terceirizada, exceto temporária","17.05","000001"],["1.1801.22.00","Serviços de fornecimento de mão de obra temporária","17.05","000001"],["1.1801.29.00","Serviços de fornecimento de mão de obra não classificados em subposições anteriores","17.05","000001"],["1.1406.11.00","Serviços de campanhas publicitárias","17.06","000001"],["1.1406.12.00","Serviços de marketing direto e mala direta","17.06","000001"],["1.1406.19.00","Serviços de propaganda não classificados em subposições anteriores","17.06","000001"],["1.1110.00.00","Franquia","17.08","000001"],["1.1404.41.00","Serviços de análise laboratorial de solos, sementes e outros materiais propagativos, fitossanitários,água de produção, bromatologia e sanidade animal","17.09","200038"],["1.1404.42.00","Serviços de análise e de exames técnicos de propriedades físicas","17.09","000001"],["1.1404.43.00","Serviços de análise e de exames técnicos de sistemas elétricos e mecânicos","17.09","000001"],["1.1404.44.00","Serviços de inspeção técnica de veículos de transporte rodoviário","17.09","000001"],["1.1404.49.00","Serviços de análise e de exames técnicos não classificados em subposições anteriores","17.09","000001"],["1.0301.10.00","Fornecimento de refeições acompanhado de serviços de restaurante","17.11","000001"],["1.0301.31.00","Fornecimento de alimentação para eventos","17.11","000001"],["1.0301.39.00","Fornecimento de alimentação, incluindo refeições, sob contrato não classificado em subposições anteriores","17.11","000001"],["1.1001.11.00","Serviços de administração e locação de imóveis residenciais","17.12","200046"],["1.1001.12.90","Serviços de administração e locação de outros imóveis não residenciais","17.12","200046"],["1.1401.21.00","Serviços de gestão em processos de negócios","17.12","000001"],["1.1401.22.00","Serviços de gestão hospitalar","17.12","000001"],["1.1806.90.00","Serviços de apoio não classificados em subposições anteriores","17.13","000001"],["1.1301.10.00","Serviços de representação e consultoria jurídica criminal","17.14","200052"],["1.1301.20.00","Serviços de representação e consultoria jurídica em outras áreas do direito, exceto consultoria tributária","17.14","200052"],["1.1301.90.00","Serviços jurídicos não classificados em subposições anteriores","17.14","200052"],["1.1301.40.00","Serviços de arbitragem, conciliação e mediação","17.15","000001"],["1.1302.11.00","Serviços de auditoria contábil","17.16","200052"],["1.1302.19.00","Serviços de auditoria não classificados em subposições anteriores","17.16","000001"],["1.0906.30.00","Serviços atuariais","17.18","000001"],["1.1302.21.00","Serviços de contabilidade","17.19","200052"],["1.1302.22.00","Serviços de escrituração mercantil","17.19","200052"],["1.1302.23.00","Serviços de folha de pagamento","17.19","200052"],["1.0905.70.00","Serviços de classificação de risco (rating)","17.20","200052"],["1.0905.80.00","Serviços fiduciários","17.20","200052"],["1.1401.12.00","Serviços de consultoria em gestão financeira","17.20","200052"],["1.1415.00.00","Serviços profissionais, técnicos e gerenciais não classificados em posições anteriores","17.21","200052"],["1.0908.00.00","Fomento comercial (factoring)","17.23","000001"],["1.2205.14.00","Serviços de palestras e conferências","17.24","000001"],["1.1406.33.00","Venda de espaço para propaganda na rede mundial de computadores, exceto sob comissão","17.25","000001"],["1.1406.34.00","Venda de espaço para propaganda em mídia exterior, exceto sob comissão","17.25","000001"],["1.1406.39.00","Venda de espaço ou tempo para propaganda, exceto sob comissão não classificados em subposições anteriores","17.25","000001"],["1.0906.20.00","Serviços de perícia e avaliação de seguros e resseguros","18.01","000001"],["1.0605.10.00","Serviços de operação de portos e canais, exceto manuseio de cargas","20.01","000001"],["1.0605.20.00","Serviços de praticagem e de atracação","20.01","000001"],["1.0605.30.00","Serviços de salvamento de embarcações","20.01","000001"],["1.0605.40.00","Serviços de navegação de apoio","20.01","000001"],["1.0606.11.00","Serviços de operação aeroportuária, exceto manuseio de cargas","20.02","000001"],["1.0606.12.00","Serviços de controle de tráfego aéreo","20.02","000001"],["1.0606.20.00","Serviços de apoio ao transporte aeroespacial","20.02","000001"],["1.0603.00.00","Serviços de apoio ao transporte ferroviário","20.03","000001"],["1.0604.10.00","Serviços de estações rodoviárias","20.03","000001"],["1.0604.90.00","Serviços de apoio ao transporte rodoviário não classificados em subposições anteriores","20.03","000001"],["1.1304.00.00","Serviços notariais e de registro","21.01","000001"],["1.0604.21.00","Serviços de operação de rodovias","22.01","000002"],["1.0604.22.00","Serviços de operação de pontes e túneis","22.01","000002"],["1.1409.21.00","Serviços de desenho industrial de embalagens, expositores de loja e objetos promocionais para comunicação e vendas","23.01","000001"],["1.1409.22.00","Serviços de desenho industrial de produtos, utensílios, equipamentos, vestuário, calçados, ornamentos, joias e objetos pessoais","23.01","000001"],["1.1409.23.00","Serviços de desenho industrial de máquinas, equipamentos, acessórios e objetos de uso industrial de qualquer natureza","23.01","000001"],["1.1409.24.00","Serviços de desenho industrial de mobiliários e itens de decoração","23.01","000001"],["1.1409.25.00","Serviços de desenho industrial de utensílios e equipamentos eletrodomésticos e eletroeletrônicos","23.01","000001"],["1.1409.29.00","Serviços de desenho industrial não classificados em subposições anteriores","23.01","000001"],["1.1409.30.00","Serviços de design de marcas, imagens, objetos gráficos e digitais","23.01","000001"],["1.1409.90.00","Serviços especializados de design não classificados em subposições anteriores","23.01","000001"],["1.1405.30.00","Serviços funerários, de cremação e de embalsamamento de animais","25.01","000001"],["1.2603.00.00","Serviços funerários, de cremação e de embalsamamento","25.01","200029"],["1.0608.10.00","Serviços de coleta e entrega de cargas no transporte multimodal","26.01","000001"],["1.0701.00.00","Serviços postais e de telegrama","26.01","000001"],["1.0703.00.00","Serviços de remessas expressas","26.01","000001"],["1.1802.40.00","Serviços de carro-forte","26.01","000001"],["1.2304.11.00","Serviços de reabilitação vocacional para pessoas com deficiência","27.01","200052"],["1.2304.12.00","Serviços de reabilitação vocacional para desempregados","27.01","200052"],["1.2304.19.00","Serviço de reabilitação vocacional não classificados nas subposições anteriores","27.01","200052"],["1.2304.20.00","Serviços de orientação e aconselhamento relacionados a crianças e adolescentes","27.01","200052"],["1.2304.90.00","Serviços de assistência social sem acomodação não classificados em subposições anteriores","27.01","200052"],["1.0902.10.00","Serviços de valoração de ativos","28.01","000001"],["1.1705.10.00","Serviços de biblioteca","29.01","200052"],["1.1705.20.00","Serviços de arquivo","29.01","200052"],["1.0204.00.00","Serviços de despacho aduaneiro","33.01","000001"],["1.1802.10.00","Serviços de investigação","34.01","000001"],["1.1401.31.00","Serviços de assessoria de imprensa","35.01","000001"],["1.1401.32.00","Serviços de relações públicas","35.01","200052"],["1.1404.30.00","Serviços meteorológicos e de previsão do tempo","36.01","000001"],["1.1806.81.00","Serviços de agenciamento de modelos","37.01","000001"],["1.1806.82.00","Serviços de agenciamento de artistas","37.01","000001"],["1.1806.83.00","Serviços de agenciamento de atletas","37.01","000001"],["1.2506.00.00","Serviços fornecidos por atletas e desportistas, por conta própria, e serviços de apoio relacionados com desportes e recreação desportiva","37.01","000001"],["1.2504.11.00","Serviços de museus","38.01","200039"],["1.2504.12.00","Serviços de preservação e operação de locais e construções históricas","38.01","000001"],["1.1109.90.00","Cessão definitiva de outros direitos não classificada em subposições anteriores","40.01","000001"],["1.2503.20.00","Serviços de autores, compositores, escultores, pintores e outros artistas, exceto os de atuação artística","40.01","000001"]];

// [cnae, item_lc116] - 803 linhas - fonte: tabela oficial CNAE x Item da Lista (LC 116/2003)
var CNAE_ITEM_DATA = [["6201-5/00","01.01"],["6201-5/00","01.02"],["6190-6/01","01.03"],["1830-0/03","01.03"],["6311-9/00","01.03"],["6201-5/00","01.04"],["6202-3/00","01.04"],["6203-1/00","01.04"],["6202-3/00","01.05"],["6204-0/00","01.06"],["6209-1/00","01.07"],["6201-5/00","01.08"],["6319-4/00","01.08"],["7210-0/00","02.01"],["7220-7/00","02.01"],["8211-3/00","03.03"],["8230-0/02","03.03"],["9003-5/00","03.03"],["3520-4/02","03.04"],["3514-0/00","03.04"],["4911-6/00","03.04"],["5221-4/00","03.04"],["4399-1/02","03.05"],["4399-1/04","03.05"],["7732-2/02","03.05"],["7739-0/03","03.05"],["8630-5/02","04.01"],["8630-5/03","04.01"],["8630-5/99","04.01"],["8640-2/01","04.02"],["8640-2/02","04.02"],["8640-2/04","04.02"],["8640-2/05","04.02"],["8640-2/06","04.02"],["8640-2/07","04.02"],["8640-2/08","04.02"],["8640-2/09","04.02"],["8640-2/10","04.02"],["8640-2/11","04.02"],["8640-2/99","04.02"],["8610-1/01","04.03"],["8610-1/02","04.03"],["8630-5/01","04.03"],["8630-5/02","04.03"],["8630-5/06","04.03"],["8640-2/01","04.03"],["8640-2/02","04.03"],["8640-2/03","04.03"],["8640-2/13","04.03"],["8650-0/99","04.04"],["8690-9/01","04.05"],["8690-9/03","04.05"],["8650-0/01","04.06"],["4771-7/02","04.07"],["4771-7/03","04.07"],["8650-0/04","04.08"],["8650-0/05","04.08"],["8650-0/06","04.08"],["8640-2/12","04.09"],["8690-9/01","04.09"],["8640-2/99","04.09"],["8650-0/07","04.09"],["8690-9/99","04.09"],["8650-0/02","04.10"],["8690-9/99","04.11"],["8630-5/04","04.12"],["8650-0/99","04.13"],["3250-7/03","04.14"],["3250-7/06","04.14"],["8650-0/03","04.15"],["8650-0/03","04.16"],["8711-5/01","04.17"],["8711-5/02","04.17"],["8711-5/03","04.17"],["8711-5/04","04.17"],["8711-5/05","04.17"],["8720-4/01","04.17"],["8730-1/01","04.17"],["8730-1/02","04.17"],["8720-4/99","04.17"],["8730-1/99","04.17"],["8630-5/07","04.18"],["8630-5/07","04.19"],["8640-2/12","04.19"],["8640-2/14","04.19"],["8690-9/02","04.19"],["8690-9/99","04.19"],["8640-2/02","04.20"],["8621-6/01","04.21"],["8621-6/02","04.21"],["8622-4/00","04.21"],["8712-3/00","04.21"],["6550-2/00","04.22"],["6550-2/00","04.23"],["7490-1/03","05.01"],["7500-1/00","05.01"],["7500-1/00","05.02"],["7500-1/00","05.03"],["0162-8/01","05.04"],["7500-1/00","05.05"],["7500-1/00","05.06"],["7500-1/00","05.07"],["8011-1/02","05.08"],["9609-2/03","05.08"],["0162-8/02","05.08"],["0162-8/03","05.08"],["0162-8/99","05.08"],["6550-2/00","05.09"],["8690-9/04","06.01"],["9602-5/01","06.01"],["9602-5/02","06.01"],["9602-5/02","06.02"],["9609-2/01","06.02"],["9609-2/06","06.02"],["9609-2/99","06.02"],["9609-2/01","06.03"],["9609-2/05","06.03"],["8592-9/01","06.04"],["8591-1/00","06.04"],["9313-1/00","06.04"],["9609-2/01","06.05"],["2391-5/01","07.01"],["7119-7/01","07.01"],["7119-7/02","07.01"],["7111-1/00","07.01"],["7112-0/00","07.01"],["2330-3/05","07.02"],["4211-1/01","07.02"],["4211-1/02","07.02"],["4221-9/01","07.02"],["4221-9/02","07.02"],["4221-9/04","07.02"],["4222-7/01","07.02"],["4222-7/02","07.02"],["4292-8/01","07.02"],["4292-8/02","07.02"],["4299-5/01","07.02"],["4311-8/02","07.02"],["4322-3/01","07.02"],["4322-3/02","07.02"],["4322-3/03","07.02"],["4329-1/03","07.02"],["4330-4/01","07.02"],["4330-4/03","07.02"],["4330-4/04","07.02"],["4330-4/05","07.02"],["4399-1/01","07.02"],["4399-1/03","07.02"],["4399-1/05","07.02"],["3321-0/00","07.02"],["4120-4/00","07.02"],["4212-0/00","07.02"],["4213-8/00","07.02"],["4223-5/00","07.02"],["4299-5/99","07.02"],["4312-6/00","07.02"],["4313-4/00","07.02"],["4319-3/00","07.02"],["4321-5/00","07.02"],["4329-1/99","07.02"],["4330-4/99","07.02"],["4391-6/00","07.02"],["4399-1/99","07.02"],["7112-0/00","07.03"],["4311-8/01","07.04"],["4211-1/01","07.05"],["4211-1/02","07.05"],["4221-9/03","07.05"],["4221-9/05","07.05"],["4322-3/01","07.05"],["4330-4/03","07.05"],["4330-4/04","07.05"],["4330-4/05","07.05"],["4399-1/03","07.05"],["9102-3/02","07.05"],["4120-4/00","07.05"],["4212-0/00","07.05"],["4321-5/00","07.05"],["4330-4/99","07.05"],["4399-1/99","07.05"],["4330-4/02","07.06"],["4330-4/03","07.06"],["4330-4/05","07.06"],["4330-4/99","07.06"],["4330-4/05","07.07"],["4329-1/05","07.08"],["4330-4/05","07.08"],["3831-9/01","07.09"],["3839-4/01","07.09"],["3811-4/00","07.09"],["3812-2/00","07.09"],["3821-1/00","07.09"],["3822-0/00","07.09"],["3831-9/99","07.09"],["3832-7/00","07.09"],["3839-4/99","07.09"],["8129-0/00","07.09"],["3702-9/00","07.10"],["8121-4/00","07.10"],["8129-0/00","07.10"],["7410-2/02","07.11"],["8130-3/00","07.11"],["3821-1/00","07.12"],["3822-0/00","07.12"],["0161-0/01","07.13"],["0162-8/99","07.13"],["8122-2/00","07.13"],["0161-0/03","07.16"],["0220-9/06","07.16"],["0230-6/00","07.16"],["4299-5/99","07.17"],["3900-5/00","07.18"],["4291-0/00","07.18"],["7112-0/00","07.19"],["7119-7/01","07.20"],["7119-7/02","07.20"],["7119-7/99","07.20"],["7490-1/02","07.21"],["0910-6/00","07.21"],["0990-4/01","07.21"],["0990-4/02","07.21"],["0990-4/03","07.21"],["7490-1/99","07.22"],["8511-2/00","08.01"],["8512-1/00","08.01"],["8513-9/00","08.01"],["8520-1/00","08.01"],["8531-7/00","08.01"],["8532-5/00","08.01"],["8533-3/00","08.01"],["8541-4/00","08.01"],["8542-2/00","08.01"],["8550-3/02","08.02"],["8592-9/02","08.02"],["8592-9/03","08.02"],["8599-6/01","08.02"],["8599-6/02","08.02"],["8599-6/03","08.02"],["8599-6/04","08.02"],["8599-6/05","08.02"],["8592-9/99","08.02"],["8593-7/00","08.02"],["8599-6/99","08.02"],["9312-3/00","08.02"],["5510-8/01","09.01"],["5510-8/02","09.01"],["5510-8/03","09.01"],["5590-6/01","09.01"],["5590-6/02","09.01"],["5590-6/03","09.01"],["5590-6/99","09.01"],["4929-9/03","09.02"],["4929-9/04","09.02"],["7911-2/00","09.02"],["7912-1/00","09.02"],["7990-2/00","09.02"],["7912-1/00","09.03"],["6612-6/03","10.01"],["6622-3/00","10.01"],["6612-6/01","10.02"],["6612-6/02","10.02"],["6612-6/04","10.02"],["6612-6/05","10.02"],["7490-1/04","10.02"],["7490-1/05","10.02"],["8299-7/05","10.02"],["9609-2/02","10.02"],["6619-3/99","10.02"],["6622-3/00","10.02"],["6022-5/02","10.03"],["6911-7/03","10.03"],["7490-1/05","10.03"],["5811-5/00","10.03"],["7490-1/04","10.04"],["4512-9/02","10.05"],["5250-8/03","10.05"],["6612-6/05","10.05"],["6821-8/01","10.05"],["6821-8/02","10.05"],["8299-7/02","10.05"],["5232-0/00","10.06"],["6391-7/00","10.07"],["7490-1/04","10.08"],["7312-2/00","10.08"],["4512-9/01","10.09"],["4512-9/02","10.09"],["4530-7/06","10.09"],["4542-1/01","10.09"],["4618-4/01","10.09"],["4618-4/02","10.09"],["4618-4/03","10.09"],["6619-3/03","10.09"],["4611-7/00","10.09"],["4612-5/00","10.09"],["4613-3/00","10.09"],["4614-1/00","10.09"],["4615-0/00","10.09"],["4616-8/00","10.09"],["4617-6/00","10.09"],["4618-4/99","10.09"],["4619-2/00","10.09"],["5913-8/00","10.10"],["5223-1/00","11.01"],["5240-1/99","11.01"],["9329-8/99","11.01"],["8011-1/01","11.02"],["8020-0/00","11.02"],["5229-0/99","11.03"],["4930-2/04","11.04"],["5211-7/01","11.04"],["5211-7/02","11.04"],["5231-1/02","11.04"],["5250-8/04","11.04"],["5211-7/99","11.04"],["5212-5/00","11.04"],["9001-9/01","12.01"],["9001-9/99","12.01"],["5914-6/00","12.02"],["9001-9/04","12.03"],["9001-9/99","12.03"],["9001-9/99","12.04"],["9001-9/99","12.05"],["9103-1/00","12.05"],["9321-2/00","12.05"],["9329-8/01","12.06"],["9001-9/02","12.07"],["9001-9/03","12.07"],["8230-0/01","12.08"],["9329-8/02","12.09"],["9329-8/03","12.09"],["9329-8/04","12.09"],["9200-3/99","12.09"],["9001-9/05","12.10"],["9200-3/02","12.10"],["9319-1/01","12.11"],["9319-1/99","12.11"],["9001-9/02","12.12"],["9001-9/01","12.13"],["9001-9/02","12.13"],["9001-9/03","12.13"],["9001-9/04","12.13"],["5911-1/99","12.13"],["6021-7/00","12.13"],["9001-9/02","12.14"],["9001-9/06","12.14"],["9001-9/02","12.15"],["9001-9/03","12.15"],["9493-6/00","12.15"],["9001-9/02","12.16"],["5914-6/00","12.16"],["9329-8/99","12.17"],["1830-0/01","13.02"],["5912-0/01","13.02"],["5912-0/02","13.02"],["5920-1/00","13.02"],["5911-1/01","13.03"],["9609-2/04","13.03"],["1830-0/02","13.03"],["5912-0/99","13.03"],["7420-0/01","13.03"],["7420-0/02","13.03"],["7420-0/03","13.03"],["7420-0/04","13.03"],["8219-9/01","13.04"],["7420-0/05","13.04"],["1741-9/01","13.05"],["1811-3/01","13.05"],["1811-3/02","13.05"],["1812-1/00","13.05"],["1813-0/01","13.05"],["1813-0/99","13.05"],["1821-1/00","13.05"],["1822-9/01","13.05"],["2930-1/03","14.01"],["3312-1/02","14.01"],["3312-1/03","14.01"],["3312-1/04","14.01"],["3313-9/01","14.01"],["3313-9/02","14.01"],["3314-7/01","14.01"],["3314-7/02","14.01"],["3314-7/03","14.01"],["3314-7/04","14.01"],["3314-7/05","14.01"],["3314-7/06","14.01"],["3314-7/07","14.01"],["3314-7/08","14.01"],["3314-7/09","14.01"],["3314-7/10","14.01"],["3314-7/11","14.01"],["3314-7/12","14.01"],["3314-7/13","14.01"],["3314-7/14","14.01"],["3314-7/15","14.01"],["3314-7/16","14.01"],["3314-7/17","14.01"],["3314-7/18","14.01"],["3314-7/19","14.01"],["3314-7/20","14.01"],["3314-7/21","14.01"],["3314-7/22","14.01"],["3316-3/01","14.01"],["3316-3/02","14.01"],["3317-1/01","14.01"],["3317-1/02","14.01"],["4322-3/02","14.01"],["4322-3/03","14.01"],["4329-1/03","14.01"],["4329-1/04","14.01"],["4329-1/05","14.01"],["4751-2/02","14.01"],["7490-1/02","14.01"],["9529-1/01","14.01"],["9529-1/03","14.01"],["9529-1/04","14.01"],["9529-1/05","14.01"],["9529-1/06","14.01"],["3311-2/00","14.01"],["3313-9/99","14.01"],["3314-7/99","14.01"],["3315-5/00","14.01"],["3319-8/00","14.01"],["4520-0/01","14.01"],["4520-0/03","14.01"],["4520-0/04","14.01"],["4520-0/05","14.01"],["4520-0/07","14.01"],["4543-9/00","14.01"],["6190-6/99","14.01"],["9512-6/00","14.01"],["9521-5/00","14.01"],["9529-1/99","14.01"],["9609-2/99","14.01"],["9511-8/00","14.02"],["9512-6/00","14.02"],["2950-6/00","14.03"],["2212-9/00","14.04"],["4520-0/06","14.04"],["2391-5/01","14.05"],["2391-5/02","14.05"],["2391-5/03","14.05"],["2599-3/02","14.05"],["2722-8/02","14.05"],["3250-7/09","14.05"],["9002-7/02","14.05"],["9529-1/05","14.05"],["1340-5/01","14.05"],["2539-0/01","14.05"],["2539-0/02","14.05"],["4520-0/02","14.05"],["4520-0/05","14.05"],["8292-0/00","14.05"],["9609-2/99","14.05"],["4322-3/03","14.06"],["4329-1/02","14.06"],["4329-1/04","14.06"],["3321-0/00","14.06"],["3329-5/99","14.06"],["6190-6/99","14.06"],["4330-4/03","14.07"],["1629-3/01","14.07"],["1822-9/99","14.08"],["1340-5/99","14.09"],["1411-8/02","14.09"],["1412-6/02","14.09"],["1412-6/03","14.09"],["1413-4/01","14.09"],["1413-4/02","14.09"],["1413-4/03","14.09"],["1531-9/02","14.09"],["9529-1/99","14.09"],["9601-7/01","14.10"],["9601-7/02","14.10"],["9601-7/03","14.10"],["1340-5/02","14.10"],["1340-5/99","14.10"],["9529-1/05","14.11"],["4520-0/08","14.11"],["4520-0/02","14.12"],["2399-1/01","14.13"],["2599-3/01","14.13"],["3329-5/01","14.13"],["4330-4/02","14.13"],["1622-6/99","14.13"],["2512-8/00","14.13"],["2542-0/00","14.13"],["6424-7/01","15.01"],["6424-7/02","15.01"],["6424-7/03","15.01"],["6424-7/04","15.01"],["6435-2/03","15.01"],["6470-1/01","15.01"],["6470-1/02","15.01"],["6470-1/03","15.01"],["6499-9/01","15.01"],["6499-9/02","15.01"],["6619-3/05","15.01"],["6421-2/00","15.01"],["6422-1/00","15.01"],["6423-9/00","15.01"],["6431-0/00","15.01"],["6432-8/00","15.01"],["6433-6/00","15.01"],["6434-4/00","15.01"],["6613-4/00","15.01"],["6424-7/01","15.02"],["6424-7/03","15.02"],["6424-7/04","15.02"],["6435-2/02","15.02"],["6421-2/00","15.02"],["6422-1/00","15.02"],["6423-9/00","15.02"],["6433-6/00","15.02"],["6450-6/00","15.02"],["6424-7/01","15.03"],["6424-7/03","15.03"],["6424-7/04","15.03"],["6421-2/00","15.03"],["6422-1/00","15.03"],["6423-9/00","15.03"],["6424-7/01","15.04"],["6424-7/03","15.04"],["6424-7/04","15.04"],["6421-2/00","15.04"],["6422-1/00","15.04"],["6423-9/00","15.04"],["6424-7/01","15.05"],["6424-7/03","15.05"],["6424-7/04","15.05"],["6421-2/00","15.05"],["6422-1/00","15.05"],["6423-9/00","15.05"],["6424-7/01","15.06"],["6424-7/03","15.06"],["6424-7/04","15.06"],["6421-2/00","15.06"],["6422-1/00","15.06"],["6423-9/00","15.06"],["6431-0/00","15.06"],["6424-7/01","15.07"],["6424-7/03","15.07"],["6424-7/04","15.07"],["6435-2/03","15.07"],["6619-3/04","15.07"],["6421-2/00","15.07"],["6422-1/00","15.07"],["6423-9/00","15.07"],["6432-8/00","15.07"],["6433-6/00","15.07"],["6450-6/00","15.07"],["6424-7/01","15.08"],["6424-7/03","15.08"],["6424-7/04","15.08"],["6435-2/02","15.08"],["6435-2/03","15.08"],["6499-9/05","15.08"],["6421-2/00","15.08"],["6422-1/00","15.08"],["6423-9/00","15.08"],["6431-0/00","15.08"],["6432-8/00","15.08"],["6433-6/00","15.08"],["6436-1/00","15.08"],["6437-9/00","15.08"],["6424-7/01","15.09"],["6424-7/03","15.09"],["6424-7/04","15.09"],["6421-2/00","15.09"],["6422-1/00","15.09"],["6423-9/00","15.09"],["6431-0/00","15.09"],["6440-9/00","15.09"],["6424-7/01","15.10"],["6424-7/03","15.10"],["6424-7/04","15.10"],["6619-3/01","15.10"],["6619-3/02","15.10"],["6421-2/00","15.10"],["6422-1/00","15.10"],["6423-9/00","15.10"],["6424-7/01","15.11"],["6424-7/03","15.11"],["6424-7/04","15.11"],["6421-2/00","15.11"],["6422-1/00","15.11"],["6423-9/00","15.11"],["6424-7/01","15.12"],["6424-7/03","15.12"],["6424-7/04","15.12"],["6619-3/01","15.12"],["6421-2/00","15.12"],["6422-1/00","15.12"],["6423-9/00","15.12"],["6431-0/00","15.12"],["6499-9/99","15.12"],["6424-7/01","15.13"],["6424-7/03","15.13"],["6424-7/04","15.13"],["6438-7/01","15.13"],["6421-2/00","15.13"],["6422-1/00","15.13"],["6423-9/00","15.13"],["6438-7/99","15.13"],["6424-7/01","15.14"],["6424-7/03","15.14"],["6424-7/04","15.14"],["6421-2/00","15.14"],["6422-1/00","15.14"],["6423-9/00","15.14"],["6424-7/01","15.15"],["6424-7/03","15.15"],["6424-7/04","15.15"],["6421-2/00","15.15"],["6422-1/00","15.15"],["6423-9/00","15.15"],["6424-7/01","15.16"],["6424-7/03","15.16"],["6424-7/04","15.16"],["6421-2/00","15.16"],["6422-1/00","15.16"],["6423-9/00","15.16"],["6432-8/00","15.16"],["6433-6/00","15.16"],["6424-7/01","15.17"],["6424-7/03","15.17"],["6424-7/04","15.17"],["6421-2/00","15.17"],["6422-1/00","15.17"],["6423-9/00","15.17"],["6424-7/01","15.18"],["6424-7/03","15.18"],["6424-7/04","15.18"],["6435-2/01","15.18"],["6435-2/02","15.18"],["6435-2/03","15.18"],["6499-9/04","15.18"],["6421-2/00","15.18"],["6422-1/00","15.18"],["6423-9/00","15.18"],["6431-0/00","15.18"],["4912-4/02","16.01"],["4912-4/03","16.01"],["4921-3/01","16.01"],["4929-9/01","16.01"],["4930-2/01","16.01"],["4930-2/03","16.01"],["4930-2/04","16.01"],["5021-1/01","16.01"],["5091-2/01","16.01"],["5099-8/01","16.01"],["4923-0/01","16.01"],["4923-0/02","16.01"],["4924-8/00","16.01"],["4940-0/00","16.01"],["4950-7/00","16.01"],["5022-0/01","16.01"],["5112-9/99","16.01"],["5229-0/02","16.01"],["6920-6/02","17.01"],["7490-1/03","17.01"],["8299-7/01","17.01"],["0162-8/99","17.01"],["0311-6/04","17.01"],["0312-4/04","17.01"],["0321-3/05","17.01"],["0322-1/07","17.01"],["6399-2/00","17.01"],["7020-4/00","17.01"],["7319-0/04","17.01"],["7320-3/00","17.01"],["7490-1/99","17.01"],["9430-8/00","17.01"],["9609-2/99","17.01"],["7490-1/01","17.02"],["5229-0/01","17.02"],["5811-5/00","17.02"],["5812-3/00","17.02"],["5813-1/00","17.02"],["5819-1/00","17.02"],["5821-2/00","17.02"],["5822-1/00","17.02"],["5823-9/00","17.02"],["5829-8/00","17.02"],["7490-1/99","17.02"],["8211-3/00","17.02"],["8219-9/99","17.02"],["8220-2/00","17.02"],["8299-7/99","17.02"],["7020-4/00","17.03"],["8211-3/00","17.03"],["7810-8/00","17.04"],["7820-5/00","17.05"],["7830-2/00","17.05"],["8111-7/00","17.05"],["5911-1/02","17.06"],["7311-4/00","17.06"],["7319-0/01","17.06"],["7319-0/02","17.06"],["7319-0/03","17.06"],["7319-0/99","17.06"],["5310-5/02","17.08"],["7740-3/00","17.08"],["6621-5/01","17.09"],["6911-7/02","17.09"],["6920-6/02","17.09"],["7119-7/04","17.09"],["7112-0/00","17.09"],["7120-1/00","17.09"],["8230-0/01","17.10"],["5620-1/02","17.11"],["6611-8/01","17.12"],["6611-8/02","17.12"],["6611-8/03","17.12"],["6611-8/04","17.12"],["6493-0/00","17.12"],["6630-4/00","17.12"],["6822-6/00","17.12"],["7740-3/00","17.12"],["8660-7/00","17.12"],["9609-2/99","17.12"],["8299-7/04","17.13"],["6911-7/01","17.14"],["6911-7/02","17.15"],["6621-5/01","17.16"],["6621-5/02","17.16"],["6920-6/02","17.16"],["7020-4/00","17.17"],["6621-5/02","17.18"],["6920-6/01","17.19"],["6612-6/05","17.20"],["6621-5/02","17.20"],["7020-4/00","17.20"],["7320-3/00","17.21"],["7490-1/99","17.21"],["8291-1/00","17.22"],["6491-3/00","17.23"],["8599-6/99","17.24"],["6621-5/01","18.01"],["6629-1/00","18.01"],["8299-7/06","19.01"],["9200-3/01","19.01"],["9200-3/99","19.01"],["5030-1/01","20.01"],["5030-1/02","20.01"],["5211-7/01","20.01"],["5231-1/01","20.01"],["5231-1/02","20.01"],["5250-8/05","20.01"],["0311-6/04","20.01"],["0312-4/04","20.01"],["0321-3/05","20.01"],["5211-7/99","20.01"],["5212-5/00","20.01"],["5239-7/00","20.01"],["5240-1/01","20.02"],["5250-8/05","20.02"],["5240-1/99","20.02"],["5250-8/05","20.03"],["5222-2/00","20.03"],["6912-5/00","21.01"],["5221-4/00","22.01"],["7410-2/01","23.01"],["7490-1/99","23.01"],["4329-1/01","24.01"],["8299-7/03","24.01"],["9529-1/02","24.01"],["3299-0/03","24.01"],["3299-0/04","24.01"],["9603-3/03","25.01"],["9603-3/04","25.01"],["9603-3/05","25.01"],["9603-3/99","25.01"],["9603-3/02","25.02"],["6511-1/02","25.03"],["9603-3/01","25.04"],["5310-5/01","26.01"],["5310-5/02","26.01"],["5320-2/01","26.01"],["5320-2/02","26.01"],["8012-9/00","26.01"],["8800-6/00","27.01"],["6621-5/01","28.01"],["6821-8/01","28.01"],["7112-0/00","28.01"],["7490-1/99","28.01"],["9101-5/00","29.01"],["7740-3/00","3.02"],["9311-5/00","3.03"],["8640-2/02","30.01"],["7210-0/00","30.01"],["7119-7/99","31.01"],["7119-7/03",""],["5250-8/01","33.01"],["5250-8/02","33.01"],["8030-7/00","34.01"],["9002-7/01","35.01"],["7020-4/00","35.01"],["7490-1/99","36.01"],["7490-1/05","37.01"],["9102-3/01","38.01"],["3211-6/01","39.01"],["7410-2/01","39.01"],["9002-7/01","40.01"]];

// [cclasstrib, nome_categoria, reducao_pct_ou_vazio, observacao_confianca] - 20 linhas
// Pesquisa legislativa feita nesta analise: 5 categorias com redução confirmada em
// dispositivo legal especifico (ver "Observação"); as demais 15 estao marcadas
// "Verificar" - nao confirmadas, usar com cautela ate validacao juridica.
var CATEGORIAS_CCLASSTRIB = [["000001","Situações tributadas integralmente pelo IBS e CBS.",0.0,"Confirmado - Regime Geral, sem redução."],["200052","Prestação de serviços de profissões intelectuais",0.3,"Confirmado - Art. 142, LCP 214/2025 (profissões regulamentadas por conselho: advocacia, contabilidade, engenharia etc.)."],["200046","Operações com bens imóveis",0.5,"Confirmado - Art. 261, caput, LCP 214/2025 (alienação, cessão de direitos reais, administração/intermediação, construção civil)."],["200027","Operações de locação, cessão onerosa e arrendamento de bens imóveis",0.7,"Confirmado - Art. 261, parágrafo único, LCP 214/2025."],["200029","Fornecimento dos serviços de saúde humana (Anexo III)",0.6,"Confirmado - Art. 128, II e Art. 130, LCP 214/2025."],["200028","Fornecimento dos serviços de educação (Anexo II)","","Verificar - suspeita de 60%, mesma lógica de saúde, requer checagem do Anexo II específico."],["200038","Fornecimento dos insumos agropecuários e aquícolas (Anexo IX)","","Verificar - Anexo IX possui regras próprias (pode incluir alíquota zero para itens da cesta básica)."],["200039","Produções nacionais artísticas/culturais (Anexo X)","","Verificar - não confirmado nesta análise."],["200041","Fornecimento de serviço de educação desportiva","","Verificar - não confirmado nesta análise."],["200044","Segurança da informação e cibernética (sócio brasileiro, Anexo XI)","","Verificar - regime específico do Anexo XI, alíquota própria não confirmada."],["200048","Hotelaria, Parques de Diversão e Parques Temáticos","","Verificar - regime específico de turismo, não confirmado."],["200051","Agências de Turismo","","Regime específico: base de cálculo pela comissão, não é redução percentual simples."],["010002","Operações do serviço financeiro","","Regime específico: tributação sobre margem/spread, não é redução percentual simples."],["011001","Planos de assistência funerária","","Regime específico - base de cálculo diferenciada, não confirmado."],["011002","Planos de assistência à saúde","","Regime específico - base de cálculo diferenciada (Art. 34-ss), não é redução percentual simples."],["011003","Intermediação de planos de assistência à saúde","","Regime específico - não confirmado nesta análise."],["011005","Planos de assistência à saúde de animais domésticos","","Regime específico - não confirmado nesta análise."],["400001","Transporte público coletivo de passageiros rodoviário e metroviário","","Verificar - possível isenção/redução a 100%, não confirmado."],["200021","Transporte público coletivo de passageiros ferroviário e hidroviário","","Verificar - possível isenção/redução a 100%, não confirmado."],["000002","Exploração de via","","Verificar - não confirmado nesta análise."]];

// [setor_aba_origem, reducao_pct] - fallback usado enquanto CNAE/NBS nao preenchidos
var SETOR_FALLBACK_DATA = [["Comercio",0.0],["Saude",0.6],["ProfLiberal",0.3],["Servico",0.0]];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 Simulação IBS/CBS')
    .addSubMenu(ui.createMenu('📥 Extração de Dados')
        .addItem('🔍 Extrair Dados do Extrato (Aba Atual)', 'extrairDadosExtrato_AbaAtual'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📄 Geração de PDF')
        .addItem('Gerar PDFs da Aba Atual (Simples Nacional)', 'gerarPDF_AbaAtual'))
    .addSeparator()
    .addItem('🏗️ Construir Planilha Consolidada (executar 1x)', 'buildConsolidado')
    .addSeparator()
    .addItem('📝 Atualizar Modelo de Comunicado', 'reescreverModeloComunicado')
    .addToUi();
}

function criarOuLimparAba_(ss, nome) {
  var sh = ss.getSheetByName(nome);
  if (sh) { sh.clear(); sh.clearFormats(); return sh; }
  return ss.insertSheet(nome);
}

/**
 * Constrói a planilha consolidada a partir das 4 abas de produção
 * (Comercio, Saude, ProfLiberal, Servico). Cria/reconstrói as abas:
 * "Simples Nacional" (dados unificados A-T + Setor + fórmulas de redução
 * e projeção 2033), "CNAE x NBS" (auditoria de conformidade cadastral),
 * "Ref_NBS", "Ref_CNAE_Item" (bases oficiais de referência) e
 * "Banco_Dados" (parâmetros editáveis). Não apaga nem altera as 4 abas
 * originais. Idempotente: pode ser executada novamente (reconstrói as
 * abas consolidadas do zero, mas relê sempre as 4 abas originais).
 */
function buildConsolidado() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Construir Planilha Consolidada',
    "Isto vai (re)criar as abas 'Simples Nacional', 'CNAE x NBS', 'Ref_NBS', 'Ref_CNAE_Item', 'Banco_Dados' e 'DASHBOARD' a partir dos dados atuais das abas Comercio, Saude, ProfLiberal e Servico. Abas com esses nomes já existentes serão limpas e reconstruídas (as 4 abas de origem NÃO são alteradas). Deseja continuar?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Coleta as linhas das 4 abas de origem (colunas A-T = 20 colunas, mais a
  //    coluna U "Valor Híbrido Informado (R$)" - campo manual de override)
  var linhasTotais = [];
  var valoresHibridos = [];
  ABAS_ORIGEM_CONSOLIDACAO.forEach(function (nomeAba) {
    var sh = ss.getSheetByName(nomeAba);
    if (!sh) return;
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    var dados = sh.getRange(2, 1, lastRow - 1, 21).getValues();
    dados.forEach(function (row) {
      if (String(row[0]).trim() === '' && String(row[1]).trim() === '') return; // linha vazia
      linhasTotais.push(row.slice(0, 20).concat([nomeAba]));
      valoresHibridos.push([row[20] || '']);
    });
  });
  var total = linhasTotais.length;
  if (total === 0) { ui.alert('Nenhuma linha encontrada nas abas de origem (Comercio, Saude, ProfLiberal, Servico).'); return; }

  // 2. Aba "Simples Nacional"
  var sn = criarOuLimparAba_(ss, 'Simples Nacional');
  sn.getRange(1, 1, 1, SN_HEADERS.length).setValues([SN_HEADERS]);
  sn.getRange(1, 1, 1, SN_HEADERS.length).setFontWeight('bold').setBackground('#2E5395').setFontColor('#ffffff').setWrap(true);
  sn.getRange(2, 1, total, 21).setValues(linhasTotais);

  var formulasV = [], formulasW = [], formulasX = [], formulasY = [], obsZ = [];
  for (var i = 0; i < total; i++) {
    var r = i + 2;
    formulasV.push(['=IFERROR(IF(VLOOKUP($A' + r + ',\'CNAE x NBS\'!$A:$H,8,FALSE)="",VLOOKUP($U' + r + ',Banco_Dados!$A$9:$B$12,2,FALSE),VLOOKUP($A' + r + ',\'CNAE x NBS\'!$A:$H,8,FALSE)),VLOOKUP($U' + r + ',Banco_Dados!$A$9:$B$12,2,FALSE))']);
    formulasW.push(['=IFERROR(IF(VLOOKUP($A' + r + ',\'CNAE x NBS\'!$A:$H,8,FALSE)<>"","CNAE x NBS (cClassTrib "&VLOOKUP($A' + r + ',\'CNAE x NBS\'!$A:$G,7,FALSE)&")","Fallback setorial ("&$U' + r + '&")"),"Fallback setorial ("&$U' + r + '&")")']);
    formulasX.push(['=IFERROR($D' + r + '*(Banco_Dados!$B$4+Banco_Dados!$B$5)*(1-$V' + r + '),"")']);
    formulasY.push(['=IFERROR($X' + r + '-$G' + r + '*12,"")']);
    var row = linhasTotais[i];
    var pendente = !row[2] && !row[3] && !row[6];
    obsZ.push([pendente ? '⚠️ Dados pendentes de extração (usar menu "Extrair Dados do Extrato")' : '']);
  }
  sn.getRange(2, 22, total, 1).setFormulas(formulasV);
  sn.getRange(2, 23, total, 1).setFormulas(formulasW);
  sn.getRange(2, 24, total, 1).setFormulas(formulasX);
  sn.getRange(2, 25, total, 1).setFormulas(formulasY);
  sn.getRange(2, 26, total, 1).setValues(obsZ);
  sn.getRange(2, 27, total, 1).setValues(valoresHibridos);
  sn.getRange(2, 16, total, 1).insertCheckboxes();
  sn.getRange(2, 22, total, 1).setNumberFormat('0%');
  sn.getRange(2, 24, total, 2).setNumberFormat('R$ #,##0.00');
  sn.setFrozenRows(1);
  sn.setFrozenColumns(2);
  sn.autoResizeColumns(1, 27);

  // 3. Aba "Ref_NBS" (base oficial NBS x Item LC116 x cClassTrib)
  var refNbs = criarOuLimparAba_(ss, 'Ref_NBS');
  refNbs.getRange(1, 1, 1, 4).setValues([['NBS Código', 'Descrição NBS', 'Item LC 116', 'cClassTrib']]);
  refNbs.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#2E5395').setFontColor('#ffffff');
  refNbs.getRange(2, 1, REF_NBS_DATA.length, 4).setValues(REF_NBS_DATA);
  refNbs.setFrozenRows(1);
  refNbs.autoResizeColumns(1, 4);
  var NBS_LAST_ROW = REF_NBS_DATA.length + 1;

  // 4. Aba "Ref_CNAE_Item" (base oficial CNAE x Item LC116)
  var refCnae = criarOuLimparAba_(ss, 'Ref_CNAE_Item');
  refCnae.getRange(1, 1, 1, 2).setValues([['CNAE', 'Item LC 116']]);
  refCnae.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#2E5395').setFontColor('#ffffff');
  refCnae.getRange(2, 1, CNAE_ITEM_DATA.length, 2).setValues(CNAE_ITEM_DATA);
  refCnae.setFrozenRows(1);
  refCnae.autoResizeColumns(1, 2);

  // 5. Aba "Banco_Dados"
  var bd = criarOuLimparAba_(ss, 'Banco_Dados');
  bd.getRange('A1').setValue('BANCO DE DADOS - PARÂMETROS DA REFORMA TRIBUTÁRIA (LCP 214/2025 / Resolução CGIBS nº 6/2026)');
  bd.getRange('A1').setFontWeight('bold').setFontSize(13).setFontColor('#1F3864');
  bd.getRange('A1:G1').merge();
  bd.getRange('A2').setValue('Ajuste os parâmetros abaixo (células azuis) para recalcular toda a planilha automaticamente.');
  bd.getRange('A2').setFontStyle('italic').setFontColor('#7F7F7F');

  bd.getRange('A4').setValue('CBS Alíquota Padrão (Federal):');
  bd.getRange('B4').setValue(0.088).setNumberFormat('0.00%').setBackground('#DDEBF7');
  bd.getRange('A5').setValue('IBS Alíquota Padrão (estimada, alíquota plena):');
  bd.getRange('B5').setValue(0.177).setNumberFormat('0.00%').setBackground('#DDEBF7');

  bd.getRange('A7').setValue('REDUÇÃO POR SETOR DE ORIGEM (FALLBACK) - usada enquanto CNAE/NBS não preenchidos na aba "CNAE x NBS"');
  bd.getRange('A7').setFontWeight('bold').setFontColor('#1F3864');
  bd.getRange('A7:B7').merge();
  bd.getRange(8, 1, 1, 2).setValues([['Setor (aba de origem)', 'Redução CBS/IBS Fallback (%)']]);
  bd.getRange(8, 1, 1, 2).setFontWeight('bold').setBackground('#2E5395').setFontColor('#ffffff');
  bd.getRange(9, 1, SETOR_FALLBACK_DATA.length, 2).setValues(SETOR_FALLBACK_DATA);
  bd.getRange(9, 2, SETOR_FALLBACK_DATA.length, 1).setNumberFormat('0%').setBackground('#DDEBF7');

  bd.getRange('D7').setValue('CATEGORIAS DE REDUÇÃO POR TRATAMENTO TRIBUTÁRIO (cClassTrib) - LCP 214/2025 / Resolução CGIBS 6/2026');
  bd.getRange('D7').setFontWeight('bold').setFontColor('#1F3864');
  bd.getRange('D7:G7').merge();
  bd.getRange(8, 4, 1, 4).setValues([['cClassTrib', 'Nome da Categoria', 'Redução % (confirmado)', 'Observação / Confiança']]);
  bd.getRange(8, 4, 1, 4).setFontWeight('bold').setBackground('#2E5395').setFontColor('#ffffff');
  bd.getRange(9, 4, CATEGORIAS_CCLASSTRIB.length, 4).setValues(CATEGORIAS_CCLASSTRIB);
  bd.getRange(9, 6, CATEGORIAS_CCLASSTRIB.length, 1).setNumberFormat('0%').setBackground('#DDEBF7');
  bd.autoResizeColumns(1, 7);
  var CAT_FIRST = 9, CAT_LAST = 9 + CATEGORIAS_CCLASSTRIB.length - 1;
  var SETOR_FIRST = 9, SETOR_LAST = 9 + SETOR_FALLBACK_DATA.length - 1;

  // 6. Aba "CNAE x NBS"
  var cn = criarOuLimparAba_(ss, 'CNAE x NBS');
  cn.getRange(1, 1, 1, CNAENBS_HEADERS.length).setValues([CNAENBS_HEADERS]);
  cn.getRange(1, 1, 1, CNAENBS_HEADERS.length).setFontWeight('bold').setBackground('#2E5395').setFontColor('#ffffff').setWrap(true);
  var cnRows = [];
  for (var j = 0; j < total; j++) {
    var lt = linhasTotais[j];
    cnRows.push([lt[0], lt[1], lt[20], '', '']);
  }
  cn.getRange(2, 1, total, 5).setValues(cnRows);
  var INPUT_BG = '#DDEBF7';
  cn.getRange(2, 4, total, 2).setBackground(INPUT_BG);

  var formulasF = [], formulasG = [], formulasH = [], formulasI = [], formulasJ = [], formulasK = [];
  for (var k = 0; k < total; k++) {
    var rr = k + 2;
    formulasF.push(['=IF($E' + rr + '="","",IFERROR(VLOOKUP($E' + rr + ',Ref_NBS!$A:$D,2,FALSE),"NBS não encontrado na base oficial"))']);
    formulasG.push(['=IF($E' + rr + '="","",IFERROR(VLOOKUP($E' + rr + ',Ref_NBS!$A:$D,4,FALSE),""))']);
    formulasH.push(['=IF($G' + rr + '="","",IFERROR(VLOOKUP($G' + rr + ',Banco_Dados!$D$' + CAT_FIRST + ':$F$' + CAT_LAST + ',3,FALSE),""))']);
    formulasI.push(['=IF(OR($D' + rr + '="",$E' + rr + '=""),"Pendente preenchimento",IF(COUNTIFS(Ref_CNAE_Item!$A:$A,$D' + rr + ',Ref_CNAE_Item!$B:$B,VLOOKUP($E' + rr + ',Ref_NBS!$A:$D,3,FALSE))>0,"✅ Regular","⚠️ Divergente!"))']);
    formulasJ.push(['=IF($E' + rr + '="","",IFERROR(INDEX(Ref_CNAE_Item!$A:$A,MATCH(VLOOKUP($E' + rr + ',Ref_NBS!$A:$D,3,FALSE),Ref_CNAE_Item!$B:$B,0)),"-"))']);
    formulasK.push(['=IF(OR($D' + rr + '="",$E' + rr + '=""),"Preencher CNAE cadastrado e NBS de fato prestado para gerar diretriz.",IF($I' + rr + '="✅ Regular","Nenhuma ação necessária.","Urgente: solicitar alteração de CNAE para incluir o código sugerido "&$J' + rr + '&" no cadastro do CNPJ, evitando rejeições automáticas nas emissões da Reforma Tributária."))']);
  }
  cn.getRange(2, 6, total, 1).setFormulas(formulasF);
  cn.getRange(2, 7, total, 1).setFormulas(formulasG);
  cn.getRange(2, 8, total, 1).setFormulas(formulasH);
  cn.getRange(2, 8, total, 1).setNumberFormat('0%');
  cn.getRange(2, 9, total, 1).setFormulas(formulasI);
  cn.getRange(2, 10, total, 1).setFormulas(formulasJ);
  cn.getRange(2, 11, total, 1).setFormulas(formulasK);

  var dvNbs = SpreadsheetApp.newDataValidation()
    .requireValueInRange(refNbs.getRange('A2:A' + NBS_LAST_ROW), true)
    .setAllowInvalid(false)
    .setHelpText('Selecione um código NBS válido da lista oficial (aba Ref_NBS).')
    .build();
  cn.getRange(2, 5, total, 1).setDataValidation(dvNbs);
  cn.setFrozenRows(1);
  cn.setFrozenColumns(2);
  cn.autoResizeColumns(1, 11);

  // 7. Aba "DASHBOARD"
  var lastRowSN = total + 1;
  var dash = criarOuLimparAba_(ss, 'DASHBOARD');
  dash.getRange('A1').setValue('DASHBOARD - SIMULAÇÃO IBS/CBS x SIMPLES NACIONAL (Reforma Tributária)');
  dash.getRange('A1').setFontWeight('bold').setFontSize(14).setFontColor('#ffffff').setBackground('#1F3864');
  dash.getRange('A1:J1').merge();
  var kpiLabels = ['Total Analisados', 'CNAEs Regulares', 'Manter no Simples', 'CNAE/NBS Preenchido', 'Alertas Qualidade Dados'];
  var kpiCols = [2, 4, 6, 8, 10];
  for (var m = 0; m < kpiLabels.length; m++) {
    dash.getRange(7, kpiCols[m]).setValue(kpiLabels[m]).setFontWeight('bold');
  }
  dash.getRange(8, 2).setFormula('=COUNTA(\'Simples Nacional\'!$A$2:$A$' + lastRowSN + ')');
  dash.getRange(8, 4).setFormula('=COUNTIF(\'CNAE x NBS\'!$I$2:$I$' + lastRowSN + ',"✅ Regular")&" de "&B8');
  dash.getRange(8, 6).setFormula('=COUNTIF(\'Simples Nacional\'!$N$2:$N$' + lastRowSN + ',"*Manter Simples*")');
  dash.getRange(8, 8).setFormula('=COUNTIFS(\'CNAE x NBS\'!$D$2:$D$' + lastRowSN + ',"<>",\'CNAE x NBS\'!$E$2:$E$' + lastRowSN + ',"<>")&" de "&B8');
  dash.getRange(8, 10).setFormula('=COUNTIF(\'Simples Nacional\'!$Z$2:$Z$' + lastRowSN + ',"<>")&" registros"');
  dash.getRange('A10').setValue('Navegação:');
  dash.getRange('A10').setFontWeight('bold');
  var navLinks = [
    ['Simples Nacional', sn.getSheetId()],
    ['CNAE x NBS', cn.getSheetId()],
    ['Banco_Dados', bd.getSheetId()]
  ];
  for (var nlk = 0; nlk < navLinks.length; nlk++) {
    dash.getRange(11 + nlk, 1).setFormula('=HYPERLINK("#gid=' + navLinks[nlk][1] + '","Ir para ' + navLinks[nlk][0] + '")');
  }
  dash.autoResizeColumns(1, 10);

  // 8. Aba "Histórico Mensal" (vazia - alimentada mês a mês por "Extrair Dados do Extrato")
  if (!ss.getSheetByName('Histórico Mensal')) {
    var hist = ss.insertSheet('Histórico Mensal');
    hist.getRange(1, 1, 1, HIST_HEADERS.length).setValues([HIST_HEADERS]);
    hist.getRange(1, 1, 1, HIST_HEADERS.length).setFontWeight('bold').setBackground('#2E5395').setFontColor('#ffffff');
    hist.setFrozenRows(1);
    hist.autoResizeColumns(1, HIST_HEADERS.length);
  }

  ss.setActiveSheet(sn);
  ui.alert('✅ Consolidação concluída: ' + total + ' empresas em "Simples Nacional". Abas criadas/atualizadas: Simples Nacional, CNAE x NBS, Ref_NBS, Ref_CNAE_Item, Banco_Dados, Histórico Mensal, DASHBOARD.\n\nPróximo passo: preencher CNAE/NBS na aba "CNAE x NBS" (ou usar "Extrair Dados do Extrato") e depois "Gerar PDFs da Aba Atual" na aba "Simples Nacional".');
}
/**
 * Gera os PDFs de comunicado a partir da aba "Simples Nacional" (linhas com
 * a caixinha "☑️ SELECIONAR" marcada e "Status / Data de Geração" vazio).
 * Usa um único modelo de Google Doc (COMUNICADO_DOC_ID) para todos os
 * setores, e cruza dados de conformidade CNAE x NBS por CNPJ.
 */
function gerarPDF_AbaAtual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getActiveSheet();
  const nomeAba = aba.getName();

  if (nomeAba !== 'Simples Nacional') {
    SpreadsheetApp.getUi().alert('Erro: execute esta função a partir da aba "Simples Nacional" (planilha consolidada). Se ela ainda não existe, rode primeiro "🏗️ Construir Planilha Consolidada".');
    return;
  }

  const pasta = DriveApp.getFolderById(ID_PASTA_DESTINO_PDF);
  const arquivoModelo = DriveApp.getFileById(COMUNICADO_DOC_ID);

  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) { SpreadsheetApp.getUi().alert('A planilha parece estar vazia.'); return; }

  // Mapa de conformidade CNAE x NBS por CNPJ (para as variáveis do comunicado)
  var cnMap = {};
  var cnSheet = ss.getSheetByName('CNAE x NBS');
  if (cnSheet && cnSheet.getLastRow() > 1) {
    var cnData = cnSheet.getRange(2, 1, cnSheet.getLastRow() - 1, 11).getValues();
    cnData.forEach(function (r) {
      cnMap[String(r[0]).trim()] = {
        cnae: r[3] || '(não preenchido)',
        nbs: r[4] || '(não preenchido)',
        cclasstrib: r[6] || '',
        reducao: r[7],
        status: r[8] || 'Pendente preenchimento',
        acao: r[10] || ''
      };
    });
  }

  // Histórico mensal (PGDAS-D) por CNPJ, mais recente primeiro (para a tabela do comunicado)
  var historicoPorCnpj = {};
  var histSheet = ss.getSheetByName('Histórico Mensal');
  if (histSheet && histSheet.getLastRow() > 1) {
    var histData = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, HIST_HEADERS.length).getValues();
    histData.forEach(function (r) {
      var key = String(r[0]).trim();
      if (!historicoPorCnpj[key]) historicoPorCnpj[key] = [];
      historicoPorCnpj[key].push(r);
    });
    Object.keys(historicoPorCnpj).forEach(function (k) {
      historicoPorCnpj[k].sort(function (a, b) { return String(b[3]).localeCompare(String(a[3])); }); // Competência desc (AAAA-MM)
    });
  }

  const dados = aba.getRange(2, 1, ultimaLinha - 1, 27).getValues();
  const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  let processados = 0;

  for (let i = 0; i < dados.length; i++) {
    let linha = dados[i];

    let cnpj = String(linha[0]);
    let empresa = String(linha[1]);
    let anexo = String(linha[2]);
    let rbt12 = String(linha[3]);
    let fatorR = String(linha[4]);
    let aliqPis = String(linha[5]);
    let valorDas = String(linha[6]);
    let percFat = String(linha[7]);
    let percCredInsumos = String(linha[8]);
    let cbsDevDentro = String(linha[9]);
    let cbsCredDentro = String(linha[10]);
    let cbsDevFora = String(linha[11]);
    let cbsCredFora = String(linha[12]);
    let recomendacao = String(linha[13]);
    let statusData = linha[14];
    let ibsDevDentro = String(linha[16]);
    let ibsCredDentro = String(linha[17]);
    let ibsDevFora = String(linha[18]);
    let ibsCredFora = String(linha[19]);
    let setor = String(linha[20]);
    let reducaoAplicada = linha[21];
    let origemReducao = String(linha[22]);
    let pleno2033 = linha[23];
    let custoExtra2033 = linha[24];
    let qualidadeDados = String(linha[25]) || 'Sem observações.';
    let valorHibridoInformado = linha[26];

    if (empresa && statusData === '' && linha[15] === true) {
      let nomeArquivo = `SIMULACAO_${setor}_${empresa}`;
      let cn = cnMap[cnpj.trim()] || { cnae: '(não preenchido)', nbs: '(não preenchido)', status: 'Pendente preenchimento', reducao: '', acao: 'Preencher a aba "CNAE x NBS" para esta empresa.' };
      let historico = historicoPorCnpj[cnpj.trim()] || [];

      // Percentuais efetivos (para o cliente entender o impacto, não só o R$)
      var reducaoNum = typeof reducaoAplicada === 'number' ? reducaoAplicada : (parseFloat(String(reducaoAplicada).replace('%', '').replace(',', '.')) || 0);
      var cbsAliqCheia = 0.088;
      var bdSheet = ss.getSheetByName('Banco_Dados');
      if (bdSheet) { var v = bdSheet.getRange('B4').getValue(); if (typeof v === 'number' && v > 0) cbsAliqCheia = v; }
      var cbsAliqForaEfetiva = cbsAliqCheia * (1 - reducaoNum);

      // Crédito Transferido ao Cliente (PJ), por cenário:
      //  - Dentro do DAS: fração de PIS/COFINS já embutida no DAS (coluna "Crédito CBS - DENTRO"), à alíquota "VAR_ALIQ_PIS" do próprio extrato.
      //  - Fora do DAS: CBS apurada pelo regime regular (coluna "Crédito CBS - FORA"), à alíquota líquida (cheia - redução). O IBS NÃO compõe o
      //    crédito transferido em 2027 porque ainda não incide sobre a maioria dos serviços (LC 214/2025) - só a partir de 2029.
      var creditoTransferidoDentro = cbsCredDentro;
      var creditoTransferidoFora = cbsCredFora;
      var parseNum_ = function (s) { return parseFloat(String(s).replace(/[^\d,.-]/g, '').replace(',', '.')) || 0; };
      var valorDasNum = parseNum_(valorDas);
      var cbsDevDentroNum = parseNum_(cbsDevDentro);
      var cbsDevForaNum = parseNum_(cbsDevFora);
      var rbt12Num = parseNum_(rbt12);
      var receitaMensalBase = rbt12Num > 0 ? rbt12Num / 12 : 0;
      // DAS reduzido (Cenário Híbrido) = DAS atual menos a fração de CBS (PIS/COFINS) que sai do DAS
      var dasReduzidoNum = valorDasNum - cbsDevDentroNum;
      var valorHibridoInformadoNum = parseNum_(valorHibridoInformado);
      var ibsDevForaNum = parseNum_(ibsDevFora); // IBS 0,1% (teste 2027-2028) já com redução setorial
      // Se a coluna "Valor Híbrido Informado (R$)" (preenchimento manual na aba de
      // origem) tiver um valor, ele PREVALECE sobre o valor calculado.
      // DAS Híbrido = (DAS atual - PIS - COFINS) + CBS(8,8% c/ redução) + IBS(0,1% c/ redução)
      var valorLiquidoForaNum = valorHibridoInformadoNum > 0 ? valorHibridoInformadoNum : (dasReduzidoNum + cbsDevForaNum + ibsDevForaNum);
      var valorLiquidoFora = 'R$ ' + valorLiquidoForaNum.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      var aliqEfetivaDentro = receitaMensalBase > 0 ? (valorDasNum / receitaMensalBase) * 100 : 0;
      var aliqEfetivaFora = receitaMensalBase > 0 ? (valorLiquidoForaNum / receitaMensalBase) * 100 : 0;
      var notaIbs2027 = 'IBS: em 2027 e 2028 é cobrado à alíquota de teste de 0,1% (0,05% estadual + 0,05% municipal - art. 344, LC 214/2025), já com a redução setorial aplicável. A transição para a alíquota plena (referência ~17,7%) ocorre gradualmente de 2029 a 2033. O IBS de 0,1% já está incluído no valor do "Simples Híbrido".';

      try {
        let copiaDoc = arquivoModelo.makeCopy('TEMP_' + nomeArquivo, pasta);
        let docId = copiaDoc.getId();
        Utilities.sleep(1500);

        let doc = DocumentApp.openById(docId);
        let corpo = doc.getBody();

        corpo.replaceText('VAR_EMPRESA', empresa);
        corpo.replaceText('VAR_CNPJ', cnpj);
        corpo.replaceText('VAR_SETOR', setor);
        corpo.replaceText('VAR_ANEXO', anexo);
        corpo.replaceText('VAR_RBT12', rbt12);
        corpo.replaceText('VAR_FATOR_R', fatorR);
        corpo.replaceText('VAR_ALIQ_PIS', aliqPis);
        corpo.replaceText('VAR_VALOR_DAS', valorDas);
        corpo.replaceText('VAR_PERC_FAT', percFat);
        corpo.replaceText('VAR_PERC_CRED_INSUMOS', percCredInsumos || '(não informado)');
        corpo.replaceText('VAR_CBS_ALIQ_FORA', (cbsAliqForaEfetiva * 100).toFixed(2).replace('.', ',') + '%');
        corpo.replaceText('VAR_ALIQ_EFETIVA_DENTRO', aliqEfetivaDentro.toFixed(3).replace('.', ',') + '%');
        corpo.replaceText('VAR_ALIQ_EFETIVA_FORA', aliqEfetivaFora.toFixed(3).replace('.', ',') + '%');
        corpo.replaceText('VAR_CBS_DEV_DENTRO', cbsDevDentro);
        corpo.replaceText('VAR_CBS_DEV_FORA', cbsDevFora);
        corpo.replaceText('VAR_CBS_CRED_DENTRO', String(creditoTransferidoDentro));
        corpo.replaceText('VAR_CBS_CRED_FORA', String(creditoTransferidoFora));
        corpo.replaceText('VAR_IBS_DEV_DENTRO', ibsDevDentro);
        corpo.replaceText('VAR_IBS_DEV_FORA', ibsDevFora);
        corpo.replaceText('VAR_IBS_CRED_DENTRO', ibsCredDentro);
        corpo.replaceText('VAR_IBS_CRED_FORA', ibsCredFora);
        corpo.replaceText('VAR_IBS_NOTA_2027', notaIbs2027);
        corpo.replaceText('VAR_VALOR_LIQUIDO_FORA', valorLiquidoFora);
        corpo.replaceText('VAR_CNAE_CADASTRADO', String(cn.cnae));
        corpo.replaceText('VAR_NBS_FATO', String(cn.nbs));
        corpo.replaceText('VAR_STATUS_CONFORMIDADE', String(cn.status));
        corpo.replaceText('VAR_REDUCAO_APLICADA', cn.reducao === '' || cn.reducao === undefined ? String(reducaoAplicada) : String(cn.reducao));
        corpo.replaceText('VAR_ORIGEM_REDUCAO', origemReducao);
        corpo.replaceText('VAR_ACAO_CADASTRAL', String(cn.acao) || 'Nenhuma ação necessária.');
        corpo.replaceText('VAR_PLENO_2033', String(pleno2033));
        corpo.replaceText('VAR_CUSTO_EXTRA_2033', String(custoExtra2033));
        corpo.replaceText('VAR_RECOMENDACAO', recomendacao);
        corpo.replaceText('VAR_QUALIDADE_DADOS', qualidadeDados);
        corpo.replaceText('VAR_DATA', dataHoje);

        inserirTabelaHistorico_(corpo, historico);

        doc.saveAndClose();
        Utilities.sleep(2000);

        let pdfBlob = copiaDoc.getAs(MimeType.PDF);
        pasta.createFile(pdfBlob).setName(nomeArquivo + '.pdf');

        DriveApp.getFileById(docId).setTrashed(true);

        aba.getRange(i + 2, 15).setValue(dataHoje);

        processados++;

      } catch (e) {
        console.log('Erro: ' + e.message);
        SpreadsheetApp.getUi().alert(`Erro ao gerar PDF da empresa: ${empresa}\n${e.message}`);
      }
    }
  }

  if (processados > 0) {
    SpreadsheetApp.getUi().alert(`Sucesso! ${processados} documentos gerados a partir da aba 'Simples Nacional'.`);
  } else {
    SpreadsheetApp.getUi().alert('Tudo atualizado! Nenhum PDF pendente.');
  }
}

// =============================================================================
// MÓDULO DE EXTRAÇÃO: DADOS DO EXTRATO (PGDASD) POR COMPETÊNCIA
// Preenche Anexo, RBT12, Fator R, Valor no DAS e (melhor esforço) Alíquota
// PIS/COFINS a partir do extrato da competência informada. Os campos
// % Faturamento PJ, % Crédito CBS, CBS Devido/Crédito e Recomendação
// continuam manuais, pois não constam no extrato do Simples Nacional.
// =============================================================================

/**
 * Idêntica à versão de produção, exceto: a redução de CBS/IBS aplicada por
 * empresa deixou de vir de um dicionário fixo por aba (REDUCAO_CBS_POR_ABA)
 * e passa a ser lida, linha a linha, da coluna "Redução Aplicada (%)" da
 * própria planilha (que já calcula CNAE x NBS -> fallback setorial).
 */
function extrairDadosExtrato_AbaAtual() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var CBS_ALIQUOTA_CHEIA = 0.088; // 8,8% - premissa CBS informada (obs.: art. 347 LC 214/2025 reduz 0,1 p.p. em 2027-2028 -> 8,7%)
  var IBS_ALIQUOTA_TESTE_2027 = 0.001; // 0,1% - alíquota de teste do IBS em 2027-2028 (art. 344 LC 214/2025), c/ redução setorial proporcional (art. 344, p.u.)
  var PERC_FATURAMENTO_PJ = 0.10; // mesmo valor gravado na coluna H
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  garantirColuna_IBSCBS(sheet, headers, '☑️ SELECIONAR');
  var idxIbsDevidoDentro = garantirColuna_IBSCBS(sheet, headers, 'IBS Devido - DENTRO do DAS (R$)');
  var idxCreditoIbsDentro = garantirColuna_IBSCBS(sheet, headers, 'Crédito IBS - DENTRO do DAS (R$)');
  var idxIbsDevidoFora = garantirColuna_IBSCBS(sheet, headers, 'IBS Devido - FORA do DAS (R$)');
  var idxCreditoIbsFora = garantirColuna_IBSCBS(sheet, headers, 'Crédito IBS - FORA do DAS (R$)');
  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var linhasSel = obterLinhasParaProcessar_IBSCBS(sheet, headers);
  if (linhasSel.length === 0) {
    ui.alert("❌ Nenhuma linha selecionada. Marque as caixinhas em '☑️ SELECIONAR' ou selecione as linhas com o mouse.");
    return;
  }

  var compDados = obterDadosCompetencia_IBSCBS();
  if (!compDados) return;

  var idxCnpj = buscarIndiceColuna(headers, 'CNPJ');
  var idxEmpresa = buscarIndiceColuna(headers, 'Razão Social');
  var idxAnexo = buscarIndiceColuna(headers, 'Anexo do Simples Nacional');
  var idxRbt12 = buscarIndiceColuna(headers, 'RBT12');
  var idxFatorR = buscarIndiceColuna(headers, 'Fator R');
  var idxAliqPis = buscarIndiceColuna(headers, 'Alíquota PIS/COFINS');
  var idxValorDas = buscarIndiceColuna(headers, 'Valor no DAS');
  var idxFaturamento = buscarIndiceColuna(headers, 'Faturamento');
  var idxCbsDevidoDentro = buscarIndiceColuna(headers, 'CBS Devido - DENTRO');
  var idxCreditoCbsDentro = buscarIndiceColuna(headers, 'Crédito CBS - DENTRO');
  var idxCbsDevidoFora = buscarIndiceColuna(headers, 'CBS Devido - FORA');
  var idxCreditoCbsFora = buscarIndiceColuna(headers, 'Crédito CBS - FORA');
  var idxRecomendacao = buscarIndiceColuna(headers, 'Recomendação');
  var idxReducaoAplicada = buscarIndiceColuna(headers, 'Redução Aplicada');

  if (idxCnpj === -1 || idxEmpresa === -1 || idxAnexo === -1 || idxRbt12 === -1 || idxFatorR === -1 || idxValorDas === -1) {
    ui.alert('❌ Erro: não localizei uma ou mais colunas base (CNPJ, Razão Social, Anexo, RBT12, Fator R, Valor no DAS).');
    return;
  }

  var pastaPai = obterPastaPrincipal_IBSCBS();
  if (!pastaPai) { ui.alert('❌ Erro de acesso à pasta raiz no Drive.'); return; }

  var allData = sheet.getDataRange().getValues();
  var erros = [];
  var avisosPis = [];
  var processados = 0;

  for (var i = 0; i < linhasSel.length; i++) {
    var row = linhasSel[i];
    var linha = allData[row - 1];

    var cnpjPlanilha = String(linha[idxCnpj]).trim();
    var empresa = String(linha[idxEmpresa]).trim();

    // Redução aplicada a esta linha (CNAE x NBS -> fallback setorial), lida
    // diretamente da planilha em vez de um dicionário fixo por aba.
    var reducaoCbsRow = 0;
    if (idxReducaoAplicada > -1) {
      var rv = linha[idxReducaoAplicada];
      reducaoCbsRow = typeof rv === 'number' ? rv : (parseFloat(String(rv).replace('%', '').replace(',', '.')) || 0);
      if (reducaoCbsRow > 1) reducaoCbsRow = reducaoCbsRow / 100;
    }
    var CBS_ALIQUOTA_EFETIVA = CBS_ALIQUOTA_CHEIA * (1 - reducaoCbsRow);

    if (!empresa) { erros.push('Linha ' + row + ': Razão Social não preenchida.'); continue; }

    var tempFileId = null;
    try {
      var pastaEmpresa = buscarPastaEmpresa_IBSCBS(pastaPai, empresa);
      if (!pastaEmpresa) { erros.push('Linha ' + row + ' (' + empresa + '): pasta da empresa não encontrada no Drive.'); continue; }

      var pdf = buscarExtrato_IBSCBS(pastaPai, pastaEmpresa, cnpjPlanilha, compDados.compLimpa, compDados.variacoes, String(compDados.anoPA));
      if (!pdf) { erros.push('Linha ' + row + ' (' + empresa + '): extrato da competência ' + compDados.inputOriginal + ' não encontrado.'); continue; }

      var tempFile = Drive.Files.copy(
        { title: 'Temp_IBSCBS_' + empresa, mimeType: MimeType.GOOGLE_DOCS },
        pdf.getId(),
        { convert: true, supportsAllDrives: true }
      );
      tempFileId = tempFile.id;

      var doc = DocumentApp.openById(tempFileId);
      var texto = doc.getBody().getText();

      var cnpjExtratoMatch = texto.match(/(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})/);
      if (cnpjExtratoMatch && cnpjPlanilha) {
        var cnpjLimpoPlanilha = cnpjPlanilha.replace(/[^\d]/g, '');
        var cnpjLimpoExtrato = cnpjExtratoMatch[1].replace(/[^\d]/g, '');
        if (cnpjLimpoPlanilha && cnpjLimpoExtrato && cnpjLimpoPlanilha !== cnpjLimpoExtrato) {
          erros.push('Linha ' + row + ' (' + empresa + '): CNPJ do extrato (' + cnpjExtratoMatch[1] + ') não corresponde ao CNPJ da planilha (' + cnpjPlanilha + '). Linha não preenchida.');
          continue;
        }
      }

      var rbt12 = 0, rpa = 0;
      var mRPA = texto.match(/Receita Bruta do PA \(RPA\)[\s\S]{0,50}?(?:Competência|Caixa)[\s\S]{0,30}?([\d\.]+(?:,\d{2}))/i);
      if (!mRPA) mRPA = texto.match(/Receita Bruta do PA \(RPA\)[\s\S]{0,80}?([\d\.]+(?:,\d{2}))/i);
      if (mRPA) rpa = parseMoney_IBSCBS(mRPA[1]);

      var mRBT = texto.match(/RBT12\)[\s\S]{0,40}?([\d\.]+(?:,\d{2}))/i);
      if (mRBT) rbt12 = parseMoney_IBSCBS(mRBT[1]);

      var idadeMeses = 13;
      var matchAbertura = texto.match(/abertura(?: no CNPJ)?:?\s*(\d{2})\/(\d{2})\/(\d{4})/i);
      if (matchAbertura) {
        var mesAb = parseInt(matchAbertura[2], 10);
        var anoAb = parseInt(matchAbertura[3], 10);
        idadeMeses = (compDados.anoPA * 12 + compDados.mesPA) - (anoAb * 12 + mesAb) + 1;
      }

      var rbt12Final = 0;
      if (idadeMeses < 12) {
        if (idadeMeses <= 0) idadeMeses = 1;
        rbt12Final = ((rbt12 + rpa) / idadeMeses) * 12;
      } else {
        var antigo = 0;
        var blocoReceitas = texto.match(/Receitas Brutas Anteriores[\s\S]{0,1500}/i);
        if (blocoReceitas) {
          var escapedAbat = compDados.textoAbatimento.replace('/', '\\/');
          var rAnt = new RegExp(escapedAbat + '[\\s\\S]{0,40}?([\\d\\.]+(?:,\\d{2}))', 'i');
          var m = rAnt.exec(blocoReceitas[0]);
          if (m) antigo = parseMoney_IBSCBS(m[1]);
        }
        rbt12Final = (rbt12 + rpa) - antigo;
      }

      var anexos = [];
      var mAnx, rAnx = /Anexo\s+(VI|V|IV|III|II|I)\b/gi;
      while ((mAnx = rAnx.exec(texto)) !== null) {
        var nomeAnx = mAnx[1].toUpperCase();
        if (anexos.indexOf(nomeAnx) === -1) anexos.push(nomeAnx);
      }
      var txtAnexo = anexos.length > 0 ? anexos.join('/') : '';

      var fatorRTxt = '';
      var mFatorR = texto.match(/Fator\s*r\s*=\s*([\d]+,\d+)/i);
      if (mFatorR) {
        var percFatorR = parseFloat(mFatorR[1].replace(',', '.'));
        fatorRTxt = percFatorR >= 0.28 ? 'com' : 'sem';
      } else if (/Fator\s*r\)[\s\S]{0,60}?N[ãa]o se aplica/i.test(texto)) {
        fatorRTxt = '';
      }

      var cbsDentroDas = 0;
      var cbsForaDas = 0;
      var ibsDentroDas = 0;
      var ibsForaDas = 0;
      var valorDas = 0;
      var aliqPis = null;
      var mBlocoDas = texto.match(/Informa[cç][oõ]es sobre DAS Gerado[\s\S]{0,600}?Total\s+([\d\.]+,\d{2})/i);
      if (mBlocoDas) {
        valorDas = parseMoney_IBSCBS(mBlocoDas[1]);
        var trechoDas = mBlocoDas[0];
        var mCofinsVal = trechoDas.match(/COFINS\s+([\d\.]+,\d{2})/i);
        var mPisVal = trechoDas.match(/PIS\/PASEP\s+([\d\.]+,\d{2})/i) || trechoDas.match(/PIS\s+([\d\.]+,\d{2})/i);
        var mIssVal = trechoDas.match(/ISS\s+([\d\.]+,\d{2})/i);
        var mIcmsVal = trechoDas.match(/ICMS\s+([\d\.]+,\d{2})/i);
        var mIrpjVal = trechoDas.match(/IRPJ\s+([\d\.]+,\d{2})/i);
        var mCsllVal = trechoDas.match(/CSLL\s+([\d\.]+,\d{2})/i);
        var mCppVal = trechoDas.match(/CPP\s+([\d\.]+,\d{2})/i) || trechoDas.match(/INSS\/?CPP\s+([\d\.]+,\d{2})/i);
        if (mIssVal) ibsDentroDas = parseMoney_IBSCBS(mIssVal[1]);
        else if (mIcmsVal) ibsDentroDas = parseMoney_IBSCBS(mIcmsVal[1]);
        if (mCofinsVal && mPisVal && valorDas > 0) {
          var valCofins = parseMoney_IBSCBS(mCofinsVal[1]);
          var valPis = parseMoney_IBSCBS(mPisVal[1]);
          cbsDentroDas = valCofins + valPis;
          aliqPis = ((valCofins + valPis) / valorDas * 100).toFixed(4).replace('.', ',') + '%';
        }
        var valIrpj = mIrpjVal ? parseMoney_IBSCBS(mIrpjVal[1]) : 0;
        var valCsll = mCsllVal ? parseMoney_IBSCBS(mCsllVal[1]) : 0;
        var valCpp = mCppVal ? parseMoney_IBSCBS(mCppVal[1]) : 0;
        var valIss = mIssVal ? parseMoney_IBSCBS(mIssVal[1]) : 0;
      }
      if (valorDas === 0) {
        var mDasAlt = texto.match(/Total do D[eé]bito Exig[ií]vel[\s\S]{0,40}?([\d\.]+(?:,\d{2}))/i);
        if (!mDasAlt) mDasAlt = texto.match(/Valor total do documento[\s\S]{0,40}?([\d\.]+(?:,\d{2}))/i);
        if (!mDasAlt) mDasAlt = texto.match(/Valor (?:total )?do DAS[\s\S]{0,40}?([\d\.]+(?:,\d{2}))/i);
        if (mDasAlt) valorDas = parseMoney_IBSCBS(mDasAlt[1]);
      }
      if (!aliqPis) avisosPis.push('Linha ' + row + ' (' + empresa + '): Alíquota PIS/COFINS não localizada automaticamente – preencher manualmente.');

      if (txtAnexo) sheet.getRange(row, idxAnexo + 1).setValue(txtAnexo);
      if (rbt12Final > 0) sheet.getRange(row, idxRbt12 + 1).setValue(formatMoney_IBSCBS(rbt12Final));
      if (fatorRTxt) sheet.getRange(row, idxFatorR + 1).setValue(fatorRTxt);
      if (valorDas > 0) sheet.getRange(row, idxValorDas + 1).setValue(formatMoney_IBSCBS(valorDas));
      if (aliqPis) sheet.getRange(row, idxAliqPis + 1).setValue(aliqPis);

      // Grava (upsert) o mês processado na aba "Histórico Mensal" - não sobrescreve meses anteriores.
      try {
        var idxSetorHist = buscarIndiceColuna(headers, 'Setor');
        var setorRow = idxSetorHist > -1 ? String(linha[idxSetorHist]) : '';
        registrarHistoricoMensal_(sheet.getParent(), {
          cnpj: cnpjPlanilha, empresa: empresa, setor: setorRow,
          competencia: compDados.anoPA + '-' + ('0' + compDados.mesPA).slice(-2),
          anexo: txtAnexo, fatorR: fatorRTxt, rpa: rpa, rbt12: rbt12Final,
          irpj: valIrpj || 0, csll: valCsll || 0, cofins: (typeof valCofins !== 'undefined' ? valCofins : 0),
          pis: (typeof valPis !== 'undefined' ? valPis : 0), cpp: valCpp || 0, iss: valIss || 0,
          totalDas: valorDas
        });
      } catch (eHist) { Logger.log('Histórico Mensal: ' + eHist.message); }

      var pjCellVal = idxFaturamento > -1 ? sheet.getRange(row, idxFaturamento + 1).getValue() : '';
      var percPjRow = PERC_FATURAMENTO_PJ;
      if (idxFaturamento > -1) {
        if (pjCellVal === '' || pjCellVal === null || pjCellVal === undefined) {
          sheet.getRange(row, idxFaturamento + 1).setValue('10%');
          percPjRow = PERC_FATURAMENTO_PJ;
        } else {
          var pjStr = String(pjCellVal).replace('%', '').replace(',', '.').trim();
          var pjNum = parseFloat(pjStr);
          if (!isNaN(pjNum)) percPjRow = pjNum > 1 ? pjNum / 100 : pjNum;
        }
      }

      if (idxCbsDevidoDentro > -1 && cbsDentroDas > 0) sheet.getRange(row, idxCbsDevidoDentro + 1).setValue(formatMoney_IBSCBS(cbsDentroDas));
      if (idxCreditoCbsDentro > -1 && cbsDentroDas > 0) sheet.getRange(row, idxCreditoCbsDentro + 1).setValue(formatMoney_IBSCBS(cbsDentroDas * percPjRow));
      if (idxCbsDevidoFora > -1 && rbt12Final > 0) {
        cbsForaDas = (rbt12Final / 12) * CBS_ALIQUOTA_EFETIVA;
        sheet.getRange(row, idxCbsDevidoFora + 1).setValue(formatMoney_IBSCBS(cbsForaDas));
        if (idxCreditoCbsFora > -1) sheet.getRange(row, idxCreditoCbsFora + 1).setValue(formatMoney_IBSCBS(cbsForaDas * percPjRow));
      }

      var anexoChaveIbs = anexos.length > 0 ? anexos[0] : null;
      var aliqIbsEfetiva = anexoChaveIbs ? obterAliquotaIbsEquivalente_IBSCBS(anexoChaveIbs, rbt12Final) : 0; // referência do regime pleno 2033 (NÃO usada no cálculo 2027)
      var ibsAliquota2027Efetiva = IBS_ALIQUOTA_TESTE_2027 * (1 - reducaoCbsRow); // 0,1% (teste 2027-2028) com a MESMA redução setorial da CBS
      if (idxIbsDevidoDentro > -1 && ibsDentroDas > 0) sheet.getRange(row, idxIbsDevidoDentro + 1).setValue(formatMoney_IBSCBS(ibsDentroDas));
      if (idxCreditoIbsDentro > -1 && ibsDentroDas > 0) sheet.getRange(row, idxCreditoIbsDentro + 1).setValue(formatMoney_IBSCBS(ibsDentroDas * percPjRow));
      if (idxIbsDevidoFora > -1 && rbt12Final > 0) {
        ibsForaDas = (rbt12Final / 12) * ibsAliquota2027Efetiva;
        sheet.getRange(row, idxIbsDevidoFora + 1).setValue(formatMoney_IBSCBS(ibsForaDas));
        if (idxCreditoIbsFora > -1) sheet.getRange(row, idxCreditoIbsFora + 1).setValue(formatMoney_IBSCBS(ibsForaDas * percPjRow));
      }

      if (idxRecomendacao > -1 && rbt12Final > 0) {
        var custoExtraTotal = (cbsForaDas + ibsForaDas) - (cbsDentroDas + ibsDentroDas);
        var recomendacaoTxt;
        if (custoExtraTotal < 0) {
          recomendacaoTxt = 'Avaliar Saida do Simples: mesmo com apenas ' + Math.round(percPjRow * 100) + '% da carteira PJ, sair do DAS para CBS+IBS geraria economia estimada de ~' + formatMoney_IBSCBS(Math.abs(custoExtraTotal)) + '/mes.';
        } else if (percPjRow >= 0.50) {
          recomendacaoTxt = 'Avaliar Hibrido: ' + Math.round(percPjRow * 100) + '% da carteira e PJ, credito integral pode compensar o custo extra de CBS+IBS (~' + formatMoney_IBSCBS(custoExtraTotal) + '/mes).';
        } else if (percPjRow >= 0.20) {
          recomendacaoTxt = 'Avaliar caso a caso: ' + Math.round(percPjRow * 100) + '% PJ, custo extra de CBS+IBS fora do DAS ~' + formatMoney_IBSCBS(custoExtraTotal) + '/mes - checar se repassa via preco.';
        } else {
          recomendacaoTxt = 'Manter Simples (DAS integral): apenas ' + Math.round(percPjRow * 100) + '% da carteira e PJ, nao compensa custo extra de CBS+IBS fora do DAS (~' + formatMoney_IBSCBS(custoExtraTotal) + '/mes).';
        }
        sheet.getRange(row, idxRecomendacao + 1).setValue(recomendacaoTxt);
      }

      processados++;

    } catch (e) {
      erros.push('Linha ' + row + ' (' + empresa + '): erro inesperado — ' + e.message);
    } finally {
      if (tempFileId) { try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (e2) { } }
    }
  }

  var resumo = '✅ ' + processados + ' de ' + linhasSel.length + ' linha(s) preenchidas com dados do extrato (competência ' + compDados.inputOriginal + ').';
  if (avisosPis.length > 0) resumo += '\n\n⚠️ Alíquota PIS/COFINS não localizada automaticamente:\n' + avisosPis.join('\n');
  if (erros.length > 0) resumo += '\n\n❌ Falhas:\n' + erros.join('\n');
  ui.alert(resumo);
}

/**
 * Grava (upsert) uma linha de histórico mensal (PGDAS-D) na aba "Histórico Mensal".
 * Chave = CNPJ + Competência (AAAA-MM). Se a aba não existir ainda (planilha não
 * consolidada), cria-a com os cabeçalhos padrão. Não apaga nem sobrescreve meses
 * de outras competências - apenas atualiza a linha da mesma competência, se já
 * processada antes (permite reprocessar um mês sem duplicar).
 */
function registrarHistoricoMensal_(ss, d) {
  var sh = ss.getSheetByName('Histórico Mensal');
  if (!sh) {
    sh = ss.insertSheet('Histórico Mensal');
    sh.getRange(1, 1, 1, HIST_HEADERS.length).setValues([HIST_HEADERS]);
    sh.getRange(1, 1, 1, HIST_HEADERS.length).setFontWeight('bold').setBackground('#2E5395').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  var aliqEfetiva = d.rpa > 0 ? (d.totalDas / d.rpa) * 100 : 0;
  var dataExtracao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var novaLinha = [
    d.cnpj, d.empresa, d.setor, d.competencia, d.anexo || '', d.fatorR || '',
    formatMoney_IBSCBS(d.rpa || 0), formatMoney_IBSCBS(d.rbt12 || 0),
    formatMoney_IBSCBS(d.irpj || 0), formatMoney_IBSCBS(d.csll || 0), formatMoney_IBSCBS(d.cofins || 0),
    formatMoney_IBSCBS(d.pis || 0), formatMoney_IBSCBS(d.cpp || 0), formatMoney_IBSCBS(d.iss || 0),
    formatMoney_IBSCBS(d.totalDas || 0), aliqEfetiva.toFixed(3).replace('.', ',') + '%', dataExtracao
  ];
  var lastRow = sh.getLastRow();
  if (lastRow > 1) {
    var chaves = sh.getRange(2, 1, lastRow - 1, 4).getValues();
    for (var i = 0; i < chaves.length; i++) {
      if (String(chaves[i][0]).trim() === String(d.cnpj).trim() && String(chaves[i][3]).trim() === String(d.competencia).trim()) {
        sh.getRange(i + 2, 1, 1, HIST_HEADERS.length).setValues([novaLinha]);
        return;
      }
    }
  }
  sh.appendRow(novaLinha);
}

/**
 * Substitui o parágrafo com o placeholder "VAR_HISTORICO_TABELA" por uma tabela
 * real do Google Docs com o histórico mensal (PGDAS-D) da empresa, mais recente
 * primeiro (até 12 meses). Se não houver histórico, apenas remove o placeholder
 * e escreve uma nota explicativa.
 */
function inserirTabelaHistorico_(corpo, historico) {
  var numChildren = corpo.getNumChildren();
  var idxPlaceholder = -1;
  for (var i = 0; i < numChildren; i++) {
    var el = corpo.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH && el.asParagraph().getText().indexOf('VAR_HISTORICO_TABELA') !== -1) {
      idxPlaceholder = i;
      break;
    }
  }
  if (idxPlaceholder === -1) return; // modelo não tem a seção de histórico - segue sem ela

  if (!historico || historico.length === 0) {
    corpo.getChild(idxPlaceholder).asParagraph().setText('Ainda não há histórico mensal suficiente (execute "Extrair Dados do Extrato" mensalmente para acumular a série).');
    return;
  }

  var linhas = historico.slice(0, 12);
  var cabecalho = ['Competência', 'RBT12 (R$)', 'Total DAS (R$)', 'Alíquota Efetiva (%)'];
  var tableData = [cabecalho];
  linhas.forEach(function (r) {
    tableData.push([String(r[3]), String(r[7]), String(r[14]), String(r[15])]);
  });

  var tabela = corpo.insertTable(idxPlaceholder, tableData);
  var headerRow = tabela.getRow(0);
  for (var c = 0; c < headerRow.getNumCells(); c++) {
    headerRow.getCell(c).setBackgroundColor('#2E5395');
    var textEl = headerRow.getCell(c).editAsText();
    textEl.setForegroundColor('#ffffff').setBold(true);
  }
  corpo.removeChild(corpo.getChild(idxPlaceholder + 1)); // remove o parágrafo original do placeholder
}
/**
 * Reescreve o corpo do modelo de Comunicado (COMUNICADO_DOC_ID) com o texto
 * unico definitivo (comunicado de opcao de regime tributario 2027 - DAS
 * Integral x Simples Hibrido), com nome da empresa, CNPJ e as duas
 * aliquotas efetivas calculadas por linha. Execute 1x (menu) sempre que o
 * texto-base do comunicado precisar ser atualizado; os PDFs ja gerados nao
 * sao afetados.
 */
function reescreverModeloComunicado() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Atualizar Modelo de Comunicado', 'Isto substitui todo o texto do Google Doc "Modelo Comunicado" (ID ' + COMUNICADO_DOC_ID + ') pelo novo modelo unificado. Deseja continuar?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var doc = DocumentApp.openById(COMUNICADO_DOC_ID);
  var corpo = doc.getBody();
  corpo.clear();

  function n(texto) { return corpo.appendParagraph(texto); }
  function bullet(texto) { var li = corpo.appendListItem(texto); li.setGlyphType(DocumentApp.GlyphType.BULLET); return li; }

  var titulo = corpo.appendParagraph('COMUNICADO - OPÇÃO DE REGIME TRIBUTÁRIO 2027');
  titulo.setHeading(DocumentApp.ParagraphHeading.TITLE);
  n('');

  n('Empresa: VAR_EMPRESA');
  n('CNPJ: VAR_CNPJ');
  n('');

  n('Conforme informado em nosso e-mail anterior, entre 01 e 30 de setembro de 2026, estará aberto o prazo para definir a forma de recolhimento da CBS para o 1º semestre de 2027: dentro do DAS, como é feito atualmente, ou de forma híbrida, com a CBS recolhida separadamente.');
  n('');
  n('Como não houve o agendamento da reunião disponibilizada em nossa agenda, realizamos a simulação com base no faturamento e na tributação atual da empresa.');
  n('');
  n('Os cenários estimados são:');
  bullet('DAS Integral: manutenção do pagamento dos tributos em uma única guia, com alíquota efetiva estimada de VAR_ALIQ_EFETIVA_DENTRO.');
  bullet('Simples Híbrido: pagamento dos demais tributos pelo DAS e da CBS separadamente, com carga tributária total estimada de VAR_ALIQ_EFETIVA_FORA.');
  n('');
  n('Caso a projeção de faturamento para 2027 seja diferente do valor atual, solicitamos que nos comunique para que possamos refazer os cálculos.');
  n('');
  n('Importante: para essa simulação, consideramos que os custos da empresa são irrelevantes para a geração de créditos de CBS.');
  n('');
  n('Além da diferença de carga tributária, é importante avaliar o impacto comercial. No regime híbrido, seus clientes pessoas jurídicas poderão ter maior aproveitamento de créditos de CBS. Portanto, caso seus clientes valorizem esse crédito ou exijam essa condição, esse fator deve ser considerado na decisão.');
  n('');
  n('A escolha deverá ser feita pela empresa, considerando a diferença de carga tributária e o impacto comercial junto aos seus clientes.');
  n('');
  n('Pedimos que nos informem até 10/09/2026 qual opção deverá ser adotada: DAS Integral – manutenção do regime atual ou Simples Híbrido.');
  n('');
  n('Em caso de dúvidas, estamos à disposição para uma reunião ou para esclarecimentos pelo Time de Atendimento – (XX) XXXX-XXXX ou pelo e-mail usuario8@exemplo.com.br.');
  n('');
  n('Documento emitido em VAR_DATA.');

  doc.saveAndClose();
  ui.alert('✅ Modelo de Comunicado atualizado.');
}
function buscarIndiceColuna(headers, nomeDesejado) {
  function normalizar(txt) {
    if (!txt) return "";
    return String(txt).toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]/g, "");
  }
  let nNome = normalizar(nomeDesejado);
  let idx = headers.findIndex(h => normalizar(h) === nNome);
  if (idx !== -1) return idx;
  return headers.findIndex(h => {
    let hNorm = normalizar(h);
    if (!hNorm) return false;
    return hNorm.includes(nNome) || nNome.includes(hNorm);
  });
}

function garantirColuna_IBSCBS(sheet, headers, nomeColuna) {
  let idx = buscarIndiceColuna(headers, nomeColuna);
  if (idx === -1) {
    let lastCol = sheet.getLastColumn();
    sheet.getRange(1, lastCol + 1).setValue(nomeColuna)
      .setFontWeight("bold").setBackground("#274e13").setFontColor("#ffffff");
    headers.push(nomeColuna);
    return lastCol;
  }
  return idx;
}

function obterLinhasParaProcessar_IBSCBS(sheet, headers) {
  var linhas = [];
  var allData = sheet.getDataRange().getValues();
  var idxSelecionar = buscarIndiceColuna(headers, "SELECIONAR");
  if (idxSelecionar !== -1) {
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][idxSelecionar] === true) linhas.push(i + 1);
    }
  }
  if (linhas.length === 0) {
    var rangeList = sheet.getActiveRangeList();
    if (rangeList) {
      var ranges = rangeList.getRanges();
      for (var i = 0; i < ranges.length; i++) {
        var start = ranges[i].getRow();
        var numRows = ranges[i].getNumRows();
        for (var j = 0; j < numRows; j++) {
          var row = start + j;
          if (row > 1 && linhas.indexOf(row) === -1) linhas.push(row);
        }
      }
    }
  }
  return linhas.sort(function (a, b) { return a - b; });
}

function parseMoney_IBSCBS(str) {
  if (!str) return 0;
  var limpo = str.toString().replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.');
  return parseFloat(limpo) || 0;
}

function formatMoney_IBSCBS(num) {
  return "R$ " + num.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function obterPastaPrincipal_IBSCBS() {
  var id = "ID_EXEMPLO"; // mesma pasta raiz "1. Impostos" do script de Alíquota
  try { return DriveApp.getFolderById(id); } catch (e) {
    var busca = DriveApp.getFoldersByName("1. Impostos");
    if (busca.hasNext()) return busca.next();
  }
  return null;
}

function buscarPastaEmpresa_IBSCBS(pastaPai, nomeEmpresa) {
  var nomeLimpo = nomeEmpresa.toString().replace(/[^\w\sÀ-ÿ]/gi, ' ').replace(/\s+/g, ' ').trim();
  var palavras = nomeLimpo.split(" ").slice(0, 3);
  if (palavras.length > 0) {
    var termoJunto = palavras.join(" ");
    var buscaJunta = pastaPai.searchFolders("title contains '" + termoJunto + "'");
    if (buscaJunta.hasNext()) return buscaJunta.next();
    var querySeparada = palavras.map(function (p) { return "title contains '" + p + "'"; }).join(" and ");
    var buscaSeparada = pastaPai.searchFolders(querySeparada);
    if (buscaSeparada.hasNext()) return buscaSeparada.next();
  }
  return null;
}

function encontrarPastaMes_IBSCBS(pastaEmpresa, variacoes, anoPAStr) {
  var subPastas = pastaEmpresa.getFolders();
  var pastasAnoEncontradas = [];
  while (subPastas.hasNext()) {
    var sub = subPastas.next();
    var nome = sub.getName().toUpperCase().replace(/\s+/g, '');
    for (var i = 0; i < variacoes.length; i++) {
      if (nome.indexOf(variacoes[i]) !== -1) return sub;
    }
    if (anoPAStr && nome.indexOf(anoPAStr) !== -1) pastasAnoEncontradas.push(sub);
  }
  for (var k = 0; k < pastasAnoEncontradas.length; k++) {
    var subPastasAno = pastasAnoEncontradas[k].getFolders();
    while (subPastasAno.hasNext()) {
      var subMes = subPastasAno.next();
      var nomeMes = subMes.getName().toUpperCase().replace(/\s+/g, '');
      for (var j = 0; j < variacoes.length; j++) {
        if (nomeMes.indexOf(variacoes[j]) !== -1) return subMes;
      }
    }
  }
  return null;
}

function buscarExtrato_IBSCBS(pastaPai, pastaEmpresa, cnpj, compLimpa, variacoes, anoPAStr) {
  if (pastaEmpresa) {
    var pastaMes = encontrarPastaMes_IBSCBS(pastaEmpresa, variacoes, anoPAStr);
    var alvo = pastaMes || pastaEmpresa;
    var arquivos = alvo.getFilesByType(MimeType.PDF);
    var fallback = null;
    while (arquivos.hasNext()) {
      var arq = arquivos.next();
      var nome = arq.getName().toUpperCase();
      if (nome.includes("RECIBO") && !nome.includes("EXTRATO") && !nome.includes("DECLARA")) continue;
      if (nome.includes("PGDASD") || nome.includes("EXTRATO") || nome.includes("DECLARA")) {
        if (nome.includes("PGDASD") || nome.includes("EXTRATO")) return arq;
        if (!fallback) fallback = arq;
      }
    }
    if (fallback) return fallback;
  }
  if (cnpj) {
    var cnpjLimpo = String(cnpj).replace(/[^\d]/g, "");
    if (cnpjLimpo) {
      var arqGlobal = pastaPai.searchFiles("title contains '" + cnpjLimpo + "' and mimeType = 'application/pdf'");
      while (arqGlobal.hasNext()) {
        var arqG = arqGlobal.next();
        var n = arqG.getName().toUpperCase();
        if (n.includes("PGDASD") || n.includes("EXTRATO")) return arqG;
      }
    }
  }
  return null;
}

function obterDadosCompetencia_IBSCBS() {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('📅 Competência do Extrato', 'Digite o mês/ano do extrato no formato MM/AAAA\n(Exemplo: 05/2026):', ui.ButtonSet.OK_CANCEL);

  if (prompt.getSelectedButton() !== ui.Button.OK) return null;

  var input = prompt.getResponseText().trim();
  if (!/^\d{2}\/\d{4}$/.test(input)) {
    ui.alert('❌ Erro: Formato inválido! Use o formato MM/AAAA (Ex: 05/2026).');
    return null;
  }

  var partes = input.split('/');
  var mes = partes[0];
  var ano = partes[1];

  var compLimpa = mes + ano;
  var variacoes = [compLimpa, ano + mes, mes + "-" + ano, ano + "-" + mes, mes + "." + ano];

  var mesesNomes = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
  var nomeMes = mesesNomes[parseInt(mes, 10) - 1];
  if (nomeMes) variacoes.push(nomeMes);

  var dataRef = new Date(parseInt(ano, 10), parseInt(mes, 10) - 1, 1);
  dataRef.setMonth(dataRef.getMonth() - 12);
  var mesAntigo = ("0" + (dataRef.getMonth() + 1)).slice(-2);
  var anoAntigo = dataRef.getFullYear();
  var textoAbatimento = mesAntigo + "/" + anoAntigo;

  return {
    compLimpa: compLimpa,
    variacoes: variacoes,
    textoAbatimento: textoAbatimento,
    inputOriginal: input,
    mesPA: parseInt(mes, 10),
    anoPA: parseInt(ano, 10)
  };
}
// =============================================================================
// TABELAS OFICIAIS SIMPLES NACIONAL (Anexos I-V) - aliquota nominal, deducao e
// percentual de ISS/ICMS embutido na aliquota efetiva do DAS (LC 123/2006).
// Usado para estimar o IBS "por fora do DAS" (substitui ISS/ICMS na Reforma).
// =============================================================================
var TABELAS_SN_IBSCBS = {
  "I": [
    { max: 180000,  nom: 0.040, ded: 0,      iss: 0, icms: 0.340 },
    { max: 360000,  nom: 0.073, ded: 5940,   iss: 0, icms: 0.340 },
    { max: 720000,  nom: 0.095, ded: 13860,  iss: 0, icms: 0.340 },
    { max: 1800000, nom: 0.107, ded: 22500,  iss: 0, icms: 0.340 },
    { max: 3600000, nom: 0.143, ded: 87300,  iss: 0, icms: 0.340 },
    { max: 4800000, nom: 0.190, ded: 378000, iss: 0, icms: 0     }
  ],
  "II": [
    { max: 180000,  nom: 0.045, ded: 0,      iss: 0, icms: 0.320 },
    { max: 360000,  nom: 0.078, ded: 5940,   iss: 0, icms: 0.320 },
    { max: 720000,  nom: 0.100, ded: 13860,  iss: 0, icms: 0.320 },
    { max: 1800000, nom: 0.112, ded: 22500,  iss: 0, icms: 0.320 },
    { max: 3600000, nom: 0.147, ded: 85500,  iss: 0, icms: 0.320 },
    { max: 4800000, nom: 0.300, ded: 720000, iss: 0, icms: 0     }
  ],
  "III": [
    { max: 180000,  nom: 0.060, ded: 0,      iss: 0.335, icms: 0 },
    { max: 360000,  nom: 0.112, ded: 9360,   iss: 0.320, icms: 0 },
    { max: 720000,  nom: 0.135, ded: 17640,  iss: 0.325, icms: 0 },
    { max: 1800000, nom: 0.160, ded: 35640,  iss: 0.325, icms: 0 },
    { max: 3600000, nom: 0.210, ded: 125640, iss: 0.335, icms: 0 },
    { max: 4800000, nom: 0.330, ded: 648000, iss: 0,     icms: 0 }
  ],
  "IV": [
    { max: 180000,  nom: 0.045, ded: 0,      iss: 0.445, icms: 0 },
    { max: 360000,  nom: 0.090, ded: 8100,   iss: 0.400, icms: 0 },
    { max: 720000,  nom: 0.102, ded: 12420,  iss: 0.400, icms: 0 },
    { max: 1800000, nom: 0.140, ded: 39780,  iss: 0.400, icms: 0 },
    { max: 3600000, nom: 0.220, ded: 183780, iss: 0.434, icms: 0 },
    { max: 4800000, nom: 0.330, ded: 828000, iss: 0,     icms: 0 }
  ],
  "V": [
    { max: 180000,  nom: 0.155, ded: 0,      iss: 0.140, icms: 0 },
    { max: 360000,  nom: 0.180, ded: 4500,   iss: 0.170, icms: 0 },
    { max: 720000,  nom: 0.195, ded: 9900,   iss: 0.190, icms: 0 },
    { max: 1800000, nom: 0.205, ded: 17100,  iss: 0.210, icms: 0 },
    { max: 3600000, nom: 0.230, ded: 62100,  iss: 0.235, icms: 0 },
    { max: 4800000, nom: 0.305, ded: 540000, iss: 0,     icms: 0 }
  ]
};
var TETO_IBS_IBSCBS = 0.05; // teto legal do ISS (5%) - aplicado tambem ao ICMS por consistencia

function obterAliquotaIbsEquivalente_IBSCBS(anexoChave, rbt12) {
  var tabela = TABELAS_SN_IBSCBS[anexoChave];
  if (!tabela || !rbt12 || rbt12 <= 0) return 0;
  var faixa = tabela.filter(function (f) { return rbt12 <= f.max; })[0];
  if (!faixa) return 0;
  var aliqEfetiva = ((rbt12 * faixa.nom) - faixa.ded) / rbt12;
  var perc = faixa.iss > 0 ? faixa.iss : faixa.icms;
  if (!perc) return 0;
  var aliqIbs = aliqEfetiva * perc;
  if (aliqIbs > TETO_IBS_IBSCBS) aliqIbs = TETO_IBS_IBSCBS;
  return aliqIbs;
}
function reextrairFatorR_IBSCBS(nomeAba, mes, ano) {
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nomeAba);
  if (!sheet) { Logger.log(nomeAba + ': aba nao encontrada'); return; }
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];

  var compLimpa = mes + ano;
  var variacoes = [compLimpa, ano + mes, mes + "-" + ano, ano + "-" + mes, mes + "." + ano];
  var mesesNomes = ["JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO","JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"];
  var nomeMes = mesesNomes[parseInt(mes,10)-1];
  if (nomeMes) variacoes.push(nomeMes);
  var dataRef = new Date(parseInt(ano,10), parseInt(mes,10)-1, 1);
  dataRef.setMonth(dataRef.getMonth()-12);
  var mesAntigo = ("0"+(dataRef.getMonth()+1)).slice(-2);
  var anoAntigo = dataRef.getFullYear();
  var textoAbatimento = mesAntigo + "/" + anoAntigo;
  var compDados = { compLimpa: compLimpa, variacoes: variacoes, textoAbatimento: textoAbatimento,
    inputOriginal: mes + "/" + ano, mesPA: parseInt(mes,10), anoPA: parseInt(ano,10) };

  var idxCnpj = buscarIndiceColuna(headers, "CNPJ");
  var idxEmpresa = buscarIndiceColuna(headers, "Razão Social");
  var idxAnexo = buscarIndiceColuna(headers, "Anexo do Simples Nacional");
  var idxRbt12 = buscarIndiceColuna(headers, "RBT12");
  var idxFatorR = buscarIndiceColuna(headers, "Fator R");
  var idxAliqPis = buscarIndiceColuna(headers, "Alíquota PIS/COFINS");
  var idxValorDas = buscarIndiceColuna(headers, "Valor no DAS");

  var pastaPai = obterPastaPrincipal_IBSCBS();
  if (!pastaPai) { Logger.log(nomeAba + ': erro de acesso a pasta raiz'); return; }

  var allData = sheet.getDataRange().getValues();
  var linhasAlvo = [];
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    var anyFilled = (idxAnexo>=0 && row[idxAnexo]) || (idxRbt12>=0 && row[idxRbt12]) || (idxValorDas>=0 && row[idxValorDas]) || (idxFatorR>=0 && row[idxFatorR]);
    if (anyFilled) linhasAlvo.push(i+1);
  }

  var erros = [], processados = 0, corrigidos = 0;
  var inicio = new Date().getTime();

  for (var k = 0; k < linhasAlvo.length; k++) {
    var row = linhasAlvo[k];
    var linha = allData[row-1];
    var cnpjPlanilha = String(linha[idxCnpj]).trim();
    var empresa = String(linha[idxEmpresa]).trim();
    var fatorRAntes = idxFatorR>=0 ? linha[idxFatorR] : '';
    if (!empresa) { erros.push("Linha "+row+": sem Razao Social"); continue; }

    var tempFileId = null;
    try {
      var pastaEmpresa = buscarPastaEmpresa_IBSCBS(pastaPai, empresa);
      if (!pastaEmpresa) { erros.push("Linha "+row+" ("+empresa+"): pasta nao encontrada"); continue; }
      var pdf = buscarExtrato_IBSCBS(pastaPai, pastaEmpresa, cnpjPlanilha, compDados.compLimpa, compDados.variacoes, String(compDados.anoPA));
      if (!pdf) { erros.push("Linha "+row+" ("+empresa+"): extrato "+compDados.inputOriginal+" nao encontrado"); continue; }

      var tempFile = Drive.Files.copy({ title: "Temp_IBSCBS_"+empresa, mimeType: MimeType.GOOGLE_DOCS }, pdf.getId(), { convert: true, supportsAllDrives: true });
      tempFileId = tempFile.id;
      var doc = DocumentApp.openById(tempFileId);
      var texto = doc.getBody().getText();

      var cnpjExtratoMatch = texto.match(/(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})/);
      if (cnpjExtratoMatch && cnpjPlanilha) {
        var cnpjLimpoPlanilha = cnpjPlanilha.replace(/[^\d]/g, "");
        var cnpjLimpoExtrato = cnpjExtratoMatch[1].replace(/[^\d]/g, "");
        if (cnpjLimpoPlanilha && cnpjLimpoExtrato && cnpjLimpoPlanilha !== cnpjLimpoExtrato) {
          erros.push("Linha "+row+" ("+empresa+"): CNPJ do extrato nao corresponde");
          continue;
        }
      }

      var fatorRTxt = "";
      var mFatorR = texto.match(/Fator\s*r\s*=\s*([\d]+,\d+)/i);
      if (mFatorR) {
        var percFatorR = parseFloat(mFatorR[1].replace(",", "."));
        fatorRTxt = percFatorR >= 0.28 ? "com" : "sem";
      } else if (/Fator\s*r\)[\s\S]{0,60}?N[ãa]o se aplica/i.test(texto)) {
        fatorRTxt = "";
      }

      if (fatorRTxt) {
        sheet.getRange(row, idxFatorR + 1).setValue(fatorRTxt);
        processados++;
        if (fatorRTxt !== fatorRAntes) corrigidos++;
      } else {
        erros.push("Linha "+row+" ("+empresa+"): Fator R nao localizado no extrato (item 2.4 ausente/formato diferente)");
      }
    } catch (e) {
      erros.push("Linha "+row+" ("+empresa+"): erro - "+e.message);
    } finally {
      if (tempFileId) { try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (e2) {} }
    }
  }

  var fim = new Date().getTime();
  var resumo = nomeAba + ": " + processados + "/" + linhasAlvo.length + " linhas com Fator R atualizado (" + corrigidos + " mudaram de valor). Tempo: " + Math.round((fim-inicio)/1000) + "s.";
  if (erros.length > 0) resumo += "\nFalhas (" + erros.length + "):\n" + erros.join("\n");
  Logger.log(resumo);
}