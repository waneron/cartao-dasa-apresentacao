#!/usr/bin/env python3
"""Ordena a aba BM - GCP Atividades - DP sem criar card nem alterar prazo."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
OUT = "/workspace/BM-GCP-Atividades-DP-ordem-temporal.xlsx"

# Textos copiados da planilha original. Nada aqui é card novo.
# ordem, etapa, card_id, descricao, complexidade, dependencia, cronograma,
# responsavel, observacoes, conflito, posicao, url, linha_original

ROWS = [
    (
        "HML — base concluída",
        "424950",
        "[DP] Provisionar ambiente GKE HML",
        "Concluído",
        "",
        "",
        "",
        "",
        "",
        "Status Concluído. É o único card de ambiente e o nome é HML. A documentação cita o cluster Azure azr-hml02; este card é o GKE de homologação.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/424950/details",
        40,
    ),
    (
        "HML — base concluída",
        "451661",
        "[DP] Migração GCP | Criar banco de dados",
        "Concluído",
        "",
        "",
        "",
        "",
        "",
        "Status Concluído. O card de migrar SQL Server está abortado; este é o banco que a planilha marca como criado.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451661/details",
        39,
    ),
    (
        "HML — base concluída",
        "451684",
        "[DP] Migração GCP | Redis",
        "Concluído",
        "Os pods precisam estar de pé para testar",
        "",
        "Back End",
        "",
        "",
        "Status Concluído. A própria dependência diz que o teste espera os pods. A documentação lista dois Redis de 6 GB (DP Cache PRD e Busca); a planilha não separa HML e PRD.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451684/details",
        35,
    ),
    (
        "HML — base concluída",
        "476416",
        "Continuação - Criação de Workloads do projeto BUSCA -> datp-searcher-api-commons",
        "Concluído",
        "",
        "",
        "",
        "",
        "",
        "Status Concluído. O card 470112, de criação do mesmo workload, continua na planilha com observação 'atuando'.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/476416/details",
        30,
    ),
    (
        "HML — sem data, em paralelo",
        "424949",
        "[DP] Definir conectividade",
        "",
        "",
        "",
        "",
        "",
        "",
        "Sem data e sem dependência. Os workloads de busca registram que a conectividade por IP está errada. Nenhum deles cita o número deste card. Não há prazo para terminar antes de 20/10/26.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/424949/details",
        43,
    ),
    (
        "HML — sem data, em paralelo",
        "424942",
        "[DP] Adequação de pipelines CI/CD",
        "Alta",
        "Infra, JulianoTelli",
        "",
        "Infra",
        "telli",
        "",
        "Sem data. Nenhum outro card lista este como dependência. Fica no início da HML em paralelo, sem prazo atribuído.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/424942/details",
        31,
    ),
    (
        "HML — Airflow local",
        "478326",
        "Spike -> Airflow na versao mais nova do airflow",
        "Alta",
        "Victor, subir um docker e testar ambienteLocal",
        "20/10/26",
        "Back End",
        "Aqui já vai ser possivel pegar imcompatibilidade de versao",
        "A dependência é ambiente local, não o GKE de HML. A documentação diz que a versão 2.2.5 é muito baixa e que a versão precisa ser equalizada com a do Google antes de continuar a migração.",
        "Data escrita: 20/10/26. É a POC citada pela linha sem card de adequação do Airflow.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/478326/details",
        2,
    ),
    (
        "HML — workloads com data",
        "470110",
        "Criação de Workloads do projeto BUSCA -> datp-searcher-api-brand",
        "Baixa",
        "Backend Hj a conectividade está por IP, isso é errado",
        "20/10/26",
        "Back End",
        "",
        "",
        "Data escrita: 20/10/26. Não cita Airflow nem Cognitive Search.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470110/details",
        3,
    ),
    (
        "HML — workloads com data",
        "470109",
        "Criação de Workloads do projeto BUSCA -> datp-infra-gateway -> (ZUUL)",
        "Média",
        "Backend Hj a conectividade está por IP, isso é errado",
        "27/10/26",
        "Back End",
        "",
        "",
        "Data escrita: 27/10/26. No diagrama AS IS o gateway aparece como dp-zuul-proxy.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470109/details",
        4,
    ),
    (
        "HML — workloads com data",
        "470118",
        "Criação de Workloads do projeto BUSCA -> datp-searcher-api-local",
        "Média",
        "Backend Hj a conectividade está por IP, isso é errado",
        "07/11/26",
        "Back End",
        "",
        "",
        "Data escrita: 07/11/26.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470118/details",
        5,
    ),
    (
        "HML — workloads com data",
        "470115",
        "Criação de Workloads do projeto BUSCA -> datp-searcher-api-healthinsurance",
        "Alta",
        "Backend Hj a conectividade está por IP, isso é errado",
        "20/11/26",
        "Back End",
        "",
        "",
        "Data escrita: 20/11/26.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470115/details",
        6,
    ),
    (
        "HML — workloads com data",
        "470113",
        "Criação de Workloads do projeto BUSCA -> datp-searcher-api-elegibility",
        "Alta",
        "Backend Hj a conectividade está por IP, isso é errado",
        "29/11/26",
        "Back End",
        "",
        "",
        "Data escrita: 29/11/26. O documento lista a API de agendamento; este card está escrito como elegibility. As duas grafias foram mantidas separadas.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470113/details",
        7,
    ),
    (
        "HML — workloads com data",
        "470112",
        "Criação de Workloads do projeto BUSCA -> datp-searcher-api-commons",
        "Baixa",
        "Backend Hj a conectividade está por IP, isso é errado",
        "",
        "Back End",
        "atuando",
        "",
        "Sem data. Observação: atuando. A continuação 476416 está concluída. Entra com as outras APIs de busca, antes do card de testar conectividade.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470112/details",
        17,
    ),
    (
        "HML — workloads com data",
        "424944",
        "[DP] Migração de aplicações, APIs",
        "Baixa",
        "testar conectividade",
        "15/12/26",
        "Back End",
        "",
        "",
        "Data escrita: 15/12/26. A dependência é testar conectividade, depois dos workloads de busca que têm data.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/424944/details",
        8,
    ),
    (
        "HML — índice, sem data",
        "451682",
        "[DP] Migração GCP | Index (cognitive search > vertex - adaptar na aplicação)",
        "Alta",
        "Maior Risco, descoberta",
        "",
        "Felipe Paixão | Back End",
        "",
        "No documento, Cognitive Search está como dependência externa, criticidade alta, e 'pode desacoplar' = TBD. O card já aponta Vertex e segue em descoberta. Não foi atribuída data.",
        "Sem data. O card 470114 registra dependência de Cognitive Search. O card 470119, cancelado, registra a mesma dependência.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451682/details",
        41,
    ),
    (
        "HML — índice, sem data",
        "470114",
        "Criação de Workloads do projeto BUSCA -> datp-searcher-api-exam",
        "Alta",
        "Backend Hj a conectividade está por IP, isso é errado",
        "",
        "Back End",
        "depende cognitiveSearch",
        "",
        "Sem data. A observação trava o card em Cognitive Search, por isso fica depois do 451682 e fora da fila datada de outubro e novembro.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470114/details",
        18,
    ),
    (
        "HML — front e back do DP, sem data",
        "469049",
        "[DP] Migração GCP | Migrar Backend -> front do DP (adaptações)",
        "Alta",
        "Ajuste de codigo, versao, eventual erro",
        "",
        "Back End",
        "",
        "",
        "Sem data. Não cita Airflow. A documentação aponta o repositório datp-backoffice-backend.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/469049/details",
        15,
    ),
    (
        "HML — front e back do DP, sem data",
        "469051",
        "[DP] Migração GCP | Migrar Frontend  portal do DP (adaptações)",
        "Média",
        "Ajuste de codigo, versao, eventual erro",
        "",
        "Back End",
        "só é possivel testar o front com tudo integrado.",
        "",
        "Sem data. A observação fala do teste integrado, que é o card 451689. Não cita Airflow.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/469051/details",
        14,
    ),
    (
        "Dados — sem HML nem PRD no card",
        "478328",
        "BigQuery => substituir dataBricks",
        "Alta",
        "Galerinha de dados DP",
        "29/01/27",
        "Dados",
        "Levar toda a orquestração do dataBricks para o BigQuery",
        "Nenhum outro card lista este como dependência. A decisão de levar o Databricks para a GCP não está na planilha, então não entrou linha nova. O card 451674, converter ETLs, está cancelado. O documento só lista os workspaces de PRD e de HML, sem jobs.",
        "Entrou pela data escrita, 29/01/27, entre o card de 15/12/26 e os de fevereiro de 2027. O card não diz HML nem PRD.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/478328/details",
        9,
    ),
    (
        "HML — portal e arquivo, data do card",
        "451688",
        "[DP] Migração GCP | Blob Storage (adaptações na aplicação)",
        "Média",
        "Criado, depende de airFlow, backend e frontEnd",
        "03/02/27",
        "Back End | Dados",
        "",
        "Data 03/02/27. A linha sem card de adequação do Airflow está em março–abril de 2027. Este card depende de Airflow, backend e frontend.",
        "Data escrita: 03/02/27. A documentação registra 269 GB de Blob.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451688/details",
        10,
    ),
    (
        "HML — portal e arquivo, data do card",
        "451677",
        "[DP] Migração GCP | Migrar Frontend  Protal de cadastro (adaptações)",
        "Alta",
        "Ajuste de codigo, versao, eventual erro",
        "07/02/27",
        "Back End",
        "só é possivel testar o front com tudo integrado.",
        "Data 07/02/27, anterior à adequação de Airflow escrita como março–abril de 2027. A observação diz que o front só é testável com tudo integrado.",
        "Data escrita: 07/02/27.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451677/details",
        11,
    ),
    (
        "HML — portal e arquivo, data do card",
        "451681",
        "[DP] Migração GCP | Migrar Backend -> portal de cadastro (adaptações)",
        "Alta",
        "Ajuste de codigo, versao, eventual erro",
        "10/02/27",
        "Back End",
        "",
        "Data 10/02/27, anterior à adequação de Airflow escrita como março–abril de 2027.",
        "Data escrita: 10/02/27.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451681/details",
        12,
    ),
    (
        "HML — Airflow, linha sem card",
        "",
        "Criar historia adequação airflow, vinculada ao fim da poc do Victor => necessario devido versao de airflow",
        "Alta",
        "as adequações levarão tempo, pode ser a maior data.",
        "entre março e abril de 2027",
        "",
        "",
        "Sem Card ID. A própria dependência diz que pode ser a maior data. Os cards 451688 (03/02/27), 451677 (07/02/27) e 451681 (10/02/27) estão antes, e o 451688 depende de Airflow. A planilha não resolve essa ordem. Não há card de workload de Airflow de HML; os seis workloads de Airflow têm 'prd' no nome.",
        "Janela escrita na planilha. Ligada ao fim da POC do card 478326. A documentação pede equalizar a 2.2.5 com a versão do Google antes de continuar a migração.",
        "",
        13,
    ),
    (
        "HML — cita Airflow, sem data",
        "468966",
        "Criação de Workloads do projeto BUSCA -> datp-backoffice-backend",
        "Alta",
        "Ajuste de codigo, versao, eventual erro",
        "",
        "Back End",
        "depende do airFlow",
        "Sem data. A observação diz que depende do Airflow. Não há card de Airflow de HML para apontar.",
        "Fica depois da linha de adequação do Airflow e antes da homologação. Sem data atribuída.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/468966/details",
        22,
    ),
    (
        "HML — cita Airflow, sem data",
        "470165",
        "Criação de Workloads do projeto PORTAL-CADASTRO -> portal-cadastro-web",
        "Média",
        "Depende do airFlow pronto, cargas cadastro e D+1 (dowloads Cadastro)",
        "",
        "Back End | Dados | Infra",
        "depende do airFlow",
        "A dependência cita Airflow pronto, cargas de cadastro e D+1. Não há card de cargas nem de D+1, nem card de Airflow de HML.",
        "Sem data. Permanece antes da homologação porque o teste do portal, no card 451677, pede tudo integrado.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470165/details",
        19,
    ),
    (
        "HML — cita Airflow, sem data",
        "470100",
        "Criação de Workloads do projeto BUSCA -> datp-backoffice-frontend",
        "Média",
        "Depende do airFlow pronto, cargas cadastro e D+1 (dowloads Cadastro)",
        "",
        "Back End | Dados | Infra",
        "depende do airFlow",
        "A dependência cita Airflow pronto, cargas de cadastro e D+1. Não há card de cargas nem de D+1, nem card de Airflow de HML.",
        "Sem data. Mesma dependência do 470165.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470100/details",
        21,
    ),
    (
        "HML — homologação",
        "451689",
        "[DP] Migração GCP | Testes Integrados, Homologação & Plano de Rollback",
        "Sem Definição de Escopo",
        "??",
        "",
        "",
        "",
        "Único card de teste da planilha. Não cita teste de carga. Complexidade: sem definição de escopo. Dependência: ?? . O template do documento lista teste funcional e performance sem plano preenchido.",
        "Fecha a HML. A posição segue o pedido de testar em homologação antes de planejar produção. O card não tem data.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451689/details",
        36,
    ),
    (
        "Produção — Airflow nomeado prd",
        "470150",
        "Criação de Workloads do projeto DATAPROVIDER-DATA -> dp-data-stack-airflow-prd-worker-statefulset",
        "Média",
        "infra managedAirFlow - infra => sem data depende de infra",
        "09/10/26",
        "Infra",
        "infra",
        "A data escrita é 09/10/26, anterior aos workloads de HML. O cronograma não foi reescrito. O bloco é produção porque o nome do card contém prd.",
        "Mesma data dos outros cinco workloads de Airflow. Ordem entre eles é a da planilha original.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470150/details",
        23,
    ),
    (
        "Produção — Airflow nomeado prd",
        "470149",
        "Criação de Workloads do projeto DATAPROVIDER-DATA -> dp-data-stack-airflow-prd",
        "Média",
        "infra managedAirFlow - infra => sem data depende de infra",
        "09/10/26",
        "Infra",
        "infra",
        "A data escrita é 09/10/26. O cronograma não foi reescrito. O bloco é produção porque o nome do card contém prd.",
        "Mesma data dos outros workloads de Airflow nomeados prd.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470149/details",
        24,
    ),
    (
        "Produção — Airflow nomeado prd",
        "470148",
        "Criação de Workloads do projeto DATAPROVIDER-DATA -> dp-data-stack-airflow-prd-sync-users",
        "Média",
        "infra managedAirFlow - infra => sem data depende de infra",
        "09/10/26",
        "Infra",
        "infra",
        "A data escrita é 09/10/26. O cronograma não foi reescrito. O bloco é produção porque o nome do card contém prd.",
        "Mesma data dos outros workloads de Airflow nomeados prd.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470148/details",
        25,
    ),
    (
        "Produção — Airflow nomeado prd",
        "470146",
        "Criação de Workloads do projeto DATAPROVIDER-DATA -> dp-data-stack-airflow-prd-scheduler",
        "Média",
        "infra managedAirFlow - infra => sem data depende de infra",
        "09/10/26",
        "Infra",
        "infra",
        "A data escrita é 09/10/26. O cronograma não foi reescrito. O bloco é produção porque o nome do card contém prd.",
        "Mesma data dos outros workloads de Airflow nomeados prd.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470146/details",
        26,
    ),
    (
        "Produção — Airflow nomeado prd",
        "470145",
        "Criação de Workloads do projeto DATAPROVIDER-DATA -> dp-data-stack-airflow-prd-flower",
        "Média",
        "infra managedAirFlow - infra => sem data depende de infra",
        "09/10/26",
        "Infra",
        "infra",
        "A data escrita é 09/10/26. O cronograma não foi reescrito. O bloco é produção porque o nome do card contém prd.",
        "Mesma data dos outros workloads de Airflow nomeados prd.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470145/details",
        27,
    ),
    (
        "Produção — Airflow nomeado prd",
        "468978",
        "Criação de Workloads do projeto DATAPROVIDER-DATA -> dp-data-stack-airflow-prd-db-migrations",
        "Média",
        "infra managedAirFlow - infra",
        "09/10/26",
        "Infra",
        "infra",
        "A data escrita é 09/10/26. O cronograma não foi reescrito. O bloco é produção porque o nome do card contém prd.",
        "Mesma data dos outros workloads de Airflow nomeados prd.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/468978/details",
        28,
    ),
    (
        "Produção — depois da homologação",
        "451683",
        "[DP] Migração GCP | Event Hub para Pub/Sub",
        "Média",
        "Ultima demanda, por ultimo -> falar depois",
        "",
        "BD",
        "pode ficar por ultimo",
        "",
        "Sem data. A dependência e a observação pedem para ficar por último. Entra na produção, antes do cutover, porque o cutover é o go-live.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451683/details",
        29,
    ),
    (
        "Produção — depois da homologação",
        "451691",
        "[DP] Migração GCP | Cutover & Go-Live",
        "Sem Definição de Escopo",
        "??",
        "",
        "",
        "",
        "Sem definição de escopo. O template do documento deixa em branco big bang ou gradual, DNS e janela de downtime.",
        "Última linha da produção. Sem data. O card não descreve o que entra no go-live.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451691/details",
        34,
    ),
    (
        "Fora da linha do tempo",
        "470119",
        "Criação de Workloads do projeto BUSCA -> dp-backend-api",
        "Cancelado",
        "",
        "",
        "",
        "depende cognitiveSearch",
        "",
        "Complexidade Cancelado. Fora da ordem.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470119/details",
        16,
    ),
    (
        "Fora da linha do tempo",
        "468985",
        "Criação de Workloads do projeto PORTAL-CADASTRO -> datp-backoffice-frontend-portal",
        "Cancelado",
        "",
        "",
        "",
        "depende do airFlow",
        "",
        "Complexidade Cancelado. Fora da ordem.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/468985/details",
        20,
    ),
    (
        "Fora da linha do tempo",
        "470121",
        "Criação de Workloads do projeto BUSCA -> dp-corporate-frontend",
        "Cancelado",
        "",
        "",
        "",
        "",
        "",
        "Complexidade Cancelado. Fora da ordem. A documentação ainda cita a URL dp-corporate-frontend.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470121/details",
        32,
    ),
    (
        "Fora da linha do tempo",
        "451674",
        "[DP] Migração GCP | Converter ETL's (DataBricks)",
        "Cancelado",
        "",
        "",
        "",
        "",
        "",
        "Complexidade Cancelado. Fora da ordem.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451674/details",
        38,
    ),
    (
        "Fora da linha do tempo",
        "424941",
        "[DP] PostgreSQL - Levantamento de recursos para reescrita de código",
        "Cancelado",
        "",
        "",
        "",
        "",
        "",
        "Complexidade Cancelado. Fora da ordem.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/424941/details",
        42,
    ),
    (
        "Fora da linha do tempo",
        "451676",
        "[DP] Migração GCP | Migrar banco de dados SQL Server -- ABORTADO",
        "Concluído",
        "",
        "",
        "",
        "",
        "",
        "Título ABORTADO e status Concluído. Fora do caminho de HML. A documentação diz que a réplica 3 não será migrada e está desligada.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451676/details",
        37,
    ),
    (
        "Fora da linha do tempo",
        "476063",
        "[DATP] DasaLogs -> api local - unidades",
        "Stand By",
        "",
        "",
        "",
        "",
        "",
        "Complexidade Stand By. Fora da ordem. Sem data.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/476063/details",
        33,
    ),
]

LACUNAS = [
    (
        "Teste de carga",
        "Não há card. O card 451689 cobre testes integrados, homologação e rollback, sem escopo. No documento, o item Testes do plano de migração é template e cita Performance sem conteúdo.",
        "Não foi criada atividade de teste de carga.",
    ),
    (
        "Airflow de homologação",
        "O documento lista Airflow de HML em datp-infra-database-stack.azr-hml02. Na planilha, os seis workloads de Airflow têm prd no nome e data 09/10/26.",
        "Os cards de PRD não foram duplicados para HML.",
    ),
    (
        "Cargas de cadastro e download D+1",
        "Aparecem só como dependência dos cards 470165 e 470100.",
        "Não há card. Nenhuma linha foi criada.",
    ),
    (
        "Databricks na GCP",
        "O card 478328 pede levar a orquestração do Databricks para o BigQuery, em 29/01/27. O card 451674, converter ETLs, está cancelado. O documento lista workspace de PRD e de HML e não descreve jobs.",
        "A decisão de mover o Databricks para a GCP não está na planilha. Não entrou linha nova.",
    ),
    (
        "API de agendamento e API de elegibility",
        "O documento diz que as APIs são exam, healthinsurance, commons, agendamento, local e brand. O diagrama tem busca-agendamento-backend. A planilha tem datp-searcher-api-elegibility (470113) e não tem card com o nome agendamento.",
        "Os nomes não foram unidos e nenhum card foi criado.",
    ),
    (
        "dp-api-gtw e dp-cad-proxy",
        "Os dois aparecem no diagrama AS IS. O gateway com card é o Zuul (470109). dp-backend-api (470119) está cancelado.",
        "Não foi criado card para o que o diagrama mostra e a planilha não lista.",
    ),
    (
        "Réplica 3 do SQL",
        "O documento diz que a réplica 3 não será migrada e está desligada. O card 451676 está abortado e concluído.",
        "Não foi criada atividade.",
    ),
    (
        "Identidade",
        "No documento, Azure AD é dependência interna, alta, e não pode desacoplar. O template TO BE traz a linha Azure AD para Google IAM sem preenchimento. Front do DP usa AD, portal usa SSO, APIs usam API key.",
        "Não foi criada atividade de identidade.",
    ),
    (
        "Estratégia e ondas",
        "As seções de 6Rs, roadmap, cutover e riscos do documento estão em branco.",
        "A ordem não veio desse template. Veio da data, do texto de dependência e do pedido de fazer HML antes de produção.",
    ),
]

REGRAS = [
    "A aba traz as 42 linhas com conteúdo da planilha original. A linha em branco do fim foi descartada. Nenhum Card ID, prazo, responsável ou descrição foi alterado.",
    "HML vem antes de produção. O único ambiente marcado como concluído é o GKE HML (424950). Os cards com prd no nome ficam no bloco de produção mesmo com data 09/10/26. Essa data permanece na coluna Cronograma.",
    "Dentro da HML, card com data segue a data escrita. Card sem data entra no ponto em que o próprio texto o prende: conectividade e CI/CD no início, sem prazo; exam depois do índice; portal que cita Airflow depois da linha de adequação do Airflow; homologação no fim da HML.",
    "O card 451689 é o único card de teste. Ele fecha a HML. Teste de carga não está escrito em lugar nenhum da planilha.",
    "Produção, depois desse card: os seis workloads de Airflow nomeados prd, na ordem original; Event Hub, que a planilha manda deixar por último; Cutover, sem escopo.",
    "Cancelado, Stand By e o SQL Server abortado ficam fora da linha do tempo, no fim da aba, para a lista continuar completa.",
    "Onde a data de um card é anterior à dependência que ele mesmo cita, a coluna Limite registra o conflito. A ordem não escolhe uma das duas datas no lugar da outra.",
]


def fills():
    return {
        "HML — base concluída": "E7E6E6",
        "HML — sem data, em paralelo": "D6E3F0",
        "HML — Airflow local": "D6E3F0",
        "HML — workloads com data": "DEEBF7",
        "HML — índice, sem data": "FFF2CC",
        "HML — front e back do DP, sem data": "DEEBF7",
        "Dados — sem HML nem PRD no card": "E2D5F1",
        "HML — portal e arquivo, data do card": "FCE4D6",
        "HML — Airflow, linha sem card": "F8CBAD",
        "HML — cita Airflow, sem data": "FCE4D6",
        "HML — homologação": "C6EFCE",
        "Produção — Airflow nomeado prd": "F4B183",
        "Produção — depois da homologação": "F4B183",
        "Fora da linha do tempo": "F2F2F2",
    }


def apply_header(ws, headers):
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    for col, text in enumerate(headers, 1):
        cell = ws.cell(1, col, text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = thin
    ws.row_dimensions[1].height = 32
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.left.text = "BM - GCP Atividades - DP"
    ws.oddFooter.right.text = "Página &P de &N"
    ws.page_setup.horizontalCentered = True
    ws.print_title_rows = "1:1"
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.6
    return thin


def write_main(wb):
    ws = wb.active
    ws.title = "BM - GCP Atividades - DP"
    headers = [
        "Ordem",
        "Etapa",
        "Card ID",
        "Descrição",
        "Complexidade",
        "Depêndencia",
        "Cronograma",
        "Recurso Responsável",
        "Observações",
        "Limite do que está escrito",
        "Por que nesta posição",
        "Card URL",
        "Linha na planilha original",
    ]
    thin = None
    etapa_fills = fills()
    body_font = Font(name="Calibri", size=10)
    link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")
    conflict_font = Font(name="Calibri", size=10, color="843C0C")
    for i, row in enumerate(ROWS, 1):
        (
            etapa,
            card_id,
            descricao,
            complexidade,
            dependencia,
            cronograma,
            responsavel,
            observacoes,
            conflito,
            posicao,
            url,
            linha,
        ) = row
        values = [
            i,
            etapa,
            card_id,
            descricao,
            complexidade,
            dependencia,
            cronograma,
            responsavel,
            observacoes,
            conflito,
            posicao,
            url,
            linha,
        ]
        for col, value in enumerate(values, 1):
            cell = ws.cell(i + 1, col, value if value != "" else None)
            cell.font = conflict_font if col == 10 and conflito else body_font
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.fill = PatternFill("solid", fgColor=etapa_fills[etapa])
        url_cell = ws.cell(i + 1, 12)
        if url:
            url_cell.hyperlink = url
            url_cell.font = link_font
        ws.row_dimensions[i + 1].height = 48
    thin = apply_header(ws, headers)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = thin
    widths = [8, 38, 12, 62, 24, 48, 28, 26, 42, 55, 55, 42, 16]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    ws.auto_filter.add_sort_condition("A2:A43")
    ws.print_options.horizontalCentered = True
    ws.sheet_view.showGridLines = False
    ws.oddHeader.center.text = ""


def write_rules(wb):
    ws = wb.create_sheet("Regra de ordem")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Como a aba foi ordenada"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    for i, text in enumerate(REGRAS, 3):
        ws.cell(i, 1, text)
        ws.cell(i, 1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(i, 1).font = Font(name="Calibri", size=11)
        ws.row_dimensions[i].height = 36
    ws.column_dimensions["A"].width = 140
    ws.row_dimensions[1].height = 22
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.6
    ws.oddFooter.right.text = "Página &P de &N"


def write_lacunas(wb):
    ws = wb.create_sheet("Lacunas")
    headers = ["Tema", "Onde aparece", "O que não foi feito"]
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    for i, item in enumerate(LACUNAS, 2):
        for col, value in enumerate(item, 1):
            cell = ws.cell(i, col, value)
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin
            cell.fill = PatternFill("solid", fgColor="FFF8E7")
        ws.row_dimensions[i].height = 48
    apply_header(ws, headers)
    for col in range(1, 4):
        ws.cell(1, col).border = thin
    for row in ws.iter_rows(min_row=1, max_row=1, max_col=3):
        for cell in row:
            cell.border = thin
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 88
    ws.column_dimensions["C"].width = 55
    ws.auto_filter.ref = f"A1:C{ws.max_row}"
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.6
    ws.print_title_rows = "1:1"
    ws.oddFooter.right.text = "Página &P de &N"


def main():
    wb = Workbook()
    write_main(wb)
    write_rules(wb)
    write_lacunas(wb)
    wb.save(OUT)


if __name__ == "__main__":
    main()
